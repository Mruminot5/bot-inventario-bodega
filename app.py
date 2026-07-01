import os
import json
import requests
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from flask import Flask, request, jsonify
from difflib import get_close_matches

app = Flask(__name__)

VERIFY_TOKEN      = os.environ.get("WA_VERIFY_TOKEN", "bodega_inventario_2024")
ACCESS_TOKEN      = os.environ.get("WA_ACCESS_TOKEN", "")
PHONE_ID          = os.environ.get("WA_PHONE_NUMBER_ID", "")
SPREADSHEET_ID    = os.environ.get("SPREADSHEET_ID", "")
GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDENTIALS", "{}")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

COL_ID     = "ID"
COL_NOMBRE = "Nombre del \u00cdtem"
COL_CAT    = "Categor\u00eda"
COL_TOTAL  = "Stock Total"
COL_DISP   = "Disponible"
COL_USO    = "En Uso"
COL_REP    = "En Reparaci\u00f3n"
COL_ESTADO = "Estado"
COL_UBIC   = "Ubicaci\u00f3n"

def get_sheet():
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).get_worksheet(0)

def cargar_inventario():
    try:
        sheet = get_sheet()
        rows = sheet.get_all_values()
        if len(rows) < 2:
            return None
        headers = rows[1]
        data = rows[2:]
        df = pd.DataFrame(data, columns=headers)
        df.columns = df.columns.str.strip()
        df = df[df[COL_ID].str.strip() != ""]
        for col in [COL_TOTAL, COL_DISP, COL_USO, COL_REP]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        return df
    except Exception as e:
        print(f"ERROR cargar_inventario: {e}", flush=True)
        return None

def buscar_item(df, consulta):
    consulta = consulta.lower().strip()
    match_id = df[df[COL_ID].str.lower() == consulta]
    if not match_id.empty:
        return match_id.iloc[0]
    nombres = df[COL_NOMBRE].str.lower().tolist()
    ids = df[COL_ID].str.lower().tolist()
    cercanos = get_close_matches(consulta, nombres + ids, n=1, cutoff=0.4)
    if cercanos:
        match = df[
            (df[COL_NOMBRE].str.lower() == cercanos[0]) |
            (df[COL_ID].str.lower() == cercanos[0])
        ]
        if not match.empty:
            return match.iloc[0]
    match_cont = df[df[COL_NOMBRE].str.lower().str.contains(consulta, na=False)]
    if not match_cont.empty:
        return match_cont.iloc[0]
    return None

def formato_item(item):
    return (
        f"[{item.get(COL_ESTADO, '')}] {item[COL_NOMBRE]} ({item[COL_ID]})\n"
        f"Categor\u00eda: {item.get(COL_CAT, '')}\n"
        f"Stock total: {item.get(COL_TOTAL, 0)}\n"
        f"Disponible: {item.get(COL_DISP, 0)}\n"
        f"En uso: {item.get(COL_USO, 0)}\n"
        f"En reparaci\u00f3n: {item.get(COL_REP, 0)}\n"
        f"Ubicaci\u00f3n: {item.get(COL_UBIC, '')}"
    )

def calcular_estado(disponible, total):
    if disponible == 0:
        return "Sin Stock"
    elif disponible <= max(1, int(total) // 4):
        return "Cr\u00edtico"
    else:
        return "Disponible"

def actualizar_item_sheet(id_item, nuevo_disp, nuevo_uso, nuevo_estado):
    try:
        sheet = get_sheet()
        headers = [h.strip() for h in sheet.row_values(2)]
        col_disp_idx   = headers.index(COL_DISP) + 1
        col_uso_idx    = headers.index(COL_USO) + 1
        col_estado_idx = headers.index(COL_ESTADO) + 1
        col_id_idx     = headers.index(COL_ID) + 1
        all_ids = sheet.col_values(col_id_idx)
        row_num = all_ids.index(id_item) + 1
        sheet.update_cell(row_num, col_disp_idx, nuevo_disp)
        sheet.update_cell(row_num, col_uso_idx, nuevo_uso)
        sheet.update_cell(row_num, col_estado_idx, nuevo_estado)
        return True, "OK"
    except Exception as e:
        return False, str(e)

def cmd_retirar(df, args):
    partes = args.rsplit(" ", 1)
    if len(partes) == 2 and partes[1].isdigit():
        nombre_consulta, cantidad = partes[0], int(partes[1])
    else:
        nombre_consulta, cantidad = args, 1
    item = buscar_item(df, nombre_consulta)
    if item is None:
        return f"\u274c No encontr\u00e9 '{nombre_consulta}' en el inventario."
    disponible = int(item.get(COL_DISP, 0))
    en_uso     = int(item.get(COL_USO, 0))
    total      = int(item.get(COL_TOTAL, 0))
    if cantidad > disponible:
        return f"\u274c Stock insuficiente.\nDisponible: {disponible} | Solicitado: {cantidad}"
    nuevo_disp   = disponible - cantidad
    nuevo_uso    = en_uso + cantidad
    nuevo_estado = calcular_estado(nuevo_disp, total)
    ok, msg = actualizar_item_sheet(item[COL_ID], nuevo_disp, nuevo_uso, nuevo_estado)
    if not ok:
        return f"\u274c Error al actualizar: {msg}"
    return (
        f"\u2705 Retiro registrado:\n"
        f"{item[COL_NOMBRE]} ({item[COL_ID]})\n"
        f"Retirado: {cantidad}\n"
        f"Disponible ahora: {nuevo_disp}\n"
        f"En uso ahora: {nuevo_uso}\n"
        f"Estado: {nuevo_estado}"
    )

def cmd_devolver(df, args):
    partes = args.rsplit(" ", 1)
    if len(partes) == 2 and partes[1].isdigit():
        nombre_consulta, cantidad = partes[0], int(partes[1])
    else:
        nombre_consulta, cantidad = args, 1
    item = buscar_item(df, nombre_consulta)
    if item is None:
        return f"\u274c No encontr\u00e9 '{nombre_consulta}' en el inventario."
    disponible = int(item.get(COL_DISP, 0))
    en_uso     = int(item.get(COL_USO, 0))
    total      = int(item.get(COL_TOTAL, 0))
    if cantidad > en_uso:
        return f"\u274c No se pueden devolver {cantidad}.\nEn uso: {en_uso}"
    nuevo_disp   = disponible + cantidad
    nuevo_uso    = en_uso - cantidad
    nuevo_estado = calcular_estado(nuevo_disp, total)
    ok, msg = actualizar_item_sheet(item[COL_ID], nuevo_disp, nuevo_uso, nuevo_estado)
    if not ok:
        return f"\u274c Error al actualizar: {msg}"
    return (
        f"\u2705 Devoluci\u00f3n registrada:\n"
        f"{item[COL_NOMBRE]} ({item[COL_ID]})\n"
        f"Devuelto: {cantidad}\n"
        f"Disponible ahora: {nuevo_disp}\n"
        f"En uso ahora: {nuevo_uso}\n"
        f"Estado: {nuevo_estado}"
    )

def responder(texto):
    t = texto.lower().strip()
    df = cargar_inventario()
    if df is None:
        return "\u274c No pude leer el inventario. Intenta m\u00e1s tarde."
    if t.startswith("retirar "):
        return cmd_retirar(df, t[8:].strip())
    if t.startswith("devolver ") or t.startswith("retornar "):
        return cmd_devolver(df, t[9:].strip())
    if t in ["hola", "hi", "ayuda", "help", "menu", "men\u00fa"]:
        return (
            "Hola! Soy *Maquina*, tu bot de inventario.\n\n"
            "*Consultas:*\n"
            "  stock [item] - ver disponibilidad\n"
            "  buscar [item] - buscar por nombre o ID\n"
            "  reparacion - items en reparaci\u00f3n\n"
            "  en uso - items en uso\n"
            "  critico - stock cr\u00edtico\n"
            "  resumen - dashboard general\n"
            "  herramientas / insumos / epp / repuestos / mobiliario / energia\n\n"
            "*Movimientos:*\n"
            "  retirar [item] [cantidad]\n"
            "  devolver [item] [cantidad]\n\n"
            "Ej: retirar taladro 1  |  stock H001"
        )
    if t in ["resumen", "total", "dashboard"]:
        sin_stock = (df[COL_DISP] == 0).sum()
        criticos  = (df[COL_ESTADO] == "Cr\u00edtico").sum()
        en_uso    = df[COL_USO].sum()
        cats = df[COL_CAT].value_counts()
        lineas = [f"*Resumen de inventario*", f"Total items: {len(df)}", f"En uso: {en_uso} unid.", f"Sin stock: {sin_stock}", f"Cr\u00edticos: {criticos}\n"]
        for cat, n in cats.items():
            lineas.append(f"  {cat}: {n} items")
        return "\n".join(lineas)
    if t in ["reparacion", "en reparacion", "reparaci\u00f3n"]:
        sub = df[df[COL_REP] > 0]
        if sub.empty:
            return "No hay items en reparaci\u00f3n actualmente."
        lineas = [f"Items en reparaci\u00f3n ({len(sub)}):"]
        for _, row in sub.iterrows():
            lineas.append(f"  - {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_REP]} unid.")
        return "\n".join(lineas)
    if t in ["en uso", "uso"]:
        sub = df[df[COL_USO] > 0]
        if sub.empty:
            return "No hay items en uso."
        lineas = [f"Items en uso ({len(sub)}):"]
        for _, row in sub.iterrows():
            lineas.append(f"  - {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_USO]} unid.")
        return "\n".join(lineas)
    if t in ["critico", "cr\u00edtico", "stock critico"]:
        sub = df[df[COL_ESTADO].isin(["Cr\u00edtico", "Sin Stock"])]
        if sub.empty:
            return "No hay items en estado cr\u00edtico."
        lineas = [f"Items cr\u00edticos ({len(sub)}):"]
        for _, row in sub.iterrows():
            lineas.append(f"  - {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_DISP]} disp.")
        return "\n".join(lineas)
    for key, cat in {"herramientas": "Herramientas", "insumos": "Insumos", "epp": "EPP", "repuestos": "Repuestos", "mobiliario": "Mobiliario", "energia": "Energ\u00eda"}.items():
        if t == key:
            sub = df[df[COL_CAT] == cat]
            if sub.empty:
                return f"No hay items en {cat}."
            lineas = [f"*{cat}* ({len(sub)} items):"]
            for _, row in sub.iterrows():
                lineas.append(f"  - {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_DISP]} disp.")
            return "\n".join(lineas)
    for prefix in ["stock ", "buscar ", "ver ", "info "]:
        if t.startswith(prefix):
            item = buscar_item(df, t[len(prefix):])
            if item is not None:
                return formato_item(item)
            return f"\u274c No encontr\u00e9 '{t[len(prefix):]}'."
    item = buscar_item(df, t)
    if item is not None:
        return formato_item(item)
    return "No entend\u00ed esa consulta. Escribe *ayuda* para ver los comandos."

def send_whatsapp_message(to, body):
    url = f"https://graph.facebook.com/v19.0/{PHONE_ID}/messages"
    r = requests.post(
        url,
        headers={"Authorization": f"Bearer {ACCESS_TOKEN}", "Content-Type": "application/json"},
        json={"messaging_product": "whatsapp", "to": to, "type": "text", "text": {"body": body}},
        timeout=10
    )
    return r.status_code, r.json()

@app.route("/webhook", methods=["GET"])
def verify():
    mode      = request.args.get("hub.mode")
    token     = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")
    if mode == "subscribe" and token == VERIFY_TOKEN:
        return challenge, 200
    return "Forbidden", 403

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json()
    print(f"WEBHOOK IN: {json.dumps(data)[:300]}", flush=True)
    try:
        msg  = data["entry"][0]["changes"][0]["value"]["messages"][0]
        body = msg.get("text", {}).get("body", "")
        print(f"MSG from={msg['from']} body={body}", flush=True)
        if body:
            respuesta = responder(body)
            print(f"RESPUESTA: {respuesta[:100]}", flush=True)
            status, resp = send_whatsapp_message(msg["from"], respuesta)
            print(f"WA_API status={status} resp={json.dumps(resp)[:200]}", flush=True)
    except Exception as e:
        print(f"ERROR webhook: {e}", flush=True)
    return '{"status":"ok"}', 200

@app.route("/", methods=["GET"])
def index():
    return "Maquina - Bot de inventario activo \u2705", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
