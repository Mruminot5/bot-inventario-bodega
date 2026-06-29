import os
import json
import requests
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from flask import Flask, request, jsonify
from difflib import get_close_matches

app = Flask(__name__)

VERIFY_TOKEN = os.environ.get("WA_VERIFY_TOKEN", "bodega_inventario_2024")
ACCESS_TOKEN   = os.environ.get("WA_ACCESS_TOKEN", "")
PHONE_ID       = os.environ.get("WA_PHONE_NUMBER_ID", "")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDENTIALS", "{}")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

COL_ID     = "ID"
COL_NOMBRE = "Nombre del \u00cdtem"
COL_CAT    = "Categor\u00eda"
COL_TOTAL = "Stock Total"
COL_DISP   = "Disponible"
COL_USO    = "En Uso"
COL_REP    = "En Reparaci\u00f3n"
COL_ESTADO = "Estado"
COL_UBIC   = "Ubicaci\u00f3n"

COLUMNAS_ORDEN = [COL_ID, COL_NOMBRE, COL_CAT, COL_TOTAL, COL_DISP, COL_USO, COL_REP, COL_ESTADO, COL_UBIC]

def get_sheet():
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).get_worksheet(0)

def cargar_inventario():
    try:
        sheet = get_sheet()
        data = sheet.get_all_values()
        if len(data) < 2:
            return None
        headers = data[1]
        rows = data[2:]
        df = pd.DataFrame(rows, columns=headers)
        df.columns = df.columns.str.strip()
        df = df[df[COL_ID].str.strip() != ""]
        for col in [COL_TOTAL, COL_DISP, COL_USO, COL_REP]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        return df
    except Exception as e:
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
    estado = str(item.get(COL_ESTADO, ""))
    emoji = {"OK": "\u2705", "En uso": "\u1f535", "Parcial": "\u1f7e1", "Cr\u00edtico": "\u1f534"}.get(estado, "\u26fb")
    return (
        f"{emoji} *{item[COL_NOMBRE]}* ({item[COL_ID]})\n"
        f"\u2022 Categor\u00eda: {item.get(COL_CAT, '')}\n"
        f"\u2022 Stock total: {item.get(COL_TOTAL, 0)}\n"
        f"\u2022 Disponible: {item.get(COL_DISP, 0)}\n"
        f"\u2022 En uso: {item.get(COL_USO, 0)}\n"
        f"\u2022 En reparaci\u00f3n: {item.get(COL_REP, 0)}\n"
        f"\u2022 Estado: {estado}\n"
        f"\u2022 Ubicaci\u00f3n: {item.get(COL_UBIC, '')}"
    )

AYUDA_AGREGAR = (
    "\u1f4cb *Formato para agregar un \u00edtem:*\n\n"
    "agregar ID | Nombre | Categor\u00eda | Stock Total | Disponible | En Uso | En Reparaci\u00f3n | Estado | Ubicaci\u00f3n\n\n"
    "\u1f4cd *Ejemplo:*\n"
    "_agregar H012 | Martillo grande | Herramientas | 5 | 4 | 1 | 0 | OK | Estante A_\n\n"
    "\u1f4c2 *Categor\u00edas v\u00e1lidas:* Herramientas, Insumos, EPP, Repuestos\n"
    "\u1f516 *Estados v\u00e1lidos:* OK, En uso, Parcial, Cr\u00edtico"
)

def agregar_item(texto):
    partes = [p.strip() for p in texto.split("|")]
    if len(partes) < 9:
        return f"\u274c Faltan campos. Se necesitan exactamente 9 campos.\n\n" + AYUDA_AGREGAR

    item_id, nombre, categoria, stock_str, disp_str, uso_str, rep_str, estado, ubicacion = \
        partes[0], partes[1], partes[2], partes[3], partes[4], partes[5], partes[6], partes[7], partes[8]

    campos = [("ID", item_id), ("Nombre", nombre), ("Categor\u00eda", categoria),
              ("Stock Total", stock_str), ("Disponible", disp_str), ("En Uso", uso_str),
              ("En Reparaci\u00f3n", rep_str), ("Estado", estado), ("Ubicaci\u00f3n", ubicacion)]
    for campo, valor in campos:
        if not valor:
            return f"\u274c El campo *{campo}* no puede estar vac\u00edo.\n\n" + AYUDA_AGREGAR

    try:
        stock = int(stock_str)
        disp  = int(disp_str)
        uso   = int(uso_str)
        rep   = int(rep_str)
    except ValueError:
        return "\u274c Stock Total, Disponible, En Uso y En Reparaci\u00f3n deben ser n\u00fameros enteros."

    cats_validas = ["Herramientas", "Insumos", "EPP", "Repuestos"]
    if categoria not in cats_validas:
        return f"\u274c Categor\u00eda '{categoria}' no v\u00e1lida.\nUsa: {', '.join(cats_validas)}"
    estados_validos = ["OK", "En uso", "Parcial", "Cr\u00edtico"]
    if estado not in estados_validos:
        return f"\u274c Estado '{estado}' no v\u00e1lido.\nUsa: {', '.join(estados_validos)}"

    try:
        sheet = get_sheet()
        sheet.append_row([item_id, nombre, categoria, stock, disp, uso, rep, estado, ubicacion])
        return (
            f"\u2745 *\u00edtem agregado al inventario:*\n"
            f"\u2022 ID: {item_id}\n"
            f"\u2022 Nombre: {nombre}\n"
            f"\u2022 Categor\u00eda: {categoria}\n"
            f"\u2022 Stock Total: {stock}\n"
            f"\u2022 Disponible: {disp}\n"
            f"\u2022 En Uso: {uso}\n"
            f"\u2022 En Reparaci\u00f3n: {rep}\n"
            f"\u2022 Estado: {estado}\n"
            f"\u2022 Ubicaci\u00f3n: {ubicacion}\n\n"
            f"\u1f4ca Puedes ver el inventario actualizado en Google Sheets."
        )
    except Exception as e:
        return f"\u274c Error al guardar en Google Sheets: {str(e)}"

def responder(texto):
    t = texto.lower().strip()

    if t.startswith("agregar "):
        return agregar_item(texto[8:].strip())
    if t == "agregar":
        return AYUDA_AGREGAR

    df = cargar_inventario()
    if df is None:
        return "\u274c No pude leer el inventario. Verifica la conexi\u00f3n con Google Sheets."

    if t in ["hola", "hi", "ayuda", "help", "menu", "inicio", "men\u00fa"]:
        return (
            "\u1f44b Hola! Soy *Maquina*, tu bot de inventario de bodega.\n\n"
            "Puedes preguntarme:\n"
            "\u2022 *stock [\u00edtem]* \u2192 ver disponibilidad\n"
            "\u2022 *buscar [\u00edtem]* \u2192 buscar por nombre o ID\n"
            "\u2022 *agregar* \u2192 agregar un \u00edtem nuevo\n"
            "\u2022 *reparaci\u00f3n* \u2192 \u00edtems en reparaci\u00f3n\n"
            "\u2022 *en uso* \u2192 \u00edtems en uso\n"
            "\u2022 *cr\u00edtico* \u2192 \u00edtems con stock cr\u00edtico\n"
            "\u2022 *resumen* \u2192 dashboard general\n"
            "\u2022 *herramientas / insumos / epp / repuestos*\n\n"
            "Ejemplo: _stock taladro_ o _buscar H001_"
        )

    if t in ["resumen", "total", "dashboard"]:
        total = len(df)
        ok  = (df[COL_ESTADO] == "OK").sum()
        uso = (df[COL_ESTADO] == "En uso").sum()
        rep = df[COL_ESTADO].isin(["Parcial", "Cr\u00edtico"]).sum()
        cats = df[COL_CAT].value_counts()
        lineas = [f"\u1f4ca *Resumen de inventario*\nTotal: {total} | \u2745 OK: {ok} | \u1f535 En uso: {uso} | \u1f7e1 Incidencias: {rep}\n"]
        for cat, n in cats.items():
            lineas.append(f" \u2022 {cat}: {n} \u00edtems")
        return "\n".join(lineas)

    if t in ["reparacion", "reparaci\u00f3n", "en reparacion", "en reparaci\u00f3n", "reparar"]:
        sub = df[df[COL_REP] > 0]
        if sub.empty:
            return "\u2745 No hay \u00edtems en reparaci\u00f3n actualmente."
        lineas = [f"\u1f527 *\u00edtems en reparaci\u00f3n ({len(sub)}):*"]
        for _, row in sub.iterrows():
            lineas.append(f" \u2022 {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_REP]} unid.")
        return "\n".join(lineas)

    if t in ["en uso", "uso"]:
        sub = df[df[COL_USO] > 0]
        if sub.empty:
            return "\u2139\ufe0f No hay \u00edtems actualmente en uso."
        lineas = [f"\u1f535 *\u00edtems en uso ({len(sub)}):*"]
        for _, row in sub.iterrows():
            lineas.append(f" \u2022 {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_USO]} unid.")
        return "\n".join(lineas)

    if t in ["critico", "cr\u00edtico", "stock critico", "stock cr\u00edtico"]:
        sub = df[df[COL_ESTADO].isin(["Cr\u00edtico"])]
        if sub.empty:
            return "\u2745 No hay \u00edtems en estado cr\u00edtico."
        lineas = [f"\u1f534 *\u00edtems cr\u00edticos ({len(sub)}):*"]
        for _, row in sub.iterrows():
            lineas.append(f" \u2022 {row[COL_NOMBRE]} ({row[COL_ID]}): stock {row[COL_TOTAL]}")
        return "\n".join(lineas)

    for key, cat in {"herramientas": "Herramientas", "insumos": "Insumos", "epp": "EPP", "repuestos": "Repuestos"}.items():
        if t == key:
            sub = df[df[COL_CAT] == cat]
            if sub.empty:
                return f"No hay \u00edtems en {cat}."
            lineas = [f"\u1f4e6 *{cat} ({len(sub)} \u00edtems):*"]
            for _, row in sub.iterrows():
                lineas.append(f" \u2022 {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_DISP]} disp.")
            return "\n".join(lineas)

    for prefix in ["stock ", "buscar ", "ver ", "info "]:
        if t.startswith(prefix):
            consulta = t[len(prefix):]
            item = buscar_item(df, consulta)
            if item is not None:
                return formato_item(item)
            return f"\u1f50d No encontr\u00e9 \u00edtem con '{consulta}'.\nIntenta con el ID (ej: H001) o parte del nombre."

    item = buscar_item(df, t)
    if item is not None:
        return formato_item(item)

    return "\u1f914 No entend\u00ed esa consulta.\nEscribe *ayuda* para ver los comandos disponibles."

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
    try:
        msg  = data["entry"][0]["changes"][0]["value"]["messages"][0]
        body = msg.get("text", {}).get("body", "")
        if body:
            send_whatsapp_message(msg["from"], responder(body))
    except (KeyError, IndexError):
        pass
    return '{"status":"ok"}', 200

@app.route("/", methods=["GET"])
def index():
    return "Maquina - Bot de inventario activo \u2705", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
