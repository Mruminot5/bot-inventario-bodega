import os
import json
import requests
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from flask import Flask, request, jsonify
from difflib import get_close_matches

app = Flask(__name__)

VERIFY_TOKEN     = os.environ.get("WA_VERIFY_TOKEN", "bodega_inventario_2024")
ACCESS_TOKEN     = os.environ.get("WA_ACCESS_TOKEN", "")
PHONE_ID         = os.environ.get("WA_PHONE_NUMBER_ID", "")
SPREADSHEET_ID   = os.environ.get("SPREADSHEET_ID", "")
GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDENTIALS", "{}")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Columnas del Sheet (fila 2 = encabezados)
COL_ID     = "ID"
COL_NOMBRE = "Nombre del Ãtem"
COL_CAT    = "CategorÃ­a"
COL_TOTAL  = "Stock Total"
COL_DISP   = "Disponible"
COL_USO    = "En Uso"
COL_REP    = "En ReparaciÃ³n"
COL_ESTADO = "Estado"
COL_UBIC   = "UbicaciÃ³n"

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
        # Fila 1 = tÃ­tulo de la hoja, Fila 2 = encabezados, Fila 3+ = datos
        headers = data[1]
        rows = data[2:]
        df = pd.DataFrame(rows, columns=headers)
        df.columns = df.columns.str.strip()
        df = df[df[COL_ID].str.strip() != ""]
        # Convertir columnas numÃ©ricas
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
    emoji = {"OK": "â", "En uso": "ðµ", "Parcial": "ð¡", "CrÃ­tico": "ð´"}.get(estado, "âª")
    return (
        f"{emoji} *{item[COL_NOMBRE]}* ({item[COL_ID]})\n"
        f"â¢ CategorÃ­a: {item.get(COL_CAT, '')}\n"
        f"â¢ Stock total: {item.get(COL_TOTAL, 0)}\n"
        f"â¢ Disponible: {item.get(COL_DISP, 0)}\n"
        f"â¢ En uso: {item.get(COL_USO, 0)}\n"
        f"â¢ En reparaciÃ³n: {item.get(COL_REP, 0)}\n"
        f"â¢ Estado: {estado}\n"
        f"â¢ UbicaciÃ³n: {item.get(COL_UBIC, '')}"
    )

AYUDA_AGREGAR = (
    "ð *Formato para agregar un Ã­tem:*\n\n"
    "agregar ID | Nombre | CategorÃ­a | Stock Total | Disponible | En Uso | En ReparaciÃ³n | Estado | UbicaciÃ³n\n\n"
    "ð *Ejemplo:*\n"
    "_agregar H012 | Martillo grande | Herramientas | 5 | 4 | 1 | 0 | OK | Estante A_\n\n"
    "ð *CategorÃ­as vÃ¡lidas:* Herramientas, Insumos, EPP, Repuestos\n"
    "ð *Estados vÃ¡lidos:* OK, En uso, Parcial, CrÃ­tico"
)

def agregar_item(texto):
    partes = [p.strip() for p in texto.split("|")]
    if len(partes) < 9:
        return f"â Faltan campos. Se necesitan exactamente 9 campos.\n\n" + AYUDA_AGREGAR

    item_id, nombre, categoria, stock_str, disp_str, uso_str, rep_str, estado, ubicacion = \
        partes[0], partes[1], partes[2], partes[3], partes[4], partes[5], partes[6], partes[7], partes[8]

    # Validar vacÃ­os
    campos = [("ID", item_id), ("Nombre", nombre), ("CategorÃ­a", categoria),
              ("Stock Total", stock_str), ("Disponible", disp_str), ("En Uso", uso_str),
              ("En ReparaciÃ³n", rep_str), ("Estado", estado), ("UbicaciÃ³n", ubicacion)]
    for campo, valor in campos:
        if not valor:
            return f"â El campo *{campo}* no puede estar vacÃ­o.\n\n" + AYUDA_AGREGAR

    # Validar numÃ©ricos
    try:
        stock = int(stock_str)
        disp  = int(disp_str)
        uso   = int(uso_str)
        rep   = int(rep_str)
    except ValueError:
        return "â Stock Total, Disponible, En Uso y En ReparaciÃ³n deben ser nÃºmeros enteros."

    # Validar categorÃ­a y estado
    cats_validas = ["Herramientas", "Insumos", "EPP", "Repuestos"]
    if categoria not in cats_validas:
        return f"â CategorÃ­a '{categoria}' no vÃ¡lida.\nUsa: {{', '.join(cats_validas)}"
    estados_validos = ["OK", "En uso", "Parcial", "CrÃ­tico"]
    if estado not in estados_validos:
        return f"â Estado '{estado}' no vÃ¡lido.\nUsa: {', '.join(estados_validos)}"

    try:
        sheet = get_sheet()
        sheet.append_row([item_id, nombre, categoria, stock, disp, uso, rep, estado, ubicacion])
        return (
            f"â *Ãtem agregado al inventario:*\n"
            f"â¢ ID: {item_id}\n"
            f"â¢ Nombre: {nombre}\n"
            f"â¢ CategorÃ­a: {categoria}\n"
            f"â¢ Stock Total: {stock}\n"
            f"â¢ Disponible: {disp}\n"
            f"â¢ En Uso: {uso}\n"
            f"â¢ En ReparaciÃ³n: {rep}\n"
            f"â¢ Estado: {estado}\n"
            f"â¢ UbicaciÃ³n: {ubicacion}\n\n"
            f"ð Puedes ver el inventario actualizado en Google Sheets."
        )
    except Exception as e:
        return f"â Error al guardar en Google Sheets: {str(e)}"

def responder(texto):
    t = texto.lower().strip()

    if t.startswith("agregar "):
        return agregar_item(texto[8:].strip())
    if t == "agregar":
        return AYUDA_AGREGAR

    df = cargar_inventario()
    if df is None:
        return "â No pude leer el inventario. Verifica la conexiÃ³n con Google Sheets."

    if t in ["hola", "hi", "ayuda", "help", "menu", "inicio", "menÃº"]:
        return (
            "ð Hola! Soy *Maquina*, tu bot de inventario de bodega.\n\n"
            "Puedes preguntarme:\n"
            "â¢ *stock [Ã­tem]* â ver disponibilidad\n"
            "â¢ *buscar [Ã­tem]* â buscar por nombre o ID\n"
            "â¢ *agregar* â agregar un Ã­tem nuevo\n"
            "â¢ *reparaciÃ³n* â Ã­tems en reparaciÃ³n\n"
            "â¢ *en uso* â Ã­tems en uso\n"
            "â¢ *crÃ­tico* â Ã­tems con stock crÃ­tico\n"
            "â¢ *resumen* â dashboard general\n"
            "â¢ *herramientas / insumos / epp / repuestos*\n\n"
            "Ejemplo: _stock taladro_ o _buscar H001_"
        )

    if t in ["resumen", "total", "dashboard"]:
        total = len(df)
        ok  = (df[COL_ESTADO] == "OK").sum()
        uso = (df[COL_ESTADO] == "En uso").sum()
        rep = df[COL_ESTADO].isin(["Parcial", "CrÃ­tico"]).sum()
        cats = df[COL_CAT]value_counts()
        lineas = [f"ð *Resumen de inventario*\nTotal: {total} | â OK: {ok} | ðµ En uso: {uso} | ð¡ Incidencias: {rep}\n"]
        for cat, n in cats.items():
            lineas.append(f"  â¢ {cat}: {n} Ã­tems")
        return "\n".join(lineas)

    if t in ["reparacion", "reparaciÃ³n", "en reparacion", "en reparaciÃ³n", "reparar"]:
        sub = df[df[COL_REP] > 0]
        if sub.empty:
            return "â No hay Ã­tems en reparaciÃ³n actualmente."
        lineas = [f"ð§ *Ãtems en reparaciÃ³n ({len(sub)}):*"]
        for _, row in sub.iterrows():
            lineas.append(f"  â¢ {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_REP]} unid.")
        return "\n".join(lineas)

    if t in ["en uso", "uso"]:
        sub = df[df[COL_USO] > 0]
        if sub.empty:
            return "â¹ï¸ No hay Ã­tems actualmente en uso."
        lineas = [f"ðµ *Ãtems en uso ({len(sub)}):*"]
        for _, row in sub.iterrows():
            lineas.append(f"  â¢ {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_USO]} unid.")
        return "\n".join(lineas)

    if t in ["critico", "crÃ­tico", "stock critico", "stock crÃ­tico"]:
        sub = df[df[COL_ESTADO].isin(["CrÃ­tico"])]
        if sub.empty:
            return "â No hay Ã­tems en estado crÃ­tico."
        lineas = [f"ð´ *Ãtems crÃ­ticos ({len(sub)}):*"]
        for _, row in sub.iterrows():
            lineas.append(f"  â¢ {row[COL_NOMBRE]} ({row[COL_ID]}): stock {row[COL_TOTAL]}")
        return "\n".join(lineas)

    for key, cat in {"herramientas": "Herramientas", "insumos": "Insumos", "epp": "EPP", "repuestos": "Repuestos"}.items():
        if t == key:
            sub = df[df[COL_CAT] == cat]
            if sub.empty:
                return f"No hay Ã­tems en {cat}."
            lineas = [f"ð¦ *{cat} ({len(sub)} Ã­tems):*"]
            for _, row in sub.iterrows():
                lineas.append(f"  â¢ {row[COL_NOMBRE]} ({row[COL_ID]}): {row[COL_DISP]} disp.")
            return "\n".join(lineas)

    for prefix in ["stock ", "buscar ", "ver ", "info "]:
        if t.startswith(prefix):
            consulta = t[len(prefix):]
            item = buscar_item(df, consulta)
            if item is not None:
                return formato_item(item)
            return f"ð No encontrÃ© Ã­tem con '{consulta}'.\nIntenta con el ID (ej: H001) o parte del nombre."

    item = buscar_item(df, t)
    if item is not None:
        return formato_item(item)

    return "ð¤ No entendÃ­ esa consulta.\nEscribe *ayuda* para ver los comandos disponibles."

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
    mode = request.args.get("hub.mode")
    token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")
    if mode == "subscribe" and token == VERIFY_TOKEN:
        return challenge, 200
    return "Forbidden", 403

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json()
    try:
        msg = data["entry"][0]["changes"][0]["value"]["messages"][0]
        body = msg.get("text", {}).get("body", "")
        if body:
            send_whatsapp_message(msg["from"], responder(body))
    except (KeyError, IndexError):
        pass
    return '{"status":"ok"}', 200

@app.route("/", methods=["GET"])
def index():
    return "Maquina - Bot de inventario activo â", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
