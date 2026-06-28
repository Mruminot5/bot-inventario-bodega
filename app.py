import os
import requests
import pandas as pd
from flask import Flask, request, jsonify
from difflib import get_close_matches

app = Flask(__name__)

VERIFY_TOKEN = os.environ.get("WA_VERIFY_TOKEN", "bodega_inventario_2024")
ACCESS_TOKEN = os.environ.get("WA_ACCESS_TOKEN", "")
PHONE_ID     = os.environ.get("WA_PHONE_NUMBER_ID", "")
EXCEL_PATH   = os.environ.get("EXCEL_PATH", "Inventario_Bodega.xlsx")

def cargar_inventario():
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name="Inventario", header=1)
        df.columns = df.columns.str.strip()
        return df.dropna(subset=["ID"])
    except Exception:
        return None

def buscar_item(df, consulta):
    consulta = consulta.lower().strip()
    match_id = df[df["ID"].str.lower() == consulta]
    if not match_id.empty:
        return match_id.iloc[0]
    nombres = df["Nombre del Item"].str.lower().tolist()
    ids = df["ID"].str.lower().tolist()
    cercanos = get_close_matches(consulta, nombres + ids, n=1, cutoff=0.4)
    if cercanos:
        match = df[
            (df["Nombre del Item"].str.lower() == cercanos[0]) |
            (df["ID"].str.lower() == cercanos[0])
        ]
        if not match.empty:
            return match.iloc[0]
    match_cont = df[df["Nombre del Item"].str.lower().str.contains(consulta, na=False)]
    if not match_cont.empty:
        return match_cont.iloc[0]
    return None

def formato_item(item):
    estado = str(item.get("Estado", ""))
    disp = item.get("Disponible", 0)
    try:
        disp = int(float(disp)) if pd.notna(disp) else 0
    except Exception:
        disp = 0
    return (
        f"[{estado}] {item['Nombre del Item']} ({item['ID']})"
        f"\nCategoria: {item.get('Categoria', '')}"
        f"\nStock total: {int(item.get('Stock Total', 0))}"
        f"\nDisponible: {disp}"
        f"\nEn uso: {int(item.get('En Uso', 0))}"
        f"\nEn reparacion: {int(item.get('En Reparacion', 0))}"
        f"\nUbicacion: {item.get('Ubicacion', '')}"
    )

def responder(texto):
    t = texto.lower().strip()
    df = cargar_inventario()
    if df is None:
        return "No pude leer el inventario. Verifica el archivo Excel."
    if t in ["hola", "hi", "ayuda", "help", "menu"]:
        return (
            "Hola! Soy el bot de inventario de bodega.\n\n"
            "Comandos disponibles:\n"
            "  stock [item] - ver disponibilidad\n"
            "  buscar [item] - buscar por nombre o ID\n"
            "  reparacion - items en reparacion\n"
            "  en uso - items actualmente en uso\n"
            "  critico - items con stock critico\n"
            "  resumen - dashboard general\n"
            "  herramientas / insumos / epp / repuestos\n\n"
            "Ejemplo: stock taladro  o  buscar H001"
        )
    if t in ["resumen", "total", "dashboard"]:
        total = len(df)
        ok  = (df["Estado"] == "OK").sum()
        uso = (df["Estado"] == "En uso").sum()
        rep = df["Estado"].isin(["Parcial", "Critico"]).sum()
        cats = df["Categoria"].value_counts()
        lineas = [f"Resumen de inventario\nTotal: {total} | OK: {ok} | En uso: {uso} | Incidencias: {rep}\n"]
        for cat, n in cats.items():
            lineas.append(f"  {cat}: {n} items")
        return "\n".join(lineas)
    if t in ["reparacion", "en reparacion", "reparar"]:
        sub = df[df["En Reparacion"] > 0]
        if sub.empty:
            return "No hay items en reparacion actualmente."
        lineas = [f"Items en reparacion ({len(sub)}):"]
        for _, row in sub.iterrows():
            lineas.append(f"  - {row['Nombre del Item']} ({row['ID']}): {int(row['En Reparacion'])} unid.")
        return "\n".join(lineas)
    if t in ["en uso", "uso"]:
        sub = df[df["En Uso"] > 0]
        if sub.empty:
            return "No hay items actualmente en uso."
        lineas = [f"Items en uso ({len(sub)}):"]
        for _, row in sub.iterrows():
            lineas.append(f"  - {row['Nombre del Item']} ({row['ID']}): {int(row['En Uso'])} unid.")
        return "\n".join(lineas)
    if t in ["critico", "stock critico"]:
        sub = df[df["Estado"].isin(["Critico"])]
        if sub.empty:
            return "No hay items en estado critico."
        lineas = [f"Items criticos ({len(sub)}):"]
        for _, row in sub.iterrows():
            lineas.append(f"  - {row['Nombre del Item']} ({row['ID']}): stock {int(row['Stock Total'])}")
        return "\n".join(lineas)
    for key, cat in {"herramientas": "Herramientas", "insumos": "Insumos", "epp": "EPP", "repuestos": "Repuestos"}.items():
        if t == key:
            sub = df[df["Categoria"] == cat]
            if sub.empty:
                return f"No hay items en {cat}."
            lineas = [f"{cat} ({len(sub)} items):"]
            for _, row in sub.iterrows():
                disp = int(row.get("Disponible", 0)) if pd.notna(row.get("Disponible", 0)) else 0
                lineas.append(f"  - {row['Nombre del Item']} ({row['ID']}): {disp} disp.")
            return "\n".join(lineas)
    for prefix in ["stock ", "buscar ", "ver ", "info "]:
        if t.startswith(prefix):
            consulta = t[len(prefix):]
            item = buscar_item(df, consulta)
            if item is not None:
                return formato_item(item)
            return f"No encontre item con '{consulta}'. Intenta con el ID o parte del nombre."
    item = buscar_item(df, t)
    if item is not None:
        return formato_item(item)
    return "No entendi esa consulta. Escribe 'ayuda' para ver los comandos."

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
    return "Bot de inventario activo", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

