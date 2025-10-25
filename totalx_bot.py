import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import Workbook, load_workbook
from datetime import datetime

# Carica variabili d'ambiente
load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

FILE_EXCEL = "estratto_conto.xlsx"

# Carica o crea il file Excel
if os.path.exists(FILE_EXCEL):
    wb = load_workbook(FILE_EXCEL)
    ws = wb.active
    if "Admins" not in wb.sheetnames:
        ws_admin = wb.create_sheet("Admins")
        ws_admin.append(["username"])
        # Admin di default
        ws_admin.append(["Ela036"])
        ws_admin.append(["Roby123"])
else:
    wb = Workbook()
    ws = wb.active
    ws.title = "Movimenti"
    ws.append(["user_id", "username", "movimento", "data_ora"])
    # Crea foglio admin
    ws_admin = wb.create_sheet("Admins")
    ws_admin.append(["username"])
    ws_admin.append(["Ela036"])
    ws_admin.append(["Roby123"])
    wb.save(FILE_EXCEL)

# Funzioni utilità
def salva_movimento(user_id: int, username: str, valore: float):
    data_ora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([user_id, username, valore, data_ora])
    wb.save(FILE_EXCEL)

def leggi_movimenti():
    movimenti = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        movimenti.append({
            "user_id": row[0],
            "username": row[1],
            "movimento": row[2],
            "data_ora": row[3]
        })
    return movimenti

def estratto_conto():
    movimenti = leggi_movimenti()
    totale = sum([m["movimento"] for m in movimenti])
    return movimenti, totale

def is_admin(username: str):
    if "Admins" not in wb.sheetnames:
        return False
    ws_admin = wb["Admins"]
    for row in ws_admin.iter_rows(min_row=2, values_only=True):
        if row[0] == username:
            return True
    return False

def add_admin(username: str):
    if not is_admin(username):
        ws_admin = wb["Admins"]
        ws_admin.append([username])
        wb.save(FILE_EXCEL)
        return True
    return False

def remove_admin(username: str):
    if is_admin(username):
        ws_admin = wb["Admins"]
        for idx, row in enumerate(ws_admin.iter_rows(min_row=2, values_only=False), start=2):
            if row[0].value == username:
                ws_admin.delete_rows(idx)
                wb.save(FILE_EXCEL)
                return True
    return False

def list_admins():
    ws_admin = wb["Admins"]
    admins = [row[0].value for row in ws_admin.iter_rows(min_row=2, values_only=True)]
    return admins

# Comandi bot
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Ciao! Sono TotalX Estratto Conto Bot Avanzato.\n"
        "Comandi:\n"
        "/add numero - aggiunge un'entrata\n"
        "/subtract numero - aggiunge un'uscita\n"
        "/total - mostra il saldo totale\n"
        "/report - mostra l'estratto conto dettagliato\n"
        "/export - ricevi un file Excel con il tuo estratto conto\n"
        "/undo - annulla l'ultima operazione\n"
        "/reset - azzera tutto e crea un nuovo foglio\n"
        "/setadmin username - aggiungi/rimuovi admin (solo admin)\n"
        "/adminlist - mostra la lista admin (solo admin)"
    )

async def add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = float(context.args[0].replace(",", "."))
        user_id = update.message.from_user.id
        username = update.message.from_user.username
        salva_movimento(user_id, username, value)
        _, saldo = estratto_conto()
        await update.message.reply_text(f"Entrata registrata: +{value}\nSaldo attuale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /add numero, esempio /add 100 o /add 0,05")

async def subtract(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = float(context.args[0].replace(",", "."))
        user_id = update.message.from_user.id
        username = update.message.from_user.username
        salva_movimento(user_id, username, -value)
        _, saldo = estratto_conto()
        await update.message.reply_text(f"Uscita registrata: -{value}\nSaldo attuale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /subtract numero, esempio /subtract 50 o /subtract 0,07")

async def total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    _, saldo = estratto_conto()
    await update.message.reply_text(f"Saldo totale: {saldo}")

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    movimenti, saldo = estratto_conto()
    if not movimenti:
        await update.message.reply_text("Nessun movimento registrato.")
        return
    report_text = "📄 Estratto Conto\n\n"
    for m in movimenti:
        tipo = "Entrata" if m["movimento"] > 0 else "Uscita"
        report_text += f"{tipo}: {m['movimento']} ({m['username']} {m['data_ora']})\n"
    report_text += f"\nSaldo Totale: {saldo}"
    await update.message.reply_text(report_text)

async def setadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    if not is_admin(username):
        await update.message.reply_text("Solo admin possono modificare la lista admin.")
        return
    try:
        target = context.args[0].replace("@", "")
        if add_admin(target):
            await update.message.reply_text(f"Admin {target} aggiunto.")
        else:
            if remove_admin(target):
                await update.message
