import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import Workbook, load_workbook
from datetime import datetime

# Carica variabili d'ambiente
load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# File principale condiviso tra admin
FILE_EXCEL = "estratto_conto.xlsx"

# Lista admin iniziale (username)
ADMINS = ["tuo_username", "roby_username"]

# Carica o crea il file Excel principale
if os.path.exists(FILE_EXCEL):
    wb = load_workbook(FILE_EXCEL)
    ws = wb.active
else:
    wb = Workbook()
    ws = wb.active
    ws.append(["user_id", "username", "movimento", "data_ora"])
    wb.save(FILE_EXCEL)

# --- Funzioni di utilità ---
def salva_movimento(user_id: int, username: str, valore: float):
    ora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if username.lower() in [a.lower() for a in ADMINS]:
        # Scrive sul file principale
        ws.append([user_id, username, valore, ora])
        wb.save(FILE_EXCEL)
    else:
        # File separato per utente non admin
        filename = f"estratto_conto_{username}.xlsx"
        if os.path.exists(filename):
            wb_user = load_workbook(filename)
            ws_user = wb_user.active
        else:
            wb_user = Workbook()
            ws_user = wb_user.active
            ws_user.append(["user_id", "username", "movimento", "data_ora"])
        ws_user.append([user_id, username, valore, ora])
        wb_user.save(filename)

def leggi_movimenti(user_id: int, username: str, admin: bool):
    if admin:
        wb_data = wb
        ws_data = ws
    else:
        filename = f"estratto_conto_{username}.xlsx"
        if not os.path.exists(filename):
            return []
        wb_data = load_workbook(filename)
        ws_data = wb_data.active
    
    movimenti = []
    for row in ws_data.iter_rows(min_row=2, values_only=True):
        movimenti.append({"valore": row[2], "username": row[1], "data_ora": row[3]})
    return movimenti

def estratto_conto(user_id: int, username: str, admin: bool):
    movimenti = leggi_movimenti(user_id, username, admin)
    totale_entrate = sum([m["valore"] for m in movimenti if m["valore"] > 0])
    totale_uscite = sum([m["valore"] for m in movimenti if m["valore"] < 0])
    saldo = totale_entrate + totale_uscite
    return movimenti, totale_entrate, totale_uscite, saldo

def annulla_ultimo(user_id: int, username: str, admin: bool):
    if admin:
        ws_data = ws
        wb_data = wb
    else:
        filename = f"estratto_conto_{username}.xlsx"
        if not os.path.exists(filename):
            return False
        wb_data = load_workbook(filename)
        ws_data = wb_data.active
    
    rows = list(ws_data.iter_rows(min_row=2))
    for row in reversed(rows):
        if row[0].value == user_id:
            ws_data.delete_rows(row[0].row, 1)
            wb_data.save(FILE_EXCEL if admin else filename)
            return True
    return False

def reset_tutto(username: str, admin: bool):
    if admin:
        global ws
        wb.remove(ws)
        ws = wb.create_sheet("Sheet1")
        ws.append(["user_id", "username", "movimento", "data_ora"])
        wb.save(FILE_EXCEL)
    else:
        filename = f"estratto_conto_{username}.xlsx"
        if os.path.exists(filename):
            os.remove(filename)

def crea_file_excel(user_id: int, username: str, admin: bool):
    movimenti, totale_entrate, totale_uscite, saldo = estratto_conto(user_id, username, admin)
    wb_user = Workbook()
    ws_user = wb_user.active
    ws_user.title = "Estratto Conto"
    ws_user.append(["Tipo", "Importo", "Utente", "Data/Ora"])
    for m in movimenti:
        tipo = "Entrata" if m["valore"] > 0 else "Uscita"
        ws_user.append([tipo, m["valore"], m["username"], m["data_ora"]])
    ws_user.append([])
    ws_user.append(["Totale Entrate", totale_entrate])
    ws_user.append(["Totale Uscite", totale_uscite])
    ws_user.append(["Saldo Finale", saldo])
    filename = f"estratto_conto_export_{username}.xlsx" if not admin else "estratto_conto_completo.xlsx"
    wb_user.save(filename)
    return filename

# --- Comandi del bot ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Ciao! Sono TotalX Estratto Conto Bot Avanzato.\n"
        "Comandi disponibili:\n"
        "/add numero - aggiunge un'entrata\n"
        "/subtract numero - aggiunge un'uscita\n"
        "/total - mostra il saldo totale\n"
        "/report - mostra l'estratto conto completo\n"
        "/export - ricevi un file Excel con l'estratto conto\n"
        "/undo - annulla l'ultima operazione\n"
        "/reset - azzera tutto e crea un nuovo foglio\n"
        "/setadmin username - aggiunge/rimuove un admin (solo admin)\n"
        "/adminlist - mostra la lista degli admin (solo admin)"
    )

async def add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        raw_value = context.args[0].replace(",", ".")
        value = float(raw_value)
        user_id = update.message.from_user.id
        username = update.message.from_user.username or update.message.from_user.first_name
        admin = username.lower() in [a.lower() for a in ADMINS]
        salva_movimento(user_id, username, value)
        _, totale_entrate, totale_uscite, saldo = estratto_conto(user_id, username, admin)
        await update.message.reply_text(f"Entrata registrata: +{value}\nSaldo totale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /add numero, esempio /add 100 o /add 0,05")

async def subtract(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        raw_value = context.args[0].replace(",", ".")
        value = float(raw_value)
        user_id = update.message.from_user.id
        username = update.message.from_user.username or update.message.from_user.first_name
        admin = username.lower() in [a.lower() for a in ADMINS]
        salva_movimento(user_id, username, -value)
        _, totale_entrate, totale_uscite, saldo = estratto_conto(user_id, username, admin)
        await update.message.reply_text(f"Uscita registrata: -{value}\nSaldo totale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /subtract numero, esempio /subtract 50 o /subtract 0,05")

async def total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    username = update.message.from_user.username or update.message.from_user.first_name
    admin = username.lower() in [a.lower() for a in ADMINS]
    _, totale_entrate, totale_uscite, saldo = estratto_conto(user_id, username, admin)
    await update.message.reply_text(f"Saldo totale: {saldo}")

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    username = update.message.from_user.username or update.message.from_user.first_name
    admin = username.lower() in [a.lower() for a in ADMINS]
    movimenti, totale_entrate, totale_uscite, saldo = estratto_conto(user_id, username, admin)
    if not movimenti:
        await update.message.reply_text("Nessun movimento registrato.")
        return
    report_text = "📄 Estratto Conto\n\n"
    for m in movimenti:
        tipo = "Entrata" if m["valore"] > 0 else "Uscita"
        report_text += f"{tipo}: {m['valore']} ({m['username']} {m['data_ora']})\n"
    report_text += f"\nTotale Entrate: {totale_entrate}\nTotale Uscite: {totale_uscite}\nSaldo Totale: {saldo}"
    await update.message.reply_text(report_text)

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    username = update.message.from_user.username or update.message.from_user.first_name
    admin = username.lower() in [a.lower() for a in ADMINS]
    filename = crea_file_excel(user_id, username, admin)
    with open(filename, "rb") as file:
        await update.message.reply_document(file, filename=filename)

async def undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    username = update.message.from_user.username or update.message.from_user.first_name
    admin = username.lower() in [a.lower() for a in ADMINS]
    success = annulla_ultimo(user_id, username, admin)
    await update.message.reply_text("Ultima operazione annullata." if success else "Nessuna operazione da annullare.")

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    admin = username.lower() in [a.lower() for a in ADMINS]
    reset_tutto(username, admin)
    await update.message.reply_text("Tutto azzerato. Nuovo foglio creato.")

async def setadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    if username.lower() not in [a.lower() for a in ADMINS]:
        await update.message.reply_text("Solo admin possono modificare la lista admin.")
        return
    try:
        target = context.args[0].lower()
        if target in [a.lower() for a in ADMINS]:
            ADMINS.remove(next(a for a in ADMINS if a.lower() == target))
            await update.message.reply_text(f"{target} rimosso dagli admin.")
        else:
            ADMINS.append(target)
            await update.message.reply_text(f"{target} aggiunto come admin.")
    except IndexError:
        await update.message.reply_text("Usa /setadmin username")

async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    if username.lower() not in [a.lower() for a in ADMINS]:
        await update.message.reply_text("Solo admin possono vedere la lista admin.")
        return
    await update.message.reply_text("Lista admin: " + ", ".join(ADMINS))

# --- Main ---
def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("add", add))
    app.add_handler(CommandHandler("subtract", subtract))
    app.add_handler(CommandHandler("total", total))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("export", export))
    app.add_handler(CommandHandler("undo", undo))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("setadmin", setadmin))
    app.add_handler(CommandHandler("adminlist", adminlist))
    print("Bot avviato...")
    app.run_polling()

if __name__ == "__main__":
    main()
