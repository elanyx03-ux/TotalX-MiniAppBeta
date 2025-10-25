import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import Workbook, load_workbook
from datetime import datetime

# Carica variabili d'ambiente
load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# Nome del file principale degli admin
ADMIN_FILE = "Movimenti_Admin.xlsx"

# Lista admin fissi
FIXED_ADMINS = ["@Ela036"]

# Controlla se il file admin esiste
if os.path.exists(ADMIN_FILE):
    wb_admin = load_workbook(ADMIN_FILE)
    ws_admin = wb_admin.active
else:
    wb_admin = Workbook()
    ws_admin = wb_admin.active
    ws_admin.append(["User", "Movimento", "Data"])
    wb_admin.save(ADMIN_FILE)

# Funzioni generiche
def get_user_file(username: str):
    """Restituisce il nome del file utente, crea se non esiste"""
    filename = f"Movimenti_{username}.xlsx"
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["User", "Movimento", "Data"])
        wb.save(filename)
    return filename, wb, ws

def salva_movimento(username: str, value: float, admin=False):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if admin:
        ws_admin.append([username, round(value,2), now])
        wb_admin.save(ADMIN_FILE)
    else:
        filename, wb, ws = get_user_file(username)
        ws.append([username, round(value,2), now])
        wb.save(filename)

def leggi_movimenti(username: str, admin=False):
    movimenti = []
    if admin:
        for row in ws_admin.iter_rows(min_row=2, values_only=True):
            movimenti.append(row)
    else:
        filename, wb, ws = get_user_file(username)
        for row in ws.iter_rows(min_row=2, values_only=True):
            movimenti.append(row)
    return movimenti

def estratto_conto(username: str, admin=False):
    movimenti = leggi_movimenti(username, admin)
    entrate = [m[1] for m in movimenti if m[1] > 0]
    uscite = [m[1] for m in movimenti if m[1] < 0]
    totale_entrate = round(sum(entrate),2)
    totale_uscite = round(sum(uscite),2)
    saldo = round(totale_entrate + sum(uscite),2)
    return entrate, uscite, totale_entrate, totale_uscite, saldo, movimenti

# Lista admin modificabile
admins = set(FIXED_ADMINS)

# Comandi bot
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Ciao! Sono TotalX Pro.\n"
        "Comandi:\n"
        "/add numero - aggiunge un'entrata\n"
        "/subtract numero - aggiunge un'uscita\n"
        "/total - mostra il saldo totale\n"
        "/report - mostra l'estratto conto completo\n"
        "/export - ricevi un file Excel con il tuo estratto conto\n"
        "/undo - annulla l'ultima operazione\n"
        "/reset - azzera tutto e crea un nuovo foglio\n"
        "/setadmin username - aggiungi/rimuovi admin (solo admin)\n"
        "/adminlist - mostra la lista admin (solo admin)"
    )

async def add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = float(context.args[0].replace(",","."))
        username = update.message.from_user.username or update.message.from_user.first_name
        admin_mode = f"@{username}" in admins
        salva_movimento(f"@{username}", value, admin=admin_mode)
        _, _, _, _, saldo, _ = estratto_conto(f"@{username}", admin=admin_mode)
        await update.message.reply_text(f"Entrata registrata: +{round(value,2)}\nSaldo attuale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /add numero, esempio /add 100 o /add 0,05")

async def subtract(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = float(context.args[0].replace(",","."))
        username = update.message.from_user.username or update.message.from_user.first_name
        admin_mode = f"@{username}" in admins
        salva_movimento(f"@{username}", -value, admin=admin_mode)
        _, _, _, _, saldo, _ = estratto_conto(f"@{username}", admin=admin_mode)
        await update.message.reply_text(f"Uscita registrata: -{round(value,2)}\nSaldo attuale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /subtract numero, esempio /subtract 50 o /subtract 0,07")

async def total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    admin_mode = f"@{username}" in admins
    _, _, _, _, saldo, _ = estratto_conto(f"@{username}", admin=admin_mode)
    await update.message.reply_text(f"Saldo totale: {saldo}")

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    admin_mode = f"@{username}" in admins
    entrate, uscite, totale_entrate, totale_uscite, saldo, movimenti = estratto_conto(f"@{username}", admin=admin_mode)
    if not movimenti:
        await update.message.reply_text("Nessun movimento registrato.")
        return
    report_text = "📄 Estratto Conto\n\n"
    for m in movimenti:
        tipo = "Entrata" if m[1]>0 else "Uscita"
        report_text += f"{tipo}: {m[1]} ({m[0]} {m[2]})\n"
    report_text += f"\nTotale Entrate: {totale_entrate}\nTotale Uscite: {totale_uscite}\nSaldo Totale: {saldo}"
    await update.message.reply_text(report_text)

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    admin_mode = f"@{username}" in admins
    filename, wb, _ = get_user_file(f"@{username}") if not admin_mode else (ADMIN_FILE, wb_admin, ws_admin)
    with open(filename, "rb") as file:
        await update.message.reply_document(file, filename=filename)

async def undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    admin_mode = f"@{username}" in admins
    filename, wb, ws = get_user_file(f"@{username}") if not admin_mode else (ADMIN_FILE, wb_admin, ws_admin)
    if ws.max_row>1:
        ws.delete_rows(ws.max_row)
        wb.save(filename)
        await update.message.reply_text("Ultima operazione annullata.")
    else:
        await update.message.reply_text("Nessuna operazione da annullare.")

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    admin_mode = f"@{username}" in admins
    filename, wb, ws = get_user_file(f"@{username}") if not admin_mode else (ADMIN_FILE, wb_admin, ws_admin)
    wb = Workbook()
    ws = wb.active
    ws.append(["User", "Movimento", "Data"])
    wb.save(filename)
    await update.message.reply_text("Foglio azzerato e nuovo file creato.")

async def setadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    if f"@{username}" not in admins:
        await update.message.reply_text("Solo admin possono modificare la lista admin.")
        return
    try:
        target = context.args[0]
        if target in FIXED_ADMINS:
            await update.message.reply_text(f"{target} è un admin fisso e non può essere rimosso.")
            return
        if target in admins:
            admins.remove(target)
            await update.message.reply_text(f"{target} rimosso dagli admin.")
        else:
            admins.add(target)
            await update.message.reply_text(f"{target} aggiunto come admin.")
    except IndexError:
        await update.message.reply_text("Errore! Usa /setadmin username")

async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username or update.message.from_user.first_name
    if f"@{username}" not in admins:
        await update.message.reply_text("Solo admin possono vedere la lista admin.")
        return
    lista = "\n".join(admins)
    await update.message.reply_text(f"Lista admin:\n{lista}")

# Avvio bot
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
