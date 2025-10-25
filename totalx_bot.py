import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import Workbook, load_workbook
from datetime import datetime

# Carica variabili d'ambiente
load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# File principali
FILE_EXCEL = "Movimenti_Admin.xlsx"
FILE_MOVIMENTI = "Estratto_Conto_Admin.xlsx"

# Admin principale fisso
IMMUTABLE_ADMINS = ["Ela036"]

# Crea o carica il file movimenti admin
if not os.path.exists(FILE_MOVIMENTI):
    wb = Workbook()
    ws = wb.active
    ws.title = "Movimenti"
    ws.append(["username", "movimento", "data_ora"])
    wb.save(FILE_MOVIMENTI)

# Crea o carica il file admin
if not os.path.exists(FILE_EXCEL):
    wb_admin = Workbook()
    ws_admin = wb_admin.active
    ws_admin.title = "Admins"
    ws_admin.append(["username"])
    wb_admin.save(FILE_EXCEL)

# Funzioni utilità
def is_admin(username: str):
    username = username.lower()
    # Controlla admin principale
    for admin in IMMUTABLE_ADMINS:
        if username == admin.lower():
            return True
    # Controlla admin aggiunti nel foglio Excel
    if os.path.exists(FILE_EXCEL):
        wb_admin = load_workbook(FILE_EXCEL)
        if "Admins" in wb_admin.sheetnames:
            ws_admin = wb_admin["Admins"]
            for row in ws_admin.iter_rows(min_row=2, values_only=True):
                if row[0] and username == row[0].lower():
                    return True
    return False

def salva_movimento(username: str, valore: float):
    wb = load_workbook(FILE_MOVIMENTI)
    ws = wb["Movimenti"]
    ws.append([username, round(valore,2), datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    wb.save(FILE_MOVIMENTI)

def leggi_movimenti():
    movimenti = []
    wb = load_workbook(FILE_MOVIMENTI)
    ws = wb["Movimenti"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1] is not None:
            movimenti.append((row[0], float(row[1]), row[2]))
    return movimenti

def estratto_conto():
    movimenti = leggi_movimenti()
    saldo = sum([m[1] for m in movimenti])
    return movimenti, saldo

def crea_file_excel():
    movimenti, saldo = estratto_conto()
    wb_user = Workbook()
    ws_user = wb_user.active
    ws_user.title = "Estratto Conto"
    ws_user.append(["Username", "Tipo", "Importo", "Data/Ora"])
    for m in movimenti:
        tipo = "Entrata" if m[1] > 0 else "Uscita"
        ws_user.append([m[0], tipo, m[1], m[2]])
    ws_user.append([])
    ws_user.append(["Saldo Totale", saldo])
    filename = "Estratto_Conto.xlsx"
    wb_user.save(filename)
    return filename

# Comandi bot
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Ciao! Sono TotalX Pro Bot.\n"
        "Comandi:\n"
        "/add numero - aggiunge un'entrata (es. /add 100 o /add 0,05)\n"
        "/subtract numero - aggiunge un'uscita (es. /subtract 50 o /subtract 0,07)\n"
        "/total - mostra il saldo totale\n"
        "/report - mostra l'estratto conto completo\n"
        "/export - ricevi un file Excel con l'estratto conto\n"
        "/undo - annulla l'ultima operazione\n"
        "/reset - azzera tutto e crea un nuovo foglio\n"
        "/setadmin username - aggiungi/rimuovi un admin (solo admin)\n"
        "/adminlist - mostra la lista admin (solo admin)"
    )

async def add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = float(context.args[0].replace(",","."))
        username = update.message.from_user.username
        salva_movimento(username, value)
        movimenti, saldo = estratto_conto()
        await update.message.reply_text(f"Entrata registrata: +{round(value,2)}\nSaldo Totale: {round(saldo,2)}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /add numero, es. /add 100 o /add 0,05")

async def subtract(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = float(context.args[0].replace(",","."))
        username = update.message.from_user.username
        salva_movimento(username, -value)
        movimenti, saldo = estratto_conto()
        await update.message.reply_text(f"Uscita registrata: -{round(value,2)}\nSaldo Totale: {round(saldo,2)}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /subtract numero, es. /subtract 50 o /subtract 0,07")

async def total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    movimenti, saldo = estratto_conto()
    await update.message.reply_text(f"Saldo Totale: {round(saldo,2)}")

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    movimenti, saldo = estratto_conto()
    if not movimenti:
        await update.message.reply_text("Nessun movimento registrato.")
        return
    report_text = "📄 Estratto Conto Completo\n\n"
    for m in movimenti:
        tipo = "Entrata" if m[1] > 0 else "Uscita"
        report_text += f"{tipo}: {m[1]} ({m[0]} {m[2]})\n"
    report_text += f"\nSaldo Totale: {round(saldo,2)}"
    await update.message.reply_text(report_text)

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    filename = crea_file_excel()
    with open(filename, "rb") as file:
        await update.message.reply_document(file, filename=filename)

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    if not is_admin(username):
        await update.message.reply_text("❌ Solo admin possono resettare i dati.")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Movimenti"
    ws.append(["username", "movimento", "data_ora"])
    wb.save(FILE_MOVIMENTI)
    await update.message.reply_text("🗑️ Tutto azzerato. Nuovo foglio creato.")

async def undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    if not is_admin(username):
        await update.message.reply_text("❌ Solo admin possono annullare l'ultima operazione.")
        return
    wb = load_workbook(FILE_MOVIMENTI)
    ws = wb["Movimenti"]
    if ws.max_row > 1:
        ws.delete_rows(ws.max_row)
        wb.save(FILE_MOVIMENTI)
        await update.message.reply_text("↩️ Ultima operazione annullata.")
    else:
        await update.message.reply_text("❌ Nessuna operazione da annullare.")

async def setadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    if not is_admin(username):
        await update.message.reply_text("❌ Solo admin possono modificare la lista admin.")
        return
    try:
        target = context.args[0].replace("@","")
    except IndexError:
        await update.message.reply_text("Errore! Usa /setadmin username")
        return
    if target in IMMUTABLE_ADMINS:
        await update.message.reply_text("❌ Non puoi rimuovere l'admin principale.")
        return
    wb_admin = load_workbook(FILE_EXCEL)
    ws_admin = wb_admin["Admins"]
    # Verifica se già presente
    presenti = [row[0] for row in ws_admin.iter_rows(min_row=2, values_only=True)]
    if target in presenti:
        # Rimuovi admin
        for idx, row in enumerate(ws_admin.iter_rows(min_row=2, values_only=False), start=2):
            if row[0].value == target:
                ws_admin.delete_rows(idx)
                wb_admin.save(FILE_EXCEL)
                await update.message.reply_text(f"❌ Admin {target} rimosso.")
                return
    # Aggiungi admin
    ws_admin.append([target])
    wb_admin.save(FILE_EXCEL)
    await update.message.reply_text(f"✅ Admin {target} aggiunto.")

async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    if not is_admin(username):
        await update.message.reply_text("❌ Solo admin possono vedere la lista admin.")
        return
    admin_list = IMMUTABLE_ADMINS.copy()
    # Aggiunge altri admin dal foglio Excel
    wb_admin = load_workbook(FILE_EXCEL)
    ws_admin = wb_admin["Admins"]
    for row in ws_admin.iter_rows(min_row=2, values_only=True):
        if row[0] and row[0] not in IMMUTABLE_ADMINS:
            admin_list.append(row[0])
    await update.message.reply_text("👑 Lista Admin:\n" + "\n".join(admin_list))

# Main
def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("add", add))
    app.add_handler(CommandHandler("subtract", subtract))
    app.add_handler(CommandHandler("total", total))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("export", export))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("undo", undo))
    app.add_handler(CommandHandler("setadmin", setadmin))
    app.add_handler(CommandHandler("adminlist", adminlist))
    print("Bot avviato...")
    app.run_polling()

if __name__ == "__main__":
    main()
