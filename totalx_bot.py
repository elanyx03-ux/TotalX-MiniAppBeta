import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import Workbook, load_workbook
from datetime import datetime

# Carica variabili d'ambiente
load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# Nome del file principale admin
ADMIN_FILE = "estratto_conto_admin.xlsx"
ADMIN_USERNAME = "@Elanyx03"

# Admin fissi
FIXED_ADMINS = {ADMIN_USERNAME}

# Lista admin dinamica (altri admin aggiunti da admin fisso)
admins = set(FIXED_ADMINS)

# Carica o crea file admin
if os.path.exists(ADMIN_FILE):
    wb_admin = load_workbook(ADMIN_FILE)
    ws_admin = wb_admin.active
else:
    wb_admin = Workbook()
    ws_admin = wb_admin.active
    ws_admin.append(["user", "movimento", "data"])
    wb_admin.save(ADMIN_FILE)

# Funzioni di utilità
def salva_movimento(user, valore, admin=False):
    filename = ADMIN_FILE if admin else f"estratto_{user}.xlsx"
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["user", "movimento", "data"])

    ws.append([user, round(valore, 2), datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    wb.save(filename)

def leggi_movimenti(user, admin=False):
    filename = ADMIN_FILE if admin else f"estratto_{user}.xlsx"
    movimenti = []
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            movimenti.append(row)
    return movimenti

def estratto_conto(user, admin=False):
    movimenti = leggi_movimenti(user, admin)
    entrate = [m[1] for m in movimenti if m[1] > 0]
    uscite = [m[1] for m in movimenti if m[1] < 0]
    totale_entrate = round(sum(entrate), 2)
    totale_uscite = round(sum(uscite), 2)
    saldo = round(totale_entrate + totale_uscite, 2)
    return movimenti, entrate, uscite, totale_entrate, totale_uscite, saldo

def crea_file_excel_utente(user, admin=False):
    movimenti, entrate, uscite, totale_entrate, totale_uscite, saldo = estratto_conto(user, admin)
    wb_user = Workbook()
    ws_user = wb_user.active
    ws_user.title = "Estratto Conto"
    ws_user.append(["Tipo", "Importo", "Utente", "Data"])
    for m in movimenti:
        tipo = "Entrata" if m[1] > 0 else "Uscita"
        ws_user.append([tipo, round(m[1], 2), m[0], m[2]])
    ws_user.append([])
    ws_user.append(["Totale Entrate", totale_entrate])
    ws_user.append(["Totale Uscite", totale_uscite])
    ws_user.append(["Saldo Finale", saldo])
    filename = f"estratto_{user}.xlsx" if not admin else "estratto_conto_admin_export.xlsx"
    wb_user.save(filename)
    return filename

# Comandi bot
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Ciao! Sono TotalX Pro Bot.\n"
        "Comandi:\n"
        "/add numero - aggiunge un'entrata\n"
        "/subtract numero - aggiunge un'uscita\n"
        "/total - mostra il saldo totale\n"
        "/report - mostra l'estratto conto completo\n"
        "/export - ricevi un file Excel con l'estratto conto\n"
        "/undo - annulla l'ultima operazione\n"
        "/reset - azzera tutto e crea un nuovo foglio\n"
        "/setadmin username - aggiunge/rimuove un admin (solo admin)\n"
        "/adminlist - mostra la lista admin (solo admin)"
    )

async def add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = float(context.args[0].replace(",", "."))
        user = update.message.from_user.username
        admin_flag = user in admins
        salva_movimento(user, value, admin=admin_flag)
        _, _, _, _, _, saldo = estratto_conto(user, admin=admin_flag)
        await update.message.reply_text(f"Entrata registrata: +{round(value,2)}\nSaldo attuale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /add numero, esempio /add 100 o /add 0,05")

async def subtract(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = float(context.args[0].replace(",", "."))
        user = update.message.from_user.username
        admin_flag = user in admins
        salva_movimento(user, -value, admin=admin_flag)
        _, _, _, _, _, saldo = estratto_conto(user, admin=admin_flag)
        await update.message.reply_text(f"Uscita registrata: -{round(value,2)}\nSaldo attuale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /subtract numero, esempio /subtract 50 o /subtract 0,07")

async def total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username
    admin_flag = user in admins
    _, _, _, _, _, saldo = estratto_conto(user, admin=admin_flag)
    await update.message.reply_text(f"Saldo totale: {saldo}")

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username
    admin_flag = user in admins
    movimenti, entrate, uscite, totale_entrate, totale_uscite, saldo = estratto_conto(user, admin=admin_flag)
    if not movimenti:
        await update.message.reply_text("Nessun movimento registrato.")
        return
    text = "📄 Estratto Conto\n\n"
    for m in movimenti:
        tipo = "Entrata" if m[1] > 0 else "Uscita"
        text += f"{tipo}: {round(m[1],2)} ({m[0]} {m[2]})\n"
    text += f"\nTotale Entrate: {totale_entrate}\nTotale Uscite: {totale_uscite}\nSaldo Totale: {saldo}"
    await update.message.reply_text(text)

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username
    admin_flag = user in admins
    filename = crea_file_excel_utente(user, admin=admin_flag)
    with open(filename, "rb") as file:
        await update.message.reply_document(file, filename=filename)

async def undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username
    admin_flag = user in admins
    filename = ADMIN_FILE if admin_flag else f"estratto_{user}.xlsx"
    if not os.path.exists(filename):
        await update.message.reply_text("Nessun movimento da annullare.")
        return
    wb = load_workbook(filename)
    ws = wb.active
    if ws.max_row <= 1:
        await update.message.reply_text("Nessun movimento da annullare.")
        return
    ws.delete_rows(ws.max_row)
    wb.save(filename)
    await update.message.reply_text("Ultimo movimento annullato.")

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username
    admin_flag = user in admins
    filename = ADMIN_FILE if admin_flag else f"estratto_{user}.xlsx"
    if os.path.exists(filename):
        os.remove(filename)
    wb = Workbook()
    ws = wb.active
    ws.append(["user", "movimento", "data"])
    wb.save(filename)
    await update.message.reply_text("Foglio azzerato e ricreato a saldo zero.")

async def setadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username
    if user not in admins:
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
        await update.message.reply_text("Errore! Usa /setadmin @username")

async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.username
    if user not in admins:
        await update.message.reply_text("Solo admin possono vedere la lista admin.")
        return
    text = "Admin attuali:\n" + "\n".join(admins)
    await update.message.reply_text(text)

# Funzione principale
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
