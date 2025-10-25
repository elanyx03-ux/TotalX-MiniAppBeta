import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import Workbook, load_workbook
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# File principale admin
FILE_ADMIN = "estratto_conto_admin.xlsx"

# Admin fisso e dinamico
FIXED_ADMIN = "@Elanyx03"
dynamic_admins = set()

# Carica o crea file admin
if os.path.exists(FILE_ADMIN):
    wb_admin = load_workbook(FILE_ADMIN)
    ws_admin = wb_admin.active
else:
    wb_admin = Workbook()
    ws_admin = wb_admin.active
    ws_admin.append(["username", "movimento", "data"])
    wb_admin.save(FILE_ADMIN)

# Funzione utilità: arrotondamento decimali
def round_decimal(value):
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

# Funzioni Excel
def salva_movimento(username, valore, admin=False):
    valore = round_decimal(valore)
    if admin:
        ws = ws_admin
        wb = wb_admin
        filename = FILE_ADMIN
    else:
        filename = f"estratto_conto_{username}.xlsx"
        if os.path.exists(filename):
            wb = load_workbook(filename)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(["username", "movimento", "data"])
    ws.append([username, valore, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    wb.save(filename)

def leggi_movimenti(username, admin=False):
    movimenti = []
    ws = ws_admin if admin else load_workbook(f"estratto_conto_{username}.xlsx").active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == username or admin:
            movimenti.append(Decimal(str(row[1])))
    return movimenti

def estratto_conto(username, admin=False):
    movimenti = leggi_movimenti(username, admin)
    entrate = [m for m in movimenti if m > 0]
    uscite = [m for m in movimenti if m < 0]
    totale_entrate = sum(entrate)
    totale_uscite = sum(uscite)
    saldo = totale_entrate + totale_uscite
    return entrate, uscite, totale_entrate, totale_uscite, saldo

def crea_file_excel(username, admin=False):
    entrate, uscite, totale_entrate, totale_uscite, saldo = estratto_conto(username, admin)
    wb_user = Workbook()
    ws_user = wb_user.active
    ws_user.title = "Estratto Conto"
    ws_user.append(["Tipo", "Importo"])
    for m in entrate:
        ws_user.append(["Entrata", float(m)])
    for m in uscite:
        ws_user.append(["Uscita", float(m)])
    ws_user.append([])
    ws_user.append(["Totale Entrate", float(totale_entrate)])
    ws_user.append(["Totale Uscite", float(totale_uscite)])
    ws_user.append(["Saldo Finale", float(saldo)])
    filename = f"estratto_conto_{username}.xlsx" if not admin else FILE_ADMIN
    wb_user.save(filename)
    return filename

# Funzione admin check
def is_admin(username):
    return username == FIXED_ADMIN or username in dynamic_admins

# Comandi bot
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Ciao! Sono TotalX Pro Bot.\nComandi:\n"
        "/add numero - aggiunge un'entrata\n"
        "/subtract numero - aggiunge un'uscita\n"
        "/total - mostra il saldo totale\n"
        "/report - mostra l'estratto conto completo\n"
        "/export - ricevi un file Excel\n"
        "/undo - annulla l'ultima operazione\n"
        "/reset - azzera tutto e crea un nuovo foglio\n"
        "/setadmin username - aggiungi/rimuovi admin (solo admin principale)\n"
        "/adminlist - mostra lista admin (solo admin)"
    )

async def add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = round_decimal(float(context.args[0].replace(",", ".")))
        username = update.message.from_user.username
        admin = is_admin(username)
        salva_movimento(username, value, admin=admin)
        _, _, _, _, saldo = estratto_conto(username, admin=admin)
        await update.message.reply_text(f"Entrata registrata: +{value}\nSaldo attuale: {saldo}")
    except:
        await update.message.reply_text("Errore! Usa /add numero (es. /add 100 o /add 0,05)")

async def subtract(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = round_decimal(float(context.args[0].replace(",", ".")))
        username = update.message.from_user.username
        admin = is_admin(username)
        salva_movimento(username, -value, admin=admin)
        _, _, _, _, saldo = estratto_conto(username, admin=admin)
        await update.message.reply_text(f"Uscita registrata: -{value}\nSaldo attuale: {saldo}")
    except:
        await update.message.reply_text("Errore! Usa /subtract numero (es. /subtract 50)")

async def total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    admin = is_admin(username)
    _, _, _, _, saldo = estratto_conto(username, admin=admin)
    await update.message.reply_text(f"Saldo totale: {saldo}")

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    admin = is_admin(username)
    entrate, uscite, tot_entrate, tot_uscite, saldo = estratto_conto(username, admin=admin)
    text = "📄 Estratto Conto\n"
    if entrate:
        text += "Entrate:\n" + "\n".join([f"+{float(m)}" for m in entrate]) + f"\nTotale Entrate: {float(tot_entrate)}\n"
    if uscite:
        text += "Uscite:\n" + "\n".join([f"{float(m)}" for m in uscite]) + f"\nTotale Uscite: {float(tot_uscite)}\n"
    text += f"Saldo Totale: {float(saldo)}"
    await update.message.reply_text(text)

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    admin = is_admin(username)
    filename = crea_file_excel(username, admin=admin)
    with open(filename, "rb") as file:
        await update.message.reply_document(file, filename=filename)

async def undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    admin = is_admin(username)
    filename = FILE_ADMIN if admin else f"estratto_conto_{username}.xlsx"
    wb = wb_admin if admin else load_workbook(filename)
    ws = wb.active
    if ws.max_row > 1:
        ws.delete_rows(ws.max_row)
        wb.save(filename)
        await update.message.reply_text("Ultima operazione annullata.")
    else:
        await update.message.reply_text("Nessuna operazione da annullare.")

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    admin = is_admin(username)
    filename = FILE_ADMIN if admin else f"estratto_conto_{username}.xlsx"
    if os.path.exists(filename):
        os.remove(filename)
    wb_new = Workbook()
    ws_new = wb_new.active
    ws_new.append(["username", "movimento", "data"])
    wb_new.save(filename)
    await update.message.reply_text("Foglio resettato correttamente, saldo azzerato.")

# Admin commands
async def setadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    if username != FIXED_ADMIN:
        await update.message.reply_text("Solo l'admin principale può modificare la lista admin.")
        return
    try:
        target = context.args[0]
    except:
        await update.message.reply_text("Usa: /setadmin @username")
        return
    if target == FIXED_ADMIN:
        await update.message.reply_text("Non puoi rimuovere l'admin principale.")
        return
    if target in dynamic_admins:
        dynamic_admins.remove(target)
        await update.message.reply_text(f"{target} rimosso dagli admin.")
    else:
        dynamic_admins.add(target)
        await update.message.reply_text(f"{target} aggiunto come admin.")

async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    if not is_admin(username):
        await update.message.reply_text("Solo gli admin possono vedere la lista admin.")
        return
    admins = [FIXED_ADMIN] + list(dynamic_admins)
    await update.message.reply_text("Admin attuali:\n" + "\n".join(admins))

# Main
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
