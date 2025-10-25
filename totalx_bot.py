import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import Workbook, load_workbook
from datetime import datetime
from decimal import Decimal, getcontext
import shutil

# Precisione decimali
getcontext().prec = 10

# Carica token
load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# File admin
FILE_ADMIN = "Movimenti_Admin.xlsx"

# Admin bloccati
IMMUTABLE_ADMINS = ["Ela036", "NyX0369"]

# Inizializza file Excel
def init_excel(filename, is_admin=False):
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Movimenti"
        ws.append(["user_id", "username", "movimento", "data_ora"])
        if is_admin:
            if "Admins" not in wb.sheetnames:
                ws_admin = wb.create_sheet("Admins")
                ws_admin.append(["username"])
                for admin in IMMUTABLE_ADMINS:
                    ws_admin.append([admin])
        wb.save(filename)
    return wb, wb.active

# Funzioni utilità
def salva_movimento(user_id: int, username: str, valore: Decimal, filename: str):
    wb, ws = init_excel(filename)
    data_ora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([user_id, username, str(valore), data_ora])
    wb.save(filename)

def leggi_movimenti(filename: str):
    wb, ws = init_excel(filename)
    movimenti = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        movimenti.append({
            "user_id": row[0],
            "username": row[1],
            "movimento": Decimal(row[2]),
            "data_ora": row[3]
        })
    return movimenti

def estratto_conto(filename: str):
    movimenti = leggi_movimenti(filename)
    totale = sum([m["movimento"] for m in movimenti])
    return movimenti, totale

def is_admin(username: str):
    if username in IMMUTABLE_ADMINS:
        return True
    wb, ws = init_excel(FILE_ADMIN, is_admin=True)
    if "Admins" not in wb.sheetnames:
        return False
    ws_admin = wb["Admins"]
    for row in ws_admin.iter_rows(min_row=2, values_only=True):
        if row[0] == username:
            return True
    return False

def add_admin(username: str):
    if username in IMMUTABLE_ADMINS:
        return False
    wb, ws = init_excel(FILE_ADMIN, is_admin=True)
    ws_admin = wb["Admins"]
    if not is_admin(username):
        ws_admin.append([username])
        wb.save(FILE_ADMIN)
        return True
    return False

def remove_admin(username: str):
    if username in IMMUTABLE_ADMINS:
        return False
    wb, ws = init_excel(FILE_ADMIN, is_admin=True)
    ws_admin = wb["Admins"]
    for idx, row in enumerate(ws_admin.iter_rows(min_row=2, values_only=False), start=2):
        if row[0].value == username:
            ws_admin.delete_rows(idx)
            wb.save(FILE_ADMIN)
            return True
    return False

def list_admins():
    wb, ws = init_excel(FILE_ADMIN, is_admin=True)
    ws_admin = wb["Admins"]
    admins = [row[0].value for row in ws_admin.iter_rows(min_row=2, values_only=True)]
    return admins

def get_user_file(username: str):
    if is_admin(username):
        return FILE_ADMIN
    else:
        return f"Movimenti_{username}.xlsx"

# Comandi bot
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Ciao! Sono TotalX Estratto Conto Bot Avanzato.\n"
        "Comandi:\n"
        "/add numero - aggiunge un'entrata\n"
        "/subtract numero - aggiunge un'uscita\n"
        "/total - mostra il saldo totale\n"
        "/report - mostra l'estratto conto completo\n"
        "/export - ricevi un file Excel con l'estratto conto\n"
        "/undo - annulla l'ultima operazione\n"
        "/reset - azzera tutto e crea un nuovo foglio\n"
        "/setadmin username - aggiunge/rimuove admin (solo admin)\n"
        "/adminlist - mostra la lista degli admin (solo admin)"
    )

async def add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = Decimal(context.args[0].replace(",", "."))
        user_id = update.message.from_user.id
        username = update.message.from_user.username
        filename = get_user_file(username)
        salva_movimento(user_id, username, value, filename)
        _, saldo = estratto_conto(filename)
        await update.message.reply_text(f"Entrata registrata: +{value:.2f}\nSaldo attuale: {saldo:.2f}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /add numero, esempio /add 100 o /add 0,05")

async def subtract(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = Decimal(context.args[0].replace(",", "."))
        user_id = update.message.from_user.id
        username = update.message.from_user.username
        filename = get_user_file(username)
        salva_movimento(user_id, username, -value, filename)
        _, saldo = estratto_conto(filename)
        await update.message.reply_text(f"Uscita registrata: -{value:.2f}\nSaldo attuale: {saldo:.2f}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /subtract numero, esempio /subtract 50 o /subtract 0,07")

async def total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    filename = get_user_file(username)
    _, saldo = estratto_conto(filename)
    await update.message.reply_text(f"Saldo totale: {saldo:.2f}")

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    filename = get_user_file(username)
    movimenti, saldo = estratto_conto(filename)
    if not movimenti:
        await update.message.reply_text("Nessun movimento registrato.")
        return
    report_text = "📄 Estratto Conto\n\n"
    for m in movimenti:
        tipo = "Entrata" if m["movimento"] > 0 else "Uscita"
        report_text += f"{tipo}: {m['movimento']:.2f} ({m['username']} {m['data_ora']})\n"
    report_text += f"\nSaldo Totale: {saldo:.2f}"
    await update.message.reply_text(report_text)

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    filename = get_user_file(username)
    temp_filename = f"export_{username}.xlsx"
    shutil.copy(filename, temp_filename)
    with open(temp_filename, "rb") as file:
        await update.message.reply_document(file, filename=temp_filename)
    os.remove(temp_filename)

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    filename = get_user_file(username)
    if os.path.exists(filename):
        os.remove(filename)
    init_excel(filename, is_admin=is_admin(username))
    await update.message.reply_text("Foglio azzerato e ricreato con successo!")

async def undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    filename = get_user_file(username)
    wb, ws = init_excel(filename)
    max_row = ws.max_row
    if max_row > 1:
        ws.delete_rows(max_row)
        wb.save(filename)
        await update.message.reply_text("Ultima operazione annullata.")
    else:
        await update.message.reply_text("Nessuna operazione da annullare.")

async def setadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    if not is_admin(username):
        await update.message.reply_text("Solo admin possono modificare la lista admin.")
        return
    try:
        target = context.args[0].replace("@", "")
        if is_admin(target):
            removed = remove_admin(target)
            if removed:
                await update.message.reply_text(f"{target} rimosso dagli admin.")
            else:
                await update.message.reply_text(f"{target} non può essere rimosso.")
        else:
            added = add_admin(target)
            if added:
                await update.message.reply_text(f"{target} aggiunto come admin.")
            else:
                await update.message.reply_text(f"{target} è già admin.")
    except IndexError:
        await update.message.reply_text("Errore! Usa /setadmin username")

async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = update.message.from_user.username
    if not is_admin(username):
        await update.message.reply_text("Solo admin possono vedere la lista admin.")
        return
    admins = list_admins()
    admins.extend(IMMUTABLE_ADMINS)
    await update.message.reply_text("Admin attuali:\n" + "\n".join(admins))

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
