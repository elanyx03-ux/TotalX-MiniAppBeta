
import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import Workbook, load_workbook
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

# Carica variabili d'ambiente
load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# Nome del file principale per gli admin
FILE_ADMIN = "estratto_conto_admin.xlsx"

# Admin fissi
FIXED_ADMINS = ["@Elanyx03"]

# Inizializza file admin se non esiste
if os.path.exists(FILE_ADMIN):
    wb_admin = load_workbook(FILE_ADMIN)
    ws_admin = wb_admin.active
else:
    wb_admin = Workbook()
    ws_admin = wb_admin.active
    ws_admin.append(["user", "movimento", "data"])
    wb_admin.save(FILE_ADMIN)

# Dizionario dei file utente normali: user_id -> filename
user_files = {}

# Funzione per ottenere il file Excel di un utente (admin o utente normale)
def get_user_file(username, user_id=None):
    if username in FIXED_ADMINS:
        return FILE_ADMIN, True  # admin
    else:
        # File separato per ogni utente non admin
        filename = f"estratto_conto_{username}.xlsx"
        if not os.path.exists(filename):
            wb = Workbook()
            ws = wb.active
            ws.append(["user", "movimento", "data"])
            wb.save(filename)
        return filename, False

# Funzione per salvare un movimento
def salva_movimento(username, user_id, valore):
    filename, is_admin = get_user_file(username, user_id)
    wb = load_workbook(filename)
    ws = wb.active
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Usa Decimal per evitare floating point issues
    valore = Decimal(str(valore)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    ws.append([username, float(valore), now])
    wb.save(filename)

# Legge tutti i movimenti di un file
def leggi_movimenti(username):
    filename, _ = get_user_file(username)
    wb = load_workbook(filename)
    ws = wb.active
    movimenti = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        movimenti.append({
            "user": row[0],
            "movimento": Decimal(str(row[1])).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP),
            "data": row[2]
        })
    return movimenti

# Calcola saldo, entrate e uscite
def estratto_conto(username):
    movimenti = leggi_movimenti(username)
    entrate = [m for m in movimenti if m["movimento"] > 0]
    uscite = [m for m in movimenti if m["movimento"] < 0]
    totale_entrate = sum([m["movimento"] for m in entrate])
    totale_uscite = sum([m["movimento"] for m in uscite])
    saldo = totale_entrate + totale_uscite
    return entrate, uscite, totale_entrate, totale_uscite, saldo

# Crea file Excel esportabile
def crea_file_excel(username):
    movimenti = leggi_movimenti(username)
    wb_user = Workbook()
    ws_user = wb_user.active
    ws_user.title = "Estratto Conto"
    ws_user.append(["Tipo", "Importo", "Utente", "Data"])
    for m in movimenti:
        tipo = "Entrata" if m["movimento"] > 0 else "Uscita"
        ws_user.append([tipo, float(m["movimento"]), m["user"], m["data"]])
    entrate, uscite, totale_entrate, totale_uscite, saldo = estratto_conto(username)
    ws_user.append([])
    ws_user.append(["Totale Entrate", float(totale_entrate)])
    ws_user.append(["Totale Uscite", float(totale_uscite)])
    ws_user.append(["Saldo Finale", float(saldo)])
    filename = f"estratto_conto_export_{username}.xlsx"
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
        "/export - ricevi un file Excel\n"
        "/undo - annulla l'ultima operazione\n"
        "/reset - azzera tutto e crea un nuovo foglio\n"
        "/setadmin username - aggiunge/rimuove un admin (solo admin)\n"
        "/adminlist - mostra la lista admin (solo admin)"
    )

async def add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        valore = context.args[0].replace(",", ".")
        valore = Decimal(valore)
        username = "@" + update.message.from_user.username
        user_id = update.message.from_user.id
        salva_movimento(username, user_id, valore)
        _, _, _, _, saldo = estratto_conto(username)
        await update.message.reply_text(f"Entrata registrata: +{valore}\nSaldo attuale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /add numero, esempio /add 100 o /add 0,05")

async def subtract(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        valore = context.args[0].replace(",", ".")
        valore = Decimal(valore)
        username = "@" + update.message.from_user.username
        user_id = update.message.from_user.id
        salva_movimento(username, user_id, -valore)
        _, _, _, _, saldo = estratto_conto(username)
        await update.message.reply_text(f"Uscita registrata: -{valore}\nSaldo attuale: {saldo}")
    except (IndexError, ValueError):
        await update.message.reply_text("Errore! Usa /subtract numero, esempio /subtract 50 o /subtract 0,07")

async def total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = "@" + update.message.from_user.username
    _, _, _, _, saldo = estratto_conto(username)
    await update.message.reply_text(f"Saldo totale: {saldo}")

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = "@" + update.message.from_user.username
    entrate, uscite, totale_entrate, totale_uscite, saldo = estratto_conto(username)
    if not entrate and not uscite:
        await update.message.reply_text("Nessun movimento registrato.")
        return
    report_text = "📄 Estratto Conto\n\n"
    if entrate:
        report_text += "Entrate:\n" + "\n".join([f"+{m['movimento']} ({m['user']} {m['data']})" for m in entrate]) + f"\nTotale Entrate: {totale_entrate}\n\n"
    if uscite:
        report_text += "Uscite:\n" + "\n".join([f"{m['movimento']} ({m['user']} {m['data']})" for m in uscite]) + f"\nTotale Uscite: {totale_uscite}\n\n"
    report_text += f"Saldo Totale: {saldo}"
    await update.message.reply_text(report_text)

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = "@" + update.message.from_user.username
    filename = crea_file_excel(username)
    with open(filename, "rb") as file:
        await update.message.reply_document(file, filename=filename)

async def undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = "@" + update.message.from_user.username
    filename, _ = get_user_file(username)
    wb = load_workbook(filename)
    ws = wb.active
    if ws.max_row > 1:
        ws.delete_rows(ws.max_row)
        wb.save(filename)
        await update.message.reply_text("Ultima operazione annullata.")
    else:
        await update.message.reply_text("Nessuna operazione da annullare.")

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = "@" + update.message.from_user.username
    filename, is_admin = get_user_file(username)
    # Ricrea il file vuoto
    wb = Workbook()
    ws = wb.active
    ws.append(["user", "movimento", "data"])
    wb.save(filename)
    await update.message.reply_text("Foglio azzerato e ricreato con saldo 0.")

async def setadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = "@" + update.message.from_user.username
    if username not in FIXED_ADMINS:
        await update.message.reply_text("Solo admin possono modificare la lista admin.")
        return
    try:
        target = context.args[0]
        if target in FIXED_ADMINS:
            await update.message.reply_text(f"{target} è un admin fisso e non può essere rimosso.")
            return
        # Qui si potrebbe aggiungere logica per aggiungere/rimuovere admin in un file separato
        await update.message.reply_text(f"Operazione admin su {target} eseguita (simulata).")
    except IndexError:
        await update.message.reply_text("Errore! Usa /setadmin @username")

async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    username = "@" + update.message.from_user.username
    if username not in FIXED_ADMINS:
        await update.message.reply_text("Solo admin possono vedere la lista admin.")
        return
    await update.message.reply_text("Admin attuali:\n" + "\n".join(FIXED_ADMINS))

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

