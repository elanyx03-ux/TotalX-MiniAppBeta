import os
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import Workbook, load_workbook

load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# File principale admin
ADMIN_FILE = "estratto_conto_admin.xlsx"

# Admin fissi
FIXED_ADMINS = ["@Ela036", "@NyX0369"]
admins = FIXED_ADMINS.copy()

def round_decimal(value):
    return float(Decimal(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def load_or_create_file(filename):
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["User", "Movimento", "Data"])
        wb.save(filename)
    return wb, ws

def salva_movimento(username, valore, admin_mode=False):
    filename = ADMIN_FILE if admin_mode else f"Movimenti_{username}.xlsx"
    wb, ws = load_or_create_file(filename)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([username, round_decimal(valore), now])
    wb.save(filename)

def leggi_movimenti(username, admin_mode=False):
    filename = ADMIN_FILE if admin_mode else f"Movimenti_{username}.xlsx"
    wb, ws = load_or_create_file(filename)
    movimenti = [(row[0], row[1], row[2]) for row in ws.iter_rows(min_row=2, values_only=True)]
    return movimenti

def estratto_conto(username, admin_mode=False):
    movimenti = leggi_movimenti(username, admin_mode)
    totale_entrate = sum(m[1] for m in movimenti if m[1] > 0)
    totale_uscite = sum(m[1] for m in movimenti if m[1] < 0)
    saldo = totale_entrate + totale_uscite
    return movimenti, totale_entrate, totale_uscite, saldo

# Comandi
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Ciao! Sono TotalX Pro.
