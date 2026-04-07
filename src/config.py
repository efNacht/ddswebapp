import os

# Webapp root = parent of src/
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(PROJECT_ROOT, "uploads")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")

# These will be set dynamically after upload
BANK_STATEMENT_FILE = os.path.join(DATA_DIR, "bank_statement.xlsx")
PAYMENTS_FILE = os.path.join(DATA_DIR, "payments.xlsx")
SALES_REPORT_FILE = os.path.join(DATA_DIR, "sales_report.xlsx")
DDS_TEMPLATE_FILE = os.path.join(DATA_DIR, "dds_template.xlsx")
PL_TEMPLATE_FILE = os.path.join(DATA_DIR, "pl_template.xlsx")

# Sheet to skip in bank statement
BANK_SKIP_SHEETS = ["Table 1"]

# Column indices for bank statement (0-based)
BANK_COLUMNS = {
    "Cuenta": 0,
    "Fecha": 1,
    "Hora": 2,
    "Sucursal": 3,
    "Descripción": 4,
    "Importe Cargo": 5,
    "Importe Abono": 6,
    "Saldo": 7,
    "Referencia": 8,
    "Concepto": 9,
    "Descripción Larga": 10,
    "Комментарий": 11,
    "Статья ДДС": 12,
}

# USD/MXN monthly average exchange rates
USD_MXN_RATES = {
    "Февраль 2025": 20.489,
    "Март 2025": 20.242,
    "Апрель 2025": 20.037,
    "Май 2025": 19.442,
    "Июнь 2025": 19.056,
    "Июль 2025": 18.672,
    "Август 2025": 18.703,
    "Сентябрь 2025": 18.499,
    "Октябрь 2025": 18.432,
    "Ноябрь 2025": 18.424,
    "Декабрь 2025": 18.068,
    "Январь 2026": 17.686,
    "Февраль 2026": 17.239,
    "Март 2026": 17.791,
}

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
