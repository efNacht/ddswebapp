"""
FL Cosmetics — Financial Report Web App
Flask backend with Gemini AI integration.
"""

import os
import sys
import json
import time
import pickle
import traceback
from datetime import datetime, timezone, timedelta

MSK = timezone(timedelta(hours=3))

def now_msk():
    return datetime.now(MSK).strftime("%d.%m.%Y %H:%M МСК")

from flask import (
    Flask, render_template, request, jsonify, send_file,
    redirect, url_for, session
)
from werkzeug.utils import secure_filename

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max upload

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output", "state.pkl")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# In-memory state
app_state = {
    "transactions": [],
    "categorized": [],
    "dds_data": None,
    "pl_data": None,
    "tax_data": None,
    "reconciliation": None,
    "monthly_agg": None,
    "balances": None,
    "has_data": False,
    "has_sales": False,
    "has_dds_template": False,
    "has_pl_template": False,
    "processing": False,
    "generated_at": None,
    "source_filename": None,
}

def save_state():
    """Persist app_state to disk so it survives worker restarts."""
    try:
        saveable = {k: v for k, v in app_state.items() if k != "processing"}
        with open(STATE_FILE, "wb") as f:
            pickle.dump(saveable, f)
    except Exception as e:
        print(f"[STATE] Save failed: {e}", flush=True)

def load_state():
    """Restore app_state from disk on startup."""
    if not os.path.exists(STATE_FILE):
        return
    try:
        with open(STATE_FILE, "rb") as f:
            saved = pickle.load(f)
        app_state.update(saved)
        app_state["processing"] = False
        print(f"[STATE] Restored: {len(app_state['categorized'])} transactions, has_data={app_state['has_data']}", flush=True)
    except Exception as e:
        print(f"[STATE] Restore failed: {e}", flush=True)

# Restore state on startup
load_state()

# Claude API — key must be set via ANTHROPIC_API_KEY environment variable
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# Valid categories for dropdown
VALID_CATEGORIES = [
    "ОП", "ОС лидерам", "ЗП офис", "ЗП склад",
    "Стимулирование продаж", "Услуги банка", "Связь",
    "Бухгалтерские услуги", "Аренда склада", "Коммунальные платежи",
    "Расходные материалы", "Транспортные в регион", "Хозяйственные расходы",
    "Реклама каталог", "НДС", "Налог на сотрудников",
    "Налог на выплаты лидерам", "Услуги по таможенному оформлению",
    "Закупка товара (контейнер)", "Упаковка", "Прочие расходы",
    "Взнос в УК", "Возврат средств", "Внутренний перевод",
    "Налоги в порту", "Автопарк", "UNKNOWN",
]


def _claude_complete(prompt):
    """Call Claude Haiku with a prompt, return text response."""
    import anthropic
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=256,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text.strip()


def gemini_categorize(transactions):
    """Use Claude Haiku to categorize UNKNOWN transactions."""
    if not ANTHROPIC_API_KEY:
        return [{"error": "ANTHROPIC_API_KEY not set"}]
    try:
        examples = []
        for t in app_state["categorized"][:50]:
            cat = t.get("predicted_category", "")
            if cat and cat != "UNKNOWN":
                examples.append(f"  concept='{t.get('concept','')}' desc='{t.get('description','')}' amount={t.get('cargo',0) or t.get('abono',0)} → {cat}")

        categories_str = ", ".join(VALID_CATEGORIES[:-1])
        examples_str = "\n".join(examples[:30])

        results = []
        for txn in transactions:
            prompt = f"""Ты финансовый аналитик компании FL Cosmetics Mexico (Faberlic, Santander bank).
Определи категорию ДДС для транзакции. Допустимые категории: {categories_str}

Примеры:
{examples_str}

Транзакция:
  Concept: {txn.get('concept', '')}
  Description: {txn.get('description', '')}
  Long description: {txn.get('long_description', '')}
  Cargo (расход): {txn.get('cargo', 0)}
  Abono (приход): {txn.get('abono', 0)}
  Month: {txn.get('month', '')}

Ответь ТОЛЬКО JSON (без markdown): {{"category": "название", "confidence": 0.0-1.0, "reason": "кратко"}}"""

            text = _claude_complete(prompt)
            if text.startswith("```"):
                text = text.split("\n", 1)[1] if "\n" in text else text[3:]
                if text.endswith("```"):
                    text = text[:-3]
                text = text.strip()

            try:
                result = json.loads(text)
                results.append({
                    "index": txn.get("_index"),
                    "category": result.get("category", "UNKNOWN"),
                    "confidence": result.get("confidence", 0.5),
                    "reason": result.get("reason", ""),
                })
            except json.JSONDecodeError:
                results.append({
                    "index": txn.get("_index"),
                    "category": "UNKNOWN",
                    "confidence": 0,
                    "reason": f"parse error: {text[:80]}",
                })
        return results
    except Exception as e:
        return [{"error": str(e)}]


def gemini_anomalies():
    """Use Claude Haiku to detect anomalies in financial data."""
    if not ANTHROPIC_API_KEY or not app_state["monthly_agg"]:
        return []
    try:
        monthly = app_state["monthly_agg"]
        summary_lines = []
        for month, cats in sorted(monthly.items()):
            total_in = sum(d.get("abono", 0) for k, d in cats.items() if k != "_total" and isinstance(d, dict))
            total_out = sum(d.get("cargo", 0) for k, d in cats.items() if k != "_total" and isinstance(d, dict))
            top_cats = sorted(
                [(k, v["cargo"]) for k, v in cats.items() if k != "_total" and isinstance(v, dict) and v.get("cargo", 0) > 0],
                key=lambda x: -x[1]
            )[:4]
            top_str = ", ".join(f"{k}={v:,.0f}" for k, v in top_cats)
            summary_lines.append(f"{month}: доход={total_in:,.0f} расход={total_out:,.0f} | {top_str}")

        prompt = f"""Ты финансовый аналитик FL Cosmetics Mexico. Найди до 5 аномалий в данных.

{chr(10).join(summary_lines)}

Аномалии: резкие скачки, необычные суммы, пропуски. Ответь ТОЛЬКО JSON массивом (без markdown):
[{{"month":"...","type":"spike|drop|missing|unusual","category":"...","description":"кратко по-русски","severity":"high|medium|low"}}]"""

        text = _claude_complete(prompt)
        if text.startswith("```"):
            text = text.split("\n", 1)[1] if "\n" in text else text[3:]
            if text.endswith("```"):
                text = text[:-3]
        return json.loads(text.strip())
    except Exception as e:
        return [{"type": "error", "description": str(e), "severity": "low"}]


def run_pipeline(bank_path):
    """Run full pipeline: parse → categorize → generate all reports."""
    import logging
    log = logging.getLogger(__name__)
    log.setLevel(logging.DEBUG)

    print(f"[PIPELINE] Starting. bank_path={bank_path}", flush=True)
    print(f"[PIPELINE] sys.path={sys.path}", flush=True)
    print(f"[PIPELINE] OUTPUT_DIR={OUTPUT_DIR}", flush=True)
    print(f"[PIPELINE] os.path.exists(bank_path)={os.path.exists(bank_path)}", flush=True)

    import src.config as cfg
    import sys as _sys
    # Ensure 'config' and 'src.config' resolve to the same module instance
    _sys.modules['config'] = cfg
    cfg.BANK_STATEMENT_FILE = bank_path

    # Check for optional files
    payments_path = os.path.join(UPLOAD_DIR, "payments.xlsx")
    sales_path = os.path.join(UPLOAD_DIR, "sales_report.xlsx")
    dds_tpl_path = os.path.join(UPLOAD_DIR, "dds_template.xlsx")
    pl_tpl_path = os.path.join(UPLOAD_DIR, "pl_template.xlsx")

    if os.path.exists(payments_path):
        cfg.PAYMENTS_FILE = payments_path
    if os.path.exists(sales_path):
        cfg.SALES_REPORT_FILE = sales_path
        app_state["has_sales"] = True
    if os.path.exists(dds_tpl_path):
        cfg.DDS_TEMPLATE_FILE = dds_tpl_path
        app_state["has_dds_template"] = True
    if os.path.exists(pl_tpl_path):
        cfg.PL_TEMPLATE_FILE = pl_tpl_path
        app_state["has_pl_template"] = True

    cfg.OUTPUT_DIR = OUTPUT_DIR

    print("[PIPELINE] Importing parser...", flush=True)
    from parser import parse_bank_statement
    print("[PIPELINE] Importing categorizer...", flush=True)
    from categorizer import (
        build_dictionary_from_bank, build_dictionary_from_william,
        merge_dictionaries, categorize_all
    )
    print("[PIPELINE] Imports OK", flush=True)

    # Parse
    print("[PIPELINE] Parsing bank statement...", flush=True)
    txns = parse_bank_statement(bank_path)
    print(f"[PIPELINE] Parsed {len(txns)} transactions", flush=True)
    app_state["transactions"] = txns

    if not txns:
        raise ValueError("Не удалось разобрать выписку — 0 транзакций. Проверьте формат файла.")

    # Categorize
    print("[PIPELINE] Categorizing...", flush=True)
    bank_dict = build_dictionary_from_bank(txns)
    try:
        william_dict = build_dictionary_from_william(cfg.PAYMENTS_FILE)
    except Exception as e:
        print(f"[PIPELINE] William dict skipped: {e}", flush=True)
        william_dict = {"by_concept": {}, "by_description": {}, "by_counterparty": {}}
    merged = merge_dictionaries(bank_dict, william_dict)
    categorized = categorize_all(txns, merged)
    print(f"[PIPELINE] Categorized {len(categorized)} transactions", flush=True)

    # Add index for frontend reference
    for i, t in enumerate(categorized):
        t["_index"] = i

    app_state["categorized"] = categorized

    # AI categorize UNKNOWNs automatically — non-fatal if Gemini unavailable
    if ANTHROPIC_API_KEY:
        unknowns = [t for t in categorized if t.get("predicted_category") == "UNKNOWN"]
        print(f"[PIPELINE] AI categorizing {len(unknowns)} UNKNOWNs...", flush=True)
        if unknowns:
            try:
                ai_results = gemini_categorize(unknowns)
                for res in ai_results:
                    if "error" not in res and res.get("category") != "UNKNOWN":
                        idx = res["index"]
                        categorized[idx]["predicted_category"] = res["category"]
                        categorized[idx]["confidence"] = res.get("confidence", 0.5)
                        categorized[idx]["ai_reason"] = res.get("reason", "")
                        categorized[idx]["ai_categorized"] = True
                ai_done = sum(1 for t in categorized if t.get("ai_categorized"))
                print(f"[PIPELINE] AI categorized {ai_done} transactions", flush=True)
            except Exception as e:
                print(f"[PIPELINE] AI categorization failed (non-fatal): {e}", flush=True)
    else:
        print("[PIPELINE] ANTHROPIC_API_KEY not set — skipping AI categorization", flush=True)

    # Generate reports
    print("[PIPELINE] Generating reports...", flush=True)
    _generate_reports(categorized)
    print("[PIPELINE] Reports done!", flush=True)

    app_state["has_data"] = True
    app_state["generated_at"] = now_msk()
    save_state()


def _generate_reports(categorized):
    """Generate all reports from categorized transactions."""
    import src.config as cfg

    print("[REPORTS] Importing generators...", flush=True)
    from dds_generator import generate_dds_data, write_dds_excel, aggregate_by_month
    from tax_summary import generate_tax_summary, write_tax_excel

    # Aggregate
    print("[REPORTS] Aggregating by month...", flush=True)
    monthly_agg, balances = aggregate_by_month(categorized)
    app_state["monthly_agg"] = dict(monthly_agg)
    app_state["balances"] = dict(balances)
    print(f"[REPORTS] {len(monthly_agg)} months aggregated", flush=True)

    # DDS
    print("[REPORTS] Generating DDS...", flush=True)
    dds_data = generate_dds_data(categorized)
    app_state["dds_data"] = dds_data
    dds_path = os.path.join(OUTPUT_DIR, "ДДС_факт.xlsx")
    write_dds_excel(dds_data, dds_path)
    print(f"[REPORTS] DDS written: {os.path.exists(dds_path)}", flush=True)

    # DDS template fill
    dds_tpl = os.path.join(UPLOAD_DIR, "dds_template.xlsx")
    print(f"[REPORTS] has_dds_template={app_state['has_dds_template']}, dds_tpl exists={os.path.exists(dds_tpl)}", flush=True)
    if app_state["has_dds_template"] or os.path.exists(dds_tpl):
        try:
            from template_filler import fill_dds_template
            fill_dds_template(dds_data, os.path.join(OUTPUT_DIR, "ДДС_план_факт.xlsx"), template_path=dds_tpl)
            app_state["has_dds_template"] = True
            print("[REPORTS] DDS template filled", flush=True)
        except Exception as e:
            print(f"[REPORTS] DDS template fill error: {e}", flush=True)
            traceback.print_exc()

    # PL
    print("[REPORTS] Generating PL...", flush=True)
    try:
        from pl_generator import generate_pl_data, write_pl_excel, parse_sales_report
        sales = None
        if app_state["has_sales"]:
            sales = parse_sales_report(cfg.SALES_REPORT_FILE)
        pl_data = generate_pl_data(categorized, sales, cfg.USD_MXN_RATES)
        app_state["pl_data"] = pl_data
        pl_path = os.path.join(OUTPUT_DIR, "PL_факт.xlsx")
        write_pl_excel(pl_data, pl_path)
        print(f"[REPORTS] PL written: {os.path.exists(pl_path)}", flush=True)
    except Exception as e:
        print(f"[REPORTS] PL generation error: {e}", flush=True)
        traceback.print_exc()
        app_state["pl_data"] = None

    # PL template fill
    pl_tpl = os.path.join(UPLOAD_DIR, "pl_template.xlsx")
    print(f"[REPORTS] has_pl_template={app_state['has_pl_template']}, pl_tpl exists={os.path.exists(pl_tpl)}", flush=True)
    if app_state["has_pl_template"] or os.path.exists(pl_tpl):
        try:
            from template_filler import fill_pl_template
            fill_pl_template(categorized, os.path.join(OUTPUT_DIR, "PL_план_факт.xlsx"), template_path=pl_tpl)
            app_state["has_pl_template"] = True
            print("[REPORTS] PL template filled", flush=True)
        except Exception as e:
            print(f"[REPORTS] PL template fill error: {e}", flush=True)
            traceback.print_exc()

    # Reconciliation
    print("[REPORTS] Reconciliation...", flush=True)
    try:
        from reconciliation import parse_provider_data, reconcile, write_reconciliation_excel
        if app_state["has_sales"]:
            providers, totals = parse_provider_data(cfg.SALES_REPORT_FILE)
            if providers:
                report = reconcile(providers, totals, categorized)
                app_state["reconciliation"] = report
                write_reconciliation_excel(report, os.path.join(OUTPUT_DIR, "Сверка_провайдеров.xlsx"))
                print("[REPORTS] Reconciliation written", flush=True)
    except Exception as e:
        print(f"[REPORTS] Reconciliation error: {e}", flush=True)
        traceback.print_exc()

    # Tax
    print("[REPORTS] Generating tax summary...", flush=True)
    tax_data = generate_tax_summary(categorized)
    app_state["tax_data"] = tax_data
    tax_path = os.path.join(OUTPUT_DIR, "Налоговая_сводка.xlsx")
    write_tax_excel(tax_data, tax_path)
    print(f"[REPORTS] Tax written: {os.path.exists(tax_path)}", flush=True)


# ─────────────────────────────────────────────────
# ROUTES — Pages
# ─────────────────────────────────────────────────

@app.route("/")
def index():
    # Always show upload page — user can load a new file anytime
    return render_template("upload.html", has_data=app_state["has_data"],
                           generated_at=app_state.get("generated_at"),
                           source_filename=app_state.get("source_filename"))


@app.route("/dashboard")
def dashboard():
    if not app_state["has_data"]:
        return redirect(url_for("index"))
    return render_template("dashboard.html")


@app.route("/transactions")
def transactions():
    if not app_state["has_data"]:
        return redirect(url_for("index"))
    return render_template("transactions.html", categories=VALID_CATEGORIES)


@app.route("/reports")
def reports():
    if not app_state["has_data"]:
        return redirect(url_for("index"))
    return render_template("reports.html")


# ─────────────────────────────────────────────────
# API — Upload & Process
# ─────────────────────────────────────────────────

@app.route("/api/upload", methods=["POST"])
def api_upload():
    if "bank_statement" not in request.files:
        return jsonify({"error": "No bank statement file"}), 400

    bank_file = request.files["bank_statement"]
    if not bank_file.filename:
        return jsonify({"error": "Empty filename"}), 400

    original_filename = bank_file.filename
    bank_path = os.path.join(UPLOAD_DIR, "bank_statement.xlsx")
    bank_file.save(bank_path)
    app_state["source_filename"] = original_filename

    # Save optional files
    for key, name in [("sales_report", "sales_report.xlsx"),
                       ("payments", "payments.xlsx"),
                       ("dds_template", "dds_template.xlsx"),
                       ("pl_template", "pl_template.xlsx")]:
        if key in request.files and request.files[key].filename:
            request.files[key].save(os.path.join(UPLOAD_DIR, name))

    try:
        app_state["processing"] = True
        start = time.time()
        run_pipeline(bank_path)
        elapsed = time.time() - start
        app_state["processing"] = False

        categorized = app_state["categorized"]
        total = len(categorized)
        unknown = sum(1 for t in categorized if t.get("predicted_category") == "UNKNOWN")
        ai_done = sum(1 for t in categorized if t.get("ai_categorized"))

        # Category breakdown — top 10 by transaction count
        cat_counts = {}
        for t in categorized:
            c = t.get("predicted_category", "UNKNOWN")
            cat_counts[c] = cat_counts.get(c, 0) + 1
        top_cats = sorted(cat_counts.items(), key=lambda x: -x[1])[:10]

        # Months covered
        months = sorted(set(t.get("month", "") for t in categorized if t.get("month")))

        # Generated files with sizes
        report_files = [
            ("ДДС_факт.xlsx", "ДДС — движение средств"),
            ("PL_факт.xlsx", "P&L — прибыли и убытки"),
            ("ДДС_план_факт.xlsx", "ДДС — шаблон компании"),
            ("PL_план_факт.xlsx", "P&L — шаблон компании"),
            ("Налоговая_сводка.xlsx", "Налоговый отчёт"),
            ("Сверка_провайдеров.xlsx", "Сверка платёжных систем"),
        ]
        generated_files = []
        for fname, title in report_files:
            fpath = os.path.join(OUTPUT_DIR, fname)
            if os.path.exists(fpath):
                sz = os.path.getsize(fpath)
                generated_files.append({"name": title, "filename": fname, "size_kb": round(sz / 1024, 1)})

        return jsonify({
            "success": True,
            "source_file": original_filename,
            "generated_at": app_state["generated_at"],
            "transactions": total,
            "unknown": unknown,
            "ai_categorized": ai_done,
            "elapsed": round(elapsed, 1),
            "months": months,
            "category_breakdown": [{"category": c, "count": n} for c, n in top_cats],
            "generated_files": generated_files,
            "redirect": "/dashboard",
        })
    except Exception as e:
        app_state["processing"] = False
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────────
# API — Dashboard data
# ─────────────────────────────────────────────────

@app.route("/api/dashboard")
def api_dashboard():
    if not app_state["has_data"]:
        return jsonify({"error": "No data"}), 404

    categorized = app_state["categorized"]
    monthly = app_state["monthly_agg"]

    from dds_generator import MONTH_ORDER
    sorted_months = sorted(monthly.keys(), key=lambda m: MONTH_ORDER.get(m, 99))

    # KPIs
    total_income = sum(t.get("abono", 0) or 0 for t in categorized)
    total_expense = sum(t.get("cargo", 0) or 0 for t in categorized)
    net_flow = total_income - total_expense

    # Monthly data for charts
    months_data = []
    for month in sorted_months:
        cats = monthly[month]
        # Exclude _total key — it's a pre-aggregated sum of all categories, would double-count
        income = sum(d.get("abono", 0) for k, d in cats.items() if k != "_total")
        expense = sum(d.get("cargo", 0) for k, d in cats.items() if k != "_total")
        months_data.append({
            "month": month,
            "income": round(income, 2),
            "expense": round(expense, 2),
            "net": round(income - expense, 2),
        })

    # Category breakdown
    cat_totals = {}
    for t in categorized:
        cat = t.get("predicted_category", "UNKNOWN")
        cargo = t.get("cargo", 0) or 0
        if cargo > 0:
            cat_totals[cat] = cat_totals.get(cat, 0) + cargo

    cat_breakdown = sorted(cat_totals.items(), key=lambda x: -x[1])

    # DDS plan vs fact (if template was loaded)
    plan_fact = None
    if app_state["dds_data"]:
        plan_fact = []
        for month in sorted_months:
            md = app_state["dds_data"].get(month, {})
            plan_fact.append({
                "month": month,
                "fact_income": round(md.get("ОП", 0), 2),
                "fact_expense": round(sum(
                    md.get(k, 0) for k in md if k not in ("ОП", "total_abono", "opening_balance", "closing_balance", "Взнос в УК", "Возврат средств", "_кв_бонус")
                ), 2),
            })

    return jsonify({
        "kpis": {
            "total_income": round(total_income, 2),
            "total_expense": round(total_expense, 2),
            "net_flow": round(net_flow, 2),
            "transaction_count": len(categorized),
            "months_count": len(sorted_months),
        },
        "monthly": months_data,
        "category_breakdown": [{"category": c, "amount": round(a, 2)} for c, a in cat_breakdown[:15]],
        "plan_fact": plan_fact,
    })


# ─────────────────────────────────────────────────
# API — Transactions
# ─────────────────────────────────────────────────

@app.route("/api/transactions")
def api_transactions():
    if not app_state["has_data"]:
        return jsonify({"error": "No data"}), 404

    categorized = app_state["categorized"]

    # Filters
    month_filter = request.args.get("month")
    cat_filter = request.args.get("category")
    search = request.args.get("search", "").lower()
    page = int(request.args.get("page", 1))
    per_page = int(request.args.get("per_page", 100))

    filtered = categorized
    if month_filter:
        filtered = [t for t in filtered if t.get("month") == month_filter]
    if cat_filter:
        filtered = [t for t in filtered if t.get("predicted_category") == cat_filter]
    if search:
        filtered = [t for t in filtered if
                    search in (t.get("concept", "") or "").lower() or
                    search in (t.get("description", "") or "").lower() or
                    search in (t.get("long_description", "") or "").lower()]

    total = len(filtered)
    start = (page - 1) * per_page
    page_data = filtered[start:start + per_page]

    # Serialize for JSON
    result = []
    for t in page_data:
        result.append({
            "index": t.get("_index"),
            "month": t.get("month", ""),
            "date": t.get("date", ""),
            "description": t.get("description", ""),
            "concept": t.get("concept", ""),
            "long_description": t.get("long_description", ""),
            "cargo": t.get("cargo", 0) or 0,
            "abono": t.get("abono", 0) or 0,
            "balance": t.get("balance", 0) or 0,
            "category": t.get("predicted_category", "UNKNOWN"),
            "confidence": t.get("confidence", 0),
            "ai_categorized": t.get("ai_categorized", False),
            "ai_reason": t.get("ai_reason", ""),
        })

    # Available months for filter
    all_months = sorted(set(t.get("month", "") for t in categorized))

    return jsonify({
        "transactions": result,
        "total": total,
        "page": page,
        "per_page": per_page,
        "pages": (total + per_page - 1) // per_page,
        "months": all_months,
    })


@app.route("/api/transactions/<int:idx>", methods=["PUT"])
def api_update_transaction(idx):
    if not app_state["has_data"]:
        return jsonify({"error": "No data"}), 404

    data = request.get_json()
    new_cat = data.get("category")

    if new_cat not in VALID_CATEGORIES:
        return jsonify({"error": f"Invalid category: {new_cat}"}), 400

    if 0 <= idx < len(app_state["categorized"]):
        app_state["categorized"][idx]["predicted_category"] = new_cat
        app_state["categorized"][idx]["confidence"] = 1.0
        app_state["categorized"][idx]["manually_edited"] = True
        return jsonify({"success": True, "index": idx, "category": new_cat})

    return jsonify({"error": "Index out of range"}), 404


# ─────────────────────────────────────────────────
# API — Recalculate
# ─────────────────────────────────────────────────

@app.route("/api/recalculate", methods=["POST"])
def api_recalculate():
    if not app_state["has_data"]:
        return jsonify({"error": "No data"}), 404

    try:
        _generate_reports(app_state["categorized"])
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────────
# API — AI
# ─────────────────────────────────────────────────

@app.route("/api/ai/categorize", methods=["POST"])
def api_ai_categorize():
    """AI-categorize UNKNOWN transactions."""
    unknowns = [t for t in app_state["categorized"] if t.get("predicted_category") == "UNKNOWN"]
    if not unknowns:
        return jsonify({"message": "No UNKNOWN transactions", "results": []})

    results = gemini_categorize(unknowns)

    # Apply results
    applied = 0
    for res in results:
        if "error" not in res and res.get("category") != "UNKNOWN":
            idx = res["index"]
            app_state["categorized"][idx]["predicted_category"] = res["category"]
            app_state["categorized"][idx]["confidence"] = res.get("confidence", 0.5)
            app_state["categorized"][idx]["ai_reason"] = res.get("reason", "")
            app_state["categorized"][idx]["ai_categorized"] = True
            applied += 1

    return jsonify({"applied": applied, "total": len(unknowns), "results": results})


@app.route("/api/ai/suggest", methods=["POST"])
def api_ai_suggest():
    """Get AI suggestion for a single transaction."""
    data = request.get_json()
    idx = data.get("index")

    if idx is None or idx < 0 or idx >= len(app_state["categorized"]):
        return jsonify({"error": "Invalid index"}), 400

    txn = app_state["categorized"][idx]
    results = gemini_categorize([txn])

    if results and "error" not in results[0]:
        return jsonify(results[0])
    return jsonify({"category": "UNKNOWN", "confidence": 0, "reason": "AI unavailable"})


@app.route("/api/ai/anomalies")
def api_ai_anomalies():
    """Detect anomalies in financial data using Gemini."""
    anomalies = gemini_anomalies()
    return jsonify({"anomalies": anomalies})


# ─────────────────────────────────────────────────
# API — Report Downloads
# ─────────────────────────────────────────────────

@app.route("/api/reports/<report_type>")
def api_download_report(report_type):
    report_files = {
        "dds": "ДДС_факт.xlsx",
        "pl": "PL_факт.xlsx",
        "dds-template": "ДДС_план_факт.xlsx",
        "pl-template": "PL_план_факт.xlsx",
        "reconciliation": "Сверка_провайдеров.xlsx",
        "tax": "Налоговая_сводка.xlsx",
    }

    filename = report_files.get(report_type)
    if not filename:
        return jsonify({"error": "Unknown report type"}), 404

    filepath = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        return jsonify({"error": f"Report not generated: {filename}"}), 404

    return send_file(filepath, as_attachment=True, download_name=filename)


@app.route("/api/reports/list")
def api_reports_list():
    """List available reports."""
    reports = []
    report_info = [
        ("dds", "ДДС_факт.xlsx", "ДДС — движение средств"),
        ("pl", "PL_факт.xlsx", "P&L — прибыли и убытки"),
        ("dds-template", "ДДС_план_факт.xlsx", "ДДС — шаблон компании (план/факт)"),
        ("pl-template", "PL_план_факт.xlsx", "P&L — шаблон компании (план/факт)"),
        ("reconciliation", "Сверка_провайдеров.xlsx", "Сверка платёжных систем"),
        ("tax", "Налоговая_сводка.xlsx", "Налоговый отчёт"),
    ]

    for key, filename, title in report_info:
        filepath = os.path.join(OUTPUT_DIR, filename)
        exists = os.path.exists(filepath)
        size = os.path.getsize(filepath) if exists else 0
        mtime = os.path.getmtime(filepath) if exists else None
        generated_at = datetime.fromtimestamp(mtime, tz=MSK).strftime("%d.%m.%Y %H:%M МСК") if mtime else None
        reports.append({
            "key": key,
            "filename": filename,
            "title": title,
            "available": exists,
            "size_kb": round(size / 1024, 1),
            "generated_at": generated_at,
        })

    return jsonify({"reports": reports})


# ─────────────────────────────────────────────────
# API — Debug / Health
# ─────────────────────────────────────────────────

@app.route("/api/dashboard/matrix")
def api_dashboard_matrix():
    """Category × month expense matrix + top transactions."""
    if not app_state["has_data"]:
        return jsonify({"error": "No data"}), 404

    from dds_generator import MONTH_ORDER
    monthly = app_state["monthly_agg"]
    categorized = app_state["categorized"]

    sorted_months = sorted(monthly.keys(), key=lambda m: MONTH_ORDER.get(m, 99))

    # Build expense matrix: category → {month → amount}
    all_cats = set()
    for cats in monthly.values():
        all_cats.update(k for k in cats if not k.startswith("_"))

    # Exclude income/transfer categories
    skip = {"ОП", "Взнос в УК", "Возврат средств", "Внутренний перевод"}
    expense_cats = sorted(all_cats - skip)

    matrix = {}
    cat_totals = {}
    for cat in expense_cats:
        row = {}
        total = 0
        for month in sorted_months:
            val = monthly[month].get(cat, {})
            amount = val.get("cargo", 0) if isinstance(val, dict) else 0
            row[month] = round(amount, 0)
            total += amount
        if total > 0:
            matrix[cat] = row
            cat_totals[cat] = total

    # Sort by total desc
    sorted_cats = sorted(matrix.keys(), key=lambda c: -cat_totals[c])

    # Month totals
    month_totals = {}
    for month in sorted_months:
        month_totals[month] = round(sum(
            monthly[month].get(cat, {}).get("cargo", 0) if isinstance(monthly[month].get(cat, {}), dict) else 0
            for cat in sorted_cats
        ), 0)

    # Top 5 single transactions by amount
    top_txns = sorted(
        [t for t in categorized if (t.get("cargo") or 0) > 0],
        key=lambda t: -(t.get("cargo") or 0)
    )[:5]

    # Monthly income
    monthly_income = {m: round(sum(
        d.get("abono", 0) if isinstance(d, dict) else 0
        for k, d in monthly[m].items() if not k.startswith("_")
    ), 0) for m in sorted_months}

    return jsonify({
        "months": sorted_months,
        "categories": sorted_cats,
        "matrix": matrix,
        "cat_totals": {c: round(cat_totals[c], 0) for c in sorted_cats},
        "month_totals": month_totals,
        "monthly_income": monthly_income,
        "top_transactions": [{
            "date": t.get("date", ""),
            "description": t.get("description", "")[:60],
            "concept": t.get("concept", "")[:40],
            "category": t.get("predicted_category", ""),
            "amount": t.get("cargo", 0),
            "month": t.get("month", ""),
        } for t in top_txns],
    })


@app.route("/api/debug")
def api_debug():
    """Health check + environment diagnostics."""
    import importlib

    src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src")
    modules_to_check = ["parser", "categorizer", "dds_generator", "pl_generator",
                        "reconciliation", "tax_summary", "template_filler"]

    module_status = {}
    for mod in modules_to_check:
        try:
            importlib.import_module(mod)
            module_status[mod] = "OK"
        except Exception as e:
            module_status[mod] = f"ERROR: {e}"

    upload_files = os.listdir(UPLOAD_DIR) if os.path.exists(UPLOAD_DIR) else []
    output_files = os.listdir(OUTPUT_DIR) if os.path.exists(OUTPUT_DIR) else []

    return jsonify({
        "status": "running",
        "python_version": sys.version,
        "gemini_key_set": bool(GEMINI_API_KEY),
        "src_dir_exists": os.path.exists(src_dir),
        "src_dir": src_dir,
        "sys_path": sys.path[:5],
        "upload_dir": UPLOAD_DIR,
        "output_dir": OUTPUT_DIR,
        "upload_files": upload_files,
        "output_files": output_files,
        "modules": module_status,
        "app_state_summary": {
            "has_data": app_state["has_data"],
            "transactions_count": len(app_state["categorized"]),
            "processing": app_state["processing"],
        }
    })


# ─────────────────────────────────────────────────
# API — Reset
# ─────────────────────────────────────────────────

@app.route("/api/reset", methods=["POST"])
def api_reset():
    """Reset all data — start fresh."""
    app_state.update({
        "transactions": [],
        "categorized": [],
        "dds_data": None,
        "pl_data": None,
        "tax_data": None,
        "reconciliation": None,
        "monthly_agg": None,
        "balances": None,
        "has_data": False,
        "has_sales": False,
        "has_dds_template": False,
        "has_pl_template": False,
        "processing": False,
        "generated_at": None,
        "source_filename": None,
    })
    # Clean upload/output dirs
    for d in [UPLOAD_DIR, OUTPUT_DIR]:
        for f in os.listdir(d):
            fp = os.path.join(d, f)
            if os.path.isfile(fp):
                os.remove(fp)
    return jsonify({"success": True})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
