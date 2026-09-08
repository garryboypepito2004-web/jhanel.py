FULL_DAY_RATES = {
    "Labor": 500.0,
    "Skill": 650.0,
    "Forman": 800.0,
}

TIER_TABLE = FULL_DAY_RATES.copy()


def get_partial_rate(days, role):
    rate = FULL_DAY_RATES.get(role, FULL_DAY_RATES["Labor"])
    if days <= 0:
        return 0.0
    if days >= 1:
        return rate
    return rate * float(days)


def calculate_labor_pay(days, role):
    rate = FULL_DAY_RATES.get(role, FULL_DAY_RATES["Labor"])
    days = float(days or 0.0)
    if days <= 0:
        return 0.0, 0.0, 0.0
    gross_pay = rate * days
    full_pay = rate * max(1.0, days)
    partial_pay = get_partial_rate(days, role)
    return gross_pay, full_pay, partial_pay


def budget_alert_status(budget, used):
    """Return a simple budget alert summary for the dashboard."""
    budget = float(budget or 0.0)
    used = float(used or 0.0)
    remaining = budget - used
    if budget <= 0:
        return {"severity": "info", "message": "No budget is currently set for this project."}
    ratio = used / budget if budget else 0.0
    if ratio >= 0.9:
        return {"severity": "danger", "message": f"Budget is nearly exhausted. {remaining:,.2f} remaining."}
    if ratio >= 0.75:
        return {"severity": "warning", "message": f"Budget is approaching the warning threshold. {remaining:,.2f} remaining."}
    return {"severity": "info", "message": f"Budget is healthy. {remaining:,.2f} remaining."}


def monthly_trend_summary(records, labor_records, payroll_expenses, months=6):
    """Summarize recent monthly spend by category for the dashboard."""
    from collections import defaultdict
    from datetime import datetime

    def month_key(record):
        value = record.get("date") or record.get("month") or record.get("recorded_at")
        if not value:
            return None
        try:
            if "T" in str(value):
                return datetime.fromisoformat(str(value)).strftime("%Y-%m")
            return datetime.strptime(str(value), "%b %d, %Y").strftime("%Y-%m")
        except ValueError:
            try:
                return datetime.fromisoformat(str(value)).strftime("%Y-%m")
            except ValueError:
                return None

    totals = defaultdict(lambda: {"Materials": 0.0, "Construction": 0.0, "Labor": 0.0, "Payroll": 0.0})

    for record in records or []:
        key = month_key(record)
        if not key:
            continue
        amount = float(record.get("amount", 0) or 0.0)
        kind = record.get("type")
        if kind == "material":
            totals[key]["Materials"] += amount
        elif kind == "expense":
            totals[key]["Construction"] += amount
        elif kind == "excess":
            totals[key]["Construction"] += amount

    for record in labor_records or []:
        key = month_key(record)
        if key:
            totals[key]["Labor"] += float(record.get("net", 0) or 0.0)

    for record in payroll_expenses or []:
        key = month_key(record)
        if key:
            totals[key]["Payroll"] += float(record.get("price", 0) or 0.0)

    current = datetime.now()
    rows = []
    for offset in range(months - 1, -1, -1):
        target = current.replace(day=1)
        for _ in range(offset):
            if target.month == 1:
                target = target.replace(year=target.year - 1, month=12)
            else:
                target = target.replace(month=target.month - 1)
        month_name = target.strftime("%Y-%m")
        data = totals.get(month_name, {"Materials": 0.0, "Construction": 0.0, "Labor": 0.0, "Payroll": 0.0})
        row = {
            "Month": month_name,
            "Materials": float(data["Materials"]),
            "Construction": float(data["Construction"]),
            "Labor": float(data["Labor"]),
            "Payroll": float(data["Payroll"]),
        }
        row["Total"] = row["Materials"] + row["Construction"] + row["Labor"] + row["Payroll"]
        rows.append(row)
    return rows
