def close_month(closes, month, summary=None, closed_at=None, note="", closed_by="System"):
    if closes is None:
        closes = {}
    data = dict(closes)
    data[month] = {
        "closed": True,
        "closed_at": closed_at or "",
        "closed_by": closed_by,
        "note": note,
        "summary": summary or {},
    }
    return data


def is_month_closed(closes, month):
    if not closes:
        return False
    entry = closes.get(month)
    return bool(entry and entry.get("closed"))


def month_summary(state, month):
    records = state.get("records", []) if isinstance(state, dict) else []
    labor_records = state.get("labor_records", []) if isinstance(state, dict) else []
    payroll_expenses = state.get("payroll_expenses", []) if isinstance(state, dict) else []

    def month_key(record):
        value = record.get("date") or record.get("month") or record.get("recorded_at")
        if not value:
            return None
        return str(value).split("T", 1)[0][:7] if "T" in str(value) else str(value)

    material = 0.0
    expense = 0.0
    labor = 0.0
    payroll = 0.0

    for record in records:
        if month_key(record) != month:
            continue
        amount = float(record.get("amount", 0) or 0)
        if record.get("type") == "material":
            material += amount
        elif record.get("type") in {"expense", "excess"}:
            expense += amount

    for record in labor_records:
        if month_key(record) == month:
            labor += float(record.get("net", 0) or 0)

    for record in payroll_expenses:
        if month_key(record) == month:
            payroll += float(record.get("price", 0) or 0)

    spent = material + expense + labor + payroll
    return {
        "month": month,
        "material": material,
        "expense": expense,
        "labor": labor,
        "payroll": payroll,
        "spent": spent,
    }
