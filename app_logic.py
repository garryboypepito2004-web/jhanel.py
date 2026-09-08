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
