import re


def parse_scanned_receipt(text):
    full_text = (text or "").strip()
    if not full_text:
        return {"name": "", "qty": 1, "price": 0.0, "delivery": 0.0}

    lines = [line.strip() for line in full_text.splitlines() if line.strip()]
    name = ""
    qty = 1
    price = 0.0
    delivery = 0.0

    for line in lines:
        lowered = line.lower()
        if not name and len(line) > 2 and not re.fullmatch(r"[0-9\s.,/\-]+", line):
            name = line
        match_qty = re.search(r"qty\s*[:=]?\s*(\d+)", lowered)
        if match_qty and qty == 1:
            qty = int(match_qty.group(1))
        match_price = re.search(r"(?:total|amount|price|net|amt)\s*[:=]?\s*php\s*([0-9,]+(?:\.\d+)?)", lowered)
        if match_price and price == 0.0:
            price = float(match_price.group(1).replace(",", ""))
        match_del = re.search(r"delivery\s*[:=]?\s*php\s*([0-9,]+(?:\.\d+)?)", lowered)
        if match_del and delivery == 0.0:
            delivery = float(match_del.group(1).replace(",", ""))

    if price == 0.0:
        for line in lines:
            number_match = re.search(r"([0-9]+(?:,[0-9]{3})*(?:\.\d+)?)", line)
            if number_match:
                candidate = float(number_match.group(1).replace(",", ""))
                if candidate > 0 and candidate > delivery:
                    price = candidate
                    break

    if not name:
        name = "Scanned material"

    return {
        "name": name,
        "qty": qty,
        "price": price,
        "delivery": delivery,
    }
