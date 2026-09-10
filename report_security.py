import html


def escape_report_text(value):
    if value is None:
        return ""
    return html.escape(str(value), quote=True)


def report_download_name(title):
    safe = (str(title or "report")).strip()
    safe = safe.replace("/", "-").replace("\\", "-")
    safe = "".join(ch if ch.isalnum() or ch in "-_ ." else "_" for ch in safe)
    return safe or "report"
