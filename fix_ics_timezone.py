import re
from datetime import datetime, timedelta

import requests

OUTLOOK_ICS_URL = r"https://outlook.office365.com/owa/calendar/496dcdb8ed5d41ae8b4b04480d79b834@Wyzer.nl/b8f043a309054357980ac45103f7509818297619634366220568/S-1-8-1254520679-68914898-3497378015-486722855/reachcalendar.ics"
OUTPUT_FILE = "fixed_calendar.ics"
TZID = "Europe/Amsterdam"

# Minimal VTIMEZONE block (voldoende voor Google; de events krijgen TZID=Europe/Amsterdam)
VTIMEZONE_BLOCK = f"""BEGIN:VTIMEZONE
TZID:{TZID}
BEGIN:STANDARD
DTSTART:19701025T030000
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:19700329T020000
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
END:DAYLIGHT
END:VTIMEZONE
"""

def unfold_ics(text: str) -> str:
    # RFC 5545 line unfolding: regels die beginnen met spatie/tab horen bij vorige regel
    return re.sub(r"\r?\n[ \t]", "", text)

def fold_ics(text: str) -> str:
    # Voor eenvoud: geen folding. (Google kan lange regels doorgaans prima aan.)
    return text

def last_sunday(year: int, month: int) -> datetime:
    d = datetime(year, month, 31)
    while d.weekday() != 6:  # zondag = 6
        d -= timedelta(days=1)
    return d

def is_dst_europe_amsterdam(dt_local: datetime) -> bool:
    """
    EU zomertijd: laatste zondag maart t/m laatste zondag oktober.
    Praktische benadering in lokale tijd:
    - start: laatste zondag maart 02:00
    - einde: laatste zondag oktober 03:00
    """
    start = last_sunday(dt_local.year, 3).replace(hour=2, minute=0, second=0)
    end = last_sunday(dt_local.year, 10).replace(hour=3, minute=0, second=0)
    return start <= dt_local < end

def utc_z_to_local_eu_amsterdam(value: str) -> str:
    """
    Converteer 'YYYYMMDDTHHMMSSZ' (UTC) naar lokale tijd Europe/Amsterdam.
    Output zonder 'Z' (zodat Google het niet als UTC blijft lezen).
    """
    dt_utc = datetime.strptime(value, "%Y%m%dT%H%M%SZ")

    # Eerst CET (+1), daarna CEST (+2) als dat in DST valt
    dt_local = dt_utc + timedelta(hours=1)
    if is_dst_europe_amsterdam(dt_local):
        dt_local = dt_utc + timedelta(hours=2)

    return dt_local.strftime("%Y%m%dT%H%M%S")

def ensure_calendar_x_wr_timezone(lines: list[str]) -> list[str]:
    """
    Voeg X-WR-TIMEZONE toe direct na BEGIN:VCALENDAR.
    Robuust tegen BOM/rare tekens in BEGIN:VCALENDAR.
    """
    if any(l.strip().startswith("X-WR-TIMEZONE:") for l in lines):
        return lines

    out = []
    inserted = False

    for line in lines:
        # verwijder BOM als die er vooraan zit
        clean = line.lstrip("\ufeff").strip()

        out.append(line)

        if (not inserted) and clean.upper() == "BEGIN:VCALENDAR":
            out.append(f"X-WR-TIMEZONE:{TZID}")
            inserted = True

    # Veiligheidsnet: als BEGIN:VCALENDAR niet exact gevonden is
    if not inserted:
        out.insert(0, f"X-WR-TIMEZONE:{TZID}")

    return out

def ensure_vtimezone(lines: list[str]) -> list[str]:
    if any(l.strip() == "BEGIN:VTIMEZONE" for l in lines):
        return lines

    out = []
    inserted = False

    for line in lines:
        clean = line.lstrip("\ufeff").strip()
        out.append(line)

        if (not inserted) and clean.upper() == "BEGIN:VCALENDAR":
            out.append(VTIMEZONE_BLOCK.strip())
            inserted = True

    if not inserted:
        out.insert(0, VTIMEZONE_BLOCK.strip())

    return out

def fix_dt_line(line: str) -> str:
    """
    Fix DTSTART/DTEND:
    - all-day events (VALUE=DATE) laten we met rust
    - als TZID ontbreekt: toevoegen
    - als tijd in UTC staat met trailing 'Z': omzetten naar lokale tijd en 'Z' verwijderen
    """
    if not (line.startswith("DTSTART") or line.startswith("DTEND")):
        return line

    if "VALUE=DATE" in line:
        return line

    # KEY;params:VALUE  (params kan leeg zijn)
    m = re.match(r"^(DTSTART|DTEND)([^:]*)\:(.*)$", line)
    if not m:
        return line

    key, params, value = m.groups()

    # UTC Z? → lokale tijd
    if value.endswith("Z") and re.fullmatch(r"\d{8}T\d{6}Z", value):
        value = utc_z_to_local_eu_amsterdam(value)

    # TZID toevoegen als ontbreekt
    if "TZID=" not in params:
        params = f"{params};TZID={TZID}"

    return f"{key}{params}:{value}"

def main():
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(OUTLOOK_ICS_URL, headers=headers, timeout=30)
    r.raise_for_status()

    raw = unfold_ics(r.text)
    lines = raw.splitlines()

    # Fix DTSTART/DTEND regels
    fixed = [fix_dt_line(l) for l in lines]

    # Voeg Google-kritische kalender timezone header toe
    fixed = ensure_calendar_x_wr_timezone(fixed)

    # Zorg voor VTIMEZONE (Google gebruikt dit soms i.c.m. TZID)
    fixed = ensure_vtimezone(fixed)

    final_text = fold_ics("\r\n".join(fixed) + "\r\n")

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(final_text)

    print(f"✅ Klaar: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
