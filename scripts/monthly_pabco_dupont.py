from __future__ import annotations

import json
import math
import os
import smtplib
from collections import defaultdict
from datetime import UTC, date, datetime, time, timedelta
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from geotab_connect import GeotabClient


LOCAL_TZ = ZoneInfo("America/Los_Angeles")
PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", PROJECT_ROOT / "outputs"))
AUDIT_DIR = Path(os.environ.get("OUTPUT_AUDIT_DIR", PROJECT_ROOT / "work"))
OPEN_TIME = time(5)
CLOSE_TIME = time(17)
MIN_DURATION = timedelta(minutes=25)
MERGE_GAP = timedelta(minutes=25)
LONG_LOAD = timedelta(hours=1, minutes=30)
MAX_DURATION = timedelta(hours=4)
PABCO_ZONE_ID = "b1D3B9"
ZONE_TOLERANCE_MILES = float(os.environ.get("ZONE_TOLERANCE_MILES", "0.2"))


def prior_month() -> tuple[date, date]:
    override = os.environ.get("REPORT_MONTH")
    if override:
        start = date.fromisoformat(f"{override}-01")
        next_month = (start.replace(day=28) + timedelta(days=4)).replace(day=1)
        return start, next_month - timedelta(days=1)
    current = datetime.now(LOCAL_TZ).date().replace(day=1)
    end = current - timedelta(days=1)
    return end.replace(day=1), end


REPORT_START, REPORT_END = prior_month()
MONTH_LABEL = REPORT_START.strftime("%B %Y")
OUTPUT_PDF = OUTPUT_DIR / f"PABCO Tacoma and Basalite Dupont Loading Times for {MONTH_LABEL}.pdf"
OUTPUT_JSON = AUDIT_DIR / f"pabco_tacoma_basalite_dupont_{REPORT_START.strftime('%Y_%m')}.json"


def require_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value


def iso_z(value: datetime) -> str:
    return value.astimezone(UTC).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def parse_dt(value: str | None) -> datetime | None:
    if not value:
        return None
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=UTC)
    return parsed.astimezone(LOCAL_TZ).replace(tzinfo=None)


def ref_id(value: object) -> str:
    return str(value.get("id") or "") if isinstance(value, dict) else str(value or "")


def fetch_lookup(client: GeotabClient, type_name: str, fields: list[str]) -> dict[str, dict]:
    rows = client.call("Get", {"typeName": type_name, "credentials": client.credentials, "resultsLimit": 50000,
                                "propertySelector": {"fields": fields, "isIncluded": True}})
    return {row["id"]: row for row in rows if row.get("id")}


def day_bounds(day: date) -> tuple[datetime, datetime]:
    return datetime.combine(day, OPEN_TIME), datetime.combine(day, CLOSE_TIME)


def clamp(start: datetime, end: datetime) -> tuple[datetime, datetime] | None:
    open_at, close_at = day_bounds(start.date())
    bounded_start, bounded_end = max(start, open_at), min(end, close_at)
    return (bounded_start, bounded_end) if bounded_end > bounded_start else None


def point_in_poly(lat: float, lon: float, points: list[dict]) -> bool:
    inside, j = False, len(points) - 1
    for i, point in enumerate(points):
        yi, xi, yj, xj = point.get("y"), point.get("x"), points[j].get("y"), points[j].get("x")
        if None not in (yi, xi, yj, xj) and ((xi > lon) != (xj > lon)):
            if lat < (yj - yi) * (lon - xi) / ((xj - xi) or 1e-12) + yi:
                inside = not inside
        j = i
    return inside


def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = math.radians(lat2 - lat1), math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * 3958.7613 * math.asin(math.sqrt(a))


def pabco_sessions(client: GeotabClient, devices: dict[str, dict]) -> tuple[list[dict], dict]:
    zone = client.call("Get", {"typeName": "Zone", "credentials": client.credentials,
                               "search": {"id": PABCO_ZONE_ID}, "resultsLimit": 1})[0]
    trips = client.call("Get", {"typeName": "Trip", "credentials": client.credentials,
                                "search": {"fromDate": iso_z(datetime.combine(REPORT_START, time.min, LOCAL_TZ)),
                                           "toDate": iso_z(datetime.combine(REPORT_END + timedelta(days=1), time.min, LOCAL_TZ))},
                                "resultsLimit": 50000})
    points = zone.get("points") or []
    center = (sum(float(p["y"]) for p in points) / len(points), sum(float(p["x"]) for p in points) / len(points))
    stops = []
    for trip in trips:
        point = trip.get("stopPoint") or {}
        if point.get("x") is None or point.get("y") is None:
            continue
        lat, lon = float(point["y"]), float(point["x"])
        if not point_in_poly(lat, lon, points) and haversine(lat, lon, *center) > ZONE_TOLERANCE_MILES:
            continue
        start, end = parse_dt(trip.get("stop")), parse_dt(trip.get("nextTripStart"))
        if not start or not end or not (REPORT_START <= start.date() <= REPORT_END):
            continue
        bounded = clamp(start, end)
        if not bounded:
            continue
        device_id = ref_id(trip.get("device"))
        stops.append({"device": devices.get(device_id, {}).get("name") or device_id, "start": bounded[0], "end": bounded[1]})
    grouped = defaultdict(list)
    for stop in stops:
        grouped[(stop["device"], stop["start"].date())].append(stop)
    sessions, ignored_short, probable_errors = [], 0, 0
    for group in grouped.values():
        group.sort(key=lambda x: x["start"])
        current = dict(group[0])
        for stop in group[1:] + [None]:
            if stop and stop["start"] - current["end"] <= MERGE_GAP:
                current["end"] = max(current["end"], stop["end"])
                continue
            duration = current["end"] - current["start"]
            if duration > MAX_DURATION:
                probable_errors += 1
            elif duration >= MIN_DURATION:
                sessions.append({**current, "duration": duration})
            else:
                ignored_short += 1
            if stop:
                current = dict(stop)
    sessions.sort(key=lambda x: x["start"])
    return sessions, {"rawCount": len(trips), "ignoredUnder25Count": ignored_short, "probableErrorCount": probable_errors}


def dupont_sessions(client: GeotabClient, devices: dict[str, dict]) -> tuple[list[dict], dict]:
    rules = client.call("Get", {"typeName": "Rule", "credentials": client.credentials, "resultsLimit": 5000,
                                "propertySelector": {"fields": ["id", "name"], "isIncluded": True}})
    selected = {r["id"]: r.get("name", "") for r in rules if "DUPONT" in r.get("name", "").upper()}
    raw = []
    for rule_id in selected:
        raw.extend(client.call("Get", {"typeName": "ExceptionEvent", "credentials": client.credentials,
                                       "search": {"fromDate": iso_z(datetime.combine(REPORT_START, time.min, LOCAL_TZ)),
                                                  "toDate": iso_z(datetime.combine(REPORT_END + timedelta(days=1), time.min, LOCAL_TZ)),
                                                  "ruleSearch": {"id": rule_id}}, "resultsLimit": 50000}))
    events = []
    for item in raw:
        stamp = parse_dt(item.get("activeFrom"))
        if not stamp:
            continue
        name = selected.get(ref_id(item.get("rule")), "").upper()
        device_id = ref_id(item.get("device"))
        events.append({"device": devices.get(device_id, {}).get("name") or device_id, "timestamp": stamp,
                       "type": "exit" if "EXIT" in name else "entry"})
    grouped = defaultdict(list)
    for event in events:
        grouped[event["device"]].append(event)
    sessions, ignored, probable_errors = [], 0, 0
    for device, group in grouped.items():
        group.sort(key=lambda x: x["timestamp"])
        current = None
        for event in group:
            if event["type"] == "entry":
                if current is None:
                    current = event
                else:
                    ignored += 1
                continue
            if current is None:
                ignored += 1
                continue
            if current["timestamp"].date() != event["timestamp"].date():
                ignored += 2
                current = None
                continue
            bounded = clamp(current["timestamp"], event["timestamp"])
            current = None
            if not bounded:
                ignored += 1
                continue
            duration = bounded[1] - bounded[0]
            if duration > MAX_DURATION:
                probable_errors += 1
            elif duration >= MIN_DURATION:
                sessions.append({"device": device, "start": bounded[0], "end": bounded[1], "duration": duration})
            else:
                ignored += 1
    sessions.sort(key=lambda x: x["start"])
    return sessions, {"rawCount": len(raw), "ignoredEventCount": ignored, "probableErrorCount": probable_errors}


def duration_text(value: timedelta) -> str:
    minutes = int(value.total_seconds() // 60)
    return f"{minutes // 60}:{minutes % 60:02d}"


def build_pdf(reports: list[tuple[str, list[dict]]]) -> None:
    doc = SimpleDocTemplate(str(OUTPUT_PDF), pagesize=letter, leftMargin=.45*inch, rightMargin=.45*inch,
                            topMargin=.4*inch, bottomMargin=.4*inch)
    styles, story = getSampleStyleSheet(), []
    for report_index, (label, sessions) in enumerate(reports):
        story.extend([Paragraph(f"{label} Loading Times - {MONTH_LABEL}", styles["Title"]), Spacer(1, .15*inch)])
        rows = [["Device", "Date", "Start Time", "End Time", "Duration"]]
        for item in sessions:
            rows.append([item["device"], item["start"].strftime("%b %-d, %Y"), item["start"].strftime("%-I:%M:%S %p"),
                         item["end"].strftime("%-I:%M:%S %p"), duration_text(item["duration"])])
        average = timedelta(seconds=sum(x["duration"].total_seconds() for x in sessions)/len(sessions)) if sessions else None
        rows.append(["", "", "", "Average", duration_text(average) if average else "-"])
        table = Table(rows, colWidths=[1.65*inch, 1.2*inch, 1.4*inch, 1.4*inch, .85*inch], repeatRows=1)
        commands = [("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")), ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                    ("GRID",(0,0),(-1,-1),.35,colors.HexColor("#D9D9D9")), ("FONTSIZE",(0,0),(-1,-1),8),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"), ("ALIGN",(1,1),(-1,-1),"CENTER"),
                    ("BACKGROUND",(0,-1),(-1,-1),colors.HexColor("#E7F0F7")),
                    ("TOPPADDING",(0,0),(-1,-1),4), ("BOTTOMPADDING",(0,0),(-1,-1),4)]
        for row_index, item in enumerate(sessions, 1):
            if item["duration"] > LONG_LOAD:
                commands.append(("BACKGROUND",(0,row_index),(-1,row_index),colors.yellow))
        table.setStyle(TableStyle(commands))
        story.append(table)
        if report_index < len(reports)-1:
            story.append(PageBreak())
    doc.build(story)


def email_pdf() -> None:
    required = ["SMTP_HOST", "SMTP_USERNAME", "SMTP_PASSWORD", "EMAIL_TO"]
    missing = [name for name in required if not os.environ.get(name)]
    if missing:
        if os.environ.get("REQUIRE_EMAIL") == "1":
            raise SystemExit(f"Missing email settings: {', '.join(missing)}")
        print("Email skipped")
        return
    msg = EmailMessage()
    msg["From"] = os.environ.get("EMAIL_FROM") or os.environ["SMTP_USERNAME"]
    msg["To"] = os.environ["EMAIL_TO"]
    msg["Subject"] = f"PABCO Tacoma and Basalite Dupont Loading Times for {MONTH_LABEL}"
    msg.set_content(f"Attached is the combined monthly loading-time report for {MONTH_LABEL}.")
    msg.add_attachment(OUTPUT_PDF.read_bytes(), maintype="application", subtype="pdf", filename=OUTPUT_PDF.name)
    with smtplib.SMTP(os.environ["SMTP_HOST"], int(os.environ.get("SMTP_PORT", "587")), timeout=60) as smtp:
        smtp.starttls(); smtp.login(os.environ["SMTP_USERNAME"], os.environ["SMTP_PASSWORD"]); smtp.send_message(msg)
    print(f"emailed={OUTPUT_PDF}")


def main() -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True); AUDIT_DIR.mkdir(parents=True, exist_ok=True)
    client = GeotabClient(os.environ.get("GEOTAB_SERVER", "my.geotab.com"), timeout=int(os.environ.get("GEOTAB_TIMEOUT", "120")))
    client.authenticate(require_env("GEOTAB_USERNAME"), require_env("GEOTAB_PASSWORD"), require_env("GEOTAB_DATABASE"))
    devices = fetch_lookup(client, "Device", ["id", "name"])
    pabco, pabco_meta = pabco_sessions(client, devices)
    dupont, dupont_meta = dupont_sessions(client, devices)
    reports = [("PABCO Tacoma", pabco), ("Basalite Dupont", dupont)]
    build_pdf(reports)
    audit = {"reportStart": REPORT_START.isoformat(), "reportEnd": REPORT_END.isoformat(), "files": {"combinedPdf": str(OUTPUT_PDF)}, "plants": {}}
    for (label, sessions), meta in zip(reports, (pabco_meta, dupont_meta)):
        audit["plants"][label] = {**meta, "cleanedLoadCount": len(sessions),
                                    "longLoadCount": sum(x["duration"] > LONG_LOAD for x in sessions),
                                    "withinOperatingWindow": all(OPEN_TIME <= x["start"].time() <= CLOSE_TIME and OPEN_TIME <= x["end"].time() <= CLOSE_TIME for x in sessions),
                                    "maxDurationSeconds": max((x["duration"].total_seconds() for x in sessions), default=0),
                                    "sortedStartTimes": [x["start"] for x in sessions] == sorted(x["start"] for x in sessions)}
    OUTPUT_JSON.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(json.dumps(audit, indent=2)); email_pdf(); return 0


if __name__ == "__main__":
    raise SystemExit(main())
