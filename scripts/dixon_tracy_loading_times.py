from __future__ import annotations

import json
import math
import os
import sys
import smtplib
from collections import defaultdict
from datetime import UTC, date, datetime, time, timedelta
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table as PdfTable, TableStyle

from geotab_connect import GeotabClient


LOCAL_TZ = ZoneInfo("America/Los_Angeles")
PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", PROJECT_ROOT / "outputs"))
OUTPUT_AUDIT_DIR = Path(os.environ.get("OUTPUT_AUDIT_DIR", PROJECT_ROOT / "work"))
MIN_DURATION = timedelta(minutes=25)
MERGE_GAP = timedelta(minutes=25)
LONG_LOAD = timedelta(hours=1, minutes=30)
MAX_DURATION = timedelta(hours=4)
OPERATING_START = time(5, 0)
OPERATING_END = time(17, 0)
ZONE_TOLERANCE_MILES = float(os.environ.get("ZONE_TOLERANCE_MILES", "0.2"))

ZONES = [
    {"label": "BB Dixon", "title": "Dixon", "id": "b1D3B0", "sheet": "BB Dixon"},
    {"label": "BB Tracy", "title": "Tracy", "id": "bC1E", "sheet": "BB Tracy"},
]


def report_day() -> date:
    override = os.environ.get("REPORT_DATE")
    if override:
        return date.fromisoformat(override)
    today = datetime.now(LOCAL_TZ).date()
    if today.weekday() == 0:
        return today - timedelta(days=3)
    return today - timedelta(days=1)


REPORT_DAY = report_day()
OUTPUT_XLSX = OUTPUT_DIR / f"bb_dixon_tracy_loading_times_{REPORT_DAY.isoformat()}.xlsx"
OUTPUT_JSON = OUTPUT_AUDIT_DIR / f"bb_dixon_tracy_loading_times_{REPORT_DAY.isoformat()}.json"


def display_date(value: date) -> str:
    return f"{value.strftime('%B')} {value.day}"


OUTPUT_PDF = OUTPUT_DIR / f"BB Dixon and BB Tracy Loading Times for {display_date(REPORT_DAY)}.pdf"


def require_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value


def iso_z(dt: datetime) -> str:
    return dt.astimezone(UTC).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def parse_api_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt.astimezone(LOCAL_TZ).replace(tzinfo=None)


def point_in_poly(lat: float, lon: float, points: list[dict]) -> bool:
    inside = False
    j = len(points) - 1
    for i, point in enumerate(points):
        yi, xi = point.get("y"), point.get("x")
        yj, xj = points[j].get("y"), points[j].get("x")
        if yi is None or xi is None or yj is None or xj is None:
            j = i
            continue
        if ((xi > lon) != (xj > lon)) and (lat < (yj - yi) * (lon - xi) / ((xj - xi) or 1e-12) + yi):
            inside = not inside
        j = i
    return inside


def haversine_miles(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius_miles = 3958.7613
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * radius_miles * math.asin(math.sqrt(a))


def zone_center(points: list[dict]) -> tuple[float, float]:
    usable = [point for point in points if "x" in point and "y" in point]
    return (
        sum(float(point["y"]) for point in usable) / len(usable),
        sum(float(point["x"]) for point in usable) / len(usable),
    )


def ref_id(value: object) -> str:
    if isinstance(value, dict):
        return str(value.get("id") or "")
    return str(value or "")


def fetch_lookup(client: GeotabClient, type_name: str, fields: list[str], limit: int = 50000) -> dict[str, dict]:
    rows = client.call(
        "Get",
        {
            "typeName": type_name,
            "credentials": client.credentials,
            "resultsLimit": limit,
            "propertySelector": {"fields": fields, "isIncluded": True},
        },
    )
    return {row["id"]: row for row in rows if row.get("id")}


def fetch_inputs(client: GeotabClient) -> tuple[dict[str, dict], list[dict], dict[str, dict]]:
    zone_map = {}
    for zone_config in ZONES:
        rows = client.call("Get", {"typeName": "Zone", "credentials": client.credentials, "search": {"id": zone_config["id"]}, "resultsLimit": 1})
        if not rows:
            raise SystemExit(f"Zone not found: {zone_config['label']} ({zone_config['id']})")
        zone_map[zone_config["id"]] = rows[0]

    start_local = datetime.combine(REPORT_DAY, time.min, LOCAL_TZ)
    end_local = start_local + timedelta(days=1)
    trips = client.call(
        "Get",
        {
            "typeName": "Trip",
            "credentials": client.credentials,
            "search": {"fromDate": iso_z(start_local), "toDate": iso_z(end_local)},
            "resultsLimit": 50000,
        },
    )
    devices = fetch_lookup(client, "Device", ["id", "name"])
    return zone_map, trips, devices


def sessions_for_zone(zone: dict, trips: list[dict], devices: dict[str, dict]) -> tuple[list[dict], list[dict]]:
    points = zone.get("points") or []
    center_lat, center_lon = zone_center(points)
    stops = []
    ignored = []

    for trip in trips:
        stop_point = trip.get("stopPoint") or {}
        lon = stop_point.get("x")
        lat = stop_point.get("y")
        if lon is None or lat is None:
            continue
        lat = float(lat)
        lon = float(lon)
        distance = haversine_miles(lat, lon, center_lat, center_lon)
        if not point_in_poly(lat, lon, points) and distance > ZONE_TOLERANCE_MILES:
            continue
        start = parse_api_datetime(trip.get("stop"))
        end = parse_api_datetime(trip.get("nextTripStart"))
        if start is None:
            continue
        if end is None:
            end = start
        window_start = datetime.combine(REPORT_DAY, OPERATING_START)
        window_end = datetime.combine(REPORT_DAY, OPERATING_END)
        if end <= window_start or start >= window_end:
            continue
        start = max(start, window_start)
        end = min(end, window_end)
        device_id = ref_id(trip.get("device"))
        device = devices.get(device_id, {})
        stops.append(
            {
                "device": device.get("name") or device_id,
                "device_id": device_id,
                "start": start,
                "end": end,
                "trip_id": trip.get("id", ""),
                "distance_miles": distance,
            }
        )

    grouped = defaultdict(list)
    for stop in stops:
        grouped[stop["device"]].append(stop)

    sessions = []
    for device, group in grouped.items():
        group.sort(key=lambda item: item["start"])
        current = None
        trip_ids = []
        for stop in group:
            if current is None:
                current = dict(stop)
                trip_ids = [stop["trip_id"]]
                continue
            gap = stop["start"] - current["end"]
            if gap <= MERGE_GAP:
                if stop["end"] > current["end"]:
                    current["end"] = stop["end"]
                trip_ids.append(stop["trip_id"])
                continue
            duration = current["end"] - current["start"]
            if duration > MAX_DURATION:
                ignored.append({**current, "duration": duration, "reason": "Probable error: duration over 4 hours"})
            elif duration >= MIN_DURATION:
                sessions.append({**current, "duration": duration, "trip_ids": trip_ids[:]})
            else:
                ignored.append({**current, "duration": duration, "reason": "Stop under 25 minutes"})
            current = dict(stop)
            trip_ids = [stop["trip_id"]]
        if current is not None:
            duration = current["end"] - current["start"]
            if duration > MAX_DURATION:
                ignored.append({**current, "duration": duration, "reason": "Probable error: duration over 4 hours"})
            elif duration >= MIN_DURATION:
                sessions.append({**current, "duration": duration, "trip_ids": trip_ids[:]})
            else:
                ignored.append({**current, "duration": duration, "reason": "Stop under 25 minutes"})

    sessions.sort(key=lambda item: item["start"])
    ignored.sort(key=lambda item: (item["start"], item["device"]))
    return sessions, ignored


def write_zone_sheet(wb: Workbook, sheet_name: str, sessions: list[dict]) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(["Device", "Date", "Start Time", "End Time", "Duration"])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    yellow_fill = PatternFill("solid", fgColor="FFFF00")
    grid = Side(style="thin", color="D9D9D9")
    border = Border(left=grid, right=grid, top=grid, bottom=grid)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", size=14)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = border

    for session in sessions:
        ws.append([session["device"], session["start"].date(), session["start"], session["end"], session["duration"]])

    first_data_row = 2
    last_data_row = 1 + len(sessions)
    average_row = last_data_row + 1
    ws.cell(average_row, 4, "Average")
    ws.cell(average_row, 5, timedelta(seconds=sum(item["duration"].total_seconds() for item in sessions) / len(sessions)) if sessions else None)

    if sessions:
        table_name = "".join(ch for ch in sheet_name if ch.isalnum()) + "LoadingTimes"
        table = Table(displayName=table_name, ref=f"A1:E{last_data_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False)
        ws.add_table(table)

    for row in ws.iter_rows(min_row=2, max_row=average_row, max_col=5):
        for cell in row:
            cell.font = Font(size=12)
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    for row_idx, session in enumerate(sessions, start=first_data_row):
        if session["duration"] > LONG_LOAD:
            for cell in ws[row_idx][0:5]:
                cell.fill = yellow_fill

    for row_idx in range(first_data_row, average_row + 1):
        ws.cell(row_idx, 2).number_format = "mmm d, yyyy"
        ws.cell(row_idx, 3).number_format = "h:mm:ss AM/PM"
        ws.cell(row_idx, 4).number_format = "h:mm:ss AM/PM"
        ws.cell(row_idx, 5).number_format = "[h]:mm"

    ws.cell(average_row, 4).alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 16
    ws.row_dimensions[1].height = 24
    for row_idx in range(2, average_row + 1):
        ws.row_dimensions[row_idx].height = 21
    ws.freeze_panes = "A2"


def save_single_zone_workbook(zone_config: dict, sessions: list[dict]) -> Path:
    wb = Workbook()
    del wb[wb.active.title]
    write_zone_sheet(wb, zone_config["sheet"], sessions)
    output = OUTPUT_DIR / f"{zone_config['title']} Loading Times for {display_date(REPORT_DAY)}.xlsx"
    wb.save(output)
    return output


def duration_text(value: timedelta) -> str:
    total_minutes = int(value.total_seconds() // 60)
    return f"{total_minutes // 60}:{total_minutes % 60:02d}"


def create_combined_pdf(zone_sessions: list[tuple[dict, list[dict]]]) -> Path:
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(
        str(OUTPUT_PDF),
        pagesize=letter,
        leftMargin=0.65 * inch,
        rightMargin=0.65 * inch,
        topMargin=0.45 * inch,
        bottomMargin=0.45 * inch,
    )
    story = []
    for index, (zone_config, sessions) in enumerate(zone_sessions):
        title = Paragraph(
            f"{zone_config['label']} Loading Times - {REPORT_DAY.strftime('%B')} {REPORT_DAY.day}, {REPORT_DAY.year}",
            styles["Title"],
        )
        story.extend([title, Spacer(1, 0.2 * inch)])
        rows = [["Device", "Date", "Start Time", "End Time", "Duration"]]
        for session in sessions:
            rows.append(
                [
                    session["device"],
                    session["start"].strftime("%b %d, %Y").replace(" 0", " "),
                    session["start"].strftime("%-I:%M:%S %p"),
                    session["end"].strftime("%-I:%M:%S %p"),
                    duration_text(session["duration"]),
                ]
            )
        average = timedelta(seconds=sum(item["duration"].total_seconds() for item in sessions) / len(sessions)) if sessions else timedelta(0)
        rows.append(["", "", "", "Average", duration_text(average) if sessions else "-"])
        table = PdfTable(rows, colWidths=[2.05 * inch, 1.18 * inch, 1.5 * inch, 1.5 * inch, 1.02 * inch], repeatRows=1)
        commands = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("ALIGN", (1, 1), (-1, -1), "CENTER"),
            ("ALIGN", (0, 0), (-1, 0), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D9D9")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#FAFAFA")]),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E7F0F7")),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]
        for row_index, session in enumerate(sessions, start=1):
            if session["duration"] > LONG_LOAD:
                commands.append(("BACKGROUND", (0, row_index), (-1, row_index), colors.yellow))
        table.setStyle(TableStyle(commands))
        story.append(table)
        if index < len(zone_sessions) - 1:
            story.append(PageBreak())
    doc.build(story)
    return OUTPUT_PDF


def email_output(file_path: Path) -> None:
    host = os.environ.get("SMTP_HOST")
    port = int(os.environ.get("SMTP_PORT") or "587")
    username = os.environ.get("SMTP_USERNAME")
    password = os.environ.get("SMTP_PASSWORD")
    sender = os.environ.get("EMAIL_FROM") or username
    recipients = [item.strip() for item in os.environ.get("EMAIL_TO", "").split(",") if item.strip()]

    if not all([host, username, password, sender]) or not recipients:
        message = "Email skipped: SMTP_HOST, SMTP_USERNAME, SMTP_PASSWORD, EMAIL_FROM/SMTP_USERNAME, and EMAIL_TO are required."
        if os.environ.get("CI") or os.environ.get("REQUIRE_EMAIL") == "1":
            raise SystemExit(message)
        print(message)
        return

    subject_date = display_date(REPORT_DAY)
    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = f"Dixon and Tracy Loading Times for {subject_date}"
    message.set_content(
        "Attached is the combined Dixon and Tracy loading time report for "
        f"{subject_date}.\n\nThis was generated automatically from Geotab."
    )
    message.add_attachment(file_path.read_bytes(), maintype="application", subtype="pdf", filename=file_path.name)

    with smtplib.SMTP(host, port, timeout=60) as smtp:
        smtp.starttls()
        smtp.login(username, password)
        smtp.send_message(message)
    print(f"emailed={file_path}")


def validate_outputs(audit: dict, plant_files: list[Path]) -> None:
    missing_files = [str(path) for path in plant_files if not path.exists() or path.stat().st_size == 0]
    if missing_files:
        raise SystemExit(f"Missing or empty report files: {', '.join(missing_files)}")

    for zone_label, zone_audit in audit["zones"].items():
        if zone_audit["cleanedLoadCount"] < 0:
            raise SystemExit(f"{zone_label} has an invalid cleaned load count.")
        if zone_audit["ignoredUnder25Count"] < 0:
            raise SystemExit(f"{zone_label} has an invalid ignored stop count.")
        if zone_audit["longLoadCount"] > zone_audit["cleanedLoadCount"]:
            raise SystemExit(f"{zone_label} long-load count exceeds cleaned load count.")
        if not zone_audit["withinOperatingWindow"]:
            raise SystemExit(f"{zone_label} contains time outside 5:00 AM-5:00 PM.")
        if zone_audit["maxDurationSeconds"] > MAX_DURATION.total_seconds():
            raise SystemExit(f"{zone_label} contains a duration over 4 hours.")


def main() -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_AUDIT_DIR.mkdir(parents=True, exist_ok=True)
    client = GeotabClient(os.environ.get("GEOTAB_SERVER", "my.geotab.com"), timeout=int(os.environ.get("GEOTAB_TIMEOUT", "120")))
    client.authenticate(require_env("GEOTAB_USERNAME"), require_env("GEOTAB_PASSWORD"), require_env("GEOTAB_DATABASE"))
    zone_map, trips, devices = fetch_inputs(client)

    wb = Workbook()
    del wb[wb.active.title]
    audit = {
        "reportDate": REPORT_DAY.isoformat(),
        "weekday": REPORT_DAY.strftime("%A"),
        "method": "Trip stopPoint inside zone polygon or within tolerance; clamp to 5:00 AM-5:00 PM; stop to nextTripStart; merge re-entry gaps <= 25 minutes; exclude durations under 25 minutes and probable errors over 4 hours",
        "zoneToleranceMiles": ZONE_TOLERANCE_MILES,
        "rawTripCount": len(trips),
        "files": {},
        "zones": {},
    }
    plant_files = []
    zone_sessions = []

    for zone_config in ZONES:
        sessions, ignored = sessions_for_zone(zone_map[zone_config["id"]], trips, devices)
        zone_sessions.append((zone_config, sessions))
        write_zone_sheet(wb, zone_config["sheet"], sessions)
        single_output = save_single_zone_workbook(zone_config, sessions)
        plant_files.append(single_output)
        sorted_start_times = [item["start"] for item in sessions] == sorted(item["start"] for item in sessions)
        audit["files"][zone_config["label"]] = str(single_output)
        audit["zones"][zone_config["label"]] = {
            "cleanedLoadCount": len(sessions),
            "ignoredUnder25Count": len(ignored),
            "longLoadCount": sum(1 for item in sessions if item["duration"] > LONG_LOAD),
            "probableErrorCount": sum(1 for item in ignored if item["reason"].startswith("Probable error")),
            "averageSeconds": sum(item["duration"].total_seconds() for item in sessions) / len(sessions) if sessions else 0,
            "maxDurationSeconds": max((item["duration"].total_seconds() for item in sessions), default=0),
            "withinOperatingWindow": all(
                OPERATING_START <= item["start"].time() <= OPERATING_END
                and OPERATING_START <= item["end"].time() <= OPERATING_END
                for item in sessions
            ),
            "sortedStartTimes": sorted_start_times,
        }
        if not sorted_start_times:
            raise SystemExit(f"{zone_config['label']} start times are not sorted.")

    wb.save(OUTPUT_XLSX)
    pdf_output = create_combined_pdf(zone_sessions)
    audit["files"]["combinedPdf"] = str(pdf_output)
    validate_outputs(audit, [OUTPUT_XLSX, *plant_files, pdf_output])
    OUTPUT_JSON.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"output={OUTPUT_XLSX}")
    print(json.dumps(audit["zones"], indent=2))
    email_output(pdf_output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
