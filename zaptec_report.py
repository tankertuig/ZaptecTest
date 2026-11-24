#!/usr/bin/env python3
"""
Zaptec monthly charge-session report for a Home Assistant OS local add-on.

Schedule:
  - First day of each month, at the configured local time.
  - Reports sessions whose local startDateTime belongs to the previous month.
Manual test:
  - Create /share/zaptec_maandrapport/nu-uitvoeren.txt
    (for example from the Terminal add-on: touch /share/zaptec_maandrapport/nu-uitvoeren.txt)
  - The add-on notices it within approximately 30 seconds and sends the report.
"""

from __future__ import annotations

import json
import logging
import os
import smtplib
import time
from datetime import datetime, date, time as dtime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


API_BASE = "https://api.zaptec.com"
TOKEN_URL = f"{API_BASE}/oauth/token"
ARCHIVE_URL = f"{API_BASE}/api/sessions/archived"

DATA_DIR = Path("/data")
SHARE_DIR = Path("/share/zaptec_maandrapport")
OPTIONS_FILE = DATA_DIR / "options.json"
LAST_RUN_FILE = DATA_DIR / "last_successful_month.txt"
MANUAL_TRIGGER_FILE = SHARE_DIR / "nu-uitvoeren.txt"

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("zaptec_monthly_report")


def load_options() -> dict[str, Any]:
    if not OPTIONS_FILE.exists():
        raise FileNotFoundError(
            f"{OPTIONS_FILE} ontbreekt. Controleer de configuratie van de add-on."
        )
    with OPTIONS_FILE.open("r", encoding="utf-8") as f:
        opts = json.load(f)

    required = [
        "zaptec_username",
        "zaptec_password",
        "installation_id",
        "gmail_sender",
        "gmail_app_password",
        "recipient",
    ]
    missing = [key for key in required if not str(opts.get(key, "")).strip()]
    if missing:
        raise ValueError("Vul deze verplichte add-on-opties in: " + ", ".join(missing))

    try:
        ZoneInfo(opts.get("timezone", "Europe/Amsterdam"))
    except ZoneInfoNotFoundError as exc:
        raise ValueError("Ongeldige tijdzone in de add-on-configuratie.") from exc

    hour = int(opts.get("report_hour", 8))
    minute = int(opts.get("report_minute", 0))
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        raise ValueError("report_hour moet 0-23 en report_minute 0-59 zijn.")
    return opts


def previous_month_bounds(now_local: datetime) -> tuple[datetime, datetime, str]:
    first_this_month = now_local.replace(
        day=1, hour=0, minute=0, second=0, microsecond=0
    )
    last_month_end = first_this_month
    last_month_start = (first_this_month - timedelta(days=1)).replace(day=1)
    month_label = last_month_start.strftime("%Y-%m")
    return last_month_start, last_month_end, month_label


def parse_api_datetime(value: str | None, tz: ZoneInfo) -> datetime | None:
    if not value:
        return None
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        # Zaptec normally returns ISO timestamps with an offset. If one is absent,
        # treat it as UTC rather than silently using the add-on container timezone.
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(tz)


def request_with_retry(
    session: requests.Session,
    method: str,
    url: str,
    *,
    max_attempts: int = 5,
    **kwargs: Any,
) -> requests.Response:
    for attempt in range(max_attempts):
        response = session.request(method, url, timeout=(15, 60), **kwargs)
        if response.status_code not in (429, 500, 502, 503, 504):
            return response

        retry_after = response.headers.get("Retry-After")
        try:
            delay = min(120, max(1, int(retry_after))) if retry_after else min(60, 2 ** attempt)
        except ValueError:
            delay = min(60, 2 ** attempt)

        log.warning(
            "Zaptec/Gmail tijdelijk niet beschikbaar (HTTP %s); opnieuw proberen over %s s.",
            response.status_code,
            delay,
        )
        if attempt == max_attempts - 1:
            response.raise_for_status()
        time.sleep(delay)

    raise RuntimeError("Onverwachte fout in retry-loop.")


def get_access_token(opts: dict[str, Any]) -> str:
    with requests.Session() as http:
        response = request_with_retry(
            http,
            "POST",
            TOKEN_URL,
            data={
                "grant_type": "password",
                "username": opts["zaptec_username"],
                "password": opts["zaptec_password"],
                "scope": "openid",
            },
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        if response.status_code >= 400:
            raise RuntimeError(
                f"Zaptec-aanmelding mislukt (HTTP {response.status_code}): "
                f"{response.text[:500]}"
            )
        payload = response.json()
        token = payload.get("access_token")
        if not token:
            raise RuntimeError("Zaptec-tokenantwoord bevat geen access_token.")
        return token


def get_archived_sessions(
    opts: dict[str, Any],
    start_local: datetime,
    now_local: datetime,
) -> list[dict[str, Any]]:
    tz = ZoneInfo(opts.get("timezone", "Europe/Amsterdam"))
    token = get_access_token(opts)
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    # The archived endpoint filters on session END time. Query from the first
    # instant of the target month through now, then filter by local START time.
    # This also includes sessions that began in the target month and ended
    # shortly after midnight on the first day of the current month.
    from_utc = start_local.astimezone(timezone.utc)
    to_utc = now_local.astimezone(timezone.utc)

    params: dict[str, Any] = {
        "From": from_utc.isoformat(timespec="seconds").replace("+00:00", "Z"),
        "To": to_utc.isoformat(timespec="seconds").replace("+00:00", "Z"),
        "InstallationId": opts["installation_id"].strip(),
        "PageSize": 200,
        "Order": 0,
    }

    all_sessions: list[dict[str, Any]] = []
    seen_cursors: set[str] = set()

    with requests.Session() as http:
        while True:
            response = request_with_retry(
                http, "GET", ARCHIVE_URL, headers=headers, params=params
            )
            if response.status_code == 401:
                # One fresh-token retry in case the token expired during pagination.
                headers["Authorization"] = f"Bearer {get_access_token(opts)}"
                response = request_with_retry(
                    http, "GET", ARCHIVE_URL, headers=headers, params=params
                )
            if response.status_code >= 400:
                raise RuntimeError(
                    f"Zaptec-laadgeschiedenis ophalen mislukt "
                    f"(HTTP {response.status_code}): {response.text[:1000]}"
                )

            page = response.json()
            if not isinstance(page, dict) or "sessions" not in page:
                raise RuntimeError(
                    "Onverwachte API-respons: verwachtte een object met 'sessions'. "
                    f"Ontvangen: {json.dumps(page, ensure_ascii=False)[:1000]}"
                )

            sessions = page.get("sessions") or []
            all_sessions.extend(s for s in sessions if isinstance(s, dict))

            if not page.get("hasMore"):
                break

            cursor = page.get("cursor")
            if not cursor or cursor in seen_cursors:
                raise RuntimeError("Zaptec-paginering gaf een lege of herhaalde cursor.")
            seen_cursors.add(cursor)
            params["Cursor"] = cursor

    # De API-range filtert op eindtijd; voor het rapport selecteren we op
    # starttijd in de lokale kalendermaand.
    month_start = start_local
    next_month_start = (start_local.replace(day=28) + timedelta(days=4)).replace(day=1)
    unique: dict[str, dict[str, Any]] = {}

    for item in all_sessions:
        session_id = str(item.get("id") or "")
        if not session_id:
            continue
        if item.get("voided") or item.get("aborted") or item.get("externallyAbandoned"):
            continue
        if item.get("replacedBySessionId"):
            continue

        started = parse_api_datetime(item.get("startDateTime"), tz)
        if started is None or not (month_start <= started < next_month_start):
            continue

        unique[session_id] = item

    return sorted(
        unique.values(),
        key=lambda x: x.get("startDateTime", ""),
    )


def make_workbook(
    sessions: list[dict[str, Any]],
    start_local: datetime,
    month_label: str,
    tz: ZoneInfo,
    output_path: Path,
) -> tuple[int, float]:
    wb = Workbook()
    summary = wb.active
    summary.title = "Maandoverzicht"
    details = wb.create_sheet("Laadsessies")

    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    white_bold = Font(color="FFFFFF", bold=True, size=14)
    header_font = Font(bold=True, color="1F1F1F")

    energy_total = sum(float(s.get("energy") or 0) for s in sessions)

    summary["A1"] = f"Zaptec laadrapport - {month_label}"
    summary["A1"].font = white_bold
    summary["A1"].fill = title_fill
    summary.merge_cells("A1:B1")
    summary["A3"] = "Rapportperiode"
    summary["B3"] = start_local.strftime("%d-%m-%Y") + " t/m " + (
        (start_local.replace(day=28) + timedelta(days=4)).replace(day=1)
        - timedelta(days=1)
    ).strftime("%d-%m-%Y")
    summary["A4"] = "Aantal laadsessies"
    summary["B4"] = len(sessions)
    summary["A5"] = "Totaal geladen"
    summary["B5"] = energy_total
    summary["B5"].number_format = '0.00 "kWh"'
    summary["A7"] = "Bron"
    summary["B7"] = "Zaptec API - gearchiveerde laadsessies"
    summary["A8"] = "Tijdzone"
    summary["B8"] = str(tz)
    summary["A9"] = "Opmerking"
    summary["B9"] = (
        "Sessies worden ingedeeld op lokale startdatum. "
        "Gegevens die Zaptec pas later archiveert, kunnen in een latere export verschijnen."
    )
    summary["B9"].alignment = Alignment(wrap_text=True, vertical="top")
    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 95
    summary.row_dimensions[9].height = 42

    headers = [
        "Startdatum",
        "Starttijd",
        "Einddatum",
        "Eindtijd",
        "Energie (kWh)",
        "Laadstation",
        "Laadpas",
        "Gebruiker",
        "E-mailadres gebruiker",
        "Sessie-ID",
    ]
    details.append(headers)
    for cell in details[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for item in sessions:
        started = parse_api_datetime(item.get("startDateTime"), tz)
        ended = parse_api_datetime(item.get("endDateTime"), tz)
        user = item.get("authorizedUser") or {}
        details.append([
            started.date() if started else None,
            started.time().replace(tzinfo=None) if started else None,
            ended.date() if ended else None,
            ended.time().replace(tzinfo=None) if ended else None,
            float(item.get("energy") or 0),
            item.get("deviceName") or item.get("deviceId") or "",
            item.get("tokenName") or "",
            user.get("fullName") or "",
            user.get("email") or "",
            item.get("id") or "",
        ])

    for row in details.iter_rows(min_row=2):
        row[0].number_format = "dd-mm-yyyy"
        row[1].number_format = "hh:mm:ss"
        row[2].number_format = "dd-mm-yyyy"
        row[3].number_format = "hh:mm:ss"
        row[4].number_format = "0.00"
    details.freeze_panes = "A2"
    details.auto_filter.ref = details.dimensions

    widths = [14, 13, 14, 13, 17, 22, 22, 26, 34, 38]
    for idx, width in enumerate(widths, start=1):
        details.column_dimensions[get_column_letter(idx)].width = width

    total_row = details.max_row + 1
    details.cell(total_row, 1, "TOTAAL")
    details.cell(total_row, 1).font = Font(bold=True)
    details.cell(total_row, 5, f"=SUM(E2:E{total_row - 1})")
    details.cell(total_row, 5).font = Font(bold=True)
    details.cell(total_row, 5).number_format = "0.00"

    wb.save(output_path)
    return len(sessions), energy_total


def send_email(opts: dict[str, Any], file_path: Path, month_label: str, count: int, kwh: float) -> None:
    sender = opts["gmail_sender"].strip()
    recipient = opts["recipient"].strip()

    msg = EmailMessage()
    msg["Subject"] = f"Zaptec laadrapport {month_label}"
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content(
        f"Bijgevoegd het Zaptec-laadrapport over {month_label}.\n\n"
        f"Aantal laadsessies: {count}\n"
        f"Totaal geladen: {kwh:.2f} kWh\n\n"
        "Het Excelbestand bevat een maandoverzicht en een tabblad met alle laadsessies."
    )

    with file_path.open("rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=file_path.name,
        )

    with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(sender, opts["gmail_app_password"].replace(" ", ""))
        smtp.send_message(msg)

    log.info("Rapport verstuurd naar %s", recipient)


def create_and_send_report(opts: dict[str, Any]) -> str:
    tz = ZoneInfo(opts.get("timezone", "Europe/Amsterdam"))
    now_local = datetime.now(tz)
    start_local, _end_local, month_label = previous_month_bounds(now_local)

    log.info("Laadsessies ophalen voor %s (installatie %s)", month_label, opts["installation_id"])
    sessions = get_archived_sessions(opts, start_local, now_local)

    SHARE_DIR.mkdir(parents=True, exist_ok=True)
    filename = f"Zaptec_laadrapport_{month_label}.xlsx"
    output_path = SHARE_DIR / filename

    count, kwh = make_workbook(
        sessions, start_local, month_label, tz, output_path
    )
    log.info("Excelbestand aangemaakt: %s (%s sessies, %.2f kWh)", output_path, count, kwh)

    send_email(opts, output_path, month_label, count, kwh)
    return month_label


def already_sent(month_label: str) -> bool:
    return LAST_RUN_FILE.exists() and LAST_RUN_FILE.read_text(encoding="utf-8").strip() == month_label


def mark_sent(month_label: str) -> None:
    LAST_RUN_FILE.parent.mkdir(parents=True, exist_ok=True)
    LAST_RUN_FILE.write_text(month_label, encoding="utf-8")


def main() -> None:
    log.info("Zaptec maandrapport gestart.")
    while True:
        try:
            opts = load_options()
            tz = ZoneInfo(opts.get("timezone", "Europe/Amsterdam"))
            now = datetime.now(tz)
            target_time = dtime(
                hour=int(opts.get("report_hour", 8)),
                minute=int(opts.get("report_minute", 0)),
            )

            manual = MANUAL_TRIGGER_FILE.exists()
            scheduled = now.day == 1 and now.time().replace(tzinfo=None) >= target_time

            if manual or scheduled:
                _, _, month_label = previous_month_bounds(now)
                if already_sent(month_label) and not manual:
                    log.debug("Rapport %s is deze maand al succesvol verstuurd.", month_label)
                else:
                    log.info(
                        "Rapport starten (%s).",
                        "handmatige test" if manual else "maandplanning",
                    )
                    sent_month = create_and_send_report(opts)
                    mark_sent(sent_month)
                    if manual and MANUAL_TRIGGER_FILE.exists():
                        MANUAL_TRIGGER_FILE.unlink()
                    log.info("Rapportage %s succesvol afgerond.", sent_month)

        except Exception:
            log.exception("Rapportage mislukt. De add-on blijft actief en probeert later opnieuw.")

        time.sleep(30)


if __name__ == "__main__":
    main()
