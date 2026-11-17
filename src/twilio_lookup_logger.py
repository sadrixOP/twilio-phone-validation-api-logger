"""Twilio Phone Number Validation & REST API Error Telemetry Logger.

Supports both Live Twilio REST API mode (via TWILIO_ACCOUNT_SID / TWILIO_AUTH_TOKEN)
and high-fidelity Simulation mode (--simulate) emulating Twilio Lookup v2 &
Messaging API JSON responses across HTTP 200, 400, 404, 429, and 500/503 statuses.
"""

from __future__ import annotations

import argparse
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

import matplotlib.gridspec as gridspec
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import requests

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
SCREENSHOT_DIR = BASE_DIR / "screenshots"

CARRIER_CATALOG = {
    "IN": [
        ("Bharti Airtel Ltd", "404", "10", "mobile"),
        ("Reliance Jio Infocomm", "405", "854", "mobile"),
        ("Vodafone Idea Ltd", "404", "20", "mobile"),
        ("BSNL Landline", None, None, "landline"),
    ],
    "US": [
        ("AT&T Wireless", "310", "410", "mobile"),
        ("Verizon Wireless", "311", "480", "mobile"),
        ("T-Mobile USA", "310", "260", "mobile"),
        ("Bandwidth.com VoIP", None, None, "voip"),
        ("Twilio Toll-Free", None, None, "tollFree"),
    ],
    "GB": [
        ("EE Limited", "234", "30", "mobile"),
        ("Vodafone UK", "234", "15", "mobile"),
        ("BT Landline", None, None, "landline"),
    ],
    "AU": [
        ("Telstra Corp", "505", "01", "mobile"),
        ("Optus Mobile", "505", "02", "mobile"),
    ],
    "SG": [
        ("Singtel Mobile", "525", "01", "mobile"),
        ("StarHub Ltd", "525", "05", "mobile"),
    ],
}

OPS_RECOMMENDATIONS = {
    200: "Provision / Route via Primary Carrier",
    400: "Reject & Request Customer E.164 Format Correction",
    404: "Quarantine Unallocated Number & Verify LNP Status",
    429: "Apply Exponential Backoff & Queue Retry",
    500: "Escalate to Carrier Partner / Trigger Fallback Route",
    503: "Retry Lookup & Monitor Super Network Latency",
}


def ensure_directories() -> None:
    """Create project directories if missing."""
    for folder in (DATA_DIR, OUTPUT_DIR, SCREENSHOT_DIR):
        folder.mkdir(parents=True, exist_ok=True)


def generate_input_numbers(count: int = 250) -> pd.DataFrame:
    """Generate 250 deterministic global test phone numbers covering normal & error scenarios."""
    rng = np.random.default_rng(108)
    country_specs = [
        ("IN", "+91", 10, 0.36),
        ("US", "+1", 10, 0.28),
        ("GB", "+44", 10, 0.14),
        ("AU", "+61", 9, 0.12),
        ("SG", "+65", 8, 0.10),
    ]
    codes = [c[0] for c in country_specs]
    weights = [c[3] for c in country_specs]
    code_map = {c[0]: (c[1], c[2]) for c in country_specs}

    rows: list[dict[str, str]] = []
    for idx in range(1, count + 1):
        cc = str(rng.choice(codes, p=weights))
        dial_prefix, digits_len = code_map[cc]
        digits = "".join(str(d) for d in rng.integers(6, 10, size=1)) + "".join(
            str(d) for d in rng.integers(0, 10, size=digits_len - 1)
        )

        roll = rng.random()
        if roll < 0.74:
            scenario = "valid"
            raw_number = f"{dial_prefix}{digits}"
        elif roll < 0.85:
            scenario = "malformed_400"
            raw_number = f"{digits[:4]}-INVALID-{digits[4:7]}" if idx % 2 == 0 else digits
        elif roll < 0.92:
            scenario = "unallocated_404"
            raw_number = f"{dial_prefix}0000{digits[4:]}"
        elif roll < 0.96:
            scenario = "rate_limit_429"
            raw_number = f"{dial_prefix}{digits}"
        else:
            scenario = "upstream_500"
            raw_number = f"{dial_prefix}{digits}"

        rows.append(
            {
                "batch_seq": f"REQ-{idx:04d}",
                "account_id": f"ACC-{101 + (idx % 12)}",
                "country_code": cc,
                "phone_number_input": raw_number,
                "test_scenario": scenario,
                "api_channel": "Lookup_v2" if idx % 5 != 0 else "Lookup_v2+Messaging",
            }
        )

    df = pd.DataFrame(rows)
    df.to_csv(DATA_DIR / "phone_numbers_input.csv", index=False, encoding="utf-8")
    return df


class TwilioPhoneValidator:
    """Validates phone numbers via Twilio Lookup v2 / Messaging REST APIs or simulation."""

    LOOKUP_URL = "https://lookups.twilio.com/v2/PhoneNumbers/{phone}?Fields=line_type_intelligence,caller_name"
    MESSAGES_URL = "https://api.twilio.com/2010-04-01/Accounts/{sid}/Messages.json"
    E164_RE = re.compile(r"^\+[1-9]\d{6,14}$")

    def __init__(self, simulate: bool = True) -> None:
        self.account_sid = os.getenv("TWILIO_ACCOUNT_SID", "").strip()
        self.auth_token = os.getenv("TWILIO_AUTH_TOKEN", "").strip()
        self.from_number = os.getenv("TWILIO_FROM_NUMBER", "+15005550006").strip()
        self.simulate = simulate or not (self.account_sid and self.auth_token)
        self.rng = np.random.default_rng(2026)

        self.session = requests.Session()
        if not self.simulate:
            self.session.auth = (self.account_sid, self.auth_token)

    def validate_row(self, row: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        """Execute a single Lookup (+ optional Messaging) check and return flat log + raw JSON."""
        if self.simulate:
            return self._simulate_api_call(row)
        return self._live_api_call(row)

    def _live_api_call(self, row: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        """Invoke real Twilio Lookup v2 and optionally Messaging REST endpoints."""
        phone = str(row["phone_number_input"])
        url = self.LOOKUP_URL.format(phone=quote(phone, safe=""))
        start_ts = datetime.now(timezone.utc)
        resp = self.session.get(url, timeout=10)
        latency_ms = round((datetime.now(timezone.utc) - start_ts).total_seconds() * 1000, 1)

        try:
            payload = resp.json()
        except ValueError:
            payload = {"message": resp.text, "code": None}

        lti = payload.get("line_type_intelligence") or {}
        http_status = resp.status_code
        twilio_code = payload.get("code")
        error_msg = payload.get("message") if http_status >= 400 else ""

        log_entry = {
            "batch_seq": row["batch_seq"],
            "account_id": row["account_id"],
            "country_code": row["country_code"],
            "phone_number_input": phone,
            "normalized_e164": payload.get("phone_number", ""),
            "e164_valid": bool(payload.get("valid", False)),
            "http_status_code": http_status,
            "twilio_error_code": twilio_code if twilio_code else "",
            "error_message": error_msg,
            "carrier_name": lti.get("carrier_name", ""),
            "line_type": lti.get("type", ""),
            "mcc": lti.get("mobile_country_code", ""),
            "mnc": lti.get("mobile_network_code", ""),
            "api_endpoint": "GET /v2/PhoneNumbers",
            "latency_ms": latency_ms,
            "recommended_ops_action": OPS_RECOMMENDATIONS.get(http_status, "Inspect API Log"),
            "checked_at_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        return log_entry, {"request": url, "http_status": http_status, "response": payload}

    def _simulate_api_call(self, row: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        """Emulate authentic Twilio Lookup v2 & Messaging REST JSON responses and error codes."""
        phone = str(row["phone_number_input"])
        cc = str(row["country_code"])
        scenario = str(row["test_scenario"])
        is_e164 = bool(self.E164_RE.match(phone))

        if scenario == "malformed_400" or not is_e164:
            http_status = 400
            err_code = 21211 if "-" in phone or not phone.startswith("+") else 21614
            err_msg = (
                f"The 'To' number '{phone}' is not a valid E.164 phone number."
                if err_code == 21211
                else f"The number '{phone}' is not a valid mobile handset destination."
            )
            payload: dict[str, Any] = {
                "code": err_code,
                "message": err_msg,
                "more_info": f"https://www.twilio.com/docs/errors/{err_code}",
                "status": 400,
            }
            latency_ms = round(float(self.rng.uniform(42, 110)), 1)
            carrier, line_type, mcc, mnc, norm_e164, valid_flag = "", "", "", "", "", False

        elif scenario == "unallocated_404":
            http_status = 404
            err_code = 20404
            err_msg = f"The requested resource /v2/PhoneNumbers/{phone} was not found."
            payload = {
                "code": err_code,
                "message": err_msg,
                "more_info": "https://www.twilio.com/docs/errors/20404",
                "status": 404,
            }
            latency_ms = round(float(self.rng.uniform(65, 145)), 1)
            carrier, line_type, mcc, mnc, norm_e164, valid_flag = "", "", "", "", phone, False

        elif scenario == "rate_limit_429":
            http_status = 429
            err_code = 20429
            err_msg = "Too Many Requests: Lookup API concurrency quota exceeded."
            payload = {
                "code": err_code,
                "message": err_msg,
                "more_info": "https://www.twilio.com/docs/errors/20429",
                "status": 429,
            }
            latency_ms = round(float(self.rng.uniform(28, 70)), 1)
            carrier, line_type, mcc, mnc, norm_e164, valid_flag = "", "", "", "", phone, False

        elif scenario == "upstream_500":
            http_status = int(self.rng.choice([500, 503], p=[0.7, 0.3]))
            err_code = 30008
            err_msg = "Unknown error / Upstream carrier HLR lookup timeout."
            payload = {
                "code": err_code,
                "message": err_msg,
                "more_info": "https://www.twilio.com/docs/errors/30008",
                "status": http_status,
            }
            latency_ms = round(float(self.rng.uniform(950, 2400)), 1)
            carrier, line_type, mcc, mnc, norm_e164, valid_flag = "", "", "", "", phone, False

        else:
            http_status = 200
            err_code = None
            err_msg = ""
            options = CARRIER_CATALOG[cc]
            carrier, mcc, mnc, line_type = options[int(self.rng.integers(0, len(options)))]
            norm_e164 = phone
            valid_flag = True
            latency_ms = round(float(self.rng.uniform(85, 290)), 1)
            payload = {
                "calling_country_code": phone[1:3] if cc != "US" else "1",
                "country_code": cc,
                "phone_number": norm_e164,
                "national_format": norm_e164,
                "valid": True,
                "validation_errors": [],
                "line_type_intelligence": {
                    "carrier_name": carrier,
                    "type": line_type,
                    "mobile_country_code": mcc,
                    "mobile_network_code": mnc,
                    "error_code": None,
                },
                "url": f"https://lookups.twilio.com/v2/PhoneNumbers/{norm_e164}",
            }

        log_entry = {
            "batch_seq": row["batch_seq"],
            "account_id": row["account_id"],
            "country_code": cc,
            "phone_number_input": phone,
            "normalized_e164": norm_e164,
            "e164_valid": valid_flag,
            "http_status_code": http_status,
            "twilio_error_code": err_code if err_code else "",
            "error_message": err_msg,
            "carrier_name": carrier,
            "line_type": line_type,
            "mcc": mcc or "",
            "mnc": mnc or "",
            "api_endpoint": "GET /v2/PhoneNumbers"
            if row["api_channel"] == "Lookup_v2"
            else "GET /v2/PhoneNumbers + POST /Messages.json",
            "latency_ms": latency_ms,
            "recommended_ops_action": OPS_RECOMMENDATIONS.get(http_status, "Inspect API Log"),
            "checked_at_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        return log_entry, {
            "request": f"GET https://lookups.twilio.com/v2/PhoneNumbers/{phone}?Fields=line_type_intelligence",
            "http_status": http_status,
            "response": payload,
        }


def export_excel_report(logs_df: pd.DataFrame) -> Path:
    """Create a color-coded multi-sheet Excel audit workbook with openpyxl."""
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = "API_Summary_Metrics"

    header_fill = PatternFill("solid", fgColor="0F172A")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    accent_fill = PatternFill("solid", fgColor="E2E8F0")

    # Summary KPI table
    ws_summary["A1"] = "Twilio Phone Number Lookup & API Error Telemetry Report"
    ws_summary["A1"].font = Font(bold=True, size=14, color="0F172A")
    ws_summary.append([])

    total = len(logs_df)
    ok_count = int((logs_df["http_status_code"] == 200).sum())
    err_count = total - ok_count
    avg_lat = round(float(logs_df["latency_ms"].mean()), 1)
    p95_lat = round(float(logs_df["latency_ms"].quantile(0.95)), 1)

    metrics = [
        ("Metric", "Value"),
        ("Total Numbers Validated", total),
        ("HTTP 200 OK (Valid & Enriched)", ok_count),
        ("HTTP 4xx/5xx Errors Logged", err_count),
        ("Validation Success Rate (%)", round(100.0 * ok_count / total, 2)),
        ("Mean API Latency (ms)", avg_lat),
        ("P95 API Latency (ms)", p95_lat),
    ]
    for r_idx, (k, v) in enumerate(metrics, start=3):
        ws_summary.cell(row=r_idx, column=1, value=k)
        ws_summary.cell(row=r_idx, column=2, value=v)
        if r_idx == 3:
            for col in (1, 2):
                cell = ws_summary.cell(row=r_idx, column=col)
                cell.fill = header_fill
                cell.font = header_font
        elif r_idx % 2 == 0:
            for col in (1, 2):
                ws_summary.cell(row=r_idx, column=col).fill = accent_fill

    # Sheet 2: HTTP Error Breakdown
    ws_err = wb.create_sheet("HTTP_Error_Breakdown")
    breakdown = (
        logs_df.groupby(["http_status_code", "twilio_error_code", "recommended_ops_action"], dropna=False)
        .agg(
            request_count=("batch_seq", "count"),
            avg_latency_ms=("latency_ms", "mean"),
        )
        .reset_index()
        .sort_values("request_count", ascending=False)
    )
    breakdown["share_pct"] = (100.0 * breakdown["request_count"] / total).round(2)
    breakdown["avg_latency_ms"] = breakdown["avg_latency_ms"].round(1)

    err_headers = list(breakdown.columns)
    ws_err.append(err_headers)
    for col_idx in range(1, len(err_headers) + 1):
        c = ws_err.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
    for row_vals in breakdown.itertuples(index=False):
        ws_err.append(list(row_vals))

    # Sheet 3: Detailed Lookup Logs with HTTP status color coding
    ws_logs = wb.create_sheet("Detailed_Lookup_Logs")
    log_headers = list(logs_df.columns)
    ws_logs.append(log_headers)
    for col_idx in range(1, len(log_headers) + 1):
        c = ws_logs.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font

    status_fills = {
        200: PatternFill("solid", fgColor="DCFCE7"),
        400: PatternFill("solid", fgColor="FEF3C7"),
        404: PatternFill("solid", fgColor="FFEDD5"),
        429: PatternFill("solid", fgColor="EDE9FE"),
        500: PatternFill("solid", fgColor="FEE2E2"),
        503: PatternFill("solid", fgColor="FEE2E2"),
    }
    status_col_idx = log_headers.index("http_status_code") + 1
    for r_i, row_vals in enumerate(logs_df.itertuples(index=False), start=2):
        ws_logs.append(list(row_vals))
        code = int(getattr(row_vals, "http_status_code"))
        if code in status_fills:
            ws_logs.cell(row=r_i, column=status_col_idx).fill = status_fills[code]
            ws_logs.cell(row=r_i, column=status_col_idx).font = Font(bold=True)

    for sheet in (ws_summary, ws_err, ws_logs):
        sheet.freeze_panes = "A2" if sheet != ws_summary else "A4"
        for col in sheet.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 46)
            for cell in col:
                cell.alignment = Alignment(vertical="center")

    out_xlsx = OUTPUT_DIR / "Twilio_API_Validation_Report.xlsx"
    wb.save(out_xlsx)
    return out_xlsx


def render_telemetry_png(logs_df: pd.DataFrame) -> Path:
    """Render a 300 DPI diagnostic API telemetry dashboard screenshot."""
    plt.rcParams.update(
        {
            "figure.facecolor": "#0B1120",
            "axes.facecolor": "#111827",
            "axes.edgecolor": "#1F2937",
            "axes.labelcolor": "#E5E7EB",
            "text.color": "#F8FAFC",
            "xtick.color": "#9CA3AF",
            "ytick.color": "#D1D5DB",
            "grid.color": "#1F2937",
            "font.family": "DejaVu Sans",
        }
    )

    fig = plt.figure(figsize=(15.5, 10.8), dpi=300)
    gs = gridspec.GridSpec(3, 2, height_ratios=[0.36, 1.0, 1.0], hspace=0.40, wspace=0.25)

    total = len(logs_df)
    ok_pct = 100.0 * (logs_df["http_status_code"] == 200).mean()
    err_cnt = int((logs_df["http_status_code"] >= 400).sum())
    p95_ms = float(logs_df["latency_ms"].quantile(0.95))

    ax_top = fig.add_subplot(gs[0, :])
    ax_top.axis("off")
    ax_top.text(0.0, 0.92, "TWILIO PHONE NUMBER VALIDATION & REST API ERROR TELEMETRY", fontsize=16.5, fontweight="bold", transform=ax_top.transAxes)
    ax_top.text(
        0.0,
        0.74,
        "Endpoints: Lookup v2 (/v2/PhoneNumbers) + Programmable Messaging (/Messages.json)  |  Batch Size: 250 Global Numbers",
        fontsize=10.5,
        color="#94A3B8",
        transform=ax_top.transAxes,
    )

    cards = [
        ("API CALLS EXECUTED", f"{total}", "5 Global Regions (+91,+1,+44,+61,+65)", "#38BDF8"),
        ("HTTP 200 VALID RATE", f"{ok_pct:.1f}%", "E.164 & Carrier Enriched", "#10B981"),
        ("4xx / 5xx ERRORS", f"{err_cnt}", "21211 / 20404 / 20429 / 30008", "#EF4444"),
        ("P95 API LATENCY", f"{p95_ms:.0f} ms", "Includes Carrier HLR Outliers", "#F59E0B"),
    ]
    for i, (lbl, val, sub, color) in enumerate(cards):
        x0 = i * 0.253
        ax_top.add_patch(
            plt.Rectangle((x0, 0.04), 0.235, 0.58, transform=ax_top.transAxes, facecolor="#111827", edgecolor=color, lw=1.6, clip_on=False)
        )
        ax_top.text(x0 + 0.014, 0.48, lbl, fontsize=8.5, fontweight="bold", color="#94A3B8", transform=ax_top.transAxes)
        ax_top.text(x0 + 0.014, 0.25, val, fontsize=16, fontweight="bold", color=color, transform=ax_top.transAxes)
        ax_top.text(x0 + 0.014, 0.09, sub, fontsize=8.0, color="#CBD5E1", transform=ax_top.transAxes)

    # Panel 1: HTTP Status Breakdown
    ax1 = fig.add_subplot(gs[1, 0])
    status_counts = logs_df["http_status_code"].value_counts().sort_index()
    status_colors = {200: "#10B981", 400: "#F59E0B", 404: "#FB923C", 429: "#A78BFA", 500: "#EF4444", 503: "#DC2626"}
    labels = [f"HTTP {code}" for code in status_counts.index]
    cols = [status_colors.get(int(code), "#38BDF8") for code in status_counts.index]
    bars = ax1.bar(labels, status_counts.values, color=cols, width=0.55)
    ax1.set_title("1. HTTP Status Code Distribution (200 / 400 / 404 / 429 / 5xx)", fontsize=11, fontweight="bold", loc="left", pad=10)
    ax1.set_ylabel("Request Count")
    ax1.grid(axis="y", linestyle="--", alpha=0.45)
    ax1.set_ylim(0, status_counts.max() * 1.18)
    for b in bars:
        ax1.text(b.get_x() + b.get_width() / 2, b.get_height() + 3, f"{int(b.get_height())}", ha="center", fontsize=9.5, fontweight="bold")

    # Panel 2: Line Type Intelligence
    ax2 = fig.add_subplot(gs[1, 1])
    valid_df = logs_df[logs_df["http_status_code"] == 200]
    lt_counts = valid_df["line_type"].value_counts()
    lt_bars = ax2.barh(lt_counts.index, lt_counts.values, color=["#38BDF8", "#10B981", "#A78BFA", "#F59E0B"][: len(lt_counts)], height=0.52)
    ax2.set_title("2. Line Type Intelligence (Valid HTTP 200 Numbers)", fontsize=11, fontweight="bold", loc="left", pad=10)
    ax2.set_xlabel("Validated Numbers")
    ax2.grid(axis="x", linestyle="--", alpha=0.45)
    ax2.set_xlim(0, lt_counts.max() * 1.2)
    for b in lt_bars:
        ax2.text(b.get_width() + 2, b.get_y() + b.get_height() / 2, f"{int(b.get_width())}", va="center", fontsize=9.5)

    # Panel 3: Top Carriers Discovered
    ax3 = fig.add_subplot(gs[2, 0])
    top_carriers = valid_df["carrier_name"].value_counts().head(8).sort_values(ascending=True)
    cbars = ax3.barh(top_carriers.index, top_carriers.values, color="#38BDF8", height=0.58)
    ax3.set_title("3. Top Carrier Networks Resolved via Lookup v2", fontsize=11, fontweight="bold", loc="left", pad=10)
    ax3.set_xlabel("Count")
    ax3.grid(axis="x", linestyle="--", alpha=0.45)
    ax3.set_xlim(0, top_carriers.max() * 1.18)
    for b in cbars:
        ax3.text(b.get_width() + 0.8, b.get_y() + b.get_height() / 2, f"{int(b.get_width())}", va="center", fontsize=8.8)

    # Panel 4: API Latency Distribution
    ax4 = fig.add_subplot(gs[2, 1])
    ok_lat = logs_df[logs_df["http_status_code"] < 500]["latency_ms"]
    err5_lat = logs_df[logs_df["http_status_code"] >= 500]["latency_ms"]
    ax4.hist(ok_lat, bins=22, color="#10B981", alpha=0.85, label="2xx / 4xx Responses")
    if len(err5_lat) > 0:
        ax4.hist(err5_lat, bins=8, color="#EF4444", alpha=0.85, label="5xx Carrier Timeouts")
    ax4.set_title("4. REST API Response Latency Distribution (ms)", fontsize=11, fontweight="bold", loc="left", pad=10)
    ax4.set_xlabel("Latency (ms)")
    ax4.set_ylabel("Frequency")
    ax4.grid(True, linestyle="--", alpha=0.45)
    ax4.legend(frameon=False, fontsize=9)

    out_png = SCREENSHOT_DIR / "api_error_telemetry.png"
    fig.savefig(out_png, dpi=300, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return out_png


def main() -> None:
    """Run Twilio validation logger in simulation or live mode."""
    parser = argparse.ArgumentParser(description="Twilio Phone Number Validation & API Error Logger")
    parser.add_argument("--live", action="store_true", help="Call live Twilio REST APIs using env credentials")
    args = parser.parse_args()

    ensure_directories()
    input_df = generate_input_numbers(count=250)
    validator = TwilioPhoneValidator(simulate=not args.live)

    logs: list[dict[str, Any]] = []
    payloads: list[dict[str, Any]] = []
    for row in input_df.to_dict(orient="records"):
        entry, raw_json = validator.validate_row(row)
        logs.append(entry)
        payloads.append(raw_json)

    logs_df = pd.DataFrame(logs)
    logs_df.to_csv(OUTPUT_DIR / "api_validation_log.csv", index=False, encoding="utf-8")

    # Persist sample raw JSON audit log (first 40 payloads + all error samples)
    error_payloads = [p for p in payloads if p["http_status"] >= 400]
    ok_payloads = [p for p in payloads if p["http_status"] == 200][:25]
    audit_path = OUTPUT_DIR / "api_payload_audit.jsonl"
    audit_path.write_text(
        "\n".join(json.dumps(item) for item in (ok_payloads + error_payloads)),
        encoding="utf-8",
    )

    xlsx_path = export_excel_report(logs_df)
    png_path = render_telemetry_png(logs_df)

    mode_label = "SIMULATION" if validator.simulate else "LIVE TWILIO API"
    print(f"[Project 2] Mode: {mode_label} | Processed {len(logs_df)} phone numbers")
    print(f"[Project 2] CSV log -> {OUTPUT_DIR / 'api_validation_log.csv'}")
    print(f"[Project 2] Excel report -> {xlsx_path}")
    print(f"[Project 2] Telemetry screenshot -> {png_path}")


if __name__ == "__main__":
    main()
