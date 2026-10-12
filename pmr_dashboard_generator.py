 from __future__ import annotations

import html
import math
import re
from pathlib import Path
import tempfile

import streamlit as st
import streamlit.components.v1 as components
from typing import Any

from openpyxl import load_workbook


COLORS = {
    "navy": "#12344D",
    "blue": "#0072CE",
    "green": "#2E8540",
    "red": "#D13438",
    "amber": "#F2C811",
    "text": "#172B4D",
    "muted": "#68768A",
}


def text(value: Any) -> str:
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value).replace("\xa0", " ").strip(),
    )


def norm(value: Any) -> str:
    return re.sub(
        r"[^a-z0-9%$]+",
        " ",
        text(value).lower().replace("&", "and"),
    ).strip()


def num(value: Any):
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        numeric_value = float(value)

        if math.isfinite(numeric_value):
            return numeric_value

        return None

    value_text = text(value)

    if not value_text:
        return None

    if value_text.lower() in {
        "n/a",
        "na",
        "n.a.",
        "-",
        "—",
        "#n/a",
    }:
        return None

    negative = (
        value_text.startswith("(")
        and value_text.endswith(")")
    )

    try:
        numeric_value = float(
            value_text
            .replace(",", "")
            .replace("$", "")
            .replace("%", "")
            .replace("(", "")
            .replace(")", "")
        )

        return -numeric_value if negative else numeric_value

    except ValueError:
        return None


def fmt(value, pct=False):
    if value is None:
        return "N/A"

    if pct:
        return f"{value:.1%}"

    return f"{value:,.0f}"


def hint(value, choices):
    normalized_value = norm(value)

    if normalized_value in choices:
        return 5

    if any(
        choice in normalized_value
        or normalized_value in choice
        for choice in choices
        if normalized_value
    ):
        return 3

    return 0


def headers(ws):
    current_names = {
        "current",
        "curr",
        "actual",
        "new",
        "current p27",
        "p27",
    }

    prior_names = {
        "prior",
        "previous",
        "old",
        "prior p27",
        "rev3",
    }

    found = []

    for row_number in range(1, ws.max_row + 1):

        values = [
            ws.cell(row_number, column_number).value
            for column_number in range(
                1,
                min(ws.max_column, 15) + 1,
            )
        ]

        current_match = max(
            [
                (
                    hint(value, current_names),
                    index + 1,
                )
                for index, value in enumerate(values)
            ],
            default=(0, 0),
        )

        prior_match = max(
            [
                (
                    hint(value, prior_names),
                    index + 1,
                )
                for index, value in enumerate(values)
            ],
            default=(0, 0),
        )

        if (
            current_match[0] >= 3
            and prior_match[0] >= 3
            and current_match[1] != prior_match[1]
        ):
            found.append(
                {
                    "row": row_number,
                    "label": max(
                        1,
                        current_match[1] - 1,
                    ),
                    "current": current_match[1],
                    "prior": prior_match[1],
                }
            )

    return found


def section(ws, row, column):
    """
    Looks above each Current/Prior header to find the title
    of the corresponding Excel table.
    """

    ignored_names = {
        "current",
        "prior",
        "p27",
        "current p27",
        "prior p27",
        "var",
        "var k$",
        "var %",
        "var%",
    }

    # The title normally appears between one and four rows
    # above the Current/Prior header.
    for candidate_row in range(
        row - 1,
        max(0, row - 7),
        -1,
    ):
        possible_columns = [
            column,
            max(1, column - 1),
            max(1, column - 2),
        ]

        for candidate_column in possible_columns:
            value = text(
                ws.cell(
                    candidate_row,
                    candidate_column,
                ).value
            )

            normalized_value = norm(value)

            if (
                value
                and normalized_value not in ignored_names
                and "mm dd yy" not in normalized_value
                and "curr prior" not in normalized_value
                and not normalized_value.startswith(
                    "thousands"
                )
            ):
                return value

    return "General"


def read_excel(path):
    # read_only=False is intentional because Excel font
    # formatting is needed to preserve bold rows.
    workbook = load_workbook(
        path,
        data_only=False,
        read_only=False,
    )

    result = {}

    for sheet_name in workbook.sheetnames:

        worksheet = workbook[sheet_name]
        detected_headers = headers(worksheet)
        rows = []

        for header_index, header in enumerate(
            detected_headers
        ):
            if header_index + 1 < len(detected_headers):
                end_row = (
                    detected_headers[
                        header_index + 1
                    ]["row"]
                    - 1
                )
            else:
                end_row = worksheet.max_row

            section_name = section(
                worksheet,
                header["row"],
                header["label"],
            )

            # Give each detected table a unique identifier,
            # even when names are repeated.
            table_key = (
                f"{sheet_name}|"
                f"{header['row']}|"
                f"{section_name}"
            )

            for row_number in range(
                header["row"] + 1,
                end_row + 1,
            ):
                label_cell = worksheet.cell(
                    row_number,
                    header["label"],
                )

                label = text(label_cell.value)

                if not label:
                    continue

                if label.startswith("("):
                    continue

                current_cell = worksheet.cell(
                    row_number,
                    header["current"],
                )

                prior_cell = worksheet.cell(
                    row_number,
                    header["prior"],
                )

                current_value = num(current_cell.value)
                prior_value = num(prior_cell.value)

                if (
                    current_value is None
                    and prior_value is None
                ):
                    continue

                format_code = (
                    current_cell.number_format or ""
                )

                normalized_label = norm(label)

                is_percentage = (
                    "%" in format_code
                    or any(
                        percentage_word
                        in normalized_label
                        for percentage_word in (
                            "rate",
                            "roe",
                            "ratio",
                            "percent",
                        )
                    )
                )

                if (
                    current_value is None
                    or prior_value is None
                ):
                    variance = None
                else:
                    variance = (
                        current_value - prior_value
                    )

                if variance is None:
                    variance_percentage = None
                elif (
                    prior_value == 0
                    and current_value != 0
                ):
                    variance_percentage = None
                elif prior_value == 0:
                    variance_percentage = 0
                else:
                    variance_percentage = (
                        variance / abs(prior_value)
                    )

                excel_bold = bool(
                    label_cell.font
                    and label_cell.font.bold
                )

                rows.append(
                    {
                        "sheet": sheet_name,
                        "table_key": table_key,
                        "section": section_name,
                        "label": label,
                        "current": current_value,
                        "prior": prior_value,
                        "var": variance,
                        "varpct": variance_percentage,
                        "is_pct": is_percentage,
                        "is_bold": excel_bold,
                    }
                )

        result[sheet_name] = rows

    return workbook.sheetnames, result


def sid(value):
    safe_value = re.sub(
        r"[^A-Za-z0-9_-]+",
        "-",
        value,
    ).strip("-")

    return "s-" + safe_value


ALLOWED_RANKING_LINES = {
    norm("REVENUE"), norm("PROCUREMENTS"), norm("PERSONNALE"),
    norm("CAPITALIZED STAFF COSTS"), norm("EXTERNAL SERVICES"),
    norm("OTHER OPERATING REVENUE"), norm("LEVIES"),
    norm("AMORTISATION AND PROVISIONS"),
    norm("RESULTS NON-CURRENT ASSETS"), norm("INVESTMENT RESULTS"),
    norm("FINANCIAL EXPENSES"), norm("CORPORATE TAX"),
    norm("MINORITY SHAREHOLDERS"),
}

RANKING_LINE_ALIASES = {
    norm("REVENUES"): norm("REVENUE"),
    norm("PROCUREMENT"): norm("PROCUREMENTS"),
    norm("PERSONNEL"): norm("PERSONNALE"),
}


def ranking_table(rows):
    """Build a separate Top 5 ranking for the selected worksheet."""
    eligible = []
    for row in rows:
        label_key = norm(row["label"])
        label_key = RANKING_LINE_ALIASES.get(label_key, label_key)
        if (label_key in ALLOWED_RANKING_LINES
                and row["var"] is not None
                and not row["is_pct"]):
            eligible.append(row)

    top_five = sorted(
        eligible, key=lambda row: abs(row["var"]), reverse=True
    )[:5]
    if not top_five:
        return ('<div class="empty">No matching requested lines with numeric '
                'Current and Prior values were found in this worksheet.</div>')

    body = []
    for rank, row in enumerate(top_five, 1):
        if row["current"] > row["prior"]:
            message = "▲ Current is higher than Prior"
            css_class = "pos"
        elif row["current"] < row["prior"]:
            message = "▼ Prior is higher than Current"
            css_class = "neg"
        else:
            message = "● Current and Prior are equal"
            css_class = "neutral"
        body.append(f"""
        <tr>
            <td>{rank}</td>
            <td><strong>{html.escape(row["label"])}</strong></td>
            <td>{fmt(row["current"])}</td>
            <td>{fmt(row["prior"])}</td>
            <td class="{css_class}">{fmt(row["var"])}</td>
            <td class="{css_class}"><strong>{message}</strong></td>
        </tr>""")
    return f"""
    <div class="tw ranking-table-wrap">
      <table class="ranking-table">
        <thead><tr><th>Rank</th><th>Line</th><th>Current</th><th>Prior</th>
        <th>Variation</th><th>Which version is higher?</th></tr></thead>
        <tbody>{''.join(body)}</tbody>
      </table>
    </div>"""


def cards(rows):
    financial_rows = [
        row
        for row in rows
        if row["var"] is not None
        and not row["is_pct"]
    ]

    stable_rows = sum(
        1
        for row in rows
        if row["varpct"] is not None
        and abs(row["varpct"]) <= 0.004
    )

    largest_movement = max(
        financial_rows,
        key=lambda row: abs(row["var"]),
        default=None,
    )

    values = [
        (
            "Metrics",
            len(financial_rows),
            "Numeric Current and Prior values",
        ),
        (
            "Increases",
            sum(
                row["var"] > 0
                for row in financial_rows
            ),
            "Positive absolute variance",
        ),
        (
            "Decreases",
            sum(
                row["var"] < 0
                for row in financial_rows
            ),
            "Negative absolute variance",
        ),
        (
            "Within ±0.4%",
            stable_rows,
            "PMR green indicator",
        ),
    ]

    if largest_movement:
        values.append(
            (
                "Largest Movement",
                largest_movement["label"],
                fmt(largest_movement["var"]),
            )
        )

    return "".join(
        (
            '<div class="kpi">'
            f"<small>{html.escape(str(title))}</small>"
            f"<b>{html.escape(str(value))}</b>"
            f"<span>{html.escape(str(description))}</span>"
            "</div>"
        )
        for title, value, description in values
    )


def table(rows, sheet_name):
    """
    Creates one HTML table for each Current/Prior block
    detected in the Excel worksheet.
    """

    tables = {}

    for row in rows:
        table_key = row["table_key"]

        if table_key not in tables:
            tables[table_key] = {
                "section": row["section"],
                "rows": [],
            }

        tables[table_key]["rows"].append(row)

    tables_html = []

    for table_number, table_data in enumerate(
        tables.values(),
        start=1,
    ):
        section_name = table_data["section"]
        section_rows = table_data["rows"]

        table_id = sid(
            f"{sheet_name}-{table_number}-"
            f"{section_name}"
        )

        table_rows_html = []

        for row in section_rows:
            if row["varpct"] is None:
                dot = "gray"
            elif abs(row["varpct"]) <= 0.004:
                dot = "green"
            else:
                dot = "red"

            search_value = html.escape(
                (
                    row["section"]
                    + " "
                    + row["label"]
                ).lower(),
                quote=True,
            )

            row_class = (
                "key-row"
                if row.get("is_bold", False)
                else ""
            )
            is_material = (
                row["varpct"] is not None
                and abs(row["varpct"]) > 0.004
            )

            label_html = html.escape(
                row["label"]
            )

            if row.get("is_bold", False):
                label_html = (
                    f"<strong>{label_html}</strong>"
                )

            table_rows_html.append(
                f"""
                <tr
                    class="{row_class}"
                    data-search="{search_value}"
                    data-material="{str(is_material).lower()}"
                >
                    <td>{label_html}</td>
                    <td>
                        {fmt(
                            row["current"],
                            row["is_pct"],
                        )}
                    </td>
                    <td>
                        {fmt(
                            row["prior"],
                            row["is_pct"],
                        )}
                    </td>
                    <td>
                        {fmt(
                            row["var"],
                            row["is_pct"],
                        )}
                    </td>
                    <td>
                        {fmt(row["varpct"], True)}
                    </td>
                    <td>
                        <i class="dot {dot}"></i>
                    </td>
                </tr>
                """
            )

        tables_html.append(
            f"""
            <div class="excel-table-block">
                <div class="excel-table-title">
                    <div>
                        <h3>
                            {html.escape(section_name)}
                        </h3>

                        <small>
                            {len(section_rows)} metrics
                        </small>
                    </div>

                    <input
                        class="table-search"
                        data-table="{table_id}"
                        placeholder="Search this table..."
                    >
                </div>

                <div class="tw">
                    <table id="{table_id}">
                        <thead>
                            <tr>
                                <th>Metric</th>
                                <th>Current</th>
                                <th>Prior</th>
                                <th>Var K$</th>
                                <th>Var %</th>
                                <th>Indicator</th>
                            </tr>
                        </thead>

                        <tbody>
                            {''.join(table_rows_html)}
                        </tbody>
                    </table>
                </div>
            </div>
            """
        )

    return "".join(tables_html)


def build(path, sheets, data):
    available_sheets = [
        sheet_name
        for sheet_name in sheets
        if data.get(sheet_name)
    ]

    if not available_sheets:
        raise ValueError(
            "No Current/Prior sections with numeric "
            "data were detected."
        )

    worksheet_options = "".join(
        (
            f'<option value="{sid(sheet_name)}">'
            f"{html.escape(sheet_name)}"
            "</option>"
        )
        for sheet_name in available_sheets
    )

    panels = []
    all_rows = []

    for sheet_index, sheet_name in enumerate(
        available_sheets
    ):
        sheet_rows = data[sheet_name]
        all_rows.extend(sheet_rows)

        active_class = (
            "active"
            if sheet_index == 0
            else ""
        )

        panels.append(
            f"""
            <section
                id="{sid(sheet_name)}"
                class="sheet {active_class}"
            >
                <h2>
                    {html.escape(sheet_name)}
                </h2>

                <p>
                    {len(sheet_rows)} metrics detected
                </p>

                <div class="kpis">
                    {cards(sheet_rows)}
                </div>

                <div class="panel ranking-panel">
                    <h3>Top 5 Changes for This Worksheet</h3>
                    <p class="panel-subtitle">
                        Ranking by absolute variation, limited to the specified PMR lines.
                        Positive means Current is higher; negative means Prior is higher.
                    </p>
                    {ranking_table(sheet_rows)}
                </div>

                <div class="tables-heading">
                    <div>
                        <h2>Detailed Tables</h2>

                        <p>
                            Each table follows a separate
                            Current/Prior block from the
                            Excel workbook.
                        </p>
                    </div>

                    <input
                        class="sheet-search"
                        data-target="{sid(sheet_name)}"
                        placeholder="Search all tables..."
                    >
                </div>

                <div class="tables-container">
                    {table(sheet_rows, sheet_name)}
                </div>
            </section>
            """
        )

    css = """
    * {
        box-sizing: border-box;
    }

    body {
        margin: 0;
        background: #F4F7FA;
        color: #172B4D;
        font-family: Segoe UI, Arial, sans-serif;
    }

    header {
        padding: 27px 32px;
        color: white;
        background:
            linear-gradient(
                120deg,
                #12344D,
                #184F75
            );
    }

    header h1 {
        margin: 0 0 6px;
    }

    header p {
        margin: 0;
        opacity: 0.9;
    }

    .toolbar {
        position: sticky;
        top: 0;
        z-index: 9;
        padding: 13px 30px;
        background: rgba(255, 255, 255, 0.96);
        border-bottom: 1px solid #D9E2EC;
    }

    select,
    input {
        padding: 10px 12px;
        border: 1px solid #B8C5D1;
        border-radius: 8px;
        background: white;
        color: #172B4D;
        font: inherit;
    }

    select {
        min-width: 330px;
        margin-left: 12px;
    }

    .material-filter {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        margin-left: 22px;
        color: #29445B;
        font-weight: 600;
    }

    .material-filter input {
        width: 17px;
        height: 17px;
        margin: 0;
    }

    main {
        max-width: 1540px;
        margin: auto;
        padding: 24px 30px;
    }

    .sheet {
        display: none;
    }

    .sheet.active {
        display: block;
    }

    .sheet h2 {
        margin-bottom: 2px;
    }

    .sheet > p,
    .tables-heading p,
    .panel-subtitle {
        margin-top: 4px;
        color: #68768A;
    }

    .kpis {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 14px;
        margin: 18px 0;
    }

    .kpi,
    .panel,
    .global,
    .excel-table-block {
        background: white;
        border: 1px solid #E4EAF0;
        border-radius: 12px;
        box-shadow:
            0 3px 12px
            rgba(23, 43, 77, 0.06);
    }

    .kpi {
        padding: 16px;
        border-top: 3px solid #0072CE;
    }

    .kpi small,
    .kpi span {
        display: block;
        color: #68768A;
    }

    .kpi b {
        display: block;
        margin: 7px 0;
        overflow: hidden;
        font-size: 23px;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .panel {
        padding: 18px;
        overflow: hidden;
    }

    .panel h3 {
        margin: 0;
    }

    .ranking-panel {
        margin-bottom: 24px;
    }

    .global {
        padding: 16px 20px;
        margin-bottom: 20px;
        border-left: 4px solid #0072CE;
    }

    .global h3 {
        margin-top: 0;
    }

    .global ol {
        columns: 2;
    }

    .global li {
        break-inside: avoid;
        min-height: 48px;
        padding: 5px;
    }

    .global li span {
        display: inline-flex;
        flex-direction: column;
        max-width: 70%;
    }

    .global small {
        color: #68768A;
    }

    .global strong {
        float: right;
    }

    .pos {
        color: #107C10;
        font-weight: 700;
    }

    .neg {
        color: #C50F1F;
        font-weight: 700;
    }
    .neutral { color: #68768A; }
    .ranking-table-wrap {
        margin-top: 14px;
        max-height: none;
        border: 1px solid #E4EAF0;
        border-radius: 9px;
    }
    .ranking-table th:first-child,
    .ranking-table td:first-child { width: 60px; text-align: center; }
    .ranking-table th:last-child,
    .ranking-table td:last-child { text-align: left; }

    .tables-heading {
        display: flex;
        align-items: flex-end;
        justify-content: space-between;
        gap: 20px;
        margin: 28px 0 14px;
    }

    .sheet-search {
        min-width: 280px;
    }

    .excel-table-block {
        margin-bottom: 22px;
        overflow: hidden;
    }

    .excel-table-title {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 20px;
        padding: 15px 18px;
        color: white;
        background:
            linear-gradient(
                90deg,
                #12344D,
                #18577F
            );
    }

    .excel-table-title h3 {
        margin: 0;
        font-size: 17px;
    }

    .excel-table-title small {
        display: block;
        margin-top: 3px;
        opacity: 0.8;
    }

    .table-search {
        min-width: 240px;
        background: white;
    }

    .tw {
        max-height: 620px;
        overflow: auto;
    }

    table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
    }

    th {
        position: sticky;
        top: 0;
        z-index: 2;
        padding: 11px;
        background: #EDF3F7;
        color: #29445B;
        border-bottom: 1px solid #CBD6DF;
        text-align: left;
    }

    td {
        padding: 9px 11px;
        border-top: 1px solid #E8EDF2;
    }

    tbody tr:hover {
        background: #F3F9FD;
    }

    td:nth-child(n+2),
    th:nth-child(n+2) {
        text-align: right;
        font-variant-numeric: tabular-nums;
    }

    td:last-child,
    th:last-child {
        text-align: center;
    }

    .key-row {
        background: #EEF6FC;
        font-weight: 700;
        border-top: 2px solid #B8D8EC;
        border-bottom: 2px solid #B8D8EC;
    }

    .key-row:hover {
        background: #E1F0FA;
    }

    .dot {
        display: inline-block;
        width: 10px;
        height: 10px;
        border-radius: 50%;
    }

    .green {
        background: #2E8540;
    }

    .red {
        background: #D13438;
    }

    .gray {
        background: #CBD5DF;
    }

    .empty {
        padding: 45px;
        color: #68768A;
        text-align: center;
    }

    @media (max-width: 1050px) {
        .kpis {
            grid-template-columns:
                repeat(2, 1fr);
        }

        .global ol {
            columns: 1;
        }
    }

    @media (max-width: 700px) {
        main {
            padding: 18px 12px;
        }

        .toolbar {
            padding: 12px;
        }

        select {
            width: 100%;
            min-width: 0;
            margin-top: 8px;
            margin-left: 0;
        }

        .kpis {
            grid-template-columns: 1fr;
        }

        .tables-heading,
        .excel-table-title {
            display: block;
        }

        .sheet-search,
        .table-search {
            width: 100%;
            min-width: 0;
            margin-top: 12px;
        }
    }
    """

    javascript = """
    const selector = document.getElementById("sel");

    selector.addEventListener("change", () => {
        document
            .querySelectorAll(".sheet")
            .forEach(sheet => {
                sheet.classList.remove("active");
            });

        const selectedSheet =
            document.getElementById(
                selector.value
            );

        if (selectedSheet) {
            selectedSheet.classList.add(
                "active"
            );

            window.dispatchEvent(
                new Event("resize")
            );
        }
    });

    const materialOnly = document.getElementById("materialOnly");

    function applyMaterialFilter() {
        const activeSheet = document.querySelector(".sheet.active");
        if (!activeSheet) return;
        activeSheet.querySelectorAll("tbody tr[data-material]").forEach(row => {
            const show = !materialOnly.checked || row.dataset.material === "true";
            row.style.display = show ? "" : "none";
        });
    }

    materialOnly.addEventListener("change", applyMaterialFilter);
    selector.addEventListener("change", () => setTimeout(applyMaterialFilter, 0));

    document
        .querySelectorAll(".sheet-search")
        .forEach(input => {
            input.addEventListener(
                "input",
                () => {
                    const query =
                        input.value
                            .toLowerCase()
                            .trim();

                    const sheet =
                        document.getElementById(
                            input.dataset.target
                        );

                    sheet
                        .querySelectorAll(
                            "tbody tr"
                        )
                        .forEach(row => {
                            row.style.display =
                                row.dataset.search
                                    .includes(query)
                                ? ""
                                : "none";
                        });
                }
            );
        });

    document
        .querySelectorAll(".table-search")
        .forEach(input => {
            input.addEventListener(
                "input",
                () => {
                    const query =
                        input.value
                            .toLowerCase()
                            .trim();

                    const selectedTable =
                        document.getElementById(
                            input.dataset.table
                        );

                    if (!selectedTable) {
                        return;
                    }

                    selectedTable
                        .querySelectorAll(
                            "tbody tr"
                        )
                        .forEach(row => {
                            row.style.display =
                                row.dataset.search
                                    .includes(query)
                                ? ""
                                : "none";
                        });
                }
            );
        });
    """

    return f"""
    <!doctype html>

    <html lang="en">
        <head>
            <meta charset="utf-8">

            <meta
                name="viewport"
                content="width=device-width, initial-scale=1"
            >

            <title>
                AVANGRID PMR Executive Dashboard
            </title>

            <style>
                {css}
            </style>
        </head>

        <body>
            <header>
                <h1>
                    AVANGRID PMR Executive Dashboard
                </h1>

                <p>
                    Source:
                    {html.escape(path.name)}
                    ·
                    Interactive PMR variance analysis
                </p>
            </header>

            <div class="toolbar">
                <b>Worksheet</b>

                <select id="sel">
                    {worksheet_options}
                </select>

                <label class="material-filter">
                    <input id="materialOnly" type="checkbox">
                    Show only variances outside ±0.4%
                </label>
            </div>

            <main>

                {''.join(panels)}
            </main>

            <script>
                {javascript}
            </script>
        </body>
    </html>
    """


def streamlit_app():
    """Display the dashboard and allow the Excel source to be replaced."""
    st.set_page_config(
        page_title="AVANGRID PMR Executive Dashboard",
        page_icon="📊",
        layout="wide",
    )

    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1.2rem;
            padding-left: 1.5rem;
            padding-right: 1.5rem;
            max-width: 100%;
        }
        .avangrid-brand {
            display: flex;
            align-items: center;
            gap: 18px;
            padding: 18px 22px;
            margin-bottom: 16px;
            color: white;
            background: linear-gradient(120deg, #12344D, #18577F);
            border-radius: 14px;
            box-shadow: 0 4px 16px rgba(23, 43, 77, 0.12);
        }
        .avangrid-wordmark {
            min-width: 150px;
            font-size: 22px;
            font-weight: 800;
            letter-spacing: 1.6px;
            color: #7AC143;
        }
        .avangrid-brand h1 {
            margin: 0;
            font-size: 34px;
        }
        .avangrid-brand p {
            margin: 4px 0 0;
            opacity: 0.88;
        }
        div[data-testid="stFileUploader"] {
            background: #ffffff;
            border: 1px solid #d9e2ec;
            border-radius: 12px;
            padding: 0.8rem 1rem;
        }
        div[data-testid="stMetric"] {
            background: white;
            border: 1px solid #E4EAF0;
            border-top: 3px solid #0072CE;
            border-radius: 12px;
            padding: 14px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    logo_path = Path("avangrid_logo.png")
    if logo_path.exists():
        left, right = st.columns([1, 6], vertical_alignment="center")
        with left:
            st.image(str(logo_path), use_container_width=True)
        with right:
            st.title("PMR Executive Dashboard")
            st.caption(
                "Upload a PMR comparison workbook and review executive variance insights."
            )
    else:
        st.markdown(
            """
            <div class="avangrid-brand">
                <div class="avangrid-wordmark">AVANGRID</div>
                <div>
                    <h1>PMR Executive Dashboard</h1>
                    <p>Upload a PMR comparison workbook and review executive variance insights.</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with st.sidebar:
        st.header("PMR Navigation")
        st.caption("Upload a workbook, then use the worksheet selector inside the dashboard.")
        st.markdown("**Materiality rule**")
        st.write("Green: variance within ±0.4%")
        st.write("Red: variance outside ±0.4%")
        st.divider()
        st.caption(
            "Files are processed for the current session. The app does not intentionally save "
            "uploaded workbooks to the repository."
        )

    uploaded_file = st.file_uploader(
        "Choose PMR comparison file",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        key="pmr_excel_file",
    )

    if uploaded_file is None:
        st.info("Select an .xlsx or .xlsm file to generate the dashboard.")
        return

    suffix = Path(uploaded_file.name).suffix.lower()
    temporary_path = None

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temporary_file:
            temporary_file.write(uploaded_file.getbuffer())
            temporary_path = Path(temporary_file.name)

        sheets, data = read_excel(temporary_path)
        dashboard_html = build(Path(uploaded_file.name), sheets, data)
        available_sheets = [sheet for sheet in sheets if data.get(sheet)]
        all_rows = [row for sheet in available_sheets for row in data[sheet]]
        numeric_rows = [
            row for row in all_rows
            if row["var"] is not None and not row["is_pct"]
        ]
        material_rows = [
            row for row in all_rows
            if row["varpct"] is not None and abs(row["varpct"]) > 0.004
        ]

        st.success(f"Dashboard updated from: {uploaded_file.name}")

        metric_1, metric_2, metric_3, metric_4 = st.columns(4)
        metric_1.metric("Loaded workbook", uploaded_file.name)
        metric_2.metric("Worksheets detected", len(available_sheets))
        metric_3.metric("Numeric metrics", len(numeric_rows))
        metric_4.metric("Outside ±0.4%", len(material_rows))

        download_1, download_2 = st.columns(2)
        with download_1:
            st.download_button(
                "Download original workbook",
                data=uploaded_file.getvalue(),
                file_name=uploaded_file.name,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )
        with download_2:
            st.download_button(
                "Download dashboard HTML",
                data=dashboard_html.encode("utf-8"),
                file_name=f"{Path(uploaded_file.name).stem} - Dashboard.html",
                mime="text/html",
                use_container_width=True,
            )

        st.caption(
            "Inside the dashboard, use the worksheet dropdown and the "
            "'Show only variances outside ±0.4%' filter."
        )
        components.html(dashboard_html, height=3600, scrolling=True)

    except Exception as error:
        st.error("The dashboard could not be created from the selected file.")
        st.exception(error)

    finally:
        if temporary_path and temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                pass


def desktop_main():
    """Keep direct Python execution available as a fallback."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected_file = filedialog.askopenfilename(
            title="Select the PMR Excel file",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("Excel Macro Files", "*.xlsm"),
                ("All Files", "*.*"),
            ],
        )
        root.destroy()
    except Exception as error:
        print(f"The file-selection window could not be opened: {error}")
        return 1

    if not selected_file:
        print("No Excel file was selected.")
        return 1

    source_file = Path(selected_file)
    output_file = source_file.with_name(
        source_file.stem + " - Dashboard.html"
    )

    try:
        sheets, data = read_excel(source_file)
        dashboard_html = build(source_file, sheets, data)
        output_file.write_text(dashboard_html, encoding="utf-8")
    except Exception as error:
        print(f"The dashboard could not be created: {error}")
        return 1

    print(f"Dashboard created: {output_file}")
    return 0


if __name__ == "__main__":
    # Streamlit sets this environment while running the script. Calling the
    # Streamlit app is also harmless when the script is launched by Streamlit.
    streamlit_app()
