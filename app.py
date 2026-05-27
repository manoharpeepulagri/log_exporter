import streamlit as st
import pandas as pd
import re
import io
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
from collections import defaultdict

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="CAN Log Analyzer",
    page_icon="🚌",
    layout="wide",
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1a1a2e;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 0.95rem;
        color: #6c757d;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 1rem 1.2rem;
        color: white;
        text-align: center;
    }
    .metric-value { font-size: 1.8rem; font-weight: 700; }
    .metric-label { font-size: 0.8rem; opacity: 0.85; margin-top: 2px; }
    .stDataFrame { border-radius: 8px; }
    div[data-testid="stMetricValue"] { font-size: 1.6rem; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# PARSE FUNCTION
# ─────────────────────────────────────────────
def parse_can_log(content: str) -> tuple[dict, pd.DataFrame]:
    """Parse BUSMASTER CAN log into metadata + dataframe."""
    lines = content.splitlines()
    meta = {}
    data_rows = []

    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith("***"):
            # Extract metadata
            m = re.match(r"\*\*\*START DATE AND TIME (.+?)\*\*\*", line)
            if m:
                meta["Session Date/Time"] = m.group(1).strip()
            m = re.match(r"\*\*\*PROTOCOL (.+?)\*\*\*", line)
            if m:
                meta["Protocol"] = m.group(1).strip()
            m = re.match(r"\*\*\*(.+?)- (\d+ bps)\*\*\*", line)
            if m:
                meta["Baud Rate"] = m.group(2)
                meta["Interface"] = m.group(1).split(",")[0].strip().replace("CHANNEL 1 - ", "")
            m = re.match(r"\*\*\*(HEX|DEC)\*\*\*", line)
            if m:
                meta["Display Format"] = m.group(1)
            m = re.match(r"\*\*\*(SYSTEM|BUS) MODE\*\*\*", line)
            if m:
                meta["Mode"] = m.group(1) + " MODE"
            continue

        # Data line: Time Tx/Rx Channel CAN_ID Type DLC [bytes...]
        parts = line.split()
        if len(parts) < 6:
            continue

        timestamp_str = parts[0]
        direction     = parts[1]
        channel       = parts[2]
        can_id        = parts[3]
        frame_type    = parts[4]

        # DLC might be int or absent for remote frames
        try:
            dlc = int(parts[5])
        except (ValueError, IndexError):
            dlc = 0

        data_bytes = parts[6:]  # hex bytes (may be empty)

        # Determine frame subtype
        if "sr" in frame_type or (len(data_bytes) == 0 and dlc > 0):
            subtype = "Remote Frame"
        elif "x" in frame_type:
            subtype = "Extended Frame"
        else:
            subtype = "Standard Frame"

        # Build byte columns
        byte_dict = {}
        for i in range(8):
            byte_dict[f"Byte{i} (Hex)"] = data_bytes[i] if i < len(data_bytes) else ""

        # Combined data field
        data_hex = " ".join(data_bytes) if data_bytes else "-"

        # Parse timestamp to ms since start
        ts_match = re.match(r"(\d+):(\d+):(\d+):(\d+)", timestamp_str)
        if ts_match:
            h, m_t, s, ms = map(int, ts_match.groups())
            ts_ms = h*3600000 + m_t*60000 + s*1000 + ms
        else:
            ts_ms = None

        data_rows.append({
            "Timestamp"       : timestamp_str,
            "Timestamp (ms)"  : ts_ms,
            "Direction"       : direction,
            "Channel"         : channel,
            "CAN ID"          : can_id,
            "Frame Type"      : subtype,
            "DLC"             : dlc,
            "Data (Hex)"      : data_hex,
            **byte_dict,
        })

    df = pd.DataFrame(data_rows)
    return meta, df


# ─────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────
def build_excel(meta: dict, df: pd.DataFrame, filename: str) -> bytes:
    wb = Workbook()

    # ── colour palette ──
    CLR_HEADER_BG  = "1A1A2E"
    CLR_HEADER_FG  = "FFFFFF"
    CLR_TITLE_BG   = "16213E"
    CLR_META_BG    = "0F3460"
    CLR_META_FG    = "E0E0E0"
    CLR_ALT_ROW    = "F0F4FF"
    CLR_TX_BG      = "E8F5E9"   # green tint
    CLR_RX_BG      = "E3F2FD"   # blue tint
    CLR_ACCENT     = "764BA2"
    CLR_BORDER     = "CCCCCC"

    thin  = Side(style="thin",   color=CLR_BORDER)
    thick = Side(style="medium", color="888888")
    thin_border  = Border(left=thin, right=thin, top=thin, bottom=thin)
    thick_border = Border(left=thick, right=thick, top=thick, bottom=thick)

    def hdr_font(sz=11, bold=True, color=CLR_HEADER_FG):
        return Font(name="Arial", size=sz, bold=bold, color=color)

    def body_font(sz=10, bold=False, color="1A1A2E"):
        return Font(name="Arial", size=sz, bold=bold, color=color)

    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def center():
        return Alignment(horizontal="center", vertical="center", wrap_text=False)

    def left():
        return Alignment(horizontal="left", vertical="center")

    # ═══════════════════════════════════════
    # SHEET 1 – SUMMARY
    # ═══════════════════════════════════════
    ws_s = wb.active
    ws_s.title = "📊 Summary"
    ws_s.sheet_view.showGridLines = False
    ws_s.column_dimensions["A"].width = 30
    ws_s.column_dimensions["B"].width = 40

    # Title block
    ws_s.merge_cells("A1:B1")
    ws_s["A1"] = "🚌  CAN Bus Log — Analysis Report"
    ws_s["A1"].font   = Font(name="Arial", size=16, bold=True, color=CLR_HEADER_FG)
    ws_s["A1"].fill   = fill(CLR_TITLE_BG)
    ws_s["A1"].alignment = center()
    ws_s.row_dimensions[1].height = 36

    ws_s.merge_cells("A2:B2")
    ws_s["A2"] = f"Source file: {filename}"
    ws_s["A2"].font      = Font(name="Arial", size=10, italic=True, color=CLR_META_FG)
    ws_s["A2"].fill      = fill(CLR_META_BG)
    ws_s["A2"].alignment = center()

    # Metadata section
    row = 4
    ws_s.merge_cells(f"A{row}:B{row}")
    ws_s[f"A{row}"] = "SESSION METADATA"
    ws_s[f"A{row}"].font   = hdr_font(10)
    ws_s[f"A{row}"].fill   = fill(CLR_HEADER_BG)
    ws_s[f"A{row}"].alignment = center()
    row += 1

    for k, v in meta.items():
        ws_s[f"A{row}"] = k
        ws_s[f"B{row}"] = v
        ws_s[f"A{row}"].font      = body_font(bold=True, color="333333")
        ws_s[f"B{row}"].font      = body_font(color="1A1A2E")
        ws_s[f"A{row}"].fill      = fill("EEF0F8")
        ws_s[f"B{row}"].fill      = fill("F8F9FF")
        ws_s[f"A{row}"].alignment = left()
        ws_s[f"B{row}"].alignment = left()
        for c in ["A", "B"]:
            ws_s[f"{c}{row}"].border = thin_border
        row += 1

    # Stats section
    row += 1
    ws_s.merge_cells(f"A{row}:B{row}")
    ws_s[f"A{row}"] = "LOG STATISTICS"
    ws_s[f"A{row}"].font      = hdr_font(10)
    ws_s[f"A{row}"].fill      = fill(CLR_HEADER_BG)
    ws_s[f"A{row}"].alignment = center()
    row += 1

    tx_df = df[df["Direction"] == "Tx"]
    rx_df = df[df["Direction"] == "Rx"]
    id_counts = df["CAN ID"].value_counts()

    stats = [
        ("Total Messages",       len(df)),
        ("Transmitted (Tx)",     len(tx_df)),
        ("Received (Rx)",        len(rx_df)),
        ("Unique CAN IDs",       df["CAN ID"].nunique()),
        ("Standard Frames",      len(df[df["Frame Type"] == "Standard Frame"])),
        ("Extended Frames",      len(df[df["Frame Type"] == "Extended Frame"])),
        ("Remote Frames",        len(df[df["Frame Type"] == "Remote Frame"])),
        ("Most Active ID",       id_counts.index[0] if len(id_counts) else "-"),
        ("Most Active ID Count", id_counts.iloc[0] if len(id_counts) else 0),
    ]

    if not df.empty and df["Timestamp (ms)"].notna().any():
        ts_vals = df["Timestamp (ms)"].dropna()
        duration_s = (ts_vals.max() - ts_vals.min()) / 1000
        stats.append(("Log Duration (s)", round(duration_s, 3)))
        if duration_s > 0:
            stats.append(("Avg Msg Rate (msg/s)", round(len(df) / duration_s, 1)))

    for label, value in stats:
        ws_s[f"A{row}"] = label
        ws_s[f"B{row}"] = value
        ws_s[f"A{row}"].font      = body_font(bold=True, color="333333")
        ws_s[f"B{row}"].font      = body_font(color="1A1A2E")
        ws_s[f"A{row}"].fill      = fill("EEF0F8")
        ws_s[f"B{row}"].fill      = fill("F8F9FF")
        ws_s[f"A{row}"].alignment = left()
        ws_s[f"B{row}"].alignment = Alignment(horizontal="right", vertical="center")
        for c in ["A", "B"]:
            ws_s[f"{c}{row}"].border = thin_border
        row += 1

    # ═══════════════════════════════════════
    # SHEET 2 – ALL MESSAGES
    # ═══════════════════════════════════════
    ws_m = wb.create_sheet("📋 All Messages")
    ws_m.sheet_view.showGridLines = False
    ws_m.freeze_panes = "A3"

    cols = [
        ("Timestamp",      18),
        ("Direction",      10),
        ("Channel",        9),
        ("CAN ID",         14),
        ("Frame Type",     18),
        ("DLC",             6),
        ("Data (Hex)",     28),
        ("Byte0 (Hex)",     9),
        ("Byte1 (Hex)",     9),
        ("Byte2 (Hex)",     9),
        ("Byte3 (Hex)",     9),
        ("Byte4 (Hex)",     9),
        ("Byte5 (Hex)",     9),
        ("Byte6 (Hex)",     9),
        ("Byte7 (Hex)",     9),
    ]

    # Merged title row
    ws_m.merge_cells(f"A1:{get_column_letter(len(cols))}1")
    ws_m["A1"] = "📋  CAN Bus — All Messages"
    ws_m["A1"].font      = Font(name="Arial", size=13, bold=True, color=CLR_HEADER_FG)
    ws_m["A1"].fill      = fill(CLR_TITLE_BG)
    ws_m["A1"].alignment = center()
    ws_m.row_dimensions[1].height = 28

    # Header row
    for ci, (col_name, col_w) in enumerate(cols, start=1):
        cell = ws_m.cell(row=2, column=ci, value=col_name)
        cell.font      = hdr_font(10)
        cell.fill      = fill(CLR_HEADER_BG)
        cell.alignment = center()
        cell.border    = thin_border
        ws_m.column_dimensions[get_column_letter(ci)].width = col_w
    ws_m.row_dimensions[2].height = 22

    # Data rows
    col_map = {name: i+1 for i, (name, _) in enumerate(cols)}
    for ri, row_data in enumerate(df.itertuples(index=False), start=3):
        direction = row_data.Direction
        if direction == "Tx":
            bg = CLR_TX_BG
        elif direction == "Rx":
            bg = CLR_RX_BG
        else:
            bg = CLR_ALT_ROW if ri % 2 == 0 else "FFFFFF"

        for col_name, _ in cols:
            ci = col_map[col_name]
            val = getattr(row_data, col_name.replace(" ", "_").replace("(", "").replace(")", ""), "")
            cell = ws_m.cell(row=ri, column=ci, value=val)
            cell.font      = body_font(9)
            cell.fill      = fill(bg)
            cell.alignment = center() if col_name not in ("Timestamp", "Data (Hex)") else left()
            cell.border    = thin_border

    # Auto-filter on header
    ws_m.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{len(df)+2}"

    # ═══════════════════════════════════════
    # SHEET 3 – CAN ID BREAKDOWN
    # ═══════════════════════════════════════
    ws_b = wb.create_sheet("🔍 CAN ID Breakdown")
    ws_b.sheet_view.showGridLines = False
    ws_b.freeze_panes = "A3"

    breakdown_cols = [
        ("CAN ID",          14),
        ("Total Messages",  16),
        ("Tx Count",        12),
        ("Rx Count",        12),
        ("Frame Type",      18),
        ("% of Traffic",    14),
        ("Min DLC",         10),
        ("Max DLC",         10),
        ("First Seen",      18),
        ("Last Seen",       18),
        ("Unique Data Payloads", 22),
    ]

    ws_b.merge_cells(f"A1:{get_column_letter(len(breakdown_cols))}1")
    ws_b["A1"] = "🔍  CAN ID Traffic Breakdown"
    ws_b["A1"].font      = Font(name="Arial", size=13, bold=True, color=CLR_HEADER_FG)
    ws_b["A1"].fill      = fill(CLR_TITLE_BG)
    ws_b["A1"].alignment = center()
    ws_b.row_dimensions[1].height = 28

    for ci, (col_name, col_w) in enumerate(breakdown_cols, start=1):
        cell = ws_b.cell(row=2, column=ci, value=col_name)
        cell.font      = hdr_font(10)
        cell.fill      = fill(CLR_HEADER_BG)
        cell.alignment = center()
        cell.border    = thin_border
        ws_b.column_dimensions[get_column_letter(ci)].width = col_w
    ws_b.row_dimensions[2].height = 22

    total_msgs = len(df)
    for ri, (can_id, grp) in enumerate(df.groupby("CAN ID"), start=3):
        bg = CLR_ALT_ROW if ri % 2 == 0 else "FFFFFF"
        frame_types = grp["Frame Type"].unique()
        ft_str = ", ".join(frame_types)
        pct = round(len(grp) / total_msgs * 100, 2) if total_msgs else 0
        unique_payloads = grp["Data (Hex)"].nunique()
        first_seen = grp["Timestamp"].iloc[0] if not grp.empty else "-"
        last_seen  = grp["Timestamp"].iloc[-1] if not grp.empty else "-"

        row_vals = [
            can_id,
            len(grp),
            len(grp[grp["Direction"] == "Tx"]),
            len(grp[grp["Direction"] == "Rx"]),
            ft_str,
            f"{pct}%",
            grp["DLC"].min(),
            grp["DLC"].max(),
            first_seen,
            last_seen,
            unique_payloads,
        ]
        for ci, val in enumerate(row_vals, start=1):
            cell = ws_b.cell(row=ri, column=ci, value=val)
            cell.font      = body_font(10)
            cell.fill      = fill(bg)
            cell.alignment = center()
            cell.border    = thin_border

    ws_b.auto_filter.ref = f"A2:{get_column_letter(len(breakdown_cols))}{len(df.groupby('CAN ID'))+2}"

    # Color-scale on "Total Messages" col (col B)
    last_row = len(df.groupby("CAN ID")) + 2
    if last_row > 3:
        ws_b.conditional_formatting.add(
            f"B3:B{last_row}",
            ColorScaleRule(
                start_type="min", start_color="FFFFFF",
                end_type="max",   end_color="764BA2"
            )
        )

    # ═══════════════════════════════════════
    # SHEET 4 – TX MESSAGES
    # ═══════════════════════════════════════
    ws_tx = wb.create_sheet("📤 Tx Messages")
    ws_tx.sheet_view.showGridLines = False
    ws_tx.freeze_panes = "A3"

    tx_cols = [c for c in cols if c[0] != "Direction"]
    ws_tx.merge_cells(f"A1:{get_column_letter(len(tx_cols))}1")
    ws_tx["A1"] = "📤  Transmitted (Tx) Messages"
    ws_tx["A1"].font      = Font(name="Arial", size=13, bold=True, color=CLR_HEADER_FG)
    ws_tx["A1"].fill      = fill("1B5E20")
    ws_tx["A1"].alignment = center()
    ws_tx.row_dimensions[1].height = 28

    for ci, (col_name, col_w) in enumerate(tx_cols, start=1):
        cell = ws_tx.cell(row=2, column=ci, value=col_name)
        cell.font      = hdr_font(10)
        cell.fill      = fill("2E7D32")
        cell.alignment = center()
        cell.border    = thin_border
        ws_tx.column_dimensions[get_column_letter(ci)].width = col_w
    ws_tx.row_dimensions[2].height = 22

    tx_filtered = df[df["Direction"] == "Tx"].reset_index(drop=True)
    tx_col_map  = {name: i+1 for i, (name, _) in enumerate(tx_cols)}
    for ri, row_data in enumerate(tx_filtered.itertuples(index=False), start=3):
        bg = CLR_TX_BG if ri % 2 == 0 else "F1F8E9"
        for col_name, _ in tx_cols:
            ci  = tx_col_map[col_name]
            val = getattr(row_data, col_name.replace(" ", "_").replace("(", "").replace(")", ""), "")
            cell = ws_tx.cell(row=ri, column=ci, value=val)
            cell.font      = body_font(9)
            cell.fill      = fill(bg)
            cell.alignment = center() if col_name not in ("Timestamp", "Data (Hex)") else left()
            cell.border    = thin_border

    # ═══════════════════════════════════════
    # SHEET 5 – RX MESSAGES
    # ═══════════════════════════════════════
    ws_rx = wb.create_sheet("📥 Rx Messages")
    ws_rx.sheet_view.showGridLines = False
    ws_rx.freeze_panes = "A3"

    rx_cols = [c for c in cols if c[0] != "Direction"]
    ws_rx.merge_cells(f"A1:{get_column_letter(len(rx_cols))}1")
    ws_rx["A1"] = "📥  Received (Rx) Messages"
    ws_rx["A1"].font      = Font(name="Arial", size=13, bold=True, color=CLR_HEADER_FG)
    ws_rx["A1"].fill      = fill("0D47A1")
    ws_rx["A1"].alignment = center()
    ws_rx.row_dimensions[1].height = 28

    for ci, (col_name, col_w) in enumerate(rx_cols, start=1):
        cell = ws_rx.cell(row=2, column=ci, value=col_name)
        cell.font      = hdr_font(10)
        cell.fill      = fill("1565C0")
        cell.alignment = center()
        cell.border    = thin_border
        ws_rx.column_dimensions[get_column_letter(ci)].width = col_w
    ws_rx.row_dimensions[2].height = 22

    rx_filtered = df[df["Direction"] == "Rx"].reset_index(drop=True)
    rx_col_map  = {name: i+1 for i, (name, _) in enumerate(rx_cols)}
    for ri, row_data in enumerate(rx_filtered.itertuples(index=False), start=3):
        bg = CLR_RX_BG if ri % 2 == 0 else "E1F5FE"
        for col_name, _ in rx_cols:
            ci  = rx_col_map[col_name]
            val = getattr(row_data, col_name.replace(" ", "_").replace("(", "").replace(")", ""), "")
            cell = ws_rx.cell(row=ri, column=ci, value=val)
            cell.font      = body_font(9)
            cell.fill      = fill(bg)
            cell.alignment = center() if col_name not in ("Timestamp", "Data (Hex)") else left()
            cell.border    = thin_border

    # ─── Save to bytes ───
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.markdown('<p class="main-header">🚌 CAN Bus Log Analyzer</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload a BUSMASTER .log file — view, filter, and download a structured Excel report.</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Drop your CAN log file here",
    type=["log", "txt"],
    help="Supports BUSMASTER .log format",
)

if uploaded_file:
    raw = uploaded_file.read().decode("utf-8", errors="replace")
    with st.spinner("Parsing log…"):
        meta, df = parse_can_log(raw)

    if df.empty:
        st.error("No valid CAN messages found in the file.")
        st.stop()

    fname = uploaded_file.name

    # ── TOP METRICS ──
    st.markdown("---")
    c1, c2, c3, c4, c5 = st.columns(5)
    metrics = [
        ("Total Messages",  len(df),                      "📨"),
        ("Tx Messages",     len(df[df["Direction"]=="Tx"]), "📤"),
        ("Rx Messages",     len(df[df["Direction"]=="Rx"]), "📥"),
        ("Unique CAN IDs",  df["CAN ID"].nunique(),         "🔑"),
        ("Frame Types",     df["Frame Type"].nunique(),     "🗂️"),
    ]
    for col, (label, value, icon) in zip([c1,c2,c3,c4,c5], metrics):
        col.metric(f"{icon} {label}", value)

    st.markdown("---")

    # ── SESSION METADATA ──
    with st.expander("📋 Session Metadata", expanded=True):
        mc1, mc2 = st.columns(2)
        meta_items = list(meta.items())
        half = len(meta_items)//2 + len(meta_items)%2
        with mc1:
            for k, v in meta_items[:half]:
                st.markdown(f"**{k}:** `{v}`")
        with mc2:
            for k, v in meta_items[half:]:
                st.markdown(f"**{k}:** `{v}`")

    # ── FILTERS ──
    st.subheader("🔍 Filter Messages")
    fc1, fc2, fc3 = st.columns(3)

    with fc1:
        direction_filter = st.multiselect(
            "Direction",
            options=["Tx", "Rx"],
            default=["Tx", "Rx"],
        )
    with fc2:
        id_options = sorted(df["CAN ID"].unique().tolist())
        id_filter = st.multiselect(
            "CAN ID",
            options=id_options,
            default=id_options,
        )
    with fc3:
        ft_options = sorted(df["Frame Type"].unique().tolist())
        ft_filter = st.multiselect(
            "Frame Type",
            options=ft_options,
            default=ft_options,
        )

    filtered_df = df[
        df["Direction"].isin(direction_filter) &
        df["CAN ID"].isin(id_filter) &
        df["Frame Type"].isin(ft_filter)
    ]

    st.caption(f"Showing **{len(filtered_df):,}** of **{len(df):,}** messages")

    # ── MESSAGE TABLE ──
    st.subheader("📋 Messages")
    display_cols = ["Timestamp", "Direction", "Channel", "CAN ID", "Frame Type", "DLC", "Data (Hex)"]

    # Colour Tx green, Rx blue
    def style_direction(val):
        if val == "Tx":
            return "background-color:#E8F5E9; color:#1B5E20; font-weight:bold"
        if val == "Rx":
            return "background-color:#E3F2FD; color:#0D47A1; font-weight:bold"
        return ""

    styled = (
        filtered_df[display_cols]
        .style
        .map(style_direction, subset=["Direction"])
        .set_properties(**{"font-size": "12px"})
    )
    st.dataframe(styled, use_container_width=True, height=420)

    # ── CAN ID SUMMARY ──
    st.subheader("📊 CAN ID Traffic Breakdown")
    breakdown = (
        df.groupby("CAN ID")
          .agg(
            Total=("CAN ID","count"),
            Tx=("Direction", lambda x: (x=="Tx").sum()),
            Rx=("Direction", lambda x: (x=="Rx").sum()),
            Frame_Type=("Frame Type", lambda x: ", ".join(x.unique())),
            Unique_Payloads=("Data (Hex)", "nunique"),
          )
          .reset_index()
          .sort_values("Total", ascending=False)
    )
    breakdown["% of Traffic"] = (breakdown["Total"] / len(df) * 100).round(2).astype(str) + "%"
    st.dataframe(breakdown, use_container_width=True, height=280)

    # ── CHARTS ──
    st.subheader("📈 Visual Overview")
    ch1, ch2 = st.columns(2)

    with ch1:
        st.markdown("**Messages per CAN ID**")
        chart_data = breakdown.set_index("CAN ID")["Total"]
        st.bar_chart(chart_data)

    with ch2:
        st.markdown("**Tx vs Rx Split**")
        tx_rx = pd.DataFrame({
            "Count": [len(df[df["Direction"]=="Tx"]), len(df[df["Direction"]=="Rx"])]
        }, index=["Tx", "Rx"])
        st.bar_chart(tx_rx)

    # ── DOWNLOAD ──
    st.markdown("---")
    st.subheader("⬇️ Download Excel Report")

    with st.spinner("Building Excel…"):
        excel_bytes = build_excel(meta, df, fname)

    out_name = fname.replace(".log", "").replace(".txt", "") + "_CAN_Analysis.xlsx"
    st.download_button(
        label="📥  Download Excel Report",
        data=excel_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
    st.caption("The Excel file contains 5 sheets: **Summary**, **All Messages**, **CAN ID Breakdown**, **Tx Messages**, and **Rx Messages** — each with filters, colour coding, and freeze-panes.")

else:
    st.info("👆 Upload a BUSMASTER CAN .log file to get started.", icon="ℹ️")
    st.markdown("""
    **What this tool does:**
    - Parses BUSMASTER `.log` files (CAN bus protocol)
    - Displays all messages with live filters by Direction, CAN ID, and Frame Type
    - Shows a per-CAN-ID traffic breakdown with Tx/Rx counts
    - Generates a fully formatted, multi-sheet Excel report ready to download
    """)