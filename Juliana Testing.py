import io
import re
import base64
from collections import defaultdict
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────────────────────
# Streamlit chrome
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="uKids — Kids Scheduler (Positions M/E)", layout="wide")
st.title("uKids — Kids Scheduler (Positions M/E)")

st.markdown(
    """
    <style>
      .stApp { background: #000; color: #fff; }
      .stButton>button, .stDownloadButton>button { background:#444; color:#fff; }
      .stDataFrame { background:#111; }
      .stAlert { color:#fff; }
      a, .stDownloadButton>button:hover, .stButton>button:hover { filter: brightness(1.1); }
      .hint { opacity: .8; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Optional logo (ignore if missing)
for logo_name in ["image(1).png", "image.png", "logo.png"]:
    try:
        with open(logo_name, "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
        st.markdown(
            f"<div style='text-align:center'><img src='data:image/png;base64,{encoded}' width='420'></div>",
            unsafe_allow_html=True,
        )
        break
    except Exception:
        pass

# ──────────────────────────────────────────────────────────────────────────────
# Helpers & parsing
# ──────────────────────────────────────────────────────────────────────────────
MONTH_ALIASES = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}
YES_SET = {"yes", "y", "true", "available"}

def normalize(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()

def read_csv_robust(uploaded_file, label_for_error):
    raw = uploaded_file.getvalue()
    encodings = ["utf-8", "utf-8-sig", "cp1252", "iso-8859-1"]
    seps = [None, ",", ";", "\t", "|"]
    last_err = None
    for enc in encodings:
        for sep in seps:
            try:
                df = pd.read_csv(io.BytesIO(raw), encoding=enc, engine="python", sep=sep)
                if df.shape[1] == 0:
                    raise ValueError("Parsed 0 columns.")
                return df
            except Exception as e:
                last_err = f"{type(e).__name__}: {e}"
                continue
    st.error(
        f"Could not read {label_for_error} CSV. Last error: {last_err}. "
        "Try re-exporting as CSV (UTF-8) or remove unusual characters in headers."
    )
    st.stop()

CAP_PATTERNS = [
    re.compile(r"^(?P<base>.*?)[\s\-]*\(\s*x?\s*(?P<n>\d+)\s*\)\s*$", re.IGNORECASE),
    re.compile(r"^(?P<base>.*?)[\s\-]*\[\s*x?\s*(?P<n>\d+)\s*\]\s*$", re.IGNORECASE),
    re.compile(r"^(?P<base>.*?)[\s\-]*x\s*(?P<n>\d+)\s*$", re.IGNORECASE),
]

def parse_role_meta(header: str):
    """Return (base_label, capacity:int). Accepts '(xN)', 'xN', or '[xN]'."""
    s = str(header).replace("\u00A0", " ").replace("×", "x").strip()
    base = s
    cap = None
    for pat in CAP_PATTERNS:
        m = pat.match(s)
        if m:
            base = str(m.group("base")).strip()
            cap = int(m.group("n"))
            break
    if cap is None:
        cap = 9999  # default large cap if omitted
    return base, cap

# ──────────────────────────────────────────────────────────────────────────────
# Date/session parsing from Kids Availability CSV
# ──────────────────────────────────────────────────────────────────────────────
def detect_date_session_headers(df: pd.DataFrame):
    """
    Parse columns like '5 Oct Morning' / '05-Oct Evening' / '19 Oct M' / '26 Oct E'.
    Returns:
      - date_map: {original_col: (pd.Timestamp(date), session_str)}
      - service_slots: sorted unique list of (date, session_str)
      - sheet_name: 'Month YYYY'
    """
    pat = re.compile(r"\b(\d{1,2})\s*[-/ ]\s*([A-Za-z]{3,})\b(?:\s+([A-Za-z]{1,20}))?", re.IGNORECASE)
    avail_cols = [c for c in df.columns if isinstance(c, str) and pat.search(c)]
    if not avail_cols:
        raise ValueError("No availability columns found. Use headers like '05-Oct Morning' or '5 Oct Evening'.")

    def norm_session(tok: str) -> str:
        t = (tok or "").strip().lower()
        if t in ("am", "morning", "m", "1st", "first", "service1", "service-1", "service"):
            return "Morning"
        if t in ("pm", "evening", "e", "2nd", "second", "service2", "service-2"):
            return "Evening"
        return "Service" if t == "" else tok.title()

    info, months_found = [], set()
    for c in avail_cols:
        m = pat.search(str(c))
        if not m:
            continue
        day = int(m.group(1))
        mon_txt = m.group(2).lower()[:3]
        month = MONTH_ALIASES.get(mon_txt)
        session = norm_session(m.group(3))
        months_found.add(month)
        info.append((c, month, day, session))

    if not months_found or None in months_found:
        raise ValueError("Could not parse month from availability headers.")
    if len(months_found) > 1:
        raise ValueError(f"Multiple months detected in availability headers: {sorted(months_found)}. Upload one month at a time.")
    month = months_found.pop()

    if "Timestamp" in df.columns:
        years = pd.to_datetime(df["Timestamp"], errors="coerce").dt.year.dropna().astype(int)
        year = int(years.mode().iloc[0]) if not years.empty else date.today().year
    else:
        year = date.today().year

    date_map, service_slots, seen = {}, [], set()
    for c, mnum, d, session in info:
        dt = pd.Timestamp(datetime(year, mnum, d)).normalize()
        slot = (dt, session)
        date_map[c] = slot
        if slot not in seen:
            seen.add(slot)
            service_slots.append(slot)

    service_slots.sort(key=lambda t: (t[0], t[1]))
    sheet_name = f"{pd.Timestamp(year=year, month=month, day=1):%B %Y}"
    return date_map, service_slots, sheet_name

# ──────────────────────────────────────────────────────────────────────────────
# UI — file inputs & settings
# ──────────────────────────────────────────────────────────────────────────────
st.subheader("1) Upload files")
c1, c2 = st.columns(2)
with c1:
    positions_file = st.file_uploader("Serving positions (CSV)", type=["csv"], key="positions_csv")
with c2:
    kids_file = st.file_uploader("Kids availability (CSV)", type=["csv"], key="kids_csv")

st.caption("• Positions CSV: one column per position+service. Examples: 'Kids Production M (x4)', 'Kids Production E (x3)', 'Check-in M (x6)'.")
st.caption("• Kids CSV: columns = Name (required), optional Campus, optional 'Evening Allowed' (Yes/No), then date columns like '5 Oct Morning' / '5 Oct Evening' (or '... M' / '... E').")

st.subheader("2) Rules")
max_serves = st.number_input("Max serves per kid (this month)", min_value=1, max_value=20, value=4, step=1)

# ──────────────────────────────────────────────────────────────────────────────
# Scheduling logic (positions + sessions)
# ──────────────────────────────────────────────────────────────────────────────
def excel_autofit(ws):
    for col_idx, column_cells in enumerate(
        ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1
    ):
        max_len = 0
        for cell in column_cells:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 80)

def parse_positions(df_positions: pd.DataFrame):
    """
    Headers like 'Kids Production M (x4)' / 'Kids Production E (x3)' / 'Check-in M'.
    Returns: list of (display_label, session, capacity) where session ∈ {"Morning","Evening","Both"}.
    """
    positions = []
    for col in df_positions.columns:
        base, cap = parse_role_meta(col)
        nb = normalize(base)
        # Determine session by suffix
        if nb.endswith(" m") or nb.endswith(" morning"):
            session = "Morning"
            label = base.rsplit(" ", 1)[0] if " " in base else base
        elif nb.endswith(" e") or nb.endswith(" evening"):
            session = "Evening"
            label = base.rsplit(" ", 1)[0] if " " in base else base
        else:
            session = "Both"
            label = base
        positions.append((label.strip(), session, int(cap)))
    positions = [(lbl, sess, cap) for (lbl, sess, cap) in positions if lbl]
    if not positions:
        st.error("No position columns detected. Use headers like 'Kids Production M (x4)' or 'Kids Production E (x3)'.")
        st.stop()
    return positions

def build_kids(kids_df: pd.DataFrame):
    # Required: Name. Optional: Campus, Evening Allowed
    cols = {normalize(c): c for c in kids_df.columns}
    name_col = cols.get("name") or list(kids_df.columns)[0]
    campus_col = cols.get("campus") or cols.get("location")
    evening_allowed_col = cols.get("evening allowed") or cols.get("evening_allowed") or cols.get("allowed evening")

    def truthy(x):
        return str(x).strip().lower() in YES_SET

    kids = []
    for _, r in kids_df.iterrows():
        name = str(r.get(name_col, "")).strip()
        if not name or str(name).lower() == "nan":
            continue
        campus = str(r.get(campus_col, "")).strip() if campus_col else ""
        evening_ok = truthy(r.get(evening_allowed_col, "yes")) if evening_allowed_col else True
        kids.append({"name": name, "campus": campus, "evening_ok": evening_ok})
    if not kids:
        st.error("No kids parsed from the CSV (check Name column).")
        st.stop()
    return kids

def parse_availability(kids_df: pd.DataFrame):
    date_map, service_slots, sheet_name = detect_date_session_headers(kids_df)

    # Availability dict: {name_norm: {(date, session): True/False}}
    avail = defaultdict(dict)

    cols = {normalize(c): c for c in kids_df.columns}
    name_col = cols.get("name") or list(kids_df.columns)[0]

    for _, row in kids_df.iterrows():
        name = str(row.get(name_col, "")).strip()
        if not name:
            continue
        key = normalize(name)
        for col, slot in date_map.items():
            ans = str(row.get(col, "")).strip().lower()
            avail[key][slot] = ans in YES_SET

    return avail, service_slots, sheet_name

def assign_kids(positions, kids, availability, service_slots, max_serves=4):
    """
    positions: list of (label, session, capacity)
    kids: list of {name, campus, evening_ok}
    availability: {name_norm: {(date, session): bool}}
    service_slots: list of (date, session)

    Rules:
      - Max total assignments per kid = max_serves
      - Never assign the same kid twice on the same calendar date (even across sessions)
      - Respect Evening Allowed
    Returns: schedule dict and unscheduled list per slot.
    """
    kids_by_key = {normalize(k["name"]): k for k in kids}

    # Expand 'Both' into explicit sessions for display and assignment
    expanded = []  # (display_label, session, capacity)
    for (label, sess, cap) in positions:
        if sess == "Both":
            expanded.append((f"{label} (Morning)", "Morning", cap))
            expanded.append((f"{label} (Evening)", "Evening", cap))
        else:
            expanded.append((f"{label} ({sess})", sess, cap))

    # (display_label, (date, session)) → [names]
    schedule = {(disp, slot): [] for (disp, sess, _c) in expanded for slot in service_slots if slot[1] == sess}
    unscheduled = {slot: [] for slot in service_slots}

    # Global bookkeeping to enforce per-kid caps and per-date uniqueness
    assigned_total = defaultdict(int)     # key -> total assignments so far
    assigned_on_date = set()              # {(key, date)} pairs

    # Deterministic order for fairness
    kid_keys_sorted = sorted(kids_by_key.keys())

    # Iterate slots chronologically; enforce daily uniqueness across sessions
    for slot in service_slots:
        d, sess = slot
        assigned_this_slot = set()

        # positions for this session
        for (disp, class_sess, cap) in [p for p in expanded if p[1] == sess]:
            for key in kid_keys_sorted:
                if key in assigned_this_slot:
                    continue
                if assigned_total[key] >= max_serves:
                    continue  # monthly cap reached
                if (key, d) in assigned_on_date:
                    continue  # already serving this date in another session/position
                if not availability.get(key, {}).get(slot, False):
                    continue
                k = kids_by_key[key]
                if sess == "Evening" and not k.get("evening_ok", True):
                    continue  # respect evening restriction
                if len(schedule[(disp, slot)]) >= cap:
                    continue

                # Assign
                schedule[(disp, slot)].append(k["name"])  # keep display casing
                assigned_this_slot.add(key)
                assigned_total[key] += 1
                assigned_on_date.add((key, d))

        # Unscheduled = said yes for this slot but not assigned
        yes_keys = [key for key in availability.keys() if availability.get(key, {}).get(slot, False)]
        for key in yes_keys:
            if key not in assigned_this_slot:
                unscheduled[slot].append(kids_by_key[key])

    return schedule, unscheduled

def schedule_to_df(schedule, service_slots):
    cols = [f"{d.strftime('%Y-%m-%d')} ({s})" for (d, s) in service_slots]
    rows = sorted({disp for (disp, _slot) in schedule.keys()})
    df = pd.DataFrame(index=rows, columns=cols)
    for (disp, slot), kids in schedule.items():
        d, s = slot
        df.loc[disp, f"{d.strftime('%Y-%m-%d')} ({s})"] = ", ".join(kids)
    return df.fillna("")

def unscheduled_to_df(unscheduled, service_slots):
    per_slot_names = {}
    per_slot_meta = {}
    for (d, s) in service_slots:
        entries = unscheduled.get((d, s), [])
        entries = sorted(entries, key=lambda k: (k.get("campus", ""), normalize(k["name"])))
        per_slot_names[(d, s)] = [e["name"] for e in entries]
        tags = []
        for e in entries:
            tag = ("Evening not allowed — " if (s == "Evening" and not e.get("evening_ok", True)) else "")
            campus = f", {e['campus']}" if e.get('campus') else ""
            tags.append(tag + campus)
        per_slot_meta[(d, s)] = tags

    max_len = max((len(v) for v in per_slot_names.values()), default=0)
    data = {}
    for (d, s) in service_slots:
        key = f"{d.strftime('%Y-%m-%d')} ({s})"
        names_list = per_slot_names[(d, s)] + [""] * (max_len - len(per_slot_names[(d, s)]))
        meta_list = per_slot_meta[(d, s)] + [""] * (max_len - len(per_slot_meta[(d, s)]))
        data[key] = names_list
        data[f"{key} info"] = meta_list
    return pd.DataFrame(data)

# ──────────────────────────────────────────────────────────────────────────────
# Run
# ──────────────────────────────────────────────────────────────────────────────
if st.button("Generate Kids Schedule", type="primary"):
    if not positions_file or not kids_file:
        st.error("Please upload the Positions and Kids CSV files.")
        st.stop()

    positions_df = read_csv_robust(positions_file, "positions")
    kids_df = read_csv_robust(kids_file, "kids")

    positions = parse_positions(positions_df)
    kids = build_kids(kids_df)
    availability, service_slots, sheet_name = parse_availability(kids_df)

    schedule, unscheduled = assign_kids(positions, kids, availability, service_slots, max_serves=max_serves)

    schedule_df = schedule_to_df(schedule, service_slots)
    unscheduled_df = unscheduled_to_df(unscheduled, service_slots)

    # Stats (count actual valid position×slot cells via schedule keys)
    filled_cells = sum(1 for v in schedule.values() if len(v) > 0)
    total_valid_cells = len(schedule)  # each key is a valid (position, slot) combination
    total_kids_scheduled = sum(len(v) for v in schedule.values())

    st.success(f"Schedule generated for **{sheet_name}**")
    st.write(f"Kids scheduled: **{total_kids_scheduled}**  •  Positions with any kids: **{filled_cells} / {total_valid_cells}**")

    st.subheader("Rosters (Position × Service)")
    st.dataframe(schedule_df, use_container_width=True)

    st.subheader("Unscheduled but Available — by Service (Name + Info)")
    st.caption("Kids who said Yes for that service but were not placed (capacity reached, already serving that day, evening restriction, or monthly cap).")
    st.dataframe(unscheduled_df, use_container_width=True)

    # Excel export
    def excel_autofit(ws):
        for col_idx, column_cells in enumerate(
            ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1
        ):
            max_len = 0
            for cell in column_cells:
                val = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(val))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 80)

    wb = Workbook()
    # Remove default sheet
    if wb.active and wb.active.title == "Sheet":
        wb.remove(wb.active)
    ws = wb.create_sheet(sheet_name)

    header = ["Position"] + [f"{d.strftime('%Y-%m-%d')} ({s})" for (d, s) in service_slots]
    ws.append(header)
    row_labels = sorted({disp for (disp, _slot) in schedule.keys()})
    for disp in row_labels:
        ws.append([disp] + [", ".join(schedule.get((disp, (d, s)), [])) for (d, s) in service_slots])
    excel_autofit(ws)

    ws2 = wb.create_sheet("Unscheduled by Service")
    header_pairs = []
    for (d, s) in service_slots:
        ds = f"{d.strftime('%Y-%m-%d')} ({s})"
        header_pairs.extend([ds, f"{ds} info"])
    ws2.append([" "] + header_pairs)

    for i in range(unscheduled_df.shape[0]):
        ws2.append([i + 1] + [unscheduled_df.iloc[i, j] for j in range(unscheduled_df.shape[1])])
    excel_autofit(ws2)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    st.download_button(
        "Download Excel (.xlsx)",
        data=buf,
        file_name=f"uKids_Kids_Schedule_{sheet_name.replace(' ','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload Positions + Kids CSVs, set rules, then click **Generate Kids Schedule**.")

