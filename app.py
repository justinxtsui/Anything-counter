import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import io
import os
from datetime import datetime

st.set_page_config(page_title="Top Industries/Buzzwords Chart", layout="wide")
st.title("Top Industries/Buzzwords Chart")
st.write("Upload a CSV or Excel file. The app auto-detects either single columns or wide columns like `Industries - …` and `Buzzwords - …`.")

uploaded_file = st.file_uploader("Choose a file", type=["csv", "xlsx", "xls"])

# -------------------- Helpers --------------------
def read_any_table(file):
    name = getattr(file, "name", "") or ""
    ext = os.path.splitext(name)[1].lower()
    if ext in [".xlsx", ".xls"]:
        engine = "openpyxl" if ext == ".xlsx" else "xlrd"
        if engine == "xlrd":
            try:
                import xlrd
                if tuple(int(x) for x in xlrd.__version__.split(".")[:2]) >= (2, 0):
                    st.error("Reading legacy .xls requires xlrd==1.2.0")
                    st.stop()
            except Exception:
                st.error("Reading legacy .xls requires xlrd==1.2.0")
                st.stop()
        try:
            xls = pd.ExcelFile(file, engine=engine)
            sheet = st.selectbox("Select sheet:", xls.sheet_names, index=0)
            return pd.read_excel(file, sheet_name=sheet, engine=engine)
        except Exception as e:
            st.exception(e)
            st.stop()
    # CSV
    try:
        return pd.read_csv(file)
    except UnicodeDecodeError:
        return pd.read_csv(file, encoding="latin-1")

def money_fmt(v: float) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)) or v == 0:
        return "£0"
    if v >= 1_000_000_000:
        x = v / 1_000_000_000
        return f"£{x:.0f}b" if x >= 100 else (f"£{x:.1f}b" if x >= 10 else f"£{x:.2f}b")
    if v >= 1_000_000:
        x = v / 1_000_000
        return f"£{x:.0f}m" if x >= 100 else (f"£{x:.1f}m" if x >= 10 else f"£{x:.2f}m")
    if v >= 1_000:
        x = v / 1_000
        return f"£{x:.0f}k" if x >= 100 else (f"£{x:.1f}k" if x >= 10 else f"£{x:.2f}k")
    return f"£{v:.0f}" if v >= 100 else (f"£{v:.1f}" if v >= 10 else f"£{v:.2f}")

def int_commas(n) -> str:
    try:
        return f"{int(n):,}"
    except Exception:
        return str(n)

def coerce_bool_df(df_bool_like: pd.DataFrame) -> pd.DataFrame:
    """Coerce various truthy values (1, True, 'Y','Yes','✓','X', non-empty strings) to boolean."""
    out = df_bool_like.copy()
    # numeric -> nonzero is True
    num_cols = out.select_dtypes(include=[np.number]).columns
    out[num_cols] = out[num_cols].fillna(0) != 0

    # remaining -> string checks
    other_cols = [c for c in out.columns if c not in num_cols]
    if other_cols:
        s = out[other_cols].astype(str).str.strip().str.lower()
        truthy = s.isin(["y", "yes", "true", "1", "✓", "✔", "x"])
        nonempty = s.ne("") & s.ne("nan")
        out[other_cols] = (truthy | nonempty)  # treat any non-empty as True
    return out.fillna(False)

def plot_bar(labels, values, title, highlight_first=True, right_formatter=int_commas):
    mpl.rcParams['svg.fonttype'] = 'none'
    mpl.rcParams['pdf.fonttype'] = 42
    mpl.rcParams['font.family'] = 'Public Sans'
    mpl.rcParams['font.sans-serif'] = ['Public Sans', 'Arial', 'DejaVu Sans']
    mpl.rcParams['font.weight'] = 'normal'

    y_pos = list(range(len(labels)))
    fig, ax = plt.subplots(figsize=(10, 6))
    max_value = max(values) if values else 0

    ax.barh(y_pos, [max_value] * len(values), color='#E0E0E0', alpha=1.0, height=0.8)
    for i, (y, v) in enumerate(zip(y_pos, values)):
        color = '#4B4897' if (highlight_first and i == 0) else '#A4A2F2'
        ax.barh(y, float(v), color=color, height=0.8)

    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.xaxis.set_visible(False)
    ax.tick_params(axis='y', which='both', length=0)

    offset_data = max_value * 0.02 if max_value else 0.05
    for i, (label, v) in enumerate(zip(labels, values)):
        text_color = 'white' if (highlight_first and i == 0) else 'black'
        ax.text(offset_data, y_pos[i], str(label),
                fontsize=13, ha='left', va='center', fontweight='normal', color=text_color)
        ax.text(max_value - offset_data, y_pos[i], right_formatter(v),
                fontsize=13, ha='right', va='center', fontweight='semibold', color=text_color)

    ax.set_title(title, fontsize=15, pad=20, fontweight='normal')
    ax.invert_yaxis()
    st.pyplot(fig, use_container_width=True)
    return fig

def find_amount_columns(cols: list[str]) -> list[str]:
    lc = [c.lower() for c in cols]
    candidates = []
    for i, c in enumerate(lc):
        if ("amount" in c and "gbp" in c) or ("amount raised" in c):
            candidates.append(cols[i])
        if "total amount received by the company" in c and "converted to gbp" in c:
            candidates.append(cols[i])
    # keep order but unique
    seen, out = set(), []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out

def detect_layout(df: pd.DataFrame):
    """Return dict describing detected layout."""
    cols = list(df.columns.astype(str))

    # long/single columns
    ind_single = None
    if "Industries" in cols:
        ind_single = "Industries"
    elif "(Company) Industries" in cols:
        ind_single = "(Company) Industries"

    buzz_single = None
    if "Buzzwords" in cols:
        buzz_single = "Buzzwords"
    elif "(Company) Buzzwords" in cols:
        buzz_single = "(Company) Buzzwords"

    # wide/prefixed columns
    ind_wide = [c for c in cols if c.startswith("Industries - ") or c.startswith("(Company) Industries - ")]
    buzz_wide = [c for c in cols if c.startswith("Buzzwords - ") or c.startswith("(Company) Buzzwords - ")]

    if ind_single and buzz_single:
        return {"mode": "single", "ind_col": ind_single, "buzz_col": buzz_single}
    if ind_wide or buzz_wide:
        return {"mode": "wide", "ind_cols": ind_wide, "buzz_cols": buzz_wide}
    return {"mode": "unknown"}

# -------------------- App --------------------
if uploaded_file is not None:
    df = read_any_table(uploaded_file)

    layout = detect_layout(df)
    amount_candidates = find_amount_columns(list(df.columns.astype(str)))
    amount_choice = None
    if amount_candidates:
        amount_choice = st.selectbox(
            "Amount column (optional)",
            ["<None>"] + amount_candidates,
            index=0 if not amount_candidates else 1
        )
        if amount_choice == "<None>":
            amount_choice = None

    ranking_by = st.radio("Rank by:", ["Count", "Total Amount Raised"], horizontal=True)
    if ranking_by == "Total Amount Raised" and not amount_choice:
        st.info("No amount column selected — totals will be £0. Choose an amount column if available.")

    # ---- Build tallies (vectorised) ----
    if layout["mode"] == "single":
        industries_col = layout["ind_col"]
        buzzwords_col  = layout["buzz_col"]

        # Split+explode both, combine into one series of items
        inds = (
            df[industries_col].dropna().astype(str).str.split(",").explode().str.strip()
        )
        buzz = (
            df[buzzwords_col].dropna().astype(str).str.split(",").explode().str.strip()
        )
        items = pd.concat([inds, buzz], ignore_index=True)
        items = items[items.ne("") & items.ne("nan")]

        counts = items.value_counts()
        if amount_choice:
            amt = pd.to_numeric(df[amount_choice], errors="coerce").fillna(0)
            # For each row, distribute its amount to each item present in that row
            # Build membership per row for both columns
            def explode_with_rowkey(series, keyname):
                s = series.copy()
                s = s.where(s.notna(), "")
                s = s.astype(str).str.split(",")
                ex = s.explode()
                ex = ex.str.strip()
                mask = ex.ne("") & ex.ne("nan")
                out = pd.DataFrame({keyname: ex[mask]})
                out["__row__"] = np.repeat(np.arange(len(series)), s.str.len())[mask]
                return out

            ex_i = explode_with_rowkey(df[industries_col], "item")
            ex_b = explode_with_rowkey(df[buzzwords_col], "item")
            ex = pd.concat([ex_i, ex_b], ignore_index=True)
            ex = ex[ex["item"].ne("")]

            ex = ex.merge(pd.DataFrame({"__row__": np.arange(len(df)), "amt": amt}), on="__row__", how="left")
            amount_per_item = ex.groupby("item", as_index=True)["amt"].sum()
        else:
            amount_per_item = pd.Series(0, index=counts.index)

    elif layout["mode"] == "wide":
        ind_cols  = layout.get("ind_cols", [])
        buzz_cols = layout.get("buzz_cols", [])
        if not ind_cols and not buzz_cols:
            st.error("Could not find any columns starting with 'Industries - ' or 'Buzzwords - '.")
            st.stop()

        pieces = []
        rename_map = {}
        if ind_cols:
            rename_map.update({c: c.split(" - ", 1)[1] for c in ind_cols})
            pieces.append(df[ind_cols].rename(columns=rename_map))
        if buzz_cols:
            rename_map.update({c: c.split(" - ", 1)[1] for c in buzz_cols})
            pieces.append(df[buzz_cols].rename(columns={c: c.split(" - ", 1)[1] for c in buzz_cols}))

        M = pd.concat(pieces, axis=1)
        # If there are duplicate names between industries and buzzwords, keep both by grouping columns with same name
        M = M.groupby(level=0, axis=1).sum()

        M_bool = coerce_bool_df(M)
        counts = M_bool.sum(axis=0).sort_values(ascending=False)

        if amount_choice:
            amt = pd.to_numeric(df[amount_choice], errors="coerce").fillna(0.0)
            # Weighted amount per item: sum over rows of amt * membership
            amount_per_item = (M_bool.astype(float).multiply(amt, axis=0)).sum(axis=0)
        else:
            amount_per_item = pd.Series(0.0, index=counts.index)
    else:
        st.error("Could not detect Industries/Buzzwords columns. "
                 "Expect either single columns ('Industries', 'Buzzwords') "
                 "or wide columns starting with 'Industries - ' / 'Buzzwords - '.")
        st.stop()

    # ---- Build list for UI, exclusions, top N ----
    metric_series = counts if ranking_by == "Count" else amount_per_item
    # Align indexes in case of differences
    all_index = metric_series.index
    # Sort by chosen metric
    all_items = metric_series.sort_values(ascending=False).index.tolist()

    excluded = st.multiselect("Exclude specific industries/buzzwords:", options=all_items, default=[])
    filt_idx = [i for i in all_items if i not in excluded]
    if not filt_idx:
        st.info("Nothing to show — all values are excluded.")
        st.stop()

    max_available = len(filt_idx)
    top_n = st.number_input("Number of top industries/buzzwords to display:",
                            min_value=1, max_value=max_available, value=min(10, max_available))

    top_labels = filt_idx[:int(top_n)]
    if ranking_by == "Count":
        top_values = [int(counts.get(k, 0)) for k in top_labels]
        formatter = int_commas
    else:
        top_values = [float(amount_per_item.get(k, 0)) for k in top_labels]
        formatter = money_fmt

    chart_title = st.text_input("Chart title:", value=f"Top {top_n} Industries/Buzzwords by {ranking_by}")

    fig = plot_bar(top_labels, top_values, chart_title, highlight_first=True, right_formatter=formatter)

    # Download as SVG
    svg_buffer = io.BytesIO()
    fig.savefig(svg_buffer, format="svg", bbox_inches="tight")
    svg_buffer.seek(0)
    st.download_button(
        label="Download Chart as SVG",
        data=svg_buffer,
        file_name=f"{chart_title.replace(' ', '_').lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.svg",
        mime="image/svg+xml",
    )
