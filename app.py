import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import io, os
from collections import Counter
from datetime import datetime

# ========================= PAGE & TITLE =========================
st.set_page_config(page_title="Ranklin — The Anything Counter", layout="wide")
st.title("Ranklin — Anything Counter + Industries/Buzzwords")
st.write("Upload a CSV or Excel file. Use the Industries/Buzzwords mode for Beauhurst-style wide columns, or the generic Anything Counter.")

# ========================= UPLOAD =========================
uploaded_file = st.file_uploader("Choose a file", type=["csv", "xlsx", "xls"])

# ========================= HELPERS =========================
def read_any_table(file):
    name = getattr(file, "name", "") or ""
    ext = os.path.splitext(name)[1].lower()

    if ext in [".xlsx", ".xls"]:
        if ext == ".xlsx":
            try:
                import openpyxl  # noqa: F401
                engine = "openpyxl"
            except Exception:
                st.error("Reading .xlsx requires `openpyxl` (pip install openpyxl).")
                st.stop()
        else:  # .xls
            try:
                import xlrd  # noqa: F401
                # xlrd>=2.0 dropped .xls
                import xlrd as _xl
                if tuple(int(x) for x in _xl.__version__.split(".")[:2]) >= (2, 0):
                    st.error("Reading .xls requires xlrd==1.2.0 (pip install xlrd==1.2.0).")
                    st.stop()
                engine = "xlrd"
            except Exception:
                st.error("Reading .xls requires xlrd==1.2.0 (pip install xlrd==1.2.0).")
                st.stop()

        try:
            xls = pd.ExcelFile(file, engine=engine)
            sheet = st.selectbox("Select sheet:", xls.sheet_names, index=0)
            return pd.read_excel(file, sheet_name=sheet, engine=engine)
        except Exception as e:
            st.exception(e)
            st.stop()

    # CSV fallback
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
    try: return f"{int(n):,}"
    except Exception: return str(n)

def _clean_list_from_text(s: str) -> list[str]:
    if not s: return []
    parts = []
    for line in s.replace(",", "\n").splitlines():
        item = line.strip()
        if item: parts.append(item)
    seen, out = set(), []
    for p in parts:
        if p not in seen:
            seen.add(p); out.append(p)
    return out

def apply_exclusions_and_custom_order(labels, values, excludes=None, custom_order=None, sort_unknown_by_value_desc=True):
    excludes = set(excludes or [])
    filtered = [(l, v) for l, v in zip(labels, values) if l not in excludes]
    if not custom_order:
        if filtered: lo, vo = zip(*filtered); return list(lo), list(vo)
        return [], []
    order_index = {name: i for i, name in enumerate(custom_order)}
    present = [(l, v) for l, v in filtered if l in order_index]
    absent  = [(l, v) for l, v in filtered if l not in order_index]
    present.sort(key=lambda lv: order_index[lv[0]])
    if sort_unknown_by_value_desc:
        absent.sort(key=lambda lv: (-(float(lv[1]) if lv[1] is not None else 0), str(lv[0]).lower()))
    merged = present + absent
    if merged: lo, vo = zip(*merged); return list(lo), list(vo)
    return [], []

def coerce_bool_df(df_bool_like: pd.DataFrame) -> pd.DataFrame:
    out = df_bool_like.copy()
    num_cols = out.select_dtypes(include=[np.number]).columns
    out[num_cols] = out[num_cols].fillna(0) != 0
    other_cols = [c for c in out.columns if c not in num_cols]
    if other_cols:
        s = out[other_cols].astype(str).str.strip().str.lower()
        truthy = s.isin(["y", "yes", "true", "1", "✓", "✔", "x"])
        nonempty = s.ne("") & s.ne("nan")
        out[other_cols] = (truthy | nonempty)
    return out.fillna(False)

def find_amount_columns(cols):
    lc = [c.lower() for c in cols]
    candidates = []
    for i, c in enumerate(lc):
        if ("amount" in c and "gbp" in c) or ("amount raised" in c):
            candidates.append(cols[i])
        if "total amount received by the company" in c and "converted to gbp" in c:
            candidates.append(cols[i])
    # unique, keep order
    seen, out = set(), []
    for c in candidates:
        if c not in seen:
            seen.add(c); out.append(c)
    return out

def detect_layout(df):
    cols = list(df.columns.astype(str))
    ind_single = "Industries" if "Industries" in cols else ("(Company) Industries" if "(Company) Industries" in cols else None)
    buzz_single = "Buzzwords" if "Buzzwords" in cols else ("(Company) Buzzwords" if "(Company) Buzzwords" in cols else None)
    ind_wide = [c for c in cols if c.startswith("Industries - ") or c.startswith("(Company) Industries - ")]
    buzz_wide = [c for c in cols if c.startswith("Buzzwords - ") or c.startswith("(Company) Buzzwords - ")]
    if ind_single and buzz_single:
        return {"mode": "single", "ind_col": ind_single, "buzz_col": buzz_single}
    if ind_wide or buzz_wide:
        return {"mode": "wide", "ind_cols": ind_wide, "buzz_cols": buzz_wide}
    return {"mode": "unknown"}

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
        ax.text(offset_data, y_pos[i], str(label), fontsize=13, ha='left', va='center', fontweight='normal', color=text_color)
        ax.text(max_value - offset_data, y_pos[i], right_formatter(v), fontsize=13, ha='right', va='center', fontweight='semibold', color=text_color)

    ax.set_title(title, fontsize=15, pad=20, fontweight='normal')
    ax.invert_yaxis()
    st.pyplot(fig, use_container_width=True)
    return fig

# ========================= APP =========================
if uploaded_file is not None:
    df = read_any_table(uploaded_file)

    mode = st.radio(
        "Are you ranking **Industries/Buzzwords**?",
        ["No – use Anything Counter", "Yes – use Industries/Buzzwords"],
        horizontal=True
    )

    # -------------------- INDUSTRIES/BUZZWORDS MODE --------------------
    if mode.endswith("Industries/Buzzwords"):
        layout = detect_layout(df)

        if layout["mode"] == "unknown":
            st.error("Could not detect Industries/Buzzwords columns. Expect either single columns ('Industries','Buzzwords') or wide columns starting with 'Industries - ' / 'Buzzwords - '.")
            st.stop()

        ranking_by = st.radio("Rank by:", ["Count", "Total Amount Raised"], horizontal=True)

        amount_candidates = find_amount_columns(list(df.columns.astype(str)))
        amount_choice = st.selectbox("Amount column (optional)", ["<None>"] + amount_candidates, index=0)
        amount_choice = None if amount_choice == "<None>" else amount_choice

        # Build tallies
        if layout["mode"] == "single":
            industries_col = layout["ind_col"]
            buzzwords_col  = layout["buzz_col"]

            inds = df[industries_col].dropna().astype(str).str.split(",").explode().str.strip()
            buzz = df[buzzwords_col].dropna().astype(str).str.split(",").explode().str.strip()
            items = pd.concat([inds, buzz], ignore_index=True)
            items = items[items.ne("") & items.ne("nan")]
            counts = items.value_counts()

            if amount_choice:
                amt = pd.to_numeric(df[amount_choice], errors="coerce").fillna(0.0)

                def explode_with_rowkey(series, keyname):
                    s = series.where(series.notna(), "").astype(str).str.split(",")
                    ex = s.explode().str.strip()
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
                amount_per_item = pd.Series(0.0, index=counts.index)

        else:  # wide
            ind_cols  = layout.get("ind_cols", [])
            buzz_cols = layout.get("buzz_cols", [])
            pieces = []
            if ind_cols:
                pieces.append(pd.DataFrame({c.split(" - ", 1)[1]: df[c] for c in ind_cols}))
            if buzz_cols:
                pieces.append(pd.DataFrame({c.split(" - ", 1)[1]: df[c] for c in buzz_cols}))
            M = pd.concat(pieces, axis=1)
            M = M.groupby(level=0, axis=1).sum()
            M_bool = coerce_bool_df(M)
            counts = M_bool.sum(axis=0).sort_values(ascending=False)
            if amount_choice:
                amt = pd.to_numeric(df[amount_choice], errors="coerce").fillna(0.0)
                amount_per_item = (M_bool.astype(float).multiply(amt, axis=0)).sum(axis=0)
            else:
                amount_per_item = pd.Series(0.0, index=counts.index)

        metric_series = counts if ranking_by == "Count" else amount_per_item
        metric_series = metric_series.sort_values(ascending=False)

        labels = metric_series.index.tolist()
        values = metric_series.values.tolist()

        # Exclude + Custom ranking UI
        with st.expander("Filter & custom ranking", expanded=False):
            excluded = st.multiselect("Exclude labels", options=labels, default=[])
            custom_rank_text = st.text_area(
                "Custom rank (top → bottom)",
                placeholder="Paste names separated by commas or new lines.",
                help="Items you list appear first in this exact order. Unlisted items follow (by value, unless you untick below)."
            )
            custom_order = _clean_list_from_text(custom_rank_text)
            sort_unknown_by_value = st.checkbox("Sort unlisted items by value (desc)", True)
            top_n = st.slider("Limit to top N after filters", 1, max(10, len(labels)), min(10, len(labels)))

        labels, values = apply_exclusions_and_custom_order(
            labels, values, excludes=excluded, custom_order=custom_order,
            sort_unknown_by_value_desc=sort_unknown_by_value
        )

        if not labels:
            st.info("Nothing to show — all values are excluded.")
            st.stop()

        chart_title = st.text_input("Chart title:", f"Top {min(top_n, len(labels))} Industries/Buzzwords by {ranking_by}")
        formatter = money_fmt if ranking_by != "Count" else int_commas
        fig = plot_bar(labels[:top_n], values[:top_n], chart_title, True, formatter)

        # Download SVG
        svg_buffer = io.BytesIO()
        fig.savefig(svg_buffer, format="svg", bbox_inches="tight")
        svg_buffer.seek(0)
        st.download_button(
            label="Download Chart as SVG",
            data=svg_buffer,
            file_name=f"{chart_title.replace(' ', '_').lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.svg",
            mime="image/svg+xml",
        )

    # -------------------- ANYTHING COUNTER MODE --------------------
    else:
        st.subheader("Analysis Options")
        analysis_type = st.radio("Select analysis type:", ["Count Values", "Sum Values"], horizontal=True)

        if analysis_type == "Count Values":
            col = st.selectbox("Select column:", df.columns.tolist())
            explode = st.checkbox("Explode comma-separated values")
            if explode:
                value_list = []
                for val in df[col].dropna():
                    items = [s.strip() for s in str(val).split(",") if s.strip()]
                    value_list.extend(items)
                counts = Counter(value_list)
                labels = list(counts.keys())
                values = list(counts.values())
            else:
                vc = df[col].value_counts(dropna=False)
                labels = [("(blank)" if (isinstance(k, float) and pd.isna(k)) else str(k)) for k in vc.index.tolist()]
                values = vc.values.tolist()
            ranking_by = "Count"
            formatter = int_commas

        else:
            group_col = st.selectbox("Group by:", df.columns.tolist())
            num_cols = df.select_dtypes(include=["number"]).columns.tolist()
            if not num_cols:
                st.warning("No numeric columns found to sum.")
                st.stop()
            sum_col = st.selectbox("Sum column:", num_cols)
            is_money = st.toggle("Treat values as money (£)?", True)
            vals = pd.to_numeric(df[sum_col], errors="coerce")
            keys = df[group_col].astype(str).fillna("")
            summed = vals.groupby(keys, sort=False).sum()
            labels = summed.index.tolist()
            values = summed.values.tolist()
            ranking_by = "Amount (£)" if is_money else "Amount"
            formatter = money_fmt if is_money else int_commas

        # sort default by value desc
        labels, values = zip(*sorted(zip(labels, values), key=lambda lv: lv[1], reverse=True)) if labels else ([], [])
        labels, values = list(labels), list(values)

        # Exclude + Custom ranking UI
        with st.expander("Filter & custom ranking", expanded=False):
            excluded = st.multiselect("Exclude labels", options=labels, default=[])
            custom_rank_text = st.text_area(
                "Custom rank (top → bottom)",
                placeholder="Paste names separated by commas or new lines.",
                help="Items you list appear first in this exact order. Unlisted items follow (by value, unless you untick below)."
            )
            custom_order = _clean_list_from_text(custom_rank_text)
            sort_unknown_by_value = st.checkbox("Sort unlisted items by value (desc)", True)
            top_n = st.slider("Limit to top N after filters", 1, max(10, len(labels)), min(10, len(labels)))

        labels, values = apply_exclusions_and_custom_order(
            labels, values, excludes=excluded, custom_order=custom_order,
            sort_unknown_by_value_desc=sort_unknown_by_value
        )

        if not labels:
            st.info("Nothing to show — all values are excluded.")
            st.stop()

        chart_title = st.text_input("Chart title:", f"Top {min(top_n, len(labels))} by {ranking_by}")
        fig = plot_bar(labels[:top_n], values[:top_n], chart_title, True, formatter)

        # Download SVG
        svg_buffer = io.BytesIO()
        fig.savefig(svg_buffer, format="svg", bbox_inches="tight")
        svg_buffer.seek(0)
        st.download_button(
            label="Download Chart as SVG",
            data=svg_buffer,
            file_name=f"{chart_title.replace(' ', '_').lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.svg",
            mime="image/svg+xml",
        )
