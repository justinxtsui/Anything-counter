import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import io, os
from datetime import datetime

st.set_page_config(page_title="Anything Counter + Industries/Buzzwords", layout="wide")
st.title("Anything Counter + Industries/Buzzwords")
st.write("Upload a CSV or Excel file. Choose Industries/Buzzwords mode for Beauhurst-style wide columns, or use the generic Anything Counter.")

uploaded_file = st.file_uploader("Choose a file", type=["csv", "xlsx", "xls"])

# ========================= SHARED HELPERS =========================
def read_any_table(file):
    name = getattr(file, "name", "") or ""
    ext = os.path.splitext(name)[1].lower()

    if ext in [".xlsx", ".xls"]:
        if ext == ".xlsx":
            try:
                import openpyxl
                engine = "openpyxl"
            except Exception:
                st.error("Reading .xlsx requires `openpyxl`. Install with `pip install openpyxl`.")
                st.stop()
        else:
            try:
                import xlrd
                if tuple(int(x) for x in xlrd.__version__.split(".")[:2]) >= (2, 0):
                    st.error("Reading .xls requires `xlrd==1.2.0` (xlrd>=2.0 dropped .xls).")
                    st.stop()
                engine = "xlrd"
            except Exception:
                st.error("Reading .xls requires `xlrd==1.2.0`.")
                st.stop()

        try:
            xls = pd.ExcelFile(file, engine=engine)
            sheet = st.selectbox("Select sheet:", xls.sheet_names, index=0)
            df = pd.read_excel(file, sheet_name=sheet, engine=engine)
            return df
        except Exception as e:
            st.exception(e)
            st.stop()

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


def offer_downloads(fig, labels, values, chart_title, value_header="Value", layout="stack"):
    """Unified export with captions."""
    svg_buffer = io.BytesIO()
    fig.savefig(svg_buffer, format="svg", bbox_inches="tight")
    svg_buffer.seek(0)

    png_buffer = io.BytesIO()
    fig.savefig(png_buffer, format="png", dpi=300, bbox_inches="tight")
    png_buffer.seek(0)

    df_out = pd.DataFrame({"Label": labels, value_header: values})
    csv_bytes = df_out.to_csv(index=False).encode("utf-8")

    def render_buttons():
        st.caption("For Adobe")
        st.download_button(
            label="Download Chart as SVG",
            data=svg_buffer,
            file_name=f"{chart_title.replace(' ', '_').lower()}.svg",
            mime="image/svg+xml",
            use_container_width=True,
        )
        st.caption("For Google Slides")
        st.download_button(
            label="Download Chart as PNG",
            data=png_buffer,
            file_name=f"{chart_title.replace(' ', '_').lower()}.png",
            mime="image/png",
            use_container_width=True,
        )
        st.caption("For Tableau")
        st.download_button(
            label="Download Data as CSV",
            data=csv_bytes,
            file_name=f"{chart_title.replace(' ', '_').lower()}_data.csv",
            mime="text/csv",
            use_container_width=True,
        )

    if layout == "columns":
        c1, c2, c3 = st.columns(3)
        with c1:
            st.caption("For Adobe")
            st.download_button(
                label="Download Chart as SVG",
                data=svg_buffer,
                file_name=f"{chart_title.replace(' ', '_').lower()}.svg",
                mime="image/svg+xml",
                use_container_width=True,
            )
        with c2:
            st.caption("For Google Slides")
            st.download_button(
                label="Download Chart as PNG",
                data=png_buffer,
                file_name=f"{chart_title.replace(' ', '_').lower()}.png",
                mime="image/png",
                use_container_width=True,
            )
        with c3:
            st.caption("For Tableau")
            st.download_button(
                label="Download Data as CSV",
                data=csv_bytes,
                file_name=f"{chart_title.replace(' ', '_').lower()}_data.csv",
                mime="text/csv",
                use_container_width=True,
            )
    else:
        render_buttons()

# ---------- Industries/Buzzwords helpers ----------
def detect_layout(df: pd.DataFrame):
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


def find_amount_columns(cols: list[str]) -> list[str]:
    lc = [c.lower() for c in cols]
    candidates = []
    for i, c in enumerate(lc):
        if ("amount" in c and "gbp" in c) or ("amount raised" in c):
            candidates.append(cols[i])
        if "total amount received by the company" in c and "converted to gbp" in c:
            candidates.append(cols[i])
    seen, out = set(), []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


# ---------- Anything Counter helpers ----------
def count_values_vectorised(series: pd.Series, explode: bool) -> pd.Series:
    if explode:
        s = series.dropna().astype(str).str.split(",").explode().str.strip()
        s = s[s.ne("") & s.ne("nan")]
        return s.value_counts(dropna=False)
    else:
        return series.value_counts(dropna=False)


def group_sum_vectorised(df: pd.DataFrame, group_col: str, sum_col: str) -> pd.Series:
    vals = pd.to_numeric(df[sum_col], errors="coerce")
    keys = df[group_col].astype(str).fillna("")
    return vals.groupby(keys, sort=False).sum()

# ========================= APP LOGIC =========================
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

        if layout["mode"] == "unknown":
            st.error("Could not detect Industries/Buzzwords columns.")
            st.stop()

        # Process wide layout
        if layout["mode"] == "wide":
            ind_cols = layout.get("ind_cols", [])
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

        # Process single layout
        else:
            industries_col = layout["ind_col"]
            buzzwords_col = layout["buzz_col"]
            inds = df[industries_col].dropna().astype(str).str.split(",").explode().str.strip()
            buzz = df[buzzwords_col].dropna().astype(str).str.split(",").explode().str.strip()
            items = pd.concat([inds, buzz], ignore_index=True)
            items = items[items.ne("") & items.ne("nan")]
            counts = items.value_counts()
            if amount_choice:
                amt = pd.to_numeric(df[amount_choice], errors="coerce").fillna(0)
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
                ex = ex.merge(pd.DataFrame({"__row__": np.arange(len(df)), "amt": amt}), on="__row__", how="left")
                amount_per_item = ex.groupby("item", as_index=True)["amt"].sum()
            else:
                amount_per_item = pd.Series(0, index=counts.index)

        metric_series = counts if ranking_by == "Count" else amount_per_item
        all_items = metric_series.sort_values(ascending=False).index.tolist()

        excluded = st.multiselect("Exclude specific industries/buzzwords:", options=all_items, default=[])
        kept = [i for i in all_items if i not in excluded]
        if not kept:
            st.stop()

        top_n = st.number_input("Number of top industries/buzzwords to display:", 1, len(kept), min(10, len(kept)))
        labels = kept[:int(top_n)]
        if ranking_by == "Count":
            values = [int(counts.get(k, 0)) for k in labels]
            formatter = int_commas
            header = "Count"
        else:
            values = [float(amount_per_item.get(k, 0)) for k in labels]
            formatter = money_fmt
            header = "Amount (£)"

        chart_title = st.text_input("Chart title:", f"Top {top_n} Industries/Buzzwords by {ranking_by}")
        fig = plot_bar(labels, values, chart_title, True, formatter)
        offer_downloads(fig, labels, values, chart_title, header)

    # -------------------- ANYTHING COUNTER MODE --------------------
    else:
        st.subheader("Analysis Options")
        analysis_type = st.radio("Select analysis type:", ["Count Values", "Sum Values"], horizontal=True)

        if analysis_type == "Count Values":
            count_column = st.selectbox("Select column to count:", df.columns.tolist())
            explode_option = st.checkbox("Explode comma-separated values before counting")
            vc = count_values_vectorised(df[count_column], explode_option)
            ranking_data = {k: {"count": int(v)} for k, v in vc.to_dict().items()}
            ranking_by = "Count"
        else:
            group_column = st.selectbox("Select column to group by:", df.columns.tolist())
            numeric_columns = df.select_dtypes(include=["number"]).columns.tolist()
            if not numeric_columns:
                st.warning("No numeric columns found to sum.")
                st.stop()
            sum_column = st.selectbox("Select column to sum:", numeric_columns)
            is_money = st.toggle("Treat values as money (£)?", value=True)
            summed = group_sum_vectorised(df, group_column, sum_column)
            ranking_data = {k: {"total": v} for k, v in summed.items()}
            ranking_by = "Total Amount"

        items = list(ranking_data.keys())
        excluded = st.multiselect("Exclude specific values:", items)
        kept = [i for i in items if i not in excluded]

        top_n = st.number_input("Number of top values:", 1, len(kept), min(10, len(kept)))
        labels = kept[:int(top_n)]
        values = []
        if ranking_by == "Count":
            values = [ranking_data[k]["count"] for k in labels]
            formatter = int_commas
            header = "Count"
        else:
            values = [ranking_data[k]["total"] for k in labels]
            formatter = money_fmt if is_money else int_commas
            header = "Amount (£)" if is_money else "Amount"

        chart_title = st.text_input("Chart title:", f"Top {top_n} by {ranking_by}")
        fig = plot_bar(labels, values, chart_title, True, formatter)
        offer_downloads(fig, labels, values, chart_title, header)
