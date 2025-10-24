import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import io, os
from datetime import datetime

st.set_page_config(page_title="Ranklin", layout="wide")
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


def offer_downloads(fig, labels, values, chart_title, value_header="Value"):
    """Unified export with bold captions + colorful large buttons."""
    import io
    import pandas as pd
    import streamlit as st

    # Export files
    svg_buffer = io.BytesIO()
    fig.savefig(svg_buffer, format="svg", bbox_inches="tight")
    svg_buffer.seek(0)
    png_buffer = io.BytesIO()
    fig.savefig(png_buffer, format="png", dpi=300, bbox_inches="tight")
    png_buffer.seek(0)
    df_out = pd.DataFrame({"Label": labels, value_header: values})
    csv_bytes = df_out.to_csv(index=False).encode("utf-8")

    # --- CSS styling for pretty large buttons ---
    st.markdown("""
        <style>
        .download-caption {
            font-weight: 700;
            font-size: 1.1rem;
            color: #1E1E1E;
            margin-top: 1rem;
            margin-bottom: 0.4rem;
        }
        div[data-testid="stDownloadButton"] button {
            background-color: #4B4897;
            color: white !important;
            font-weight: 600;
            font-size: 1rem;
            border: none;
            border-radius: 8px;
            padding: 0.8rem 1rem;
            width: 100%;
            transition: all 0.2s ease-in-out;
        }
        div[data-testid="stDownloadButton"] button:hover {
            background-color: #6A67E5;
            transform: translateY(-2px);
        }
        </style>
    """, unsafe_allow_html=True)

    # --- Render stacked buttons with captions ---
    st.markdown('<p class="download-caption">For Adobe 🧑🏼‍🎨</p>', unsafe_allow_html=True)
    st.download_button("Download Chart as SVG", svg_buffer, file_name=f"{chart_title}.svg", mime="image/svg+xml")

    st.markdown('<p class="download-caption">For Google Slides 📈</p>', unsafe_allow_html=True)
    st.download_button("Download Chart as PNG", png_buffer, file_name=f"{chart_title}.png", mime="image/png")

    st.markdown('<p class="download-caption">For Tableau 🫣</p>', unsafe_allow_html=True)
    st.download_button("Download Data as CSV", csv_bytes, file_name=f"{chart_title}_data.csv", mime="text/csv")


# ---------- Helper functions ----------
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
    return list(dict.fromkeys(candidates))


def count_values_vectorised(series, explode):
    if explode:
        s = series.dropna().astype(str).str.split(",").explode().str.strip()
        s = s[s.ne("") & s.ne("nan")]
        return s.value_counts(dropna=False)
    else:
        return series.value_counts(dropna=False)


def group_sum_vectorised(df, group_col, sum_col):
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
        amount_choice = st.selectbox("Amount column (optional)", ["<None>"] + amount_candidates, index=0)
        amount_choice = None if amount_choice == "<None>" else amount_choice
        ranking_by = st.radio("Rank by:", ["Count", "Total Amount Raised"], horizontal=True)

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
        else:
            st.error("Could not detect Industries/Buzzwords columns.")
            st.stop()

        metric_series = counts if ranking_by == "Count" else amount_per_item
        labels = metric_series.index.tolist()
        values = metric_series.values.tolist()

        chart_title = st.text_input("Chart title:", f"Top {len(labels)} Industries/Buzzwords by {ranking_by}")
        fig = plot_bar(labels[:10], values[:10], chart_title, True, money_fmt if ranking_by != "Count" else int_commas)
        offer_downloads(fig, labels[:10], values[:10], chart_title, "Value")

    # -------------------- ANYTHING COUNTER MODE --------------------
    else:
        st.subheader("Analysis Options")
        analysis_type = st.radio("Select analysis type:", ["Count Values", "Sum Values"], horizontal=True)

        if analysis_type == "Count Values":
            col = st.selectbox("Select column:", df.columns.tolist())
            explode = st.checkbox("Explode comma-separated values")
            counts = count_values_vectorised(df[col], explode)
            labels = counts.index.tolist()
            values = counts.values.tolist()
            ranking_by = "Count"
            formatter = int_commas
        else:
            group_col = st.selectbox("Group by:", df.columns.tolist())
            num_cols = df.select_dtypes(include=["number"]).columns.tolist()
            sum_col = st.selectbox("Sum column:", num_cols)
            is_money = st.toggle("Treat values as money (£)?", True)
            summed = group_sum_vectorised(df, group_col, sum_col)
            labels = summed.index.tolist()
            values = summed.values.tolist()
            ranking_by = "Amount (£)" if is_money else "Amount"
            formatter = money_fmt if is_money else int_commas

        chart_title = st.text_input("Chart title:", f"Top {min(10, len(labels))} by {ranking_by}")
        fig = plot_bar(labels[:10], values[:10], chart_title, True, formatter)
        offer_downloads(fig, labels[:10], values[:10], chart_title, ranking_by)
