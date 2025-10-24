import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from collections import Counter
import io
from datetime import datetime
import os

st.set_page_config(page_title="The Anything Counter", layout="wide")
st.title('The Anything Counter')
st.write('Upload a CSV or Excel file to analyze and visualize column data.')

uploaded_file = st.file_uploader('Choose a file', type=['csv', 'xlsx', 'xls'])

def read_any_table(file):
    name = getattr(file, "name", "") or ""
    ext = os.path.splitext(name)[1].lower()

    if ext in [".xlsx", ".xls"]:
        # Excel handling with friendly engine checks
        if ext == ".xlsx":
            try:
                import openpyxl  # noqa: F401
            except ImportError:
                st.error(
                    "Reading .xlsx requires the **openpyxl** package.\n\n"
                    "Install it with:\n\n`pip install openpyxl`\n\n"
                    "Or add to requirements.txt: `openpyxl>=3.1.2`"
                )
                st.stop()
            engine = "openpyxl"
        else:  # .xls
            try:
                import xlrd  # noqa: F401
                # xlrd>=2.0 dropped .xls; 1.2.0 is needed
                if tuple(int(x) for x in xlrd.__version__.split(".")[:2]) >= (2, 0):
                    raise ImportError(
                        "xlrd>=2.0 no longer supports .xls. Install xlrd==1.2.0."
                    )
            except ImportError:
                st.error(
                    "Reading legacy .xls requires **xlrd==1.2.0**.\n\n"
                    "Install with:\n\n`pip install xlrd==1.2.0`\n\n"
                    "Or add to requirements.txt: `xlrd==1.2.0`"
                )
                st.stop()
            engine = "xlrd"

        try:
            xls = pd.ExcelFile(file, engine=engine)
            sheet = st.selectbox("Select sheet:", xls.sheet_names, index=0)
            df = pd.read_excel(file, sheet_name=sheet, engine=engine)
            return df
        except Exception as e:
            st.exception(e)
            st.stop()

    # Default to CSV
    try:
        return pd.read_csv(file)
    except UnicodeDecodeError:
        # Fallback encoding
        return pd.read_csv(file, encoding="latin-1")

if uploaded_file is not None:
    df = read_any_table(uploaded_file)

    st.subheader('Analysis Options')

    # Choose analysis type
    analysis_type = st.radio('Select analysis type:', ['Count Values', 'Sum Values'], horizontal=True)

    if analysis_type == 'Count Values':
        # Column selection for counting
        count_column = st.selectbox('Select column to count:', df.columns.tolist())

        # Option to explode comma-separated values
        explode_option = st.checkbox('Explode comma-separated values before counting')

        # Perform counting
        if explode_option:
            value_list = []
            for val in df[count_column].dropna():
                items = [item.strip() for item in str(val).split(',') if item.strip()]
                value_list.extend(items)
            value_counts = Counter(value_list)
            ranking_data = {k: {'count': v, 'total_amount': 0} for k, v in value_counts.items()}
        else:
            value_counts = df[count_column].value_counts(dropna=False).to_dict()
            ranking_data = {
                ('' if (isinstance(k, float) and pd.isna(k)) else k): {'count': v, 'total_amount': 0}
                for k, v in value_counts.items()
            }

        ranking_by = 'Count'

    else:  # Sum Values
        # Column selection for grouping and summing
        group_column = st.selectbox('Select column to group by (unique values):', df.columns.tolist())

        # Get numeric columns for summing
        numeric_columns = df.select_dtypes(include=['number']).columns.tolist()
        if not numeric_columns:
            st.warning("No numeric columns found to sum. Please upload a dataset with numeric columns or choose 'Count Values'.")
            st.stop()

        sum_column = st.selectbox('Select column to sum:', numeric_columns)

        # ---- Robust grouping to avoid TypeError on mixed types ----
        # Coerce sum column to numeric (already numeric per select, but safe)
        sum_series = pd.to_numeric(df[sum_column], errors='coerce')

        # Build safe, consistent group keys (hashable + single dtype)
        def _safe_group_key(x):
            # Convert unhashable (e.g., list) to string; keep NaN as NaN for dropna handling
            if pd.isna(x):
                return x
            try:
                hash(x)
                return x
            except TypeError:
                return str(x)

        group_keys_raw = df[group_column].map(_safe_group_key)
        # Cast to string to avoid mixed-type sorting/factorization issues
        group_keys = group_keys_raw.astype(str).fillna('')

        # Group using the Series.groupby path; disable sorting to avoid mixed type sorting
        grouped_series = sum_series.groupby(group_keys, sort=False).sum()
        grouped = grouped_series.to_dict()

        ranking_data = {
            ('' if (isinstance(k, float) and pd.isna(k)) else k): {'count': 0, 'total_amount': v}
            for k, v in grouped.items()
        }

        ranking_by = 'Total Amount'
        # ----------------------------------------------------------

    # Initial list for exclusions
    all_values = list(ranking_data.keys())

    # Exclusion multiselect
    excluded_values = st.multiselect(
        'Exclude specific values:',
        options=sorted(all_values, key=lambda x: str(x).lower()),
        default=[]
    )

    # Filter out excluded values
    filtered_data = {k: v for k, v in ranking_data.items() if k not in excluded_values}

    if not filtered_data:
        st.info("Nothing to show — all values are excluded.")
        st.stop()

    # Ranking mode
    rank_mode = st.radio(
        "Ranking mode",
        ["Highest first", "Lowest first", "Custom order"],
        help="Choose how to order the bars."
    )

    # Determine sort keys and values according to ranking_by
    def value_key(item):
        _, v = item
        return v['count'] if ranking_by == 'Count' else v['total_amount']

    # Number input for top N
    max_available = len(filtered_data)
    top_n = st.number_input(
        'Number of top values to display:',
        min_value=1,
        max_value=max_available,
        value=min(10, max_available)
    )

    # Build ordering
    if rank_mode in ["Highest first", "Lowest first"]:
        reverse_flag = (rank_mode == "Highest first")
        sorted_items = sorted(filtered_data.items(), key=value_key, reverse=reverse_flag)
        top_items = sorted_items[:top_n]
        labels = [str(k) if str(k) != "" else "(blank)" for k, _ in top_items]
        values = [v['count'] if ranking_by == 'Count' else v['total_amount'] for _, v in top_items]
        highlight_top = True
    else:
        st.markdown("**Custom order**: Edit the `Rank` column to set the display order (1 = top).")
        default_order = sorted(filtered_data.items(), key=lambda x: (-value_key(x), str(x[0]).lower()))
        df_order = pd.DataFrame({
            "Label": [("(blank)" if str(k) == "" else str(k)) for k, _ in default_order],
            ranking_by: [(v['count'] if ranking_by == 'Count' else v['total_amount']) for _, v in default_order],
            "Rank": list(range(1, len(default_order) + 1))
        })
        edited = st.data_editor(
            df_order,
            num_rows="fixed",
            use_container_width=True,
            column_config={
                "Label": st.column_config.TextColumn(disabled=True),
                "Rank": st.column_config.NumberColumn(min_value=1, max_value=len(default_order), step=1),
            },
            hide_index=True,
        )
        edited = edited.sort_values(by=["Rank", "Label"], ascending=[True, True])
        edited_top = edited.head(top_n)
        labels = edited_top["Label"].astype(str).tolist()
        # Map labels back to original keys (treat "(blank)" as empty string)
        inv_map = {("(blank)" if str(k) == "" else str(k)): (v['count'] if ranking_by == 'Count' else v['total_amount'])
                   for k, v in filtered_data.items()}
        values = [inv_map.get(lbl, 0) for lbl in labels]
        highlight_top = False  # custom mode: same colour for all bars

    chart_title = st.text_input('Chart title:', value=f'Top {top_n} by {ranking_by}')

    # Format function
    def format_value(value, is_amount=False):
        if is_amount:
            if value == 0:
                return '£0'
            if value >= 1_000_000_000:
                formatted = value / 1_000_000_000
                if formatted >= 100: return f'£{formatted:.0f}b'
                elif formatted >= 10: return f'£{formatted:.1f}b'
                else: return f'£{formatted:.2f}b'
            elif value >= 1_000_000:
                formatted = value / 1_000_000
                if formatted >= 100: return f'£{formatted:.0f}m'
                elif formatted >= 10: return f'£{formatted:.1f}m'
                else: return f'£{formatted:.2f}m'
            elif value >= 1_000:
                formatted = value / 1_000
                if formatted >= 100: return f'£{formatted:.0f}k'
                elif formatted >= 10: return f'£{formatted:.1f}k'
                else: return f'£{formatted:.2f}k'
            else:
                if value >= 100: return f'£{value:.0f}'
                elif value >= 10: return f'£{value:.1f}'
                else: return f'£{value:.2f}'
        else:
            try:
                return f'{int(value):,}'
            except Exception:
                return str(value)

    # Matplotlib styling
    mpl.rcParams['svg.fonttype'] = 'none'
    mpl.rcParams['pdf.fonttype'] = 42
    mpl.rcParams['font.family'] = 'Public Sans'
    mpl.rcParams['font.sans-serif'] = ['Public Sans', 'Arial', 'DejaVu Sans']
    mpl.rcParams['font.weight'] = 'normal'

    # Build chart
    y_pos = list(range(len(labels)))
    fig, ax = plt.subplots(figsize=(10, 6))
    max_value = max(values) if values else 0

    ax.barh(y_pos, [max_value] * len(values), color='#E0E0E0', alpha=1.0, height=0.8)

    base_color = '#A4A2F2'
    top_color = '#4B4897'
    for i, (y, value) in enumerate(zip(y_pos, values)):
        color = (top_color if (highlight_top and i == 0) else base_color)
        ax.barh(y, float(value), color=color, height=0.8)

    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.xaxis.set_visible(False)
    ax.tick_params(axis='y', which='both', length=0)

    offset_points = 5.67
    try:
        offset_data = offset_points * (max_value / (ax.get_window_extent().width * 72 / fig.dpi))
    except Exception:
        offset_data = max_value * 0.01 if max_value else 0.05

    for i, (label, value) in enumerate(zip(labels, values)):
        text_color = 'white' if (highlight_top and i == 0) else 'black'
        ax.text(offset_data, y_pos[i], str(label),
                fontsize=13, ha='left', va='center', fontweight='normal', color=text_color)
        ax.text(max_value - offset_data, y_pos[i],
                format_value(value, is_amount=(ranking_by == 'Total Amount')),
                fontsize=13, ha='right', va='center', fontweight='semibold', color=text_color)

    ax.set_title(chart_title, fontsize=15, pad=20, fontweight='normal')
    ax.invert_yaxis()
    plt.tight_layout()
    st.pyplot(fig, use_container_width=True)

    # Download as SVG
    svg_buffer = io.BytesIO()
    fig.savefig(svg_buffer, format='svg', bbox_inches='tight')
    svg_buffer.seek(0)
    st.download_button(
        label='Download Chart as SVG',
        data=svg_buffer,
        file_name=f'{chart_title.replace(" ", "_").lower()}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.svg',
        mime='image/svg+xml'
    )
