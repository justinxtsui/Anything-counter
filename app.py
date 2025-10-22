import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from collections import Counter
import io
from datetime import datetime

st.title('Dynamic Data Ranking Chart')

st.write('Upload a CSV file to analyze and visualize column data.')

uploaded_file = st.file_uploader('Choose a CSV file', type=['csv'])

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    
    st.subheader('Analysis Options')
    
    # Choose analysis type
    analysis_type = st.radio('Select analysis type:', ['Count Values', 'Sum Values'])
    
    if analysis_type == 'Count Values':
        # Column selection for counting
        count_column = st.selectbox('Select column to count:', df.columns.tolist())
        
        # Option to explode comma-separated values
        explode_option = st.checkbox('Explode comma-separated values before counting')
        
        # Perform counting
        if explode_option:
            # Explode comma-separated values
            value_list = []
            for val in df[count_column].dropna():
                items = [item.strip() for item in str(val).split(',') if item.strip()]
                value_list.extend(items)
            
            value_counts = Counter(value_list)
            ranking_data = {k: {'count': v, 'total_amount': 0} for k, v in value_counts.items()}
        else:
            # Count without exploding
            value_counts = df[count_column].value_counts().to_dict()
            ranking_data = {k: {'count': v, 'total_amount': 0} for k, v in value_counts.items()}
        
        # Sort by count
        all_values = sorted(ranking_data.keys(), key=lambda x: ranking_data[x]['count'], reverse=True)
        ranking_by = 'Count'
        
    else:  # Sum Values
        # Column selection for grouping and summing
        group_column = st.selectbox('Select column to group by (unique values):', df.columns.tolist())
        
        # Get numeric columns for summing
        numeric_columns = df.select_dtypes(include=['number']).columns.tolist()
        sum_column = st.selectbox('Select column to sum:', numeric_columns)
        
        # Perform aggregation
        grouped = df.groupby(group_column)[sum_column].sum().to_dict()
        ranking_data = {k: {'count': 0, 'total_amount': v} for k, v in grouped.items()}
        
        # Sort by sum
        all_values = sorted(ranking_data.keys(), key=lambda x: ranking_data[x]['total_amount'], reverse=True)
        ranking_by = 'Total Amount'
    
    # Exclusion multiselect
    excluded_values = st.multiselect(
        'Exclude specific values:',
        options=all_values,
        default=[]
    )
    
    # Filter out excluded values
    filtered_data = {k: v for k, v in ranking_data.items() if k not in excluded_values}
    
    # Number input for top N
    max_available = len(filtered_data)
    top_n = st.number_input(
        'Number of top values to display:',
        min_value=1,
        max_value=max_available,
        value=min(10, max_available)
    )
    
    # Get top N from filtered data
    if ranking_by == 'Count':
        topN = sorted(filtered_data.items(), key=lambda x: x[1]['count'], reverse=True)[:top_n]
        labels = [k for k, v in topN]
        values = [v['count'] for k, v in topN]
    else:
        topN = sorted(filtered_data.items(), key=lambda x: x[1]['total_amount'], reverse=True)[:top_n]
        labels = [str(k) for k, v in topN]
        values = [v['total_amount'] for k, v in topN]
    
    chart_title = st.text_input('Chart title:', value=f'Top {top_n} by {ranking_by}')
    
    # Function to format values to 3 significant figures
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
            return f'{int(value):,}'
    
    mpl.rcParams['svg.fonttype'] = 'none'
    mpl.rcParams['pdf.fonttype'] = 42
    mpl.rcParams['font.family'] = 'Public Sans'
    mpl.rcParams['font.sans-serif'] = ['Public Sans', 'Arial', 'DejaVu Sans']
    mpl.rcParams['font.weight'] = 'normal'
    
    y_pos = list(range(len(labels)))
    fig, ax = plt.subplots(figsize=(10, 6))
    max_value = max(values)
    
    ax.barh(y_pos, [max_value] * len(values), color='#E0E0E0', alpha=1.0, height=0.8)
    for i, (y, value) in enumerate(zip(y_pos, values)):
        color = '#4B4897' if i == 0 else '#A4A2F2'
        ax.barh(y, value, color=color, height=0.8)
    
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.xaxis.set_visible(False)
    ax.tick_params(axis='y', which='both', length=0)
    
    offset_points = 5.67
    offset_data = offset_points * (max_value / (ax.get_window_extent().width * 72 / fig.dpi))
    for i, (label, value) in enumerate(zip(labels, values)):
        text_color = 'white' if i == 0 else 'black'
        ax.text(offset_data, y_pos[i], str(label),
                fontsize=13, ha='left', va='center', fontweight='normal', color=text_color)
        ax.text(max_value - offset_data, y_pos[i], format_value(value, is_amount=(ranking_by == 'Total Amount')),
                fontsize=13, ha='right', va='center', fontweight='semibold', color=text_color)
    
    ax.set_title(chart_title, fontsize=15, pad=20, fontweight='normal')
    ax.invert_yaxis()
    plt.tight_layout()
    st.pyplot(fig)
    
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