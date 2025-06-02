import warnings
warnings.filterwarnings('ignore')

import pandas as pd
import streamlit as st
import numpy as np
from datetime import datetime
from io import BytesIO  # Add this import

# ================================
# CONFIGURATION SECTION
# ================================
# DATE_COL = 'Fecha documento'
# INVOICE_COL = 'Nº de pedido' 
# PRODUCT_COL = 'Material'
# QUANTITY_COL = 'Cantidad pedido'
# CATEGORY_COL = 'MARCA'
# DATA_SOURCE = './example_1.xlsx'

# ================================
# DATA PREPARATION FUNCTIONS
# ================================
def load_and_prepare_data(data_source):
    """
    Load data from Excel or CSV files and prepare it for analysis
    data_source can be: DataFrame, Excel file path, or CSV file path
    """
    if isinstance(data_source, str) and (data_source.endswith('.xlsx') or data_source.endswith('.xls')):
        df = pd.read_excel(data_source)
    elif isinstance(data_source, str) and data_source.endswith('.csv'):
        df = pd.read_csv(data_source)
    else:
        df = data_source.copy()
    
    # Let pandas automatically parse the date column
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors='coerce')
    
    # Let pandas automatically handle numeric conversion
    df[QUANTITY_COL] = pd.to_numeric(df[QUANTITY_COL], errors='coerce')
    
    # Add time-based columns for analysis
    df['Year'] = df[DATE_COL].dt.year
    df['Month'] = df[DATE_COL].dt.month
    df['Day'] = df[DATE_COL].dt.date
    df['Weekday'] = df[DATE_COL].dt.day_name()
    
    # FIXED: Proper week calculation with year
    df['Year_Week'] = df[DATE_COL].dt.strftime('%Y-W%U')
    
    # NEW: Quarterly analysis
    df['Year_Quarter'] = df[DATE_COL].dt.to_period('Q')
    
    return df

def calculate_statistics(series, name):
    """Calculate comprehensive statistics for a series with edge case handling"""
    # FIXED: Handle edge cases for statistical calculations
    if len(series) == 0:
        return {f'{name}_{stat}': np.nan for stat in 
                ['count', 'min', 'max', 'mean', 'median', 'std', 'p25', 'p75', 'p90', 'p95']}
    
    clean_series = series.dropna()
    if len(clean_series) == 0:
        return {f'{name}_{stat}': np.nan for stat in 
                ['count', 'min', 'max', 'mean', 'median', 'std', 'p25', 'p75', 'p90', 'p95']}
    
    stats = {
        f'{name}_count': len(clean_series),
        f'{name}_min': clean_series.min(),
        f'{name}_max': clean_series.max(),
        f'{name}_mean': clean_series.mean(),
        f'{name}_median': clean_series.median(),
        f'{name}_std': clean_series.std() if len(clean_series) > 1 else 0,  # FIXED: Handle single value case
        f'{name}_p25': clean_series.quantile(0.25),
        f'{name}_p75': clean_series.quantile(0.75),
        f'{name}_p90': clean_series.quantile(0.90),
        f'{name}_p95': clean_series.quantile(0.95),
    }
    return stats

def analyze_by_period(df, period_col):
    """
    Perform comprehensive analysis for a given time period
    """
    # First, aggregate by period to get totals per period
    period_totals = df.groupby(period_col).agg({
        QUANTITY_COL: 'sum',
        INVOICE_COL: 'nunique'
    }).rename(columns={
        QUANTITY_COL: 'Total_Quantity',
        INVOICE_COL: 'Invoices_Unique'
    })
    
    period_totals['Lines_Total'] = df.groupby(period_col).size()
    period_totals['Avg_Lines_per_Invoice'] = period_totals['Lines_Total'] / period_totals['Invoices_Unique']
    
    # Reset index to make period_col a regular column
    period_totals = period_totals.reset_index()
    period_totals['Period_Type'] = period_col
    
    # Reorder columns
    cols = ['Period_Type', period_col, 'Lines_Total', 'Total_Quantity', 'Invoices_Unique', 'Avg_Lines_per_Invoice']
    period_totals = period_totals[cols]
    period_totals.rename(columns={period_col: 'Period'}, inplace=True)
    
    return period_totals

def comprehensive_analysis(df, category_col=None):
    """
    Run complete analysis across all time periods
    """
    print("=" * 60)
    print("COMPREHENSIVE ORDERS ANALYSIS")
    print("=" * 60)
    
    # Basic dataset info
    print(f"\nDataset Overview:")
    print(f"Total Lines: {len(df):,}")
    print(f"Date Range: {df[DATE_COL].min().date()} to {df[DATE_COL].max().date()}")
    print(f"Unique Invoices: {df[INVOICE_COL].nunique():,}")
    print(f"Unique Products: {df[PRODUCT_COL].nunique():,}")
    print(f"Total Quantity: {df[QUANTITY_COL].sum():,}")
    
    # Analysis by different periods
    analyses = {}
    
    # Daily analysis
    daily_analysis = analyze_by_period(df, 'Day')
    analyses['Daily'] = daily_analysis
    
    # FIXED: Weekly analysis using proper year-week
    weekly_analysis = analyze_by_period(df, 'Year_Week')
    analyses['Weekly'] = weekly_analysis
    
    # Monthly analysis (simplified for line chart plotting)
    monthly_analysis = analyze_by_period(df, 'Month')
    analyses['Monthly'] = monthly_analysis
    
    # Weekday analysis
    weekday_analysis = analyze_by_period(df, 'Weekday')
    analyses['Weekday'] = weekday_analysis
    
    # Quarterly analysis (simplified for line chart plotting)
    quarterly_analysis = analyze_by_period(df, 'Year_Quarter')
    analyses['Quarterly'] = quarterly_analysis
    
    # ADDED: Statistics across days for store planning
    daily_stats = calculate_daily_statistics(daily_analysis)
    analyses['Daily_Statistics'] = daily_stats
    
    # ADDED: Statistics across weekdays for store planning
    weekday_stats = calculate_weekday_statistics(df)
    analyses['Weekday_Statistics'] = weekday_stats
    
    # NEW: Category analysis (if category column provided)
    if category_col and category_col in df.columns:
        category_analysis = analyze_category_by_day(df, category_col)
        analyses['Category_Analysis'] = category_analysis
        print(f"Category column analyzed: {category_col}")
    
    return analyses

def calculate_daily_statistics(daily_analysis):
    """
    Calculate statistics across all days for store planning
    """
    lines_totals = daily_analysis['Lines_Total']
    quantity_totals = daily_analysis['Total_Quantity']
    
    stats = {
        'Metric': [
            'Days Analyzed',
            'Avg Lines per Day', 
            'Max Lines per Day',
            'Min Lines per Day',
            'Lines 75th Percentile',
            'Lines 90th Percentile',
            'Avg Units per Day',
            'Max Units per Day', 
            'Min Units per Day',
            'Units 75th Percentile',
            'Units 90th Percentile'
        ],
        'Value': [
            len(daily_analysis),
            f"{lines_totals.mean():.1f}",
            f"{lines_totals.max():.0f}",
            f"{lines_totals.min():.0f}", 
            f"{lines_totals.quantile(0.75):.0f}",
            f"{lines_totals.quantile(0.90):.0f}",
            f"{quantity_totals.mean():.1f}",
            f"{quantity_totals.max():.0f}",
            f"{quantity_totals.min():.0f}",
            f"{quantity_totals.quantile(0.75):.0f}",
            f"{quantity_totals.quantile(0.90):.0f}"
        ]
    }
    
    return pd.DataFrame(stats)

def calculate_weekday_statistics(df):
    """
    Calculate average and 90th percentile for each weekday based on daily quantities
    """
    # Group daily data by weekday and get quantities for each day of that weekday
    daily_by_weekday = df.groupby(['Weekday', 'Day'])[QUANTITY_COL].sum().reset_index()
    
    weekday_stats = []
    weekdays_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    
    for weekday in weekdays_order:
        weekday_data = daily_by_weekday[daily_by_weekday['Weekday'] == weekday]
        if len(weekday_data) > 0:
            quantities = weekday_data[QUANTITY_COL]
            weekday_stats.append({
                'Weekday': weekday,
                'Avg_Units': f"{quantities.mean():.1f}",
                'P90_Units': f"{quantities.quantile(0.90):.0f}",
                'Days_Count': len(quantities)
            })
    
    return pd.DataFrame(weekday_stats)

def analyze_category_by_day(df, category_col):
    """
    Calculate average units per day for each category
    """
    # Group by category and day, sum quantities
    category_daily = df.groupby([category_col, 'Day'])[QUANTITY_COL].sum().reset_index()
    
    # Calculate average per day for each category
    category_avg = category_daily.groupby(category_col)[QUANTITY_COL].mean().reset_index()
    category_avg.columns = ['Category', 'Avg_Units_per_Day']
    category_avg = category_avg.round(1)
    
    return category_avg

def summary_insights(df, analyses):
    """
    Generate key insights from the analysis
    """
    print(f"\n" + "="*60)
    print("KEY INSIGHTS")
    print("="*60)
    
    # Best and worst performing days
    daily = analyses['Daily']
    best_day_idx = daily['Lines_Total'].idxmax()
    worst_day_idx = daily['Lines_Total'].idxmin()
    print(f"Highest activity day: {daily.loc[best_day_idx, 'Period']} ({daily.loc[best_day_idx, 'Lines_Total']} lines)")
    print(f"Lowest activity day: {daily.loc[worst_day_idx, 'Period']} ({daily.loc[worst_day_idx, 'Lines_Total']} lines)")
    
    # UPDATED: Daily quantity insights for store preparation using statistics across days
    daily_stats = analyses['Daily_Statistics']
    print(f"\n📦 Daily Statistics for Store Preparation:")
    for _, row in daily_stats.iterrows():
        print(f"  {row['Metric']}: {row['Value']}")
    
    # NEW: Weekday statistics for store preparation (simplified)
    weekday_stats = analyses['Weekday_Statistics']
    print(f"\n📅 Weekday Statistics for Store Preparation:")
    for _, row in weekday_stats.iterrows():
        print(f"  {row['Weekday']}: Avg {row['Avg_Units']} units, P90 {row['P90_Units']} units ({row['Days_Count']} days)")
    
    # Weekday patterns (showing individual weekdays totals)
    weekday = analyses['Weekday']
    best_weekday_idx = weekday['Lines_Total'].idxmax()
    worst_weekday_idx = weekday['Lines_Total'].idxmin()
    print(f"\nIndividual Weekday Totals (across entire period):")
    print(f"Most active weekday: {weekday.loc[best_weekday_idx, 'Period']} ({weekday.loc[best_weekday_idx, 'Lines_Total']} total lines)")
    print(f"Least active weekday: {weekday.loc[worst_weekday_idx, 'Period']} ({weekday.loc[worst_weekday_idx, 'Lines_Total']} total lines)")
    
    # Quantity insights
    print(f"\nOverall Quantity Analysis:")
    print(f"Average order quantity: {df[QUANTITY_COL].mean():.1f} units")
    print(f"Largest single order: {df[QUANTITY_COL].max()} units")
    
    # Invoice insights
    avg_lines_per_invoice = df.groupby(INVOICE_COL).size().mean()
    print(f"Average lines per invoice: {avg_lines_per_invoice:.1f}")
    
    # NEW: Category analysis (if available)
    if 'Category_Analysis' in analyses:
        category_analysis = analyses['Category_Analysis']
        print(f"\n📂 Category Analysis - Average Units per Day:")
        for _, row in category_analysis.iterrows():
            print(f"  {row['Category']}: {row['Avg_Units_per_Day']} units/day")

def export_analysis_to_excel(df, analyses):
    """
    Export all analysis results to Excel in memory and return as bytes
    """
    print(f"\n📊 Generating analysis results in memory...")
    
    # Create a BytesIO buffer to store the Excel file in memory
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define formats for better presentation
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1
        })
        
        integer_format = workbook.add_format({
            'num_format': '#,##0',
            'border': 1
        })
        
        date_format = workbook.add_format({
            'num_format': 'yyyy-mm-dd',
            'border': 1
        })
        
        # Sheet 1: Summary Overview
        summary_data = {
            'Metric': [
                'Total Lines (Orders)',
                'Date Range Start',
                'Date Range End',
                'Total Days Analyzed',
                'Unique Invoices',
                'Unique Products',
                'Total Quantity (Units)',
                'Average Quantity per Line',
                'Average Lines per Invoice',
                'Most Active Day',
                'Least Active Day',
                'Generated On'
            ],
            'Value': [
                f"{len(df):,}",
                df[DATE_COL].min().strftime('%Y-%m-%d'),
                df[DATE_COL].max().strftime('%Y-%m-%d'),
                f"{(df[DATE_COL].max() - df[DATE_COL].min()).days + 1:,}",
                f"{df[INVOICE_COL].nunique():,}",
                f"{df[PRODUCT_COL].nunique():,}",
                f"{df[QUANTITY_COL].sum():,}",
                f"{df[QUANTITY_COL].mean():.2f}",
                f"{df.groupby(INVOICE_COL).size().mean():.2f}",
                analyses['Daily'].loc[analyses['Daily']['Lines_Total'].idxmax(), 'Period'],
                analyses['Daily'].loc[analyses['Daily']['Lines_Total'].idxmin(), 'Period'],
                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        worksheet = writer.sheets['Summary']
        worksheet.set_column('A:A', 25, header_format)
        worksheet.set_column('B:B', 20)
        
        # Export all analyses
        analysis_sheets = {
            'Daily': 'Daily Analysis',
            'Weekly': 'Weekly Analysis',
            'Monthly': 'Monthly Analysis', 
            'Weekday': 'Weekday Analysis',
            'Quarterly': 'Quarterly Analysis',
            'Daily_Statistics': 'Daily Statistics Summary',
            'Weekday_Statistics': 'Weekday Statistics Summary'
        }
        
        # Add category analysis if available
        if 'Category_Analysis' in analyses:
            analysis_sheets['Category_Analysis'] = 'Category Analysis'
        
        for analysis_key, sheet_name in analysis_sheets.items():
            df_to_export = analyses[analysis_key].copy()
            df_to_export.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            worksheet.set_row(0, None, header_format)
            for col_num, col_name in enumerate(df_to_export.columns):
                if 'Period' in col_name and analysis_key == 'Daily':
                    worksheet.set_column(col_num, col_num, 12, date_format)
                elif any(x in col_name for x in ['Count', 'Total', 'Unique']):
                    worksheet.set_column(col_num, col_num, 12, integer_format)
                else:
                    worksheet.set_column(col_num, col_num, 12, number_format)
        
        # Product Analysis
        product_analysis = df.groupby(PRODUCT_COL).agg({
            QUANTITY_COL: ['sum', 'count', 'mean'],
            INVOICE_COL: 'nunique'
        }).round(2)
        
        product_analysis.columns = ['Total_Quantity', 'Total_Lines', 'Avg_Quantity_per_Line', 'Unique_Invoices']
        product_analysis = product_analysis.sort_values('Total_Quantity', ascending=False).reset_index()
        product_analysis.to_excel(writer, sheet_name='Product Analysis', index=False)
        
        # Invoice Analysis
        invoice_analysis = df.groupby(INVOICE_COL).agg({
            QUANTITY_COL: ['sum', 'count', 'mean'],
            PRODUCT_COL: 'nunique',
            DATE_COL: ['min', 'max']
        }).round(2)
        
        invoice_analysis.columns = ['Total_Quantity', 'Total_Lines', 'Avg_Quantity_per_Line', 
                                   'Unique_Products', 'First_Date', 'Last_Date']
        invoice_analysis = invoice_analysis.sort_values('Total_Quantity', ascending=False).reset_index()
        invoice_analysis.to_excel(writer, sheet_name='Invoice Analysis', index=False)
        
        # Raw Data Sample
        raw_data_sample = df.head(1000).copy()
        raw_data_sample.to_excel(writer, sheet_name='Raw Data Sample', index=False)
    
    # Get the value from the BytesIO buffer
    output.seek(0)
    excel_bytes = output.getvalue()
    
    print(f"✅ Analysis generated successfully in memory!")
    sheets_created = "Summary, Daily/Weekly/Monthly/Weekday/Quarterly Analysis, Daily & Weekday Statistics"
    if 'Category_Analysis' in analyses:
        sheets_created += ", Category Analysis"
    sheets_created += ", Product Analysis, Invoice Analysis, Raw Data Sample"
    print(f"📋 Sheets created: {sheets_created}")
    
    return excel_bytes

# ================================
# STREAMLIT APP
# ================================
st.set_page_config(
    page_title="Procesador de Archivos",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown('<h1 class="main-title">📊 Procesador de Archivos</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Carga tu archivo Excel o CSV y procesa las columnas de manera eficiente</p>', unsafe_allow_html=True)

# Initialize session state variables
if 'df' not in st.session_state:
    st.session_state.df = None
if 'sheet_names' not in st.session_state:
    st.session_state.sheet_names = None
if 'uploaded_file_name' not in st.session_state:
    st.session_state.uploaded_file_name = None

st.markdown('<div class="upload-container">', unsafe_allow_html=True)
st.markdown("### Cargar Archivo (Excel o CSV)")
uploaded_file = st.file_uploader(
    "Selecciona un archivo Excel (.xlsx, .xls) o CSV (.csv)",
    type=['xlsx', 'xls', 'csv'],
    help="Sube tu archivo para ver una vista previa"
)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file is not None:
    # Check if this is a new file (different from the one stored in session state)
    if st.session_state.uploaded_file_name != uploaded_file.name:
        # Reset session state for new file
        st.session_state.df = None
        st.session_state.sheet_names = None
        st.session_state.uploaded_file_name = uploaded_file.name
        
        # Read the new file
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            excel_file = pd.ExcelFile(uploaded_file)
            st.session_state.sheet_names = excel_file.sheet_names
            # Don't read data yet, wait for sheet selection
        else:
            # For CSV, read immediately
            st.session_state.df = pd.read_csv(uploaded_file)
    
    # Handle Excel files
    if uploaded_file.name.endswith(('.xlsx', '.xls')):
        if st.session_state.sheet_names:
            selected_sheet = st.selectbox("Seleccionar hoja:", st.session_state.sheet_names)
            
            # Only read data if we don't have it or sheet changed
            if selected_sheet and (st.session_state.df is None or 
                                   getattr(st.session_state, 'current_sheet', None) != selected_sheet):
                excel_file = pd.ExcelFile(uploaded_file)
                st.session_state.df = excel_file.parse(selected_sheet)
                st.session_state.current_sheet = selected_sheet
            
            if st.session_state.df is not None:
                st.write("Vista previa de datos:")
                st.dataframe(st.session_state.df.head())
                
                st.markdown("### Seleccionar Columnas")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    DATE_COL = st.selectbox(
                        "Columna de fecha:",
                        st.session_state.df.columns,
                    )
                    PRODUCT_COL = st.selectbox(
                        "Columna de SKU:",
                        st.session_state.df.columns,
                    )
                
                with col2:
                    INVOICE_COL = st.selectbox(
                        "Columna de factura:",
                        st.session_state.df.columns,
                    )
                    QUANTITY_COL = st.selectbox(
                        "Columna de cantidad:",
                        st.session_state.df.columns,
                    )

                with col3:
                    CATEGORY_COL = st.selectbox(
                        "Columna de categoría (opcional):",
                        [None] + list(st.session_state.df.columns),
                        help="Selecciona una columna de categoría si está disponible"
                    )
                
                if st.button("Generar Reporte"):
                    with st.spinner('Generando Reporte...'):
                        df_processed = load_and_prepare_data(st.session_state.df)
                        analyses = comprehensive_analysis(df_processed, CATEGORY_COL)

                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_filename = f'orders_analysis_results_{timestamp}.xlsx'

                        # Get Excel file as bytes (in memory)
                        excel_bytes = export_analysis_to_excel(df_processed, analyses)

                        st.success(f"📊 Reporte generado exitosamente!")

                        st.download_button(
                            label="📥 Descargar Archivo Excel",
                            data=excel_bytes,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

    # Handle CSV files
    else:
        if st.session_state.df is not None:
            st.write("Vista previa de datos:")
            st.dataframe(st.session_state.df.head())
            
            st.markdown("### Seleccionar Columnas")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                DATE_COL = st.selectbox(
                    "Columna de fecha:",
                    st.session_state.df.columns,
                )
                PRODUCT_COL = st.selectbox(
                    "Columna de SKU:",
                    st.session_state.df.columns,
                )
            
            with col2:
                INVOICE_COL = st.selectbox(
                    "Columna de factura:",
                    st.session_state.df.columns,
                )
                QUANTITY_COL = st.selectbox(
                    "Columna de cantidad:",
                    st.session_state.df.columns,
                )

            with col3:
                CATEGORY_COL = st.selectbox(
                    "Columna de categoría (opcional):",
                    [None] + list(st.session_state.df.columns),
                    help="Selecciona una columna de categoría si está disponible"
                )
            
            if st.button("Generar Reporte"):
                with st.spinner('Generando Reporte...'):
                    df_processed = load_and_prepare_data(st.session_state.df)
                    analyses = comprehensive_analysis(df_processed, CATEGORY_COL)

                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f'orders_analysis_results_{timestamp}.xlsx'

                    # Get Excel file as bytes (in memory)
                    excel_bytes = export_analysis_to_excel(df_processed, analyses)

                    st.success(f"📊 Reporte generado exitosamente!")

                    st.download_button(
                        label="📥 Descargar Archivo Excel",
                        data=excel_bytes,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )