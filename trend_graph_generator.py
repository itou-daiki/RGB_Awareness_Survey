
import pandas as pd
import numpy as np
import io
from config import COMPETENCY_MAP, SCORE_MAP, HISTORICAL_BENCHMARKS

# --- Constants ---
COMPETENCIES_FOR_GRAPH = [comp for _, comp, _ in COMPETENCY_MAP]

# --- Data Processing ---
def calculate_competency_averages(df):
    results = {}
    for _, competency, questions in COMPETENCY_MAP:
        valid_qs = [q for q in questions if q in df.columns]
        if valid_qs:
            avg = df[valid_qs].mean().mean()
            if pd.isna(avg):
                avg = 0
            results[competency] = avg
    return results

def generate_trend_graph(df_current):
    output = io.BytesIO()
    
    current_averages = calculate_competency_averages(df_current)
    
    chart_data = {"Competency": COMPETENCIES_FOR_GRAPH}
    for year in ["R4", "R5", "R6"]:
        chart_data[year] = [HISTORICAL_BENCHMARKS[comp].get(year, None) for comp in COMPETENCIES_FOR_GRAPH]
    
    chart_data["R7"] = [current_averages.get(comp, None) for comp in COMPETENCIES_FOR_GRAPH]

    df_chart = pd.DataFrame(chart_data).round(1)

    with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': False}}) as writer:
        df_chart.to_excel(writer, sheet_name='TrendData', index=False)
        workbook = writer.book
        ws = writer.sheets['TrendData']

        chart = workbook.add_chart({'type': 'line'})
        
        num_rows = len(df_chart)
        # Define a color palette for the series
        colors = ['#4F81BD', '#C0504D', '#9ABA60', '#F79646', '#8064A2', '#4BACC6', '#B2A1C7', '#7F7F7F']
        
        for i, col_name in enumerate(df_chart.columns[1:], 1):
            chart.add_series({
                'name':       ['TrendData', 0, i],
                'categories': ['TrendData', 1, 0, num_rows, 0],
                'values':     ['TrendData', 1, i, num_rows, i],
                'marker':     {'type': 'circle', 'size': 5, 'fill': {'color': colors[i-1]}, 'border': {'color': colors[i-1]}},
                'line':       {'color': colors[i-1], 'width': 1.5},
            })

        chart.set_title({'name': 'RGB Competency Trends (R4-R7)', 'name_font': {'size': 14, 'bold': True}})
        chart.set_x_axis({'name': 'Competency', 'name_font': {'size': 10, 'bold': True}})
        chart.set_y_axis({'name': 'Average Score', 'name_font': {'size': 10, 'bold': True}, 'min': 1, 'max': 4, 'major_gridlines': {'visible': True}})
        chart.set_legend({'position': 'bottom', 'font': {'size': 9}})
        chart.set_plotarea({'border': {'none': True}}) # No border around plot area
        chart.set_chartarea({'border': {'none': True}}) # No border around chart area
        
        ws.insert_chart('B2', chart, {'x_scale': 1.8, 'y_scale': 1.8})
    
    output.seek(0)
    return output
