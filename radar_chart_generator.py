
import pandas as pd
import numpy as np
import io
from config import COMPETENCY_MAP, SCORE_MAP

# --- Constants ---
COMPETENCIES_FOR_CHART = [comp for _, comp, _ in COMPETENCY_MAP]
    
# --- Data Processing ---
def calculate_competency_averages(df, competency_map):
    results = {}
    for _, competency, questions in competency_map:
        valid_qs = [q for q in questions if q in df.columns]
        if valid_qs:
            avg = df[valid_qs].mean().mean()
            if pd.isna(avg):
                avg = 0
            results[competency] = avg
    return results

def create_radar_chart_report(data, writer, sheet_name):
    """Creates a worksheet with data and a radar chart."""
    df_chart = pd.DataFrame(data).round(1)
    df_chart.to_excel(writer, sheet_name=sheet_name, index=False)
    
    workbook = writer.book
    ws = writer.sheets[sheet_name]

    chart = workbook.add_chart({'type': 'radar', 'subtype': 'filled'})
    
    num_rows = len(df_chart)
    for i, col_name in enumerate(df_chart.columns[1:], 1):
        chart.add_series({
            'name':       [sheet_name, 0, i],
            'categories': [sheet_name, 1, 0, num_rows, 0],
            'values':     [sheet_name, 1, i, num_rows, i],
            'fill':       {'color': '#C6E0B4' if i == 1 else '#BDD7EE'}, # Example colors
            'border':     {'color': '#6D9B4F' if i == 1 else '#5B9BD5'},
        })

    chart.set_title({'name': f'RGB Competency Radar Chart - {sheet_name}'})
    chart.set_x_axis({'name': 'Competency'}) # For radar, this sets category labels
    chart.set_y_axis({'min': 1, 'max': 4, 'major_gridlines': {'visible': True}})
    chart.set_legend({'position': 'right'})
    chart.set_style(11) # Apply a common chart style

    ws.insert_chart('E2', chart, {'x_scale': 1.5, 'y_scale': 1.5})


def create_summary_radar_sheet(writer, df_processed):
    """Creates a summary worksheet with a single radar chart comparing all grades."""
    sheet_name = '学年別比較'
    
    # 1. Calculate averages for each grade
    grade_averages = {}
    for grade in [1, 2, 3]:
        df_grade = df_processed[df_processed['学年'] == grade]
        grade_averages[f'{grade}年'] = calculate_competency_averages(df_grade, COMPETENCY_MAP)

    # 2. Prepare data for the DataFrame
    summary_data = {'Competency': COMPETENCIES_FOR_CHART}
    for grade_name, averages in grade_averages.items():
        summary_data[grade_name] = [averages.get(c, 0) for c in COMPETENCIES_FOR_CHART]
        
    df_summary = pd.DataFrame(summary_data)
    df_summary.to_excel(writer, sheet_name=sheet_name, index=False)

    # 3. Create the multi-series radar chart
    workbook = writer.book
    ws = writer.sheets[sheet_name]
    chart = workbook.add_chart({'type': 'radar'}) # Use standard radar for overlap visibility

    num_rows = len(df_summary)
    colors = ['#4F81BD', '#C0504D', '#9ABA60'] # Colors for Grade 1, 2, 3

    for i, col_name in enumerate(df_summary.columns[1:], 1):
        chart.add_series({
            'name':       [sheet_name, 0, i],
            'categories': [sheet_name, 1, 0, num_rows, 0],
            'values':     [sheet_name, 1, i, num_rows, i],
            'line':       {'color': colors[i-1]},
        })

    chart.set_title({'name': 'RGB Competency Radar Chart - 学年別比較'})
    chart.set_y_axis({'min': 1, 'max': 4, 'major_gridlines': {'visible': True}})
    chart.set_legend({'position': 'right'})
    chart.set_style(11)

    ws.insert_chart('F2', chart, {'x_scale': 1.5, 'y_scale': 1.5})


def generate_radar_chart(df_processed):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': False}}) as writer:
        # Always generate the "Overall" sheet
        overall_avg = calculate_competency_averages(df_processed, COMPETENCY_MAP)
        chart_data_overall = {
            "Competency": COMPETENCIES_FOR_CHART,
            "R7_Overall": [overall_avg.get(c, 0) for c in COMPETENCIES_FOR_CHART]
        }
        create_radar_chart_report(chart_data_overall, writer, "Overall")

        # Only generate per-grade and summary sheets if the '学年' column exists
        if '学年' in df_processed.columns:
            # Per-grade sheets
            for grade in [1, 2, 3]:
                df_grade = df_processed[df_processed['学年'] == grade]
                if not df_grade.empty:
                    grade_avg = calculate_competency_averages(df_grade, COMPETENCY_MAP)
                    chart_data_grade = {
                        "Competency": COMPETENCIES_FOR_CHART,
                        f"R7_Grade_{grade}": [grade_avg.get(c, 0) for c in COMPETENCIES_FOR_CHART]
                    }
                    create_radar_chart_report(chart_data_grade, writer, f"Grade_{grade}")

            # Summary sheet with all grades compared
            create_summary_radar_sheet(writer, df_processed)

    output.seek(0)
    return output
