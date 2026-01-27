import pandas as pd
import numpy as np
import io
import re
from config import COMPETENCY_MAP, SCORE_MAP, ALL_QUESTIONS

def create_dashboard_sheet(writer, df_all_data):
    ws = writer.book.add_worksheet('集計結果表示')
    
    # Formats
    header_fmt = writer.book.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    category_fmt = writer.book.add_format({'bold': True, 'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    question_fmt = writer.book.add_format({'border': 1, 'text_wrap': True, 'align': 'left', 'valign': 'top'})
    num_fmt = writer.book.add_format({'num_format': '0.0', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    percent_fmt = writer.book.add_format({'num_format': '0.00"%"', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    
    # Column headers
    headers = ["大分類", "能力指標", "質問項目", "平均値", "4(%)", "3(%)", "2(%)", "1(%)"]
    ws.write_row(0, 0, headers, header_fmt)
    
    row_idx = 1
    for cat, sub, questions in COMPETENCY_MAP:
        category_start_row = row_idx
        for q_idx, q in enumerate(questions):
            # Calculate stats for the current question
            stats = {
                'avg': df_all_data[q].mean() if q in df_all_data.columns else np.nan,
                'dist': (df_all_data[q].value_counts(normalize=True) * 100) if q in df_all_data.columns else pd.Series()
            }
            if pd.isna(stats['avg']):
                stats['avg'] = 0
            
            ws.write(row_idx, 2, q, question_fmt) # Question item
            ws.write(row_idx, 3, stats['avg'], num_fmt) # Average
            ws.write(row_idx, 4, stats['dist'].get(4, 0), percent_fmt) # 4(%)
            ws.write(row_idx, 5, stats['dist'].get(3, 0), percent_fmt) # 3(%)
            ws.write(row_idx, 6, stats['dist'].get(2, 0), percent_fmt) # 2(%)
            ws.write(row_idx, 7, stats['dist'].get(1, 0), percent_fmt) # 1(%)
            
            row_idx += 1
        
        # Merge cells for "大分類" and "能力指標"
        if row_idx > category_start_row:
            ws.merge_range(category_start_row, 0, row_idx - 1, 0, cat, category_fmt)
            ws.merge_range(category_start_row, 1, row_idx - 1, 1, sub, category_fmt)

    # Set column widths
    ws.set_column('A:A', 15)  # 大分類
    ws.set_column('B:B', 15)  # 能力指標
    ws.set_column('C:C', 60)  # 質問項目
    ws.set_column('D:D', 10)  # 平均値
    ws.set_column('E:H', 8)   # Percentages


def create_grade_report(df_all_data, grade_name, grade_filter, writer, survey_period):
    """Generates an Excel report for a specific grade for a given survey period."""
    if grade_filter != '全体':
        df_target = df_all_data[df_all_data['学年'] == grade_filter].copy()
    else:
        df_target = df_all_data.copy()

    # Extract round name like "第二回" from "9月(第二回)"
    round_name_match = re.search(r'（(.*?)）', survey_period)
    round_name = round_name_match.group(1) if round_name_match else survey_period

    sheet_name = f'{round_name}{grade_name}'
    ws = writer.book.add_worksheet(sheet_name)

    # --- 1. Summary Statistics ---
    summary_data = []
    for q in ALL_QUESTIONS:
        if q in df_target.columns:
            counts = df_target[q].value_counts()
            percentages = df_target[q].value_counts(normalize=True) * 100
            total_responses = df_target[q].count()
            avg = df_target[q].mean()
            if pd.isna(avg):
                avg = 0

            summary_data.append({
                "question": q,
                "count_4": counts.get(4, 0),
                "count_3": counts.get(3, 0),
                "count_2": counts.get(2, 0),
                "count_1": counts.get(1, 0),
                "total": total_responses,
                "percent_4": percentages.get(4, 0),
                "percent_3": percentages.get(3, 0),
                "percent_2": percentages.get(2, 0),
                "percent_1": percentages.get(1, 0),
                "average": avg
            })
    
    df_summary = pd.DataFrame(summary_data)
    
    # --- Writing to Excel ---
    # Header formats
    bold_fmt = writer.book.add_format({'bold': True})
    percent_fmt = writer.book.add_format({'num_format': '0.00"%"'})
    avg_fmt = writer.book.add_format({'num_format': '0.00'})

    # Write summary headers
    ws.write(0, 0, '人数', bold_fmt)
    ws.write(1, 1, 'とてもそう思う')
    ws.write(2, 1, 'どちらかといえばそう思う')
    ws.write(3, 1, 'どちらかといえばそう思わない')
    ws.write(4, 1, 'そう思わない')
    ws.write(5, 1, '回答人数', bold_fmt)

    ws.write(7, 0, '割合', bold_fmt)
    ws.write(8, 1, 'とてもそう思う')
    ws.write(9, 1, 'どちらかといえばそう思う')
    ws.write(10, 1, 'どちらかといえばそう思わない')
    ws.write(11, 1, 'そう思わない')

    ws.write(13, 0, '4件法による平均値', bold_fmt)

    # Write summary data
    for i, row in df_summary.iterrows():
        col = i + 2
        ws.write(1, col, row['count_4'])
        ws.write(2, col, row['count_3'])
        ws.write(3, col, row['count_2'])
        ws.write(4, col, row['count_1'])
        ws.write(5, col, row['total'])

        ws.write(8, col, row['percent_4'], percent_fmt)
        ws.write(9, col, row['percent_3'], percent_fmt)
        ws.write(10, col, row['percent_2'], percent_fmt)
        ws.write(11, col, row['percent_1'], percent_fmt)
        
        ws.write(13, col, row['average'], avg_fmt)
        # Write question text at the bottom
        ws.write(15, col, row['question'], writer.book.add_format({'text_wrap': True}))

    # --- 2. Raw Data ---
    # Write raw data starting from row 17
    df_target.to_excel(writer, sheet_name=sheet_name, startrow=17, index=False)
    
    # Adjust column widths
    ws.set_column('A:B', 15)
    for i in range(len(ALL_QUESTIONS)):
        ws.set_column(i+2, i+2, 15)


def generate_grade_reports(df_processed, survey_period):
    """
    Generates all grade-specific reports and returns them as a dictionary
    of in-memory Excel files.
    """
    # Proactive check: If '学年' column doesn't exist, no grade reports can be generated.
    if '学年' not in df_processed.columns:
        return {}

    reports = {}
    grades = {
        "1年": 1,
        "2年": 2,
        "3年": 3,
        "全体": "全体"
    }

    for name, grade_filter in grades.items():
        # For individual grades, check if there is any data before creating a report
        if grade_filter != '全体' and df_processed[df_processed['学年'] == grade_filter].empty:
            continue
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': False}}) as writer:
            # Create the grade-specific sheet
            create_grade_report(df_processed, name, grade_filter, writer, survey_period)
            # Create the dashboard sheet
            create_dashboard_sheet(writer, df_processed)
        
        output.seek(0)
        reports[name] = output

    return reports