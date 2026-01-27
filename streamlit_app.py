import streamlit as st
import pandas as pd
import io
import re

print("--- [EXECUTION] Running latest streamlit_app.py ---")

# Import the refactored generators and the new unified preprocessor
from data_processor import preprocess_data
from report_1_generator import generate_report_one
from radar_chart_generator import generate_radar_chart
from trend_graph_generator import generate_trend_graph
from grade_reports_generator import generate_grade_reports

st.set_page_config(layout="wide")

st.title("🎓 RGB意識調査 統合レポート生成システム")

# --- Sidebar for controls ---
st.sidebar.header("設定")
uploaded_file = st.sidebar.file_uploader("① アンケート結果Excelをアップロード", type=["xlsx"])
current_survey = st.sidebar.selectbox(
    "② 調査時期を選択",
    ["4月(第一回)", "9月(第二回)", "1月(第三回)"],
    index=1  # Default to 9月(第二回)
)

# --- Main app body ---
if uploaded_file is None:
    st.info("サイドバーからExcelファイルをアップロードし、調査時期を選択してください。")
else:
    try:
        # Preprocess the data and store it in the session state to avoid reprocessing
        if 'df_processed' not in st.session_state or st.session_state.get('uploaded_filename') != uploaded_file.name:
            with st.spinner("ファイルを読み込み、前処理を実行中..."):
                df_raw = pd.read_excel(uploaded_file)
                st.session_state['df_processed'] = preprocess_data(df_raw.copy())
                st.session_state['uploaded_filename'] = uploaded_file.name
                # Clear old reports when a new file is uploaded
                st.session_state['reports_generated'] = False
                st.success("ファイルの準備が完了しました。")
        
        df_processed = st.session_state['df_processed']
        
        st.header("レポートの一括生成")
        st.write(f"**調査時期:** `{current_survey}`")
        
        if st.button("全レポートを一括生成", type="primary"):
            with st.spinner("すべてのレポートを生成中です... これには数秒かかる場合があります。"):
                # 1. Generate all reports in memory, passing the survey period where needed
                print(f"--- [CALL] About to call generate_report_one with survey_period='{current_survey}' ---")
                st.session_state['report_one_bytes'] = generate_report_one(df_processed, current_survey)
                st.session_state['radar_chart_bytes'] = generate_radar_chart(df_processed)
                st.session_state['trend_graph_bytes'] = generate_trend_graph(df_processed)
                st.session_state['grade_reports'] = generate_grade_reports(df_processed, current_survey)
                st.session_state['reports_generated'] = True
                st.session_state['generated_for_survey'] = current_survey # Store which survey was generated

        # Display download buttons only after generation is complete for the current survey
        if st.session_state.get('reports_generated') and st.session_state.get('generated_for_survey') == current_survey:
            st.markdown("---")
            st.header(f"生成されたレポート (`{current_survey}`)")

            # --- Dynamically create filenames ---
            # Extract month like "9月" from "9月(第二回)"
            month_match = re.match(r'(\d+月)', current_survey)
            month_str = month_match.group(1) if month_match else "UnknownMonth"

            # Create two columns for better layout
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("会議資料")
                st.download_button(
                    label="【その１】質問項目と表",
                    data=st.session_state['report_one_bytes'],
                    file_name=f"【その１データ】 RGB意識調査の質問項目と表(職員会議用）.xlsx", # This filename seems static
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn1"
                )
                st.download_button(
                    label="【その２】RGBレーダーチャート",
                    data=st.session_state['radar_chart_bytes'],
                    file_name=f"【その２データ】RGBレーダーチャート（R7職員会議資料用）.xlsx", # This filename also seems static
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn2"
                )
                st.download_button(
                    label="【その３】RGB推移グラフ",
                    data=st.session_state['trend_graph_bytes'],
                    file_name=f"【その３データ】【R3～R7】RGB推移グラフ（R7職員会議用）.xlsx", # This filename also seems static
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn3"
                )
            
            with col2:
                st.subheader("学年別 詳細レポート")
                grade_reports = st.session_state['grade_reports']
                for i, (name, report_bytes) in enumerate(grade_reports.items()):
                    st.download_button(
                        label=f"【{name}】結果（分布あり）",
                        data=report_bytes,
                        file_name=f"1.RGB意識調査R7.{month_str}結果（{name}・分布あり）.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"btn_grade_{i}"
                    )

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
        # Clear session state on error to allow for a fresh start
        for key in list(st.session_state.keys()):
            del st.session_state[key]