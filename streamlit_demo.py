import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.markdown("## 📈 MSD 點擊排行分析")
st.info("請同時上傳 7天 與 14天 的 Excel 檔案，點擊分析後可直接下載結果。")

uploaded_files = st.file_uploader("", 
                                type=["xlsx"], 
                                help="請同時上傳兩個檔案，否則無法進行分析",
                                accept_multiple_files=True)

if uploaded_files:
    if len(uploaded_files) < 2:
        st.warning("請一次上傳兩個 Excel 檔案")
    elif len(uploaded_files) > 2:
        st.warning("請一次上傳兩個 Excel 檔案")
    else:
        st.markdown("### ⬆️ 已成功上傳檔案")
        st.success("請點擊下方按鈕進行分析。")

        if st.button("開始分析"):
            try:
                with st.spinner("資料分析中..."):
                    p1, p2 = uploaded_files

                    p1_df = pd.read_excel(p1)
                    p2_df = pd.read_excel(p2)

                    # 依照第一個數字得知是 14days 還是 7days
                    if p1_df.iloc[0, 1] > p2_df.iloc[0, 1]:
                        big, small = p1, p2
                    else:
                        big, small = p2, p1

                    # 數字較大的檔案即為 14days 的資料, 較小的為 7days 
                    df7 = pd.read_excel(small)
                    df14 = pd.read_excel(big)

                    # 進行資料分析
                    df7  = df7.rename(columns={df7.columns[1] : 'click_7days'})
                    df14 = df14.rename(columns={df14.columns[1]: 'click_14days'})

                    df = pd.merge(df7[['label','click_7days']],
                                df14[['label','click_14days']],
                                on='label', how='outer').fillna(0)

                    df['click_prev7'] = (df['click_14days'] - df['click_7days']).clip(lower=0)
                    df['rank_after']  = df['click_7days'].rank(method='min', ascending=False).astype(int)
                    df['rank_before'] = df['click_prev7'].rank(method='min', ascending=False).astype(int)
                    df['rank_change'] = df['rank_before'] - df['rank_after']

                    result_sorted = df[['label','click_prev7','click_7days','rank_before','rank_after','rank_change']] \
                                    .sort_values('rank_change', ascending=False)

                    unique_changes = sorted(df['rank_change'].unique(), reverse=True)
                    top3_vals = unique_changes[:3]
                    bottom3_vals = unique_changes[-3:]

                    result_sorted = result_sorted.rename(columns={
                        "label":"診所名稱",
                        "click_prev7":"上週點擊",
                        "click_7days":"本週點擊",
                        "rank_before":"上週排名",
                        "rank_after":"本週排名",
                        "rank_change":"排名變化"
                    })



                # 即將輸出的檔案名稱, 需包含當下時間
                now = datetime.now()
                time_str = now.strftime("%m.%d-%H.%M")     
                result_output_path = F'{time_str}-Lndata-MSD對比結果.xlsx'

                output = BytesIO()

                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook  = writer.book

                    result_sorted.to_excel(writer, sheet_name='All_Clinics', index=False)
                    worksheet1 = writer.sheets['All_Clinics']
                    yellow_fmt = workbook.add_format({'bg_color': '#FFFF00'})
                    red_fmt    = workbook.add_format({'bg_color': '#FF0000'})
                    nrows = len(result_sorted)
                    if top3_vals:
                        formula_up = 'OR(' + ','.join(f'$F2={v}' for v in top3_vals) + ')'
                        worksheet1.conditional_format(f'A2:A{nrows+1}', {
                            'type': 'formula',
                            'criteria': formula_up,
                            'format': yellow_fmt})
                    if bottom3_vals:
                        formula_down = 'OR(' + ','.join(f'$F2={v}' for v in bottom3_vals) + ')'
                        worksheet1.conditional_format(f'A2:A{nrows+1}', {
                            'type': 'formula',
                            'criteria': formula_down,
                            'format': red_fmt})
                output.seek(0)
                st.success(f"分析完成，下載檔案：{result_output_path}")
                st.download_button(
                    label="✅下載分析結果",
                    data=output,
                    file_name=result_output_path,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            except Exception as e:
                st.error(f"分析失敗：請確認上傳檔案是否正確")
