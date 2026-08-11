import streamlit as st
import pandas as pd
import datetime
import calendar
import matplotlib.pyplot as plt

# 1. 網頁頁面設定
st.set_page_config(page_title="個人雲端記帳管家", page_icon="💰", layout="wide")

# 設定 Matplotlib 中文字型支援
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial Unicode MS', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

st.title("💰 個人雲端記帳與預算監控 App")
st.write("無需安裝軟體！上傳您的 Excel 記帳檔案，即可自動分析當月花費與警報。")

# 2. 側邊欄設定與檔案上傳
st.sidebar.header("⚙️ 控制面板")
uploaded_file = st.sidebar.file_uploader("上傳 Excel 記帳檔 (.xlsx)", type=["xlsx"])
current_balance = st.sidebar.number_input("目前銀行帳戶餘額 (NTD)", value=25000, step=1000)
analysis_date = st.sidebar.date_input("分析基準日期", datetime.date.today())

if uploaded_file is not None:
    try:
        # 讀取 Excel 兩個 Sheet
        df_budget = pd.read_excel(uploaded_file, sheet_name="預算設定")
        df_trans = pd.read_excel(uploaded_file, sheet_name="收支紀錄")
        df_trans['日期'] = pd.to_datetime(df_trans['日期'])

        # 當月資料過濾
        df_month = df_trans[
            (df_trans['日期'].dt.year == analysis_date.year) & 
            (df_trans['日期'].dt.month == analysis_date.month)
        ]
        df_expense = df_month[df_month['收支類型'] == '支出']
        actual_spend = df_expense.groupby('分類名稱')['金額'].sum()

        # 頂部關鍵指標 (Metrics)
        total_income = df_month[df_month['收支類型'] == '收入']['金額'].sum()
        total_expense = actual_spend.sum()
        
        col1, col2, col3 = st.columns(3)
        col1.metric("當月總收入", f"${total_income:,.0f}")
        col2.metric("當月總支出", f"${total_expense:,.0f}")
        col3.metric("當月淨結餘", f"${total_income - total_expense:,.0f}")

        st.markdown("---")

        # 主要內容區域：左圖表、右明細
        left_col, right_col = st.columns([5, 5])

        with left_col:
            st.subheader("📊 花費分類佔比圓餅圖")
            if not df_expense.empty:
                fig, ax = plt.subplots(figsize=(6, 6))
                ax.pie(
                    actual_spend,
                    labels=actual_spend.index,
                    autopct='%1.1f%%',
                    startangle=140,
                    colors=['#1a73e8', '#34a853', '#fbbc04', '#ea4335', '#a142f4', '#ff6d01']
                )
                st.pyplot(fig)
            else:
                st.info("當月尚無支出紀錄。")

        with right_col:
            st.subheader("📋 各分類預算監控")
            _, total_days = calendar.monthrange(analysis_date.year, analysis_date.month)
            time_progress_pct = round((analysis_date.day / total_days) * 100, 1)
            
            st.caption(f"🗓️ 當月時間進度：第 {analysis_date.day}/{total_days} 天 ({time_progress_pct}%)")

            report = []
            pending_fixed_amount = 0

            for _, row in df_budget.iterrows():
                cat = row['分類名稱']
                budget = row['預算金額']
                is_fixed = (row['支出類型'] == '固定')
                pay_day = row['每月扣款日'] if pd.notna(row['每月扣款日']) else None
                
                spent = actual_spend.get(cat, 0)
                spend_pct = (spent / budget * 100) if budget > 0 else 0
                
                if is_fixed:
                    if spent > 0:
                        status = "✅ 已扣款"
                    else:
                        pending_fixed_amount += budget
                        status = "⏳ 待扣款"
                else:
                    target_today = budget * (analysis_date.day / total_days)
                    status = "⚠️ 燒錢過快" if spent > target_today * 1.15 else "🟢 正常"

                report.append({
                    "分類": cat,
                    "類型": "固定" if is_fixed else "變動",
                    "預算": f"${budget:,}",
                    "實際": f"${spent:,}",
                    "使用率": f"{spend_pct:.1f}%",
                    "狀態": status
                })

            st.dataframe(pd.DataFrame(report), use_container_width=True)

            # 預留資金缺口警報
            st.subheader("🏦 銀行帳戶預留資金試算")
            gap = pending_fixed_amount - current_balance
            st.write(f"👉 本月剩餘待扣固定支出總額：**${pending_fixed_amount:,}**")
            
            if gap > 0:
                st.error(f"🚨【預留資金不足】請最晚在扣款日前補入 **${gap:,}**！")
            else:
                st.success(f"🟢【資金充足】扣除剩餘固定支出後，預估還剩 **${abs(gap):,}**。")

    except Exception as e:
        st.error(f"讀取檔案失敗，請確保 Excel 格式包含「預算設定」與「收支紀錄」工作表：{e}")
else:
    st.info("👈 請點擊左側邊欄的「上傳 Excel 記帳檔」開始使用！")
