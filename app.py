import io
import os
import urllib.request
import calendar
import datetime
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# 1. 網頁頁面設定
st.set_page_config(page_title="個人雲端記帳管家", page_icon="💰", layout="wide")

# ----------------- 中文字型自動註冊 -----------------
font_path = "NotoSansTC-Regular.ttf"
if not os.path.exists(font_path):
    try:
        font_url = "https://raw.githubusercontent.com/google/fonts/main/ofl/notosanstc/NotoSansTC-Regular.ttf"
        urllib.request.urlretrieve(font_url, font_path)
    except Exception:
        pass

if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.sans-serif'] = ['Noto Sans TC', 'sans-serif']
else:
    for sys_font in [
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc'
    ]:
        if os.path.exists(sys_font):
            fm.fontManager.addfont(sys_font)
            plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei', 'sans-serif']
            break

plt.rcParams['axes.unicode_minus'] = False
# ----------------------------------------------------

# 2. 初始化 Session State 數據 (若無資料則建立預設)
if "df_budget" not in st.session_state:
    st.session_state.df_budget = pd.DataFrame({
        "分類名稱": ["居住房租", "水電瓦斯", "電信網路", "飲食餐飲", "交通通勤", "娛樂休閒", "日常雜項"],
        "預算金額": [11000, 1000, 500, 12000, 2800, 4000, 5000],
        "支出類型": ["固定", "固定", "固定", "變動", "變動", "變動", "變動"],
        "每月扣款日": [20, 20, 20, None, None, None, None]
    })

if "df_trans" not in st.session_state:
    st.session_state.df_trans = pd.DataFrame({
        "日期": [
            pd.to_datetime("2026-07-15"), pd.to_datetime("2026-07-20"), # 跨月歷史資料測試
            pd.to_datetime("2026-08-01"), pd.to_datetime("2026-08-05"), 
            pd.to_datetime("2026-08-05"), pd.to_datetime("2026-08-06")
        ],
        "收支類型": ["收入", "支出", "支出", "支出", "收入", "支出"],
        "分類名稱": ["薪資", "日常雜項", "飲食餐飲", "居住房租", "薪資", "飲食餐飲"],
        "金額": [0, 0, 0, 0, 0, 0],
        "備註": ["7月薪資", "購買家電", "午餐外帶", "8月房租", "8月薪資入帳", "朋友聚餐"]
    })

# 3. 匯出 Excel 備份檔
def export_to_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.df_budget.to_excel(writer, sheet_name='預算設定', index=False)
        st.session_state.df_trans.to_excel(writer, sheet_name='收支紀錄', index=False)
    return output.getvalue()

st.title("💰 個人雲端記帳與預算監控 App")

# 4. 側邊欄控制面板
st.sidebar.header("⚙️ 控制面板")

# 起始底金設定（只需設定一次，系統將自動以此為基礎滾動加總）
initial_balance = st.sidebar.number_input(
    "帳戶起始底金 (NTD)", 
    value=20000, 
    step=1000, 
    help="設定開始記帳前帳戶內的初始金額，系統會自動加減歷史所有收支進行滾動累積。"
)

analysis_date = st.sidebar.date_input("分析基準日期", datetime.date.today())

st.sidebar.markdown("---")
st.sidebar.markdown("### 💾 備份資料")
excel_bytes = export_to_excel()
st.sidebar.download_button(
    label="📥 下載當前資料為 Excel 檔",
    data=excel_bytes,
    file_name="my_budget_backup.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# 5. 主分頁切換
tab1, tab2 = st.tabs(["📊 分析儀表板", "✍️ 線上記帳與預算編輯"])

# ==================== Tab 1: 分析儀表板 ====================
with tab1:
    df_budget = st.session_state.df_budget
    df_trans = st.session_state.df_trans
    
    # 轉換日期格式
    df_trans['日期_dt'] = pd.to_datetime(df_trans['日期'])
    
    # ---------------- 核心滾動邏輯運算 ----------------
    # A. 截至「分析基準日」的全歷史累積收支
    df_history = df_trans[df_trans['日期_dt'].dt.date <= analysis_date]
    total_cum_income = df_history[df_history['收支類型'] == '收入']['金額'].sum()
    total_cum_expense = df_history[df_history['收支類型'] == '支出']['金額'].sum()
    
    # 截至分析基準日的即時總餘額 (起始底金 + 歷史總收入 - 歷史總支出)
    current_total_balance = initial_balance + total_cum_income - total_cum_expense

    # B. 當選定月份專屬收支分析
    df_month = df_trans[
        (df_trans['日期_dt'].dt.year == analysis_date.year) & 
        (df_trans['日期_dt'].dt.month == analysis_date.month)
    ]
    month_income = df_month[df_month['收支類型'] == '收入']['金額'].sum()
    df_expense = df_month[df_month['收支類型'] == '支出']
    month_expense = df_expense['金額'].sum()
    month_net_savings = month_income - month_expense  # 當月淨結餘

    actual_spend = df_expense.groupby('分類名稱')['金額'].sum()

    # C. 時間進度與預測
    _, total_days = calendar.monthrange(analysis_date.year, analysis_date.month)
    current_day = analysis_date.day
    time_progress_ratio = current_day / total_days
    time_progress_pct = round(time_progress_ratio * 100, 1)

    total_planned_budget = df_budget['預算金額'].sum()

    projected_total_expense = 0
    report_data = []
    pending_fixed_amount = 0

    for _, row in df_budget.iterrows():
        cat = row['分類名稱']
        budget = row['預算金額']
        is_fixed = (row['支出類型'] == '固定')
        
        spent = actual_spend.get(cat, 0)
        diff = budget - spent
        
        if is_fixed:
            proj = spent if spent > 0 else budget
            if spent > 0:
                status = "✅ 已扣款"
            else:
                pending_fixed_amount += budget
                status = "⏳ 待扣款"
        else:
            proj = (spent / current_day * total_days) if current_day > 0 else 0
            target_today = budget * time_progress_ratio
            if spent > budget:
                status = "🚨 已透支"
            elif spent > target_today * 1.15:
                status = "⚠️ 燒錢過快"
            else:
                status = "🟢 正常"

        projected_total_expense += proj
        report_data.append({
            "分類名稱": cat,
            "支出類型": "固定" if is_fixed else "變動",
            "月初預估預算": budget,
            "目前實際花費": spent,
            "預算差額": diff,
            "預估月底花費": round(proj),
            "狀態": status
        })

    df_report = pd.DataFrame(report_data)

    # 頂部 4 大核心 KPI
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    kpi_col1.metric("💵 截至當前帳戶總餘額", f"${current_total_balance:,.0f}", help="起始底金 + 歷年歷月截至該日總收入 - 總支出")
    kpi_col2.metric("📈 當月累積淨結餘", f"${month_net_savings:,.0f}", delta=f"收入 ${month_income:,.0f} | 支出 ${month_expense:,.0f}")
    kpi_col3.metric("🎯 當月預估月底總花費", f"${projected_total_expense:,.0f}", delta=f"${total_planned_budget - projected_total_expense:,.0f} 預算差額")
    kpi_col4.metric("🗓️ 當月時間進度", f"{time_progress_pct}%", f"第 {current_day}/{total_days} 天")

    st.markdown("---")

    # 圖表展現
    st.subheader(f"📊 {analysis_date.year} 年 {analysis_date.month} 月 預算 vs. 實際花費比較圖")
    if not df_report.empty:
        fig, ax = plt.subplots(figsize=(10, 4))
        categories = df_report['分類名稱']
        x = range(len(categories))
        width = 0.35

        ax.bar([p - width/2 for p in x], df_report['月初預估預算'], width, label='月初預估預算', color='#e0e0e0')
        colors = ['#ea4335' if "透支" in r['狀態'] or "燒錢" in r['狀態'] else '#34a853' for _, r in df_report.iterrows()]
        ax.bar([p + width/2 for p in x], df_report['目前實際花費'], width, label='目前實際花費', color=colors)
        ax.plot([p + width/2 for p in x], df_report['預估月底花費'], "r--o", label='預估月底總花費')

        ax.set_xticks(x)
        ax.set_xticklabels(categories, rotation=15)
        ax.set_ylabel("金額 (NTD)")
        ax.legend()
        st.pyplot(fig)

    # 比對數據表
    st.subheader("📋 詳細分類對比表")
    st.dataframe(df_report, use_container_width=True)

    # 預留資金試算與自動扣除評估
    st.subheader("🏦 銀行帳戶預留資金試算")
    gap = pending_fixed_amount - current_total_balance
    
    col_bank1, col_bank2 = st.columns(2)
    col_bank1.write(f"💳 截至基準日當前帳戶實際總餘額：**${current_total_balance:,.0f}**")
    col_bank1.write(f"⏳ 本月剩餘待扣固定支出總額：**${pending_fixed_amount:,.0f}**")
    
    if gap > 0:
        col_bank2.error(f"🚨【帳戶餘額不足】請最晚在扣款日前補入 **${gap:,.0f}**，否則預計扣款失敗！")
    else:
        after_deduct = current_total_balance - pending_fixed_amount
        col_bank2.success(f"🟢【資金充足】扣除本月剩餘待扣固定支出後，預估帳戶淨餘額為 **${after_deduct:,.0f}**。")


# ==================== Tab 2: 線上記帳與預算編輯 ====================
with tab2:
    st.subheader("➕ 單筆快速填寫記帳")
    
    with st.form("add_transaction_form", clear_on_submit=True):
        f_col1, f_col2, f_col3 = st.columns(3)
        t_date = f_col1.date_input("日期", datetime.date.today())
        t_type = f_col2.selectbox("收支類型", ["支出", "收入"])
        
        category_options = list(st.session_state.df_budget['分類名稱'].unique()) + ["薪資", "副業收入", "投資理財", "其他"]
        t_category = f_col3.selectbox("分類名稱", category_options)
        
        f_col4, f_col5 = st.columns(2)
        t_amount = f_col4.number_input("金額 (NTD)", min_value=1, value=100, step=50)
        t_note = f_col5.text_input("備註（選填）", "")
        
        submit_btn = st.form_submit_button("➕ 立即新增紀錄")
        
        if submit_btn:
            new_record = pd.DataFrame([{
                "日期": pd.to_datetime(t_date),
                "收支類型": t_type,
                "分類名稱": t_category,
                "金額": t_amount,
                "備註": t_note
            }])
            st.session_state.df_trans = pd.concat([st.session_state.df_trans, new_record], ignore_index=True)
            st.success(f"✅ 已成功記錄：{t_date} [{t_type}] {t_category} ${t_amount:,}")
            st.rerun()

    st.markdown("---")

    col_edit1, col_edit2 = st.columns([6, 4])

    with col_edit1:
        st.subheader("📝 編輯所有收支紀錄 (包含跨月歷史紀錄)")
        edited_trans = st.data_editor(
            st.session_state.df_trans,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD"),
                "收支類型": st.column_config.SelectboxColumn("收支類型", options=["支出", "收入"]),
                "金額": st.column_config.NumberColumn("金額 (NTD)", min_value=0, format="$%d")
            }
        )
        st.session_state.df_trans = edited_trans

    with col_edit2:
        st.subheader("🎯 調整預算設定")
        edited_budget = st.data_editor(
            st.session_state.df_budget,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "預算金額": st.column_config.NumberColumn("預算金額", min_value=0, format="$%d"),
                "支出類型": st.column_config.SelectboxColumn("支出類型", options=["固定", "變動"])
            }
        )
        st.session_state.df_budget = edited_budget
