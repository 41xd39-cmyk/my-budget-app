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

# 2. 初始化 Session State 數據
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
            pd.to_datetime("2026-07-15"), pd.to_datetime("2026-07-20"),
            pd.to_datetime("2026-08-01"), pd.to_datetime("2026-08-05"), 
            pd.to_datetime("2026-08-05"), pd.to_datetime("2026-08-06")
        ],
        "實際扣款日": [
            pd.to_datetime("2026-07-15"), pd.to_datetime("2026-08-15"), # 7月刷卡，8/15扣款範例
            pd.to_datetime("2026-08-01"), pd.to_datetime("2026-09-05"), # 8月刷卡，9/5扣款範例
            pd.to_datetime("2026-08-05"), pd.to_datetime("2026-08-06")
        ],
        "收支類型": ["收入", "支出", "支出", "支出", "收入", "支出"],
        "分類名稱": ["薪資", "日常雜項", "飲食餐飲", "居住房租", "薪資", "飲食餐飲"],
        "金額": [0, 0, 0, 0, 0, 0],
        "備註": ["7月薪資", "刷卡購買家電(8/15扣)", "午餐外帶", "8月房租(9/5扣)", "8月薪資入帳", "朋友聚餐"]
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
initial_balance = st.sidebar.number_input(
    "帳戶起始底金 (NTD)", 
    value=41721, 
    step=1000, 
    help="設定開始記帳前帳戶內的初始金額，系統會自動根據實際扣款日進行動態累積。"
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
    
    # 補全實際扣款日（若有空值則以交易日期替代）
    df_trans['實際扣款日'] = df_trans['實際扣款日'].fillna(df_trans['日期'])
    
    df_trans['日期_dt'] = pd.to_datetime(df_trans['日期'])
    df_trans['扣款日_dt'] = pd.to_datetime(df_trans['實際扣款日'])
    
    analysis_date_dt = pd.to_datetime(analysis_date)
    
    # ---------------- 核心帳戶餘額算式（依據實際扣款日） ----------------
    # 已發生且實際扣款日 <= 分析基準日 的真實現金流
    df_paid_history = df_trans[df_trans['扣款日_dt'].dt.date <= analysis_date]
    total_cum_income = df_paid_history[df_paid_history['收支類型'] == '收入']['金額'].sum()
    total_cum_paid_expense = df_paid_history[df_paid_history['收支類型'] == '支出']['金額'].sum()
    
    # 目前銀行帳戶實際餘額
    current_real_balance = initial_balance + total_cum_income - total_cum_paid_expense

    # ---------------- 當月消費與扣款狀態分析 ----------------
    # 消費發生在當月的紀錄
    df_month_consumed = df_trans[
        (df_trans['日期_dt'].dt.year == analysis_date.year) & 
        (df_trans['日期_dt'].dt.month == analysis_date.month)
    ]
    month_income = df_month_consumed[df_month_consumed['收支類型'] == '收入']['金額'].sum()
    df_month_expense = df_month_consumed[df_month_consumed['收支類型'] == '支出']
    month_total_expense = df_month_expense['金額'].sum() # 當月消費總額

    # 當月消費中：已扣款 vs 待扣款
    month_paid_expense = df_month_expense[df_month_expense['扣款日_dt'].dt.date <= analysis_date]['金額'].sum()
    month_pending_expense = df_month_expense[df_month_expense['扣款日_dt'].dt.date > analysis_date]['金額'].sum()

    # 包含跨月刷卡消費：截至基準日「所有已消費但尚未扣款」的總金額
    df_all_pending = df_trans[
        (df_trans['收支類型'] == '支出') & 
        (df_trans['日期_dt'].dt.date <= analysis_date) & 
        (df_trans['扣款日_dt'].dt.date > analysis_date)
    ]
    total_unpaid_credit_card = df_all_pending['金額'].sum()

    actual_spend = df_month_expense.groupby('分類名稱')['金額'].sum()

    # 時間進度計算
    _, total_days = calendar.monthrange(analysis_date.year, analysis_date.month)
    current_day = analysis_date.day
    time_progress_ratio = current_day / total_days
    time_progress_pct = round(time_progress_ratio * 100, 1)

    total_planned_budget = df_budget['預算金額'].sum()

    projected_total_expense = 0
    report_data = []

    for _, row in df_budget.iterrows():
        cat = row['分類名稱']
        budget = row['預算金額']
        is_fixed = (row['支出類型'] == '固定')
        
        spent = actual_spend.get(cat, 0)
        diff = budget - spent
        
        if is_fixed:
            proj = spent if spent > 0 else budget
            status = "✅ 已完成" if spent > 0 else "⏳ 預計本月發生"
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
            "當月消費金額": spent,
            "預算差額": diff,
            "預估月底花費": round(proj),
            "狀態": status
        })

    df_report = pd.DataFrame(report_data)

    # 頂部 KPI 指標
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    kpi_col1.metric("💵 當前銀行實際餘額", f"${current_real_balance:,.0f}", help="起始底金 + 已實際扣款/入帳的淨額")
    kpi_col2.metric("💳 當月消費總額", f"${month_total_expense:,.0f}", f"已扣 ${month_paid_expense:,.0f} | 待扣 ${month_pending_expense:,.0f}")
    kpi_col3.metric("🚨 跨月/刷卡未扣總額", f"${total_unpaid_credit_card:,.0f}", help="包含歷史與當月已刷卡但扣款日在未來的金額")
    kpi_col4.metric("🗓️ 當月時間進度", f"{time_progress_pct}%", f"第 {current_day}/{total_days} 天")

    st.markdown("---")

    # 圖表展現
    st.subheader(f"📊 {analysis_date.year} 年 {analysis_date.month} 月 預算 vs. 當月消費比較圖")
    if not df_report.empty:
        fig, ax = plt.subplots(figsize=(10, 4))
        categories = df_report['分類名稱']
        x = range(len(categories))
        width = 0.35

        ax.bar([p - width/2 for p in x], df_report['月初預估預算'], width, label='月初預估預算', color='#e0e0e0')
        colors = ['#ea4335' if "透支" in r['狀態'] or "燒錢" in r['狀態'] else '#34a853' for _, r in df_report.iterrows()]
        ax.bar([p + width/2 for p in x], df_report['當月消費金額'], width, label='當月消費金額', color=colors)
        ax.plot([p + width/2 for p in x], df_report['預估月底花費'], "r--o", label='預估月底總花費')

        ax.set_xticks(x)
        ax.set_xticklabels(categories, rotation=15)
        ax.set_ylabel("金額 (NTD)")
        ax.legend()
        st.pyplot(fig)

    # 比對數據表
    st.subheader("📋 詳細分類對比表")
    st.dataframe(df_report, use_container_width=True)

    # 刷卡與延遲扣款試算區
    st.subheader("🏦 信用卡與延遲扣款預留資金試算")
    
    col_bank1, col_bank2 = st.columns(2)
    col_bank1.write(f"💳 基準日當前銀行實際餘額：**${current_real_balance:,.0f}**")
    col_bank1.write(f"⏳ 截至基準日已刷卡/消費但未扣款金額：**${total_unpaid_credit_card:,.0f}**")
    
    gap = total_unpaid_credit_card - current_real_balance
    if gap > 0:
        col_bank2.error(f"🚨【預留資金不足】未來需扣款金額大於當前餘額，請最晚在扣款日前補入 **${gap:,.0f}**！")
    else:
        after_deduct = current_real_balance - total_unpaid_credit_card
        col_bank2.success(f"🟢【資金充足】扣除所有已刷卡未扣款項後，預估銀行剩餘淨額為 **${after_deduct:,.0f}**。")

    if not df_all_pending.empty:
        with st.expander("🔍 點擊查看目前未扣款明細清單"):
            st.dataframe(df_all_pending[['日期', '實際扣款日', '分類名稱', '金額', '備註']], use_container_width=True)


# ==================== Tab 2: 線上記帳與預算編輯 ====================
with tab2:
    st.subheader("➕ 單筆快速填寫記帳")
    
    # 快速填寫表單
    with st.form("add_transaction_form", clear_on_submit=True):
        f_col1, f_col2, f_col3 = st.columns(3)
        t_date = f_col1.date_input("消費/交易日期", datetime.date.today())
        
        # 實際扣款日設定（預設帶入消費日期）
        t_pay_date = f_col2.date_input("實際扣款日期 (刷卡/延遲扣款)", datetime.date.today(), help="若為刷卡消費，請選擇預計扣款日/卡費繳款日；若無延遲扣款則保持與消費日期相同即可。")
        t_type = f_col3.selectbox("收支類型", ["支出", "收入"])
        
        f_col4, f_col5, f_col6 = st.columns(3)
        category_options = list(st.session_state.df_budget['分類名稱'].unique()) + ["薪資", "副業收入", "投資理財", "其他"]
        t_category = f_col4.selectbox("分類名稱", category_options)
        t_amount = f_col5.number_input("金額 (NTD)", min_value=1, value=100, step=50)
        t_note = f_col6.text_input("備註（選填）", "")
        
        submit_btn = st.form_submit_button("➕ 立即新增紀錄")
        
        if submit_btn:
            new_record = pd.DataFrame([{
                "日期": pd.to_datetime(t_date),
                "實際扣款日": pd.to_datetime(t_pay_date),
                "收支類型": t_type,
                "分類名稱": t_category,
                "金額": t_amount,
                "備註": t_note
            }])
            st.session_state.df_trans = pd.concat([st.session_state.df_trans, new_record], ignore_index=True)
            st.success(f"✅ 已成功記錄：{t_date} [{t_type}] {t_category} ${t_amount:,} (預計扣款日: {t_pay_date})")
            st.rerun()

    st.markdown("---")

    col_edit1, col_edit2 = st.columns([6, 4])

    with col_edit1:
        st.subheader("📝 編輯所有收支紀錄 (包含實際扣款日欄位)")
        edited_trans = st.data_editor(
            st.session_state.df_trans,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "日期": st.column_config.DateColumn("消費日期", format="YYYY-MM-DD"),
                "實際扣款日": st.column_config.DateColumn("實際扣款日", format="YYYY-MM-DD"),
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
