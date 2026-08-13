import os
import urllib.request
import calendar
import datetime
import traceback
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import gspread
from google.oauth2.service_account import Credentials

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
plt.rcParams['axes.unicode_minus'] = False
# ----------------------------------------------------

# ----------------- Google Sheets 串接設定 -----------------
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_gspread_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    # 相容換行字元處理
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def load_data_from_gsheets():
    client = get_gspread_client()
    spreadsheet_id = st.secrets["general"]["spreadsheet_id"]
    sh = client.open_by_key(spreadsheet_id)
    
    existing_sheets = [w.title for w in sh.worksheets()]
    
    # 1. 讀取或自動初始化「預算設定」工作表
    if "預算設定" in existing_sheets:
        ws_budget = sh.worksheet("預算設定")
        data_b = ws_budget.get_all_records()
        if data_b:
            df_budget = pd.DataFrame(data_b)
        else:
            df_budget = pd.DataFrame({
                "分類名稱": ["居住房租", "水電瓦斯", "電信網路", "飲食餐飲", "交通通勤", "娛樂休閒", "日常雜項"],
                "預算金額": [11000, 1000, 500, 12000, 3000, 4000, 5000],
                "支出類型": ["固定", "固定", "固定", "變動", "變動", "變動", "變動"],
                "每月扣款日": [20, 20, 20, None, None, None, None]
            })
            ws_budget.update([df_budget.columns.values.tolist()] + df_budget.values.tolist())
    else:
        ws_budget = sh.add_worksheet(title="預算設定", rows="100", cols="20")
        df_budget = pd.DataFrame({
            "分類名稱": ["居住房租", "水電瓦斯", "電信網路", "飲食餐飲", "交通通勤", "娛樂休閒", "日常雜項"],
            "預算金額": [11000, 1000, 500, 12000, 3000, 4000, 5000],
            "支出類型": ["固定", "固定", "固定", "變動", "變動", "變動", "變動"],
            "每月扣款日": [20, 20, 20, None, None, None, None]
        })
        ws_budget.update([df_budget.columns.values.tolist()] + df_budget.values.tolist())

    # 2. 讀取或自動初始化「收支紀錄」工作表
    if "收支紀錄" in existing_sheets:
        ws_trans = sh.worksheet("收支紀錄")
        data_t = ws_trans.get_all_records()
        df_trans = pd.DataFrame(data_t)
    else:
        ws_trans = sh.add_worksheet(title="收支紀錄", rows="100", cols="20")
        df_trans = pd.DataFrame({
            "日期": ["2026-08-01", "2026-08-05", "2026-08-05", "2026-08-06"],
            "實際扣款日": ["2026-08-01", "2026-08-05", "2026-08-05", "2026-08-06"],
            "收支類型": ["支出", "支出", "收入", "支出"],
            "分類名稱": ["飲食餐飲", "居住房租", "薪資", "飲食餐飲"],
            "金額": [0, 0, 0, 0],
            "備註": ["午餐外帶", "8月房租", "8月薪資入帳", "朋友聚餐"]
        })
        ws_trans.update([df_trans.columns.values.tolist()] + df_trans.values.tolist())

    if not df_trans.empty and '日期' in df_trans.columns:
        df_trans['日期'] = pd.to_datetime(df_trans['日期'])
        df_trans['實際扣款日'] = pd.to_datetime(df_trans['實際扣款日'])

    return df_budget, df_trans

def save_data_to_gsheets(df_budget, df_trans):
    client = get_gspread_client()
    spreadsheet_id = st.secrets["general"]["spreadsheet_id"]
    sh = client.open_by_key(spreadsheet_id)
    
    ws_budget = sh.worksheet("預算設定")
    ws_budget.clear()
    ws_budget.update([df_budget.columns.values.tolist()] + df_budget.values.tolist())
    
    ws_trans = sh.worksheet("收支紀錄")
    ws_trans.clear()
    df_trans_save = df_trans.copy()
    if not df_trans_save.empty:
        df_trans_save['日期'] = pd.to_datetime(df_trans_save['日期']).dt.strftime('%Y-%m-%d')
        df_trans_save['實際扣款日'] = pd.to_datetime(df_trans_save['實際扣款日']).dt.strftime('%Y-%m-%d')
    ws_trans.update([df_trans_save.columns.values.tolist()] + df_trans_save.values.tolist())

# 初始化 Session State 數據
if "df_budget" not in st.session_state or "df_trans" not in st.session_state:
    try:
        df_b, df_t = load_data_from_gsheets()
        st.session_state.df_budget = df_b
        st.session_state.df_trans = df_t
    except Exception as e:
        st.error(f"❌ 串接 Google 試算表失敗：{e}")
        st.code(traceback.format_exc())
        st.stop()
# ---------------------------------------------------------

st.title("💰 個人雲端記帳與預算監控 App (雲端同步版)")

# 2. 側邊欄控制面板
st.sidebar.header("⚙️ 控制面板")
initial_balance = st.sidebar.number_input(
    "帳戶起始底金 (NTD)", 
    value=41721, 
    step=1000, 
    help="設定開始記帳前帳戶內的初始金額。"
)
analysis_date = st.sidebar.date_input("分析基準日期", datetime.date.today())

if st.sidebar.button("🔄 從 Google 試算表同步資料"):
    try:
        df_b, df_t = load_data_from_gsheets()
        st.session_state.df_budget = df_b
        st.session_state.df_trans = df_t
        st.sidebar.success("已成功同步最新資料！")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"同步失敗：{e}")

# 3. 主分頁切換
tab1, tab2 = st.tabs(["📊 分析儀表板", "✍️ 線上記帳與預算編輯"])

# ==================== Tab 1: 分析儀表板 ====================
with tab1:
    df_budget = st.session_state.df_budget
    df_trans = st.session_state.df_trans
    
    if not df_trans.empty:
        df_trans['實際扣款日'] = df_trans['實際扣款日'].fillna(df_trans['日期'])
        df_trans['日期_dt'] = pd.to_datetime(df_trans['日期'])
        df_trans['扣款日_dt'] = pd.to_datetime(df_trans['實際扣款日'])
        
        df_paid_history = df_trans[df_trans['扣款日_dt'].dt.date <= analysis_date]
        total_cum_income = df_paid_history[df_paid_history['收支類型'] == '收入']['金額'].sum()
        total_cum_paid_expense = df_paid_history[df_paid_history['收支類型'] == '支出']['金額'].sum()
        current_real_balance = initial_balance + total_cum_income - total_cum_paid_expense

        df_month_consumed = df_trans[
            (df_trans['日期_dt'].dt.year == analysis_date.year) & 
            (df_trans['日期_dt'].dt.month == analysis_date.month)
        ]
        month_total_expense = df_month_consumed[df_month_consumed['收支類型'] == '支出']['金額'].sum()
        month_paid_expense = df_month_consumed[(df_month_consumed['收支類型'] == '支出') & (df_month_consumed['扣款日_dt'].dt.date <= analysis_date)]['金額'].sum()
        month_pending_expense = df_month_consumed[(df_month_consumed['收支類型'] == '支出') & (df_month_consumed['扣款日_dt'].dt.date > analysis_date)]['金額'].sum()

        df_all_pending = df_trans[
            (df_trans['收支類型'] == '支出') & 
            (df_trans['日期_dt'].dt.date <= analysis_date) & 
            (df_trans['扣款日_dt'].dt.date > analysis_date)
        ]
        total_unpaid_credit_card = df_all_pending['金額'].sum()
        actual_spend = df_month_consumed[df_month_consumed['收支類型'] == '支出'].groupby('分類名稱')['金額'].sum()
    else:
        current_real_balance = initial_balance
        month_total_expense = month_paid_expense = month_pending_expense = total_unpaid_credit_card = 0
        actual_spend = pd.Series()
        df_all_pending = pd.DataFrame()

    _, total_days = calendar.monthrange(analysis_date.year, analysis_date.month)
    current_day = analysis_date.day
    time_progress_ratio = current_day / total_days
    time_progress_pct = round(time_progress_ratio * 100, 1)

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

    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    kpi_col1.metric("💵 當前銀行實際餘額", f"${current_real_balance:,.0f}")
    kpi_col2.metric("💳 當月消費總額", f"${month_total_expense:,.0f}", f"已扣 ${month_paid_expense:,.0f} | 待扣 ${month_pending_expense:,.0f}")
    kpi_col3.metric("🚨 刷卡未扣總額", f"${total_unpaid_credit_card:,.0f}")
    kpi_col4.metric("🗓️ 當月時間進度", f"{time_progress_pct}%", f"第 {current_day}/{total_days} 天")

    st.markdown("---")

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

    st.subheader("📋 詳細分類對比表")
    st.dataframe(df_report, use_container_width=True)

# ==================== Tab 2: 線上記帳與預算編輯 ====================
with tab2:
    st.subheader("➕ 單筆快速填寫記帳")
    
    with st.form("add_transaction_form", clear_on_submit=True):
        f_col1, f_col2, f_col3 = st.columns(3)
        t_date = f_col1.date_input("消費/交易日期", datetime.date.today())
        t_pay_date = f_col2.date_input("實際扣款日期", datetime.date.today())
        t_type = f_col3.selectbox("收支類型", ["支出", "收入"])
        
        f_col4, f_col5, f_col6 = st.columns(3)
        category_options = list(st.session_state.df_budget['分類名稱'].unique()) + ["薪資", "副業收入", "投資理財", "其他"]
        t_category = f_col4.selectbox("分類名稱", category_options)
        t_amount = f_col5.number_input("金額 (NTD)", min_value=1, value=100, step=50)
        t_note = f_col6.text_input("備註（選填）", "")
        
        submit_btn = st.form_submit_button("➕ 立即新增紀錄並同步至 Google 雲端")
        
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
            save_data_to_gsheets(st.session_state.df_budget, st.session_state.df_trans)
            st.success(f"✅ 已成功紀錄並同步寫入 Google 試算表！")
            st.rerun()

    st.markdown("---")

    st.subheader("📝 編輯所有收支紀錄")
    edited_trans = st.data_editor(
        st.session_state.df_trans,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "日期": st.column_config.DateColumn("消費日期", format="YYYY-MM-DD"),
            "實際扣款日": st.column_config.DateColumn("實際扣款日", format="YYYY-MM-DD"),
            "金額": st.column_config.NumberColumn("金額 (NTD)", min_value=0, format="$%d")
        }
    )
    
    if st.button("💾 儲存表格修改至 Google 雲端"):
        st.session_state.df_trans = edited_trans
        save_data_to_gsheets(st.session_state.df_budget, st.session_state.df_trans)
        st.success("✅ 表格變更已儲存至 Google 試算表！")
