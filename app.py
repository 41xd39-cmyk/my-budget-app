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
st.set_page_config(page_title="多用戶雲端記帳管家", page_icon="💰", layout="wide")

# ----------------- 中文字型自動註冊 -----------------
font_path = "NotoSansTC-Regular.ttf"
if not os.path.exists(font_path) or os.path.getsize(font_path) < 100000:
    font_urls = [
        "https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/notosanstc/NotoSansTC-Regular.ttf",
        "https://raw.githubusercontent.com/google/fonts/main/ofl/notosanstc/NotoSansTC-Regular.ttf"
    ]
    for url in font_urls:
        try:
            urllib.request.urlretrieve(url, font_path)
            if os.path.exists(font_path) and os.path.getsize(font_path) > 100000:
                break
        except Exception:
            continue

my_font = None
if os.path.exists(font_path) and os.path.getsize(font_path) > 100000:
    fm.fontManager.addfont(font_path)
    my_font = fm.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = my_font.get_name()

plt.rcParams['axes.unicode_minus'] = False
# ----------------------------------------------------

# ----------------- Google Sheets 串接設定 -----------------
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_gspread_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def load_all_data():
    client = get_gspread_client()
    spreadsheet_id = st.secrets["general"]["spreadsheet_id"]
    sh = client.open_by_key(spreadsheet_id)
    
    # 讀取使用者帳號表
    ws_users = sh.worksheet("使用者帳號")
    df_users = pd.DataFrame(ws_users.get_all_records())
    
    # 讀取預算設定
    ws_budget = sh.worksheet("預算設定")
    df_budget = pd.DataFrame(ws_budget.get_all_records())
    
    # 讀取收支紀錄
    ws_trans = sh.worksheet("收支紀錄")
    df_trans = pd.DataFrame(ws_trans.get_all_records())
    
    return df_users, df_budget, df_trans

def save_data_to_gsheets(df_budget_all, df_trans_all):
    client = get_gspread_client()
    spreadsheet_id = st.secrets["general"]["spreadsheet_id"]
    sh = client.open_by_key(spreadsheet_id)
    
    ws_budget = sh.worksheet("預算設定")
    ws_budget.clear()
    ws_budget.update([df_budget_all.columns.values.tolist()] + df_budget_all.fillna("").values.tolist())
    
    ws_trans = sh.worksheet("收支紀錄")
    ws_trans.clear()
    df_trans_save = df_trans_all.copy()
    if not df_trans_save.empty and '日期' in df_trans_save.columns:
        df_trans_save['日期'] = pd.to_datetime(df_trans_save['日期']).dt.strftime('%Y-%m-%d')
        df_trans_save['實際扣款日'] = pd.to_datetime(df_trans_save['實際扣款日']).dt.strftime('%Y-%m-%d')
    ws_trans.update([df_trans_save.columns.values.tolist()] + df_trans_save.fillna("").values.tolist())

def register_user(username, password, name):
    client = get_gspread_client()
    spreadsheet_id = st.secrets["general"]["spreadsheet_id"]
    sh = client.open_by_key(spreadsheet_id)
    ws_users = sh.worksheet("使用者帳號")
    ws_users.append_row([str(username), str(password), str(name)])

# ----------------- 登入機制處理 -----------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.current_user = None

try:
    df_users_all, df_budget_all, df_trans_all = load_all_data()
except Exception as e:
    st.error(f"❌ 讀取資料庫失敗：{e}")
    st.stop()

if not st.session_state.logged_in:
    st.title("🔐 多用戶雲端記帳管家 - 系統登入")
    login_tab, register_tab = st.tabs(["🔑 帳號登入", "📝 註冊新帳號"])
    
    with login_tab:
        with st.form("login_form"):
            user_input = st.text_input("帳號 (Username)")
            pass_input = st.text_input("密碼 (Password)", type="password")
            submit_login = st.form_submit_button("登入")
            
            if submit_login:
                if not df_users_all.empty:
                    matched_user = df_users_all[
                        (df_users_all['username'].astype(str) == user_input.strip()) & 
                        (df_users_all['password'].astype(str) == pass_input.strip())
                    ]
                    if not matched_user.empty:
                        st.session_state.logged_in = True
                        st.session_state.current_user = user_input.strip()
                        st.session_state.user_name = matched_user.iloc[0]['name']
                        st.success("登入成功！頁面轉導中...")
                        st.rerun()
                    else:
                        st.error("❌ 帳號或密碼錯誤！")
                else:
                    st.error("❌ 系統尚未有任何使用者，請先註冊帳號。")

    with register_tab:
        with st.form("register_form"):
            reg_user = st.text_input("設定帳號")
            reg_pass = st.text_input("設定密碼", type="password")
            reg_name = st.text_input("您的姓名/暱稱")
            submit_reg = st.form_submit_button("註冊並建立預設範本")
            
            if submit_reg:
                if reg_user and reg_pass:
                    if not df_users_all.empty and reg_user in df_users_all['username'].astype(str).values:
                        st.error("⚠️ 該帳號已被註冊，請換一個帳號！")
                    else:
                        register_user(reg_user, reg_pass, reg_name)
                        st.success("✅ 註冊成功！請切換至「帳號登入」分頁進行登入。")
                else:
                    st.warning("請填寫完整的帳號與密碼。")
    st.stop()

# ==================== 登入成功後的主介面 ====================
current_user = st.session_state.current_user

# 側邊欄登入狀態資訊
st.sidebar.markdown(f"👤 **目前登入者：** {st.session_state.user_name} (`{current_user}`)")
if st.sidebar.button("🚪 登出系統"):
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.rerun()

st.sidebar.markdown("---")

# 核心：過濾出「僅限當前登入使用者」的資料
df_user_budget = df_budget_all[df_budget_all['user_id'].astype(str) == current_user].copy()
df_user_trans = df_trans_all[df_trans_all['user_id'].astype(str) == current_user].copy()

# 若新使用者沒有預算設定，給予預設範本
if df_user_budget.empty:
    df_user_budget = pd.DataFrame({
        "user_id": [current_user]*7,
        "分類名稱": ["居住房租", "水電瓦斯", "電信網路", "飲食餐飲", "交通通勤", "娛樂休閒", "日常雜項"],
        "預算金額": [11000, 1000, 500, 12000, 3000, 4000, 5000],
        "支出類型": ["固定", "固定", "固定", "變動", "變動", "變動", "變動"],
        "每月扣款日": [20, 20, 20, "", "", "", ""]
    })

# 側邊欄控制面板
initial_balance = st.sidebar.number_input("帳戶起始底金 (NTD)", value=20000, step=1000)
analysis_date = st.sidebar.date_input("分析基準日期", datetime.date.today())

st.title(f"💰 {st.session_state.user_name} 的雲端記帳管家")

tab1, tab2 = st.tabs(["📊 分析儀表板", "✍️ 線上記帳與預算編輯"])

# ==================== Tab 1: 分析儀表板 ====================
with tab1:
    if not df_user_trans.empty:
        df_user_trans['實際扣款日'] = df_user_trans['實際扣款日'].fillna(df_user_trans['日期'])
        df_user_trans['日期_dt'] = pd.to_datetime(df_user_trans['日期'])
        df_user_trans['扣款日_dt'] = pd.to_datetime(df_user_trans['實際扣款日'])
        
        df_paid_history = df_user_trans[df_user_trans['扣款日_dt'].dt.date <= analysis_date]
        total_cum_income = df_paid_history[df_paid_history['收支類型'] == '收入']['金額'].sum()
        total_cum_paid_expense = df_paid_history[df_paid_history['收支類型'] == '支出']['金額'].sum()
        current_real_balance = initial_balance + total_cum_income - total_cum_paid_expense

        df_month_consumed = df_user_trans[
            (df_user_trans['日期_dt'].dt.year == analysis_date.year) & 
            (df_user_trans['日期_dt'].dt.month == analysis_date.month)
        ]
        month_total_expense = df_month_consumed[df_month_consumed['收支類型'] == '支出']['金額'].sum()
        month_paid_expense = df_month_consumed[(df_month_consumed['收支類型'] == '支出') & (df_month_consumed['扣款日_dt'].dt.date <= analysis_date)]['金額'].sum()
        month_pending_expense = df_month_consumed[(df_month_consumed['收支類型'] == '支出') & (df_month_consumed['扣款日_dt'].dt.date > analysis_date)]['金額'].sum()

        df_all_pending = df_user_trans[
            (df_user_trans['收支類型'] == '支出') & 
            (df_user_trans['日期_dt'].dt.date <= analysis_date) & 
            (df_user_trans['扣款日_dt'].dt.date > analysis_date)
        ]
        total_unpaid_credit_card = df_all_pending['金額'].sum()
        actual_spend = df_month_consumed[df_month_consumed['收支類型'] == '支出'].groupby('分類名稱')['金額'].sum()
    else:
        current_real_balance = initial_balance
        month_total_expense = month_paid_expense = month_pending_expense = total_unpaid_credit_card = 0
        actual_spend = pd.Series()

    _, total_days = calendar.monthrange(analysis_date.year, analysis_date.month)
    current_day = analysis_date.day
    time_progress_ratio = current_day / total_days
    time_progress_pct = round(time_progress_ratio * 100, 1)

    projected_total_expense = 0
    report_data = []

    for _, row in df_user_budget.iterrows():
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
        if my_font:
            ax.set_xticklabels(categories, rotation=15, fontproperties=my_font)
            ax.set_ylabel("金額 (NTD)", fontproperties=my_font)
            ax.legend(prop=my_font)
        else:
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
        category_options = list(df_user_budget['分類名稱'].unique()) + ["薪資", "副業收入", "投資理財", "其他"]
        t_category = f_col4.selectbox("分類名稱", category_options)
        t_amount = f_col5.number_input("金額 (NTD)", min_value=1, value=100, step=50)
        t_note = f_col6.text_input("備註（選填）", "")
        
        submit_btn = st.form_submit_button("➕ 立即新增紀錄")
        
        if submit_btn:
            new_record = pd.DataFrame([{
                "user_id": current_user,
                "日期": pd.to_datetime(t_date),
                "實際扣款日": pd.to_datetime(t_pay_date),
                "收支類型": t_type,
                "分類名稱": t_category,
                "金額": t_amount,
                "備註": t_note
            }])
            
            df_trans_updated = pd.concat([df_trans_all, new_record], ignore_index=True)
            save_data_to_gsheets(df_budget_all, df_trans_updated)
            st.success("✅ 已成功記錄並寫入個人專屬雲端庫！")
            st.rerun()

    st.markdown("---")

    st.subheader("📝 編輯您的個人收支紀錄")
    # 只呈現與編輯當前使用者的紀錄
    display_cols = ["日期", "實際扣款日", "收支類型", "分類名稱", "金額", "備註"]
    edited_user_trans = st.data_editor(
        df_user_trans[display_cols] if not df_user_trans.empty else pd.DataFrame(columns=display_cols),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "日期": st.column_config.DateColumn("消費日期", format="YYYY-MM-DD"),
            "實際扣款日": st.column_config.DateColumn("實際扣款日", format="YYYY-MM-DD"),
            "金額": st.column_config.NumberColumn("金額 (NTD)", min_value=0, format="$%d")
        }
    )
    
    if st.button("💾 儲存個人表格修改"):
        edited_user_trans['user_id'] = current_user
        # 保留其他使用者的資料，僅替換當前使用者的資料
        other_users_trans = df_trans_all[df_trans_all['user_id'].astype(str) != current_user]
        df_trans_new_all = pd.concat([other_users_trans, edited_user_trans], ignore_index=True)
        
        save_data_to_gsheets(df_budget_all, df_trans_new_all)
        st.success("✅ 個人變更已成功同步至雲端！")
        st.rerun()
