import io
import os
import urllib.request
import calendar
import datetime
import traceback
import hashlib
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

# ----------------- 安全性哈希函式 -----------------
def hash_text(text: str) -> str:
    return hashlib.sha256(text.strip().encode('utf-8')).hexdigest()

# ----------------- Excel 月報表匯出生成器 -----------------
def generate_monthly_excel_report(df_report_month, df_trans_month, user_name, analysis_date):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_report_month.to_excel(writer, sheet_name='預算與實際對比', index=False)
        if not df_trans_month.empty:
            display_cols = [c for c in ['日期', '實際扣款日', '收支類型', '分類名稱', '金額', '支付帳戶', '備註'] if c in df_trans_month.columns]
            df_trans_month[display_cols].to_excel(writer, sheet_name='當月收支明細', index=False)
    return output.getvalue()

# ----------------- Google Sheets 串接設定 -----------------
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@st.cache_resource
def get_gspread_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=300)
def load_all_data():
    client = get_gspread_client()
    spreadsheet_id = st.secrets["general"]["spreadsheet_id"]
    sh = client.open_by_key(spreadsheet_id)
    
    existing_sheets = [w.title for w in sh.worksheets()]
    
    # 1. 使用者帳號
    ws_users = sh.worksheet("使用者帳號")
    df_users = pd.DataFrame(ws_users.get_all_records())
    for c in ["username", "password", "name", "sec_question", "sec_answer"]:
        if c not in df_users.columns: df_users[c] = pd.Series(dtype=str)

    # 2. 預算設定
    ws_budget = sh.worksheet("預算設定")
    df_budget = pd.DataFrame(ws_budget.get_all_records())
    for c in ["user_id", "分類名稱", "預算金額", "支出類型", "每月扣款日"]:
        if c not in df_budget.columns: df_budget[c] = pd.Series(dtype=object)

    # 3. 收支紀錄
    ws_trans = sh.worksheet("收支紀錄")
    df_trans = pd.DataFrame(ws_trans.get_all_records())
    for c in ["user_id", "日期", "實際扣款日", "收支類型", "分類名稱", "金額", "備註", "支付帳戶"]:
        if c not in df_trans.columns: df_trans[c] = pd.Series(dtype=object)

    # 4. 支付帳戶
    if "支付帳戶" in existing_sheets:
        ws_acc = sh.worksheet("支付帳戶")
        df_acc = pd.DataFrame(ws_acc.get_all_records())
    else:
        ws_acc = sh.add_worksheet(title="支付帳戶", rows="100", cols="10")
        df_acc = pd.DataFrame(columns=["user_id", "帳戶名稱", "帳戶類型", "起始金額"])
        ws_acc.update([df_acc.columns.values.tolist()])

    for c in ["user_id", "帳戶名稱", "帳戶類型", "起始金額"]:
        if c not in df_acc.columns: df_acc[c] = pd.Series(dtype=object)

    # 5. 儲蓄目標
    if "儲蓄目標" in existing_sheets:
        ws_goals = sh.worksheet("儲蓄目標")
        df_goals = pd.DataFrame(ws_goals.get_all_records())
    else:
        ws_goals = sh.add_worksheet(title="儲蓄目標", rows="100", cols="10")
        df_goals = pd.DataFrame(columns=["user_id", "目標名稱", "目標金額", "當前累積金額", "預計完成日期"])
        ws_goals.update([df_goals.columns.values.tolist()])

    for c in ["user_id", "目標名稱", "目標金額", "當前累積金額", "預計完成日期"]:
        if c not in df_goals.columns: df_goals[c] = pd.Series(dtype=object)

    # 數據轉型
    if not df_budget.empty and '預算金額' in df_budget.columns:
        df_budget['預算金額'] = pd.to_numeric(df_budget['預算金額'], errors='coerce').fillna(0)
        
    if not df_trans.empty and '金額' in df_trans.columns:
        df_trans['金額'] = pd.to_numeric(df_trans['金額'], errors='coerce').fillna(0)

    if not df_acc.empty and '起始金額' in df_acc.columns:
        df_acc['起始金額'] = pd.to_numeric(df_acc['起始金額'], errors='coerce').fillna(0)

    if not df_goals.empty:
        if '目標金額' in df_goals.columns: df_goals['目標金額'] = pd.to_numeric(df_goals['目標金額'], errors='coerce').fillna(0)
        if '當前累積金額' in df_goals.columns: df_goals['當前累積金額'] = pd.to_numeric(df_goals['當前累積金額'], errors='coerce').fillna(0)

    return df_users, df_budget, df_trans, df_acc, df_goals

def save_all_data_to_gsheets(df_budget_all, df_trans_all, df_acc_all, df_goals_all):
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
    
    ws_acc = sh.worksheet("支付帳戶")
    ws_acc.clear()
    ws_acc.update([df_acc_all.columns.values.tolist()] + df_acc_all.fillna("").values.tolist())

    ws_goals = sh.worksheet("儲蓄目標")
    ws_goals.clear()
    df_goals_save = df_goals_all.copy()
    if not df_goals_save.empty and '預計完成日期' in df_goals_save.columns:
        df_goals_save['預計完成日期'] = pd.to_datetime(df_goals_save['預計完成日期']).dt.strftime('%Y-%m-%d')
    ws_goals.update([df_goals_save.columns.values.tolist()] + df_goals_save.fillna("").values.tolist())

    st.cache_data.clear()

def register_user(username, password, name, sec_q, sec_a):
    client = get_gspread_client()
    spreadsheet_id = st.secrets["general"]["spreadsheet_id"]
    sh = client.open_by_key(spreadsheet_id)
    ws_users = sh.worksheet("使用者帳號")
    hashed_pass = hash_text(password)
    hashed_ans = hash_text(sec_a)
    ws_users.append_row([str(username), str(hashed_pass), str(name), str(sec_q), str(hashed_ans)])
    st.cache_data.clear()

def update_user_password(username, new_password):
    client = get_gspread_client()
    spreadsheet_id = st.secrets["general"]["spreadsheet_id"]
    sh = client.open_by_key(spreadsheet_id)
    ws_users = sh.worksheet("使用者帳號")
    df_users = pd.DataFrame(ws_users.get_all_records())
    
    hashed_new_pass = hash_text(new_password)
    df_users.loc[df_users['username'].astype(str) == str(username), 'password'] = hashed_new_pass
    
    ws_users.clear()
    ws_users.update([df_users.columns.values.tolist()] + df_users.fillna("").values.tolist())
    st.cache_data.clear()

# ----------------- 登入狀態控制 -----------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.current_user = None

try:
    df_users_all, df_budget_all, df_trans_all, df_acc_all, df_goals_all = load_all_data()
except Exception as e:
    st.error(f"❌ 讀取資料庫失敗：{e}")
    st.info("💡 提示：若剛剛發送過多請求觸發上限，請稍等 1 分鐘後刷新頁面即可恢復。")
    st.stop()

# 尚未登入視圖
if not st.session_state.logged_in:
    st.title("🔐 雲端記帳管家 - 安全登入")
    login_tab, register_tab, reset_tab = st.tabs(["🔑 帳號登入", "📝 註冊新帳號", "❓ 忘記密碼"])
    
    with login_tab:
        with st.form("login_form"):
            user_input = st.text_input("帳號 (Username)")
            pass_input = st.text_input("密碼 (Password)", type="password")
            submit_login = st.form_submit_button("登入")
            
            if submit_login:
                if not df_users_all.empty and 'username' in df_users_all.columns:
                    hashed_input = hash_text(pass_input)
                    matched_user = df_users_all[
                        (df_users_all['username'].astype(str) == user_input.strip()) & 
                        ((df_users_all['password'].astype(str) == hashed_input) | (df_users_all['password'].astype(str) == pass_input.strip()))
                    ]
                    if not matched_user.empty:
                        st.session_state.logged_in = True
                        st.session_state.current_user = user_input.strip()
                        st.session_state.user_name = matched_user.iloc[0]['name']
                        st.success("登入成功！頁面載入中...")
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
            
            sec_questions = [
                "您第一隻寵物的名字？",
                "您畢業的國小名稱？",
                "您最喜歡的一部電影？",
                "您出生居住的城市名字？"
            ]
            reg_sec_q = st.selectbox("選擇安全驗證問題 (忘記密碼時驗證)", sec_questions)
            reg_sec_a = st.text_input("安全問題答案")
            
            submit_reg = st.form_submit_button("註冊並建立帳戶")
            
            if submit_reg:
                if reg_user and reg_pass and reg_sec_a:
                    if not df_users_all.empty and reg_user in df_users_all['username'].astype(str).values:
                        st.error("⚠️ 該帳號已被註冊，請換一個帳號！")
                    else:
                        register_user(reg_user, reg_pass, reg_name, reg_sec_q, reg_sec_a)
                        st.success("✅ 註冊成功！安全驗證問題已設定，請切換至「帳號登入」分頁。")
                else:
                    st.warning("請完整填寫帳號、密碼與安全問題答案。")

    with reset_tab:
        st.subheader("🔑 驗證安全問題並重設密碼")
        reset_username = st.text_input("請輸入您的帳號", key="reset_u")
        
        if reset_username:
            user_row = df_users_all[df_users_all['username'].astype(str) == reset_username.strip()]
            if not user_row.empty:
                saved_q = user_row.iloc[0].get('sec_question', '')
                saved_a_hash = user_row.iloc[0].get('sec_answer', '')
                
                if saved_q:
                    st.info(f"❓ **您的安全問題：** {saved_q}")
                    with st.form("reset_pass_form"):
                        ans_input = st.text_input("請輸入安全問題答案")
                        new_pass = st.text_input("請設定新密碼", type="password")
                        submit_reset = st.form_submit_button("重設密碼")
                        
                        if submit_reset:
                            if hash_text(ans_input) == str(saved_a_hash):
                                update_user_password(reset_username.strip(), new_pass)
                                st.success("🎉 密碼已成功更新！請使用新密碼前往「帳號登入」分頁進行登入。")
                            else:
                                st.error("❌ 安全問題答案不正確！")
                else:
                    st.warning("⚠️ 該帳號為舊版本建立，未設定安全問題。請聯繫管理員。")
            else:
                st.error("❌ 找不到該帳號，請確認輸入是否正確。")

    st.stop()

# ==================== 登入成功後的主介面 ====================
current_user = st.session_state.current_user

st.sidebar.markdown(f"👤 **目前登入者：** {st.session_state.user_name} (`{current_user}`)")

col_btn1, col_btn2 = st.sidebar.columns(2)
if col_btn1.button("🔄 刷新雲端"):
    st.cache_data.clear()
    st.rerun()

if col_btn2.button("🚪 登出"):
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.rerun()

st.sidebar.markdown("---")

df_user_budget = df_budget_all[df_budget_all['user_id'].astype(str) == current_user].copy()
df_user_trans = df_trans_all[df_trans_all['user_id'].astype(str) == current_user].copy()
df_user_acc = df_acc_all[df_acc_all['user_id'].astype(str) == current_user].copy()
df_user_goals = df_goals_all[df_goals_all['user_id'].astype(str) == current_user].copy()

if df_user_acc.empty:
    df_user_acc = pd.DataFrame({
        "user_id": [current_user]*3,
        "帳戶名稱": ["現金錢包", "主力銀行帳戶", "常用信用卡"],
        "帳戶類型": ["現金", "銀行帳戶", "信用卡"],
        "起始金額": [3000, 50000, 0]
    })

if df_user_budget.empty:
    default_categories = ["居住房租", "水電瓦斯", "電信網路", "飲食餐飲", "交通通勤", "娛樂休閒", "日常雜項"]
    df_user_budget = pd.DataFrame({
        "user_id": [current_user]*len(default_categories),
        "分類名稱": default_categories,
        "預算金額": [15000, 3000, 1000, 10000, 3000, 4000, 5000],
        "支出類型": ["固定", "固定", "固定", "變動", "變動", "變動", "變動"],
        "每月扣款日": [5, 25, 10, "", "", "", ""]
    })

analysis_date = st.sidebar.date_input("分析基準日期", datetime.date.today())

st.title(f"💰 {st.session_state.user_name} 的雲端記帳管家")

tab1, tab2, tab3 = st.tabs(["📊 分析儀表板", "✍️ 線上記帳與預算編輯", "🎯 儲蓄目標與多帳戶管理"])

# ==================== Tab 1: 分析儀表板 ====================
with tab1:
    user_accounts_list = list(df_user_acc['帳戶名稱'].unique())
    account_balances = {}
    total_assets = 0
    
    if not df_user_trans.empty:
        df_user_trans['實際扣款日'] = df_user_trans['實際扣款日'].fillna(df_user_trans['日期'])
        df_user_trans['日期_dt'] = pd.to_datetime(df_user_trans['日期'])
        df_user_trans['扣款日_dt'] = pd.to_datetime(df_user_trans['實際扣款日'])
        
        df_paid = df_user_trans[df_user_trans['扣款日_dt'].dt.date <= analysis_date]
        
        for _, acc_row in df_user_acc.iterrows():
            acc_name = acc_row['帳戶名稱']
            start_val = float(acc_row['起始金額'])
            acc_inc = df_paid[(df_paid['支付帳戶'] == acc_name) & (df_paid['收支類型'] == '收入')]['金額'].sum()
            acc_exp = df_paid[(df_paid['支付帳戶'] == acc_name) & (df_paid['收支類型'] == '支出')]['金額'].sum()
            bal = start_val + acc_inc - acc_exp
            account_balances[acc_name] = bal
            total_assets += bal

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
        for _, acc_row in df_user_acc.iterrows():
            account_balances[acc_row['帳戶名稱']] = float(acc_row['起始金額'])
            total_assets += float(acc_row['起始金額'])
        df_month_consumed = pd.DataFrame()
        month_total_expense = month_paid_expense = month_pending_expense = total_unpaid_credit_card = 0
        actual_spend = pd.Series()

    _, total_days = calendar.monthrange(analysis_date.year, analysis_date.month)
    current_day = analysis_date.day
    time_progress_ratio = current_day / total_days
    time_progress_pct = round(time_progress_ratio * 100, 1)

    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    kpi_col1.metric("💵 當前總資產 (全帳戶)", f"${total_assets:,.0f}")
    kpi_col2.metric("💳 當月消費總額", f"${month_total_expense:,.0f}", f"已扣 ${month_paid_expense:,.0f} | 待扣 ${month_pending_expense:,.0f}")
    kpi_col3.metric("🚨 刷卡待扣/延遲扣款", f"${total_unpaid_credit_card:,.0f}")
    kpi_col4.metric("🗓️ 當月時間進度", f"{time_progress_pct}%", f"第 {current_day}/{total_days} 天")

    st.markdown("---")

    st.subheader("💳 個別支付帳戶即時餘額")
    acc_cols = st.columns(min(max(len(account_balances), 1), 4))
    for idx, (acc_name, bal) in enumerate(account_balances.items()):
        acc_cols[idx % 4].metric(f"🏦 {acc_name}", f"${bal:,.0f}")

    st.markdown("---")

    projected_total_expense = 0
    report_data = []

    for _, row in df_user_budget.iterrows():
        cat = row['分類名稱']
        budget = float(row['預算金額'])
        is_fixed = (row['支出類型'] == '固定')
        spent = float(actual_spend.get(cat, 0))
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

    col_rpt1, col_rpt2 = st.columns([3, 1])
    col_rpt1.subheader(f"📊 {analysis_date.year} 年 {analysis_date.month} 月 預算 vs. 當月消費分析")
    
    excel_report_bytes = generate_monthly_excel_report(df_report, df_month_consumed, st.session_state.user_name, analysis_date)
    col_rpt2.download_button(
        label=f"📥 下載 {analysis_date.month} 月財務報表 (.xlsx)",
        data=excel_report_bytes,
        file_name=f"{st.session_state.user_name}_{analysis_date.year}_{analysis_date.month}_月報表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    chart_col1, chart_col2 = st.columns([6, 4])
    
    with chart_col1:
        if not df_report.empty:
            fig, ax = plt.subplots(figsize=(7, 4))
            categories = df_report['分類名稱']
            x = range(len(categories))
            width = 0.35

            ax.bar([p - width/2 for p in x], df_report['月初預估預算'], width, label='預算', color='#e0e0e0')
            colors = ['#ea4335' if "透支" in r['狀態'] or "燒錢" in r['狀態'] else '#34a853' for _, r in df_report.iterrows()]
            ax.bar([p + width/2 for p in x], df_report['當月消費金額'], width, label='實際消費', color=colors)

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

    with chart_col2:
        if not df_month_consumed.empty and (df_month_consumed['收支類型'] == '支出').any():
            spend_by_cat = df_month_consumed[df_month_consumed['收支類型'] == '支出'].groupby('分類名稱')['金額'].sum()
            spend_by_cat = spend_by_cat[spend_by_cat > 0]
            
            if not spend_by_cat.empty:
                fig_pie, ax_pie = plt.subplots(figsize=(5, 4))
                wedges, texts, autotexts = ax_pie.pie(
                    spend_by_cat,
                    labels=spend_by_cat.index,
                    autopct='%1.1f%%',
                    startangle=140,
                    colors=plt.cm.Set3.colors
                )
                if my_font:
                    for text in texts: text.set_fontproperties(my_font)
                    for autotext in autotexts: autotext.set_fontproperties(my_font)
                ax_pie.set_title("🍕 當月消費類別比例", fontproperties=my_font if my_font else None)
                st.pyplot(fig_pie)
        else:
            st.info("💡 當月尚無支出消費紀錄，無法產生圓餅圖。")

    st.subheader("📋 詳細分類對比表")
    st.dataframe(df_report, use_container_width=True)

    st.markdown("---")

    st.subheader("📈 跨月歷史收支與淨儲蓄趨勢圖")
    if not df_user_trans.empty:
        df_trend = df_user_trans.copy()
        df_trend['年月'] = df_trend['日期_dt'].dt.strftime('%Y-%m')
        
        trend_summary = df_trend.groupby(['年月', '收支類型'])['金額'].sum().unstack(fill_value=0).reset_index()
        if '收入' not in trend_summary.columns: trend_summary['收入'] = 0
        if '支出' not in trend_summary.columns: trend_summary['支出'] = 0
        
        trend_summary['淨儲蓄'] = trend_summary['收入'] - trend_summary['支出']
        trend_summary = trend_summary.sort_values('年月')

        fig_trend, ax_trend = plt.subplots(figsize=(10, 3.5))
        ax_trend.plot(trend_summary['年月'], trend_summary['收入'], marker='o', color='#34a853', label='總收入', linewidth=2)
        ax_trend.plot(trend_summary['年月'], trend_summary['支出'], marker='o', color='#ea4335', label='總支出', linewidth=2)
        ax_trend.plot(trend_summary['年月'], trend_summary['淨儲蓄'], marker='s', linestyle='--', color='#4285f4', label='淨儲蓄', linewidth=2)

        ax_trend.set_ylabel("金額 (NTD)", fontproperties=my_font if my_font else None)
        ax_trend.legend(prop=my_font if my_font else None)
        ax_trend.grid(True, linestyle=':', alpha=0.6)
        
        if my_font:
            ax_trend.set_xticklabels(trend_summary['年月'], rotation=15, fontproperties=my_font)
        else:
            ax_trend.set_xticklabels(trend_summary['年月'], rotation=15)
            
        st.pyplot(fig_trend)
    else:
        st.info("💡 尚無歷史記帳資料可進行跨月趨勢分析。")

    st.markdown("---")

    col_auto, col_goal_show = st.columns(2)
    
    with col_auto:
        st.subheader("⚡ 週期性固定收支快速帶入")
        st.write(f"自動帶入預算表中標記為**「固定」**的項目至 **{analysis_date.year} 年 {analysis_date.month} 月**：")
        if st.button("⚡ 一鍵帶入本月固定支出"):
            fixed_items = df_user_budget[df_user_budget['支出類型'] == '固定']
            if not fixed_items.empty:
                new_records = []
                default_acc = user_accounts_list[0] if user_accounts_list else "預設帳戶"
                for _, f_row in fixed_items.iterrows():
                    try:
                        pay_day = int(f_row['每月扣款日']) if str(f_row['每月扣款日']).isdigit() else 1
                    except Exception:
                        pay_day = 1
                    target_date = datetime.date(analysis_date.year, analysis_date.month, min(pay_day, total_days))
                    new_records.append({
                        "user_id": current_user,
                        "日期": pd.to_datetime(target_date),
                        "實際扣款日": pd.to_datetime(target_date),
                        "收支類型": "支出",
                        "分類名稱": f_row['分類名稱'],
                        "金額": f_row['預算金額'],
                        "備註": f"自動帶入本月固定支出 ({f_row['分類名稱']})",
                        "支付帳戶": default_acc
                    })
                df_trans_new_all = pd.concat([df_trans_all, pd.DataFrame(new_records)], ignore_index=True)
                save_all_data_to_gsheets(df_budget_all, df_trans_new_all, df_acc_all, df_goals_all)
                st.success(f"✅ 已成功自動帶入 {len(new_records)} 筆固定支出！")
                st.rerun()

    with col_goal_show:
        st.subheader("🎯 儲蓄目標進度")
        if not df_user_goals.empty:
            for _, g_row in df_user_goals.iterrows():
                target_amt = float(g_row['目標金額']) if '目標金額' in g_row else 0
                curr_amt = float(g_row['當前累積金額']) if '當前累積金額' in g_row else 0
                pct = min(curr_amt / target_amt, 1.0) if target_amt > 0 else 0
                st.write(f"**{g_row.get('目標名稱', '目標')}** (${curr_amt:,.0f} / ${target_amt:,.0f})")
                st.progress(pct)

# ==================== Tab 2: 線上記帳與預算編輯 ====================
with tab2:
    user_custom_categories = list(df_user_budget['分類名稱'].unique())
    income_categories = ["薪資", "副業收入", "投資理財", "其他收入"]
    all_available_categories = user_custom_categories + [c for c in income_categories if c not in user_custom_categories]

    st.subheader("➕ 單筆快速填寫記帳")
    
    with st.form("add_transaction_form", clear_on_submit=True):
        f_col1, f_col2, f_col3 = st.columns(3)
        t_date = f_col1.date_input("消費/交易日期", datetime.date.today())
        t_pay_date = f_col2.date_input("實際扣款日期", datetime.date.today())
        t_type = f_col3.selectbox("收支類型", ["支出", "收入"])
        
        f_col4, f_col5, f_col6, f_col7 = st.columns(4)
        t_category = f_col4.selectbox("分類名稱", all_available_categories)
        t_amount = f_col5.number_input("金額 (NTD)", min_value=1, value=100, step=50)
        t_account = f_col6.selectbox("支付/入帳帳戶", user_accounts_list if user_accounts_list else ["預設帳戶"])
        t_note = f_col7.text_input("備註（選填）", "")
        
        submit_btn = st.form_submit_button("➕ 欄位無誤，獨立新增紀錄")
        
        if submit_btn:
            new_record = pd.DataFrame([{
                "user_id": current_user,
                "日期": pd.to_datetime(t_date),
                "實際扣款日": pd.to_datetime(t_pay_date),
                "收支類型": t_type,
                "分類名稱": t_category,
                "金額": t_amount,
                "備註": t_note,
                "支付帳戶": t_account
            }])
            
            df_trans_updated = pd.concat([df_trans_all, new_record], ignore_index=True)
            save_all_data_to_gsheets(df_budget_all, df_trans_updated, df_acc_all, df_goals_all)
            st.success("✅ 已成功記錄並寫入個人專屬雲端庫！")
            st.rerun()

    st.markdown("---")

    col_edit1, col_edit2 = st.columns([6, 4])

    with col_edit1:
        st.subheader("📝 編輯您的個人收支紀錄")
        trans_display_cols = ["日期", "實際扣款日", "收支類型", "分類名稱", "金額", "支付帳戶", "備註"]
        
        if df_user_trans.empty:
            df_editor_input = pd.DataFrame(columns=trans_display_cols)
        else:
            df_editor_input = df_user_trans[trans_display_cols].copy()

        edited_user_trans = st.data_editor(
            df_editor_input,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "日期": st.column_config.DateColumn("消費日期", format="YYYY-MM-DD"),
                "實際扣款日": st.column_config.DateColumn("實際扣款日", format="YYYY-MM-DD"),
                "收支類型": st.column_config.SelectboxColumn("收支類型", options=["支出", "收入"]),
                "分類名稱": st.column_config.SelectboxColumn("分類名稱", options=all_available_categories),
                "支付帳戶": st.column_config.SelectboxColumn("支付帳戶", options=user_accounts_list if user_accounts_list else ["預設帳戶"]),
                "金額": st.column_config.NumberColumn("金額 (NTD)", min_value=0, format="$%d")
            }
        )
        
        if st.button("💾 儲存收支紀錄修改"):
            edited_user_trans['user_id'] = current_user
            other_users_trans = df_trans_all[df_trans_all['user_id'].astype(str) != current_user]
            df_trans_new_all = pd.concat([other_users_trans, edited_user_trans], ignore_index=True)
            
            save_all_data_to_gsheets(df_budget_all, df_trans_new_all, df_acc_all, df_goals_all)
            st.success("✅ 個人收支紀錄已同步更新！")
            st.rerun()

    with col_edit2:
        st.subheader("🎯 自訂您的預算與分類設定")
        budget_display_cols = ["分類名稱", "預算金額", "支出類型", "每月扣款日"]
        
        edited_user_budget = st.data_editor(
            df_user_budget[budget_display_cols],
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "分類名稱": st.column_config.TextColumn("分類名稱"),
                "預算金額": st.column_config.NumberColumn("預算金額", min_value=0, format="$%d"),
                "支出類型": st.column_config.SelectboxColumn("支出類型", options=["固定", "變動"])
            }
        )
        
        if st.button("💾 儲存自訂分類與預算"):
            edited_user_budget['user_id'] = current_user
            other_users_budget = df_budget_all[df_budget_all['user_id'].astype(str) != current_user]
            df_budget_new_all = pd.concat([other_users_budget, edited_user_budget], ignore_index=True)
            
            save_all_data_to_gsheets(df_budget_new_all, df_trans_all, df_acc_all, df_goals_all)
            st.success("✅ 個人預算與分類已同步更新！")
            st.rerun()

# ==================== Tab 3: 儲蓄目標與多帳戶管理 ====================
with tab3:
    col_acc_mgr, col_goal_mgr = st.columns(2)

    with col_acc_mgr:
        st.subheader("💳 管理您的支付帳戶 (多錢包)")
        acc_display_cols = ["帳戶名稱", "帳戶類型", "起始金額"]
        
        edited_user_acc = st.data_editor(
            df_user_acc[acc_display_cols],
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "帳戶名稱": st.column_config.TextColumn("帳戶名稱"),
                "帳戶類型": st.column_config.SelectboxColumn("帳戶類型", options=["現金", "銀行帳戶", "信用卡", "電子支付"]),
                "起始金額": st.column_config.NumberColumn("起始金額", format="$%d")
            }
        )
        
        if st.button("💾 儲存支付帳戶設定"):
            edited_user_acc['user_id'] = current_user
            other_users_acc = df_acc_all[df_acc_all['user_id'].astype(str) != current_user]
            df_acc_new_all = pd.concat([other_users_acc, edited_user_acc], ignore_index=True)
            
            save_all_data_to_gsheets(df_budget_all, df_trans_all, df_acc_new_all, df_goals_all)
            st.success("✅ 支付帳戶已更新！")
            st.rerun()

    with col_goal_mgr:
        st.subheader("🎯 管理您的儲蓄與財務目標")
        goals_display_cols = ["目標名稱", "目標金額", "當前累積金額", "預計完成日期"]
        
        if df_user_goals.empty:
            df_goals_input = pd.DataFrame(columns=goals_display_cols)
        else:
            df_goals_input = df_user_goals[goals_display_cols].copy()

        edited_user_goals = st.data_editor(
            df_goals_input,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "目標名稱": st.column_config.TextColumn("目標名稱"),
                "目標金額": st.column_config.NumberColumn("目標金額", min_value=0, format="$%d"),
                "當前累積金額": st.column_config.NumberColumn("當前累積金額", min_value=0, format="$%d"),
                "預計完成日期": st.column_config.DateColumn("預計完成日期", format="YYYY-MM-DD")
            }
        )
        
        if st.button("💾 儲存儲蓄目標設定"):
            edited_user_goals['user_id'] = current_user
            other_users_goals = df_goals_all[df_goals_all['user_id'].astype(str) != current_user]
            df_goals_new_all = pd.concat([other_users_goals, edited_user_goals], ignore_index=True)
            
            save_all_data_to_gsheets(df_budget_all, df_trans_all, df_acc_all, df_goals_new_all)
            st.success("✅ 儲蓄目標已更新！")
            st.rerun()
