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

# ----------------- 強制註冊中文字型 (解決口口口亂碼) -----------------
font_path = "NotoSansTC-Regular.ttf"

# 自動下載並向 Matplotlib 註冊 Google Noto Sans TC 中文字型
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
    # 備用：掃描 Linux 系統字型 (packages.txt)
    for sys_font in [
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc'
    ]:
        if os.path.exists(sys_font):
            fm.fontManager.addfont(sys_font)
            plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei', 'sans-serif']
            break

plt.rcParams['axes.unicode_minus'] = False
# ------------------------------------------------------------------

# 生成標準範本 Excel 檔的函式
def generate_template_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_budget = pd.DataFrame({
            "分類名稱": ["居住房租", "水電瓦斯", "電信網路", "飲食餐飲", "交通通勤", "娛樂休閒", "日常雜項"],
            "預算金額": [15000, 3000, 1000, 10000, 3000, 4000, 5000],
            "支出類型": ["固定", "固定", "固定", "變動", "變動", "變動", "變動"],
            "每月扣款日": [5, 25, 10, None, None, None, None]
        })
        df_budget.to_excel(writer, sheet_name='預算設定', index=False)

        df_trans = pd.DataFrame({
            "日期": ["2026-08-01", "2026-08-05", "2026-08-05", "2026-08-06", "2026-08-08", "2026-08-10"],
            "收支類型": ["支出", "支出", "收入", "支出", "支出", "支出"],
            "分類名稱": ["飲食餐飲", "居住房租", "薪資", "飲食餐飲", "交通通勤", "娛樂休閒"],
            "金額": [350, 15000, 60000, 1200, 800, 2500],
            "備註": ["午餐外帶", "8月房租", "8月薪資入帳", "朋友聚餐", "悠遊卡加值", "購買遊戲"]
        })
        df_trans.to_excel(writer, sheet_name='收支紀錄', index=False)
        
    return output.getvalue()

st.title("💰 個人雲端記帳與預算監控 App")
st.write("上傳 Excel 記帳檔案，自動分析「月初預估預算」、「目前實際花費」與「預估預留資金」。")

# 2. 側邊欄設定
st.sidebar.header("⚙️ 控制面板")

st.sidebar.markdown("### 1. 取得記帳範本")
template_bytes = generate_template_excel()
st.sidebar.download_button(
    label="📥 下載標準 Excel 記帳範本",
    data=template_bytes,
    file_name="記帳範本_預算與收支.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 2. 上傳與分析")
uploaded_file = st.sidebar.file_uploader("上傳 Excel 記帳檔 (.xlsx)", type=["xlsx"])
current_balance = st.sidebar.number_input("目前銀行帳戶餘額 (NTD)", value=25000, step=1000)
analysis_date = st.sidebar.date_input("分析基準日期", datetime.date.today())

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheets = excel_file.sheet_names
        
        if "預算設定" not in sheets or "收支紀錄" not in sheets:
            st.error("⚠️ 上傳的 Excel 格式不正確！必須包含名為 **「預算設定」** 與 **「收支紀錄」** 的兩個工作表。請使用左側欄下載的標準範本。")
        else:
            df_budget = pd.read_excel(uploaded_file, sheet_name="預算設定")
            df_trans = pd.read_excel(uploaded_file, sheet_name="收支紀錄")
            df_trans['日期'] = pd.to_datetime(df_trans['日期'])

            df_month = df_trans[
                (df_trans['日期'].dt.year == analysis_date.year) & 
                (df_trans['日期'].dt.month == analysis_date.month)
            ]
            df_expense = df_month[df_month['收支類型'] == '支出']
            actual_spend = df_expense.groupby('分類名稱')['金額'].sum()

            _, total_days = calendar.monthrange(analysis_date.year, analysis_date.month)
            current_day = analysis_date.day
            time_progress_ratio = current_day / total_days
            time_progress_pct = round(time_progress_ratio * 100, 1)

            total_planned_budget = df_budget['預算金額'].sum()
            total_actual_spend = actual_spend.sum()

            projected_total = 0
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

                projected_total += proj
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

            # 3. 頂部 KPI 指標
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("月初預估總預算", f"${total_planned_budget:,.0f}")
            col2.metric("目前實際花費", f"${total_actual_spend:,.0f}", delta=f"${total_planned_budget - total_actual_spend:,.0f} 剩餘")
            col3.metric("預估月底總花費", f"${projected_total:,.0f}")
            col4.metric("當月時間進度", f"{time_progress_pct}%", f"第 {current_day}/{total_days} 天")

            st.markdown("---")

            # 4. 月初預估 vs 實際花費對比圖
            st.subheader("📊 月初預估預算 vs. 目前實際花費 比較圖")
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

            # 5. 詳細比對數據表
            st.subheader("📋 詳細分類對比表")
            st.dataframe(df_report, use_container_width=True)

            # 6. 預留資金試算
            st.subheader("🏦 銀行帳戶預留資金試算")
            gap = pending_fixed_amount - current_balance
            st.write(f"💰 本月剩餘待扣固定支出總額：**${pending_fixed_amount:,}**")
            if gap > 0:
                st.error(f"🚨【預留資金不足】請最晚在扣款日前補入 **${gap:,}**！")
            else:
                st.success(f"🟢【資金充足】扣除剩餘固定支出後，預估還剩 **${abs(gap):,}**。")

    except Exception as e:
        st.error(f"檔案讀取失敗：{e}")
else:
    st.info("👈 請點擊左側邊欄的 **「📥 下載標準 Excel 記帳範本」** 取得範本，填寫後再上傳分析！")
