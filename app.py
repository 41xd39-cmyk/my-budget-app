import streamlit as st
import pandas as pd
import datetime
import calendar
import matplotlib.pyplot as plt

# 1. 網頁頁面設定
st.set_page_config(page_title="個人雲端記帳管家", page_icon="💰", layout="wide")

# 設定 Matplotlib 字型
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial Unicode MS', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

st.title("💰 個人雲端記帳與預算監控 App")
st.write("上傳 Excel 記帳檔案，自動分析「月初預估預算」、「目前實際花費」與「月底預算推測」。")

# 2. 側邊欄設定
st.sidebar.header("⚙️ 控制面板")
uploaded_file = st.sidebar.file_uploader("上傳 Excel 記帳檔 (.xlsx)", type=["xlsx"])
current_balance = st.sidebar.number_input("目前銀行帳戶餘額 (NTD)", value=25000, step=1000)
analysis_date = st.sidebar.date_input("分析基準日期", datetime.date.today())

if uploaded_file is not None:
    try:
        # 讀取 Excel 工作表
        df_budget = pd.read_excel(uploaded_file, sheet_name="預算設定")
        df_trans = pd.read_excel(uploaded_file, sheet_name="收支紀錄")
        df_trans['日期'] = pd.to_datetime(df_trans['日期'])

        # 過濾當月資料
        df_month = df_trans[
            (df_trans['日期'].dt.year == analysis_date.year) & 
            (df_trans['日期'].dt.month == analysis_date.month)
        ]
        df_expense = df_month[df_month['收支類型'] == '支出']
        actual_spend = df_expense.groupby('分類名稱')['金額'].sum()

        # 計算當月天數與時間進度
        _, total_days = calendar.monthrange(analysis_date.year, analysis_date.month)
        current_day = analysis_date.day
        time_progress_ratio = current_day / total_days
        time_progress_pct = round(time_progress_ratio * 100, 1)

        # 3. 計算數據彙總表
        total_planned_budget = df_budget['預算金額'].sum()
        total_actual_spend = actual_spend.sum()
        
        # 變動預算與固定預算分類計算
        fixed_cats = df_budget[df_budget['支出類型'] == '固定']['分類名稱'].tolist()
        var_budget_total = df_budget[df_budget['支出類型'] == '變動']['預算金額'].sum()
        var_spend_total = actual_spend[actual_spend.index.isin(df_budget[df_budget['支出類型'] == '變動']['分類名稱'])].sum()
        
        var_spend_pct = (var_spend_total / var_budget_total * 100) if var_budget_total > 0 else 0

        # 月底預估總支出計算 Engine
        projected_total = 0
        report_data = []
        pending_fixed_amount = 0

        for _, row in df_budget.iterrows():
            cat = row['分類名稱']
            budget = row['預算金額']
            is_fixed = (row['支出類型'] == '固定')
            pay_day = row['每月扣款日'] if pd.notna(row['每月扣款日']) else None
            
            spent = actual_spend.get(cat, 0)
            diff = budget - spent  # 正數代表剩餘，負數代表超支
            
            if is_fixed:
                # 固定支出：已扣款則以實際金額計，未扣款則預測等於原預算
                proj = spent if spent > 0 else budget
                if spent > 0:
                    status = "✅ 已完成扣款"
                else:
                    pending_fixed_amount += budget
                    status = "⏳ 待扣款"
            else:
                # 變動支出：按日均消耗速度推算月底總花費
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

        # 4. 頂部 KPI 指標列
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("月初預估總預算", f"${total_planned_budget:,.0f}")
        col2.metric("目前累積實際花費", f"${total_actual_spend:,.0f}", delta=f"${total_planned_budget - total_actual_spend:,.0f} 剩餘額度")
        col3.metric("預估月底總花費", f"${projected_total:,.0f}", delta=f"${total_planned_budget - projected_total:,.0f}", delta_color="normal")
        col4.metric("當月時間進度", f"{time_progress_pct}%", f"第 {current_day}/{total_days} 天")

        # 5. 消費節奏指示器 (Pacing Bar)
        st.markdown("### ⏱️ 變動支出燒錢節奏分析")
        st.write(f"當月時間已過 **{time_progress_pct}%** ｜ 變動預算已消耗 **{var_spend_pct:.1f}%**")
        st.progress(min(int(var_spend_pct), 100))
        
        if var_spend_pct > time_progress_pct + 10:
            st.warning("⚠️ **警告**：您目前的變動支出消耗速度快於時間天數，建議適度控制日常消費！")
        elif var_spend_pct < time_progress_pct - 10:
            st.success("🎉 **良好**：您的日常消費控制得當，低於時間進度基準！")

        st.markdown("---")

        # 6. 主要圖表：月初預估 vs 實際花費比較圖
        st.subheader("📊 月初預估預算 vs. 目前實際花費 比較圖")
        
        fig, ax = plt.subplots(figsize=(10, 4.5))
        
        categories = df_report['分類名稱']
        x = range(len(categories))
        width = 0.35

        bars1 = ax.bar([p - width/2 for p in x], df_report['月初預估預算'], width, label='月初預估預算', color='#e0e0e0')
        
        # 依狀態動態給予實際花費色彩
        colors = []
        for _, r in df_report.iterrows():
            if "透支" in r['狀態'] or "燒錢" in r['狀態']:
                colors.append('#ea4335') # 紅
            elif "待扣" in r['狀態']:
                colors.append('#fbbc04') # 黃
            else:
                colors.append('#34a853') # 綠

        bars2 = ax.bar([p + width/2 for p in x], df_report['目前實際花費'], width, label='目前實際花費', color=colors)

        # 繪製預估月底線標記
        ax.plot([p + width/2 for p in x], df_report['預估月底花費'], "r--o", label='預估月底總花費', linewidth=1.5, markersize=5)

        ax.set_xticks(x)
        ax.set_xticklabels(categories, rotation=15)
        ax.set_ylabel("金額 (NTD)")
        ax.legend()

        st.pyplot(fig)

        # 7. 詳細比對數據表格
        st.subheader("📋 詳細分類對比表")
        
        # 格式化顯示金額
        df_show = df_report.copy()
        df_show['月初預估預算'] = df_show['月初預估預算'].map("${:,.0f}".format)
        df_show['目前實際花費'] = df_show['目前實際花費'].map("${:,.0f}".format)
        df_show['預算差額'] = df_show['預算差額'].map(lambda v: f"+${v:,.0f}" if v >= 0 else f"-${abs(v):,.0f}")
        df_show['預估月底花費'] = df_show['預估月底花費'].map("${:,.0f}".format)

        st.dataframe(df_show, use_container_width=True)

        # 8. 銀行預留資金試算
        st.subheader("🏦 銀行帳戶預留資金試算")
        gap = pending_fixed_amount - current_balance
        col_a, col_b = st.columns(2)
        col_a.write(f"💰 本月剩餘待扣固定支出總額：**${pending_fixed_amount:,}**")
        
        if gap > 0:
            col_b.error(f"🚨【預留資金不足】請最晚在扣款日前補入 **${gap:,}**！")
        else:
            col_b.success(f"🟢【資金充足】扣除剩餘固定支出後，預估還剩 **${abs(gap):,}**。")

    except Exception as e:
        st.error(f"讀取檔案失敗，請確保 Excel 格式包含「預算設定」與「收支紀錄」工作表：{e}")
else:
    st.info("👈 請點擊左側邊欄的「上傳 Excel 記帳檔」開始使用！")
