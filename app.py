"""
电商店铺利润速算器 - 主程序
功能：文件上传、字段映射、利润计算、图表展示、报告下载、利润总结、自动分析
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import tempfile
import os
from datetime import datetime

# 导入自定义模块
from utils.data_processor import (
    load_file, auto_map_columns, generate_sample_excel,
    calculate_metrics, generate_report, detect_date_column
)

# ---------- 页面配置 ----------
st.set_page_config(
    page_title="电商店铺利润速算器",
    page_icon="📊",
    layout="wide"
)

# ---------- 初始化 session_state ----------
if "df" not in st.session_state:
    st.session_state.df = None
if "report_file_path" not in st.session_state:
    st.session_state.report_file_path = None
if "report_ready" not in st.session_state:
    st.session_state.report_ready = False
if "auto_map_result" not in st.session_state:
    st.session_state.auto_map_result = None
if "metrics" not in st.session_state:
    st.session_state.metrics = None

# ---------- 辅助函数：获取数据日期范围 ----------
def get_date_range(df: pd.DataFrame):
    date_col = detect_date_column(df)
    if date_col is None:
        return None, None
    try:
        dates = pd.to_datetime(df[date_col])
        start_date = dates.min().strftime("%Y-%m-%d")
        end_date = dates.max().strftime("%Y-%m-%d")
        return start_date, end_date
    except:
        return None, None

# ---------- 侧边栏：使用教程 ----------
with st.sidebar:
    st.markdown("## 📖 使用教程")
    st.markdown("""
    1. **导出报表**：从电商平台（淘宝、拼多多、抖音、京东等）导出订单报表。
    2. **上传文件**：支持 .xlsx, .xls, .csv，≤20MB。
    3. **确认字段映射**：系统自动识别，可手动调整。
    4. **查看利润**：点击「计算利润」。
    5. **下载报告**：点击「生成报告」→「下载报告」，无次数限制。
    """)
    st.markdown("### ❓ 常见问题")
    st.markdown("""
    - **支持哪些平台？** 所有可导出订单明细的平台。
    - **退款额为负数？** 系统会取绝对值处理。
    - **费用多选？** 支持多列求和（佣金+广告费+技术服务费等）。
    - **日期折线图？** 自动检测日期列并绘制。
    """)

# ---------- 页脚（请替换为你的小红书链接） ----------
st.markdown("---")
st.markdown(
    "更多电商财税干货，请关注小红书 **[@AC 双枫]** "
    "（[点击跳转](https://www.xiaohongshu.com/user/profile/66d6604c000000001d0307be)）",
    unsafe_allow_html=True
)

# ---------- 主界面 ----------
st.title("📊 电商店铺利润速算器")
st.caption("上传报表，自动计算净收入、毛利润、退款率 | 附带自动经营分析")

# ---------- 文件上传区域 ----------
col1, col2 = st.columns([3, 1])
with col1:
    uploaded_file = st.file_uploader(
        "上传Excel或CSV文件",
        type=["xlsx", "xls", "csv"],
        help="支持 .xlsx, .xls, .csv，大小≤20MB"
    )
with col2:
    sample_data = generate_sample_excel()
    st.download_button(
        label="📥 下载示例文件",
        data=sample_data,
        file_name="sample_order_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="下载示例数据，了解所需列格式"
    )

# ---------- 读取文件 ----------
if uploaded_file is not None:
    try:
        with st.spinner("正在读取文件..."):
            df = load_file(uploaded_file)
            st.session_state.df = df
        st.success(f"文件加载成功！共 {df.shape[0]} 行，{df.shape[1]} 列")
        st.subheader("数据预览（前5行）")
        st.dataframe(df.head(), use_container_width=True)

        auto_result = auto_map_columns(df)
        st.session_state.auto_map_result = auto_result
    except Exception as e:
        st.error(f"文件读取失败：{str(e)}")
        st.stop()
else:
    st.info("请上传文件或下载示例文件体验")
    st.stop()

# ---------- 字段映射 ----------
st.subheader("🔗 字段映射")
st.markdown("请将报表中的列对应到以下业务字段（支持多选平台费用）")

all_columns = st.session_state.df.columns.tolist()
auto = st.session_state.auto_map_result

sales_col = st.selectbox(
    "💰 销售额列（实收款/订单金额）",
    options=all_columns,
    index=all_columns.index(auto["sales_col"]) if auto["sales_col"] in all_columns else 0,
)

refund_options = [None] + all_columns
refund_index = 0
if auto["refund_col"] and auto["refund_col"] in all_columns:
    refund_index = all_columns.index(auto["refund_col"]) + 1
refund_col = st.selectbox(
    "🔁 退款额列（退款金额）",
    options=refund_options,
    index=refund_index,
    format_func=lambda x: "未选择" if x is None else x,
)

default_fee_cols = [col for col in auto["fee_cols"] if col in all_columns]
fee_cols = st.multiselect(
    "💸 平台费用列（可多选：佣金、广告费、技术服务费等）",
    options=all_columns,
    default=default_fee_cols,
)

# ---------- 计算利润按钮 ----------
if st.button("📈 计算利润", type="primary", use_container_width=True):
    if sales_col is None:
        st.error("请选择销售额列")
        st.stop()
    refund_selected = refund_col if refund_col is not None else None
    if refund_selected is None:
        st.warning("未选择退款列，将假定退款为0")

    with st.spinner("正在计算指标..."):
        try:
            total_sales, total_refund, net_revenue, total_fees, gross_profit, refund_rate, fee_rate = calculate_metrics(
                st.session_state.df, sales_col, refund_selected, fee_cols
            )
            st.session_state.metrics = {
                "total_sales": total_sales,
                "total_refund": total_refund,
                "net_revenue": net_revenue,
                "total_fees": total_fees,
                "gross_profit": gross_profit,
                "refund_rate": refund_rate,
                "fee_rate": fee_rate,
                "sales_col": sales_col,
                "refund_col": refund_selected,
                "fee_cols": fee_cols
            }
            st.session_state.report_ready = False
        except Exception as e:
            st.error(f"计算失败：{str(e)}")
            st.stop()

# ---------- 展示指标和图表 ----------
if st.session_state.metrics is not None:
    m = st.session_state.metrics

    st.subheader("📊 核心指标")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("净收入", f"¥{m['net_revenue']:,.2f}")
    with col2:
        st.metric("预估毛利润", f"¥{m['gross_profit']:,.2f}")
    with col3:
        st.metric("退款率", f"{m['refund_rate']:.2f}%")
    with col4:
        st.metric("费用率", f"{m['fee_rate']:.2f}%")

    # 饼图
    fig_pie = px.pie(
        names=["平台费用", "毛利润"],
        values=[m['total_fees'], m['gross_profit']],
        title="费用构成占比",
        color_discrete_sequence=["#FF6B6B", "#4ECDC4"]
    )
    st.plotly_chart(fig_pie, use_container_width=True)

    # 折线图（如果存在日期列）
    date_col = detect_date_column(st.session_state.df)
    if date_col:
        st.subheader("📅 每日净收入趋势")
        df_temp = st.session_state.df.copy()
        df_temp["净收入(行)"] = df_temp[m['sales_col']] - (df_temp[m['refund_col']] if m['refund_col'] else 0)
        df_temp["日期"] = pd.to_datetime(df_temp[date_col])
        daily_net = df_temp.groupby("日期")["净收入(行)"].sum().reset_index()
        fig_line = px.line(daily_net, x="日期", y="净收入(行)", title="每日净收入变化")
        st.plotly_chart(fig_line, use_container_width=True)
    else:
        st.info("未检测到日期列，如需每日趋势图请在报表中包含日期/时间列")

    # ---------- 利润总结（含期间） ----------
    st.subheader("📝 利润总结")
    shop_name = st.text_input("店铺名称", value="我的店铺", help="用于生成利润总结文案")

    start_date, end_date = get_date_range(st.session_state.df)
    period_str = f"{start_date} 至 {end_date}" if start_date else "未检测到日期列"

    summary_text = f"""【{shop_name}】利润速报
📅 统计期间：{period_str}
💰 净收入：¥{m['net_revenue']:,.2f}
📈 预估毛利润：¥{m['gross_profit']:,.2f}
🔁 退款率：{m['refund_rate']:.1f}%
💸 费用率：{m['fee_rate']:.1f}%

总结：店铺净收入为 {m['net_revenue']:,.0f} 元，毛利润为 {m['gross_profit']:,.0f} 元。退款率 {m['refund_rate']:.1f}%，费用占比 {m['fee_rate']:.1f}%。"""

    st.text_area("利润总结（可复制）", value=summary_text, height=220, key="profit_summary")
    if st.button("📋 复制总结", key="copy_summary"):
        st.info("请手动选中上方文本后复制（Ctrl+C）")

    # ---------- 报告生成与下载（无限制） ----------
    st.subheader("📄 报告下载")
    if st.button("生成报告（Excel）", use_container_width=True):
        with st.spinner("正在生成报告..."):
            report_path = generate_report(
                st.session_state.df,
                m['sales_col'], m['refund_col'], m['fee_cols'],
                m['total_sales'], m['total_refund'], m['net_revenue'],
                m['total_fees'], m['gross_profit'], m['refund_rate'], m['fee_rate']
            )
            st.session_state.report_file_path = report_path
            st.session_state.report_ready = True
            st.rerun()

    if st.session_state.report_ready and st.session_state.report_file_path:
        with open(st.session_state.report_file_path, "rb") as f:
            file_bytes = f.read()
        st.download_button(
            label="⬇️ 下载报告",
            data=file_bytes,
            file_name="店铺利润报告.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # ---------- 经营数据自动分析 ----------
    st.subheader("🔍 经营数据自动分析")
    net_rev = m['net_revenue']
    refund = m['refund_rate']
    fee = m['fee_rate']

    analysis = []
    if net_rev <= 0:
        analysis.append("⚠️ **净收入为负或零**：当前处于亏损状态，建议审查成本结构或提高售价。")
    elif net_rev < 1000:
        analysis.append("📉 **净收入较低**：可能处于起步阶段，建议加大推广或优化选品。")
    elif net_rev < 10000:
        analysis.append("📈 **净收入中等**：保持稳定，可尝试提升客单价或复购率。")
    else:
        analysis.append("🎉 **净收入优秀**：盈利状况良好，建议关注费用控制以进一步提升利润。")

    if refund < 3:
        analysis.append("✅ **退款率极低**（<3%）：商品质量和描述一致性很好，保持优势。")
    elif refund < 8:
        analysis.append("👍 **退款率正常**（3%-8%）：属于行业平均水平，可继续关注退款原因。")
    elif refund < 15:
        analysis.append("⚠️ **退款率偏高**（8%-15%）：建议检查商品描述、物流或客服响应。")
    else:
        analysis.append("🔴 **退款率过高**（>15%）：严重侵蚀利润，需紧急排查产品质量或发货问题。")

    if fee < 10:
        analysis.append("💰 **费用率较低**（<10%）：平台成本控制优秀，盈利空间大。")
    elif fee < 20:
        analysis.append("📊 **费用率合理**（10%-20%）：符合常规电商开销，可关注广告费优化。")
    elif fee < 30:
        analysis.append("⚡ **费用率偏高**（20%-30%）：建议审查广告投放ROI或佣金政策。")
    else:
        analysis.append("💸 **费用率过高**（>30%）：严重压缩利润，需重新评估营销支出。")

    if m['gross_profit'] <= 0:
        analysis.append("🚨 **毛利润为负**：必须立即调整定价或降低可变成本。")
    elif m['gross_profit'] / net_rev < 0.1:
        analysis.append("📉 **毛利率低于10%**：盈利脆弱，需提高售价或降低成本。")
    elif m['gross_profit'] / net_rev < 0.2:
        analysis.append("📊 **毛利率中等**（10%-20%）：可考虑增加高毛利产品。")
    else:
        analysis.append("🌟 **毛利率优秀**（>20%）：盈利能力强，可扩大规模。")

    analysis.append("\n**综合建议**：" +
                    ("当前店铺整体健康，继续保持优化。" if net_rev > 0 and refund < 10 and fee < 25 else
                     "存在改进空间，请重点关注上述红色预警项。"))

    for line in analysis:
        st.markdown(line)
    st.caption("注：以上分析基于通用电商经验，具体业务请结合实际情况判断。")