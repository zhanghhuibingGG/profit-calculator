"""
数据处理模块：文件读取、字段自动识别、指标计算、报告生成
"""
import pandas as pd
import tempfile
import io

def auto_map_columns(df: pd.DataFrame, api_key: str = None) -> dict:
    """
    根据列名关键词自动识别销售额、退款额、费用列。
    预留大模型API接口，当前使用规则匹配。
    """
    cols = df.columns.tolist()
    sales_keywords = ["销售", "实收", "订单金额", "销售额", "实收款", "交易金额", "买家实付"]
    refund_keywords = ["退款", "退单", "退款金额", "退货"]
    fee_keywords = ["佣金", "服务费", "广告费", "技术服务费", "平台费", "费用", "扣款", "营销支出"]

    sales_col = None
    refund_col = None
    fee_cols = []

    for col in cols:
        col_lower = col.lower()
        if sales_col is None and any(kw in col_lower for kw in sales_keywords):
            sales_col = col
        if refund_col is None and any(kw in col_lower for kw in refund_keywords):
            refund_col = col
        if any(kw in col_lower for kw in fee_keywords):
            fee_cols.append(col)

    if sales_col is None:
        for col in cols:
            if pd.api.types.is_numeric_dtype(df[col]):
                sales_col = col
                break

    return {"sales_col": sales_col, "refund_col": refund_col, "fee_cols": fee_cols}

def load_file(uploaded_file) -> pd.DataFrame:
    """读取上传的Excel/CSV文件"""
    file_ext = uploaded_file.name.split('.')[-1].lower()
    if file_ext == 'csv':
        df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
    elif file_ext in ['xlsx', 'xls']:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
    else:
        raise ValueError("不支持的文件格式，请上传 .xlsx, .xls 或 .csv 文件")
    return df

def generate_sample_excel():
    """生成包含丰富数据的示例文件（返回字节数据）"""
    sample_data = {
        "订单号": ["ORD001", "ORD002", "ORD003", "ORD004", "ORD005", "ORD006", "ORD007", "ORD008", "ORD009", "ORD010"],
        "订单状态": ["已完成", "已完成", "已退款", "已完成", "已完成", "已完成", "退款中", "已完成", "已完成", "已完成"],
        "商品名称": ["T恤-白色M", "牛仔裤-蓝色L", "帽子-黑色", "运动鞋-42", "连衣裙-S", "外套-L", "围巾", "手套", "袜子3双", "背包"],
        "实收款": [89.0, 199.0, 49.0, 329.0, 159.0, 399.0, 29.0, 39.0, 25.0, 129.0],
        "退款金额": [0, 0, 49.0, 0, 0, 0, 29.0, 0, 0, 0],
        "商品数量": [1, 1, 1, 1, 1, 1, 1, 2, 3, 1],
        "商品单价": [89.0, 199.0, 49.0, 329.0, 159.0, 399.0, 29.0, 19.5, 8.33, 129.0],
        "运费": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        "佣金": [2.67, 5.97, 1.47, 9.87, 4.77, 11.97, 0.87, 1.17, 0.75, 3.87],
        "广告费": [1.5, 5.0, 0.5, 12.0, 3.0, 15.0, 0.2, 0.5, 0.3, 2.0],
        "技术服务费": [0.89, 1.99, 0.49, 3.29, 1.59, 3.99, 0.29, 0.39, 0.25, 1.29],
        "营销支出": [0.5, 1.2, 0, 2.5, 1.0, 3.0, 0, 0.1, 0.1, 0.8],
        "订单日期": [
            "2025-03-01", "2025-03-01", "2025-03-02", "2025-03-02", "2025-03-03",
            "2025-03-04", "2025-03-05", "2025-03-06", "2025-03-07", "2025-03-08"
        ],
        "付款时间": [
            "2025-03-01 10:23:00", "2025-03-01 14:15:00", "2025-03-02 09:45:00", "2025-03-02 18:30:00",
            "2025-03-03 11:00:00", "2025-03-04 16:20:00", "2025-03-05 08:50:00", "2025-03-06 12:10:00",
            "2025-03-07 19:40:00", "2025-03-08 07:30:00"
        ]
    }
    df = pd.DataFrame(sample_data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="订单明细")
    output.seek(0)
    return output.getvalue()

def calculate_metrics(df: pd.DataFrame, sales_col: str, refund_col: str, fee_cols: list):
    """计算核心指标"""
    total_sales = df[sales_col].sum()
    total_refund = abs(df[refund_col].sum()) if refund_col else 0
    net_revenue = total_sales - total_refund
    total_fees = df[fee_cols].sum().sum() if fee_cols else 0
    gross_profit = net_revenue - total_fees
    refund_rate = (total_refund / total_sales * 100) if total_sales else 0
    fee_rate = (total_fees / net_revenue * 100) if net_revenue else 0
    return total_sales, total_refund, net_revenue, total_fees, gross_profit, refund_rate, fee_rate

def generate_report(df: pd.DataFrame, sales_col: str, refund_col: str, fee_cols: list,
                    total_sales, total_refund, net_revenue, total_fees, gross_profit, refund_rate, fee_rate):
    """生成报告并返回临时文件路径"""
    df_detail = df.copy()
    df_detail["净收入(行)"] = df_detail[sales_col] - (df_detail[refund_col] if refund_col else 0)

    if total_sales > 0 and fee_cols:
        df_detail["分摊费用"] = df_detail[sales_col] / total_sales * total_fees
        df_detail["预估利润(行)"] = df_detail["净收入(行)"] - df_detail["分摊费用"]
    else:
        df_detail["分摊费用"] = 0
        df_detail["预估利润(行)"] = df_detail["净收入(行)"]

    summary_data = {
        "指标名称": ["总销售额", "总退款额", "净收入", "平台费用合计", "预估毛利润", "退款率", "费用率"],
        "数值": [total_sales, total_refund, net_revenue, total_fees, gross_profit, refund_rate, fee_rate],
        "单位": ["元", "元", "元", "元", "元", "%", "%"]
    }
    df_summary = pd.DataFrame(summary_data)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name="汇总指标", index=False)
            df_detail.to_excel(writer, sheet_name="明细数据", index=False)
        return tmp.name

def detect_date_column(df: pd.DataFrame):
    """自动检测日期列"""
    for col in df.columns:
        if "日期" in col or "时间" in col or "date" in col.lower():
            try:
                pd.to_datetime(df[col])
                return col
            except:
                continue
    return None