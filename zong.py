import io
import re
import traceback
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

# --- 页面全局配置 ---
st.set_page_config(page_title="供应链欠料格式转换工具箱", layout="wide")
st.title("🛠️ 供应链欠料格式转换工具箱")

# --- 侧边栏：模式选择与通用参数 ---
st.sidebar.header("⚙️ 功能与规则设置")

app_mode = st.sidebar.radio(
    "请选择转换模式：",
    ["🔌 Cable Reel 格式 (4列明细)", "🧹 Floor Nozzle / SBD 格式 (5列明细)"],
)

st.sidebar.markdown("---")
digits_only_nozzle = st.sidebar.checkbox("Floor nozzle 仅保留数字", value=False)


# ==========================================
# 通用工具函数 (Common Helper Functions)
# ==========================================


def get_display_width(val):
    """计算字符串实际显示宽度（支持中文字符双倍宽度计算）"""
    s = str(val or "")
    return sum(2 if ord(c) > 127 else 1 for c in s)


def get_monday_date(date_str):
    """计算所在周周一的日期 (YYYY-MM-DD)"""
    if not date_str:
        return ""
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        monday_dt = dt - timedelta(days=dt.weekday())
        return monday_dt.strftime("%Y-%m-%d")
    except ValueError:
        return date_str


def parse_header_date(val, force_year=None):
    """强大的日期表头解析函数（支持标准格式、Excel序列号、中文年月日及月日短格式）"""
    if pd.isna(val):
        return None

    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%Y-%m-%d")

    try:
        num_val = float(val)
        if 30000 < num_val < 60000:
            d = pd.to_datetime(num_val, unit="D", origin="1899-12-30")
            return d.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass

    val_str = str(val).strip()
    if not val_str or val_str.lower() in ["nan", "none", "nat"]:
        return None

    curr_year = force_year if force_year else datetime.now().year

    # 正则：X年X月X日
    m_cn = re.search(
        r"(?:(\d{4})\s*年\s*)?(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?", val_str
    )
    if m_cn:
        y, m, d = m_cn.groups()
        year_num = int(y) if y else curr_year
        m_num, d_num = int(m), int(d)
        if 1 <= m_num <= 12 and 1 <= d_num <= 31:
            return f"{year_num:04d}-{m_num:02d}-{d_num:02d}"

    # 正则：YYYY-MM-DD
    m_full = re.search(
        r"(?<!\d)(\d{4})[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])(?!\d)",
        val_str,
    )
    if m_full:
        y, m, d = m_full.groups()
        return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"

    # 正则：MM-DD (短格式)
    m_short = re.search(
        r"(?<!\d)(0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])(?!\d)", val_str
    )
    if m_short:
        m, d = m_short.groups()
        return f"{curr_year:04d}-{int(m):02d}-{int(d):02d}"

    return None


def safe_convert_number(raw_val):
    """安全数值转换"""
    if pd.isna(raw_val):
        return None
    s = str(raw_val).strip().replace(",", "")
    if not s or s.lower() in ["nan", "none", "nat"]:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def clean_cell(val):
    """单元格清洗函数"""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ["nan", "none", "nat"]:
        return None
    return s


def find_target_column(df, keywords, max_scan_rows=10):
    """关键字模糊查找列索引"""
    keywords = [k.lower() for k in keywords]
    for r in range(min(max_scan_rows, len(df))):
        for c in range(len(df.columns)):
            cell = str(df.iloc[r, c]).strip().lower()
            if any(kw in cell for kw in keywords):
                return c
    return None


def read_uploaded_file(uploaded_file, key_suffix=""):
    """统一文件读取工具（自动适配 CSV 编码与 Excel 多 Sheet）"""
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file, header=None, encoding="utf-8")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(
                uploaded_file,
                header=None,
                encoding="gbk",
                encoding_errors="replace",
            )

    engine = "xlrd" if file_name.endswith(".xls") else "openpyxl"
    xl_file = pd.ExcelFile(uploaded_file, engine=engine)
    sheet_names = xl_file.sheet_names

    default_idx = 0
    match_sheet_keys = ["主计划", "cr", "cable", "reel", "sbd", "nozzle"]
    for idx, sheet_name in enumerate(sheet_names):
        if any(key in sheet_name.lower() for key in match_sheet_keys):
            default_idx = idx
            break

    selected_sheet = st.selectbox(
        "请选择要处理的工作表 (Sheet)：",
        sheet_names,
        index=default_idx,
        key=f"sheet_selector_{key_suffix}",
    )
    return pd.read_excel(xl_file, sheet_name=selected_sheet, header=None)


def build_excel_bytes(df_dict_or_df, sheet_name="Sheet1"):
    """导出 Excel 统一处理（支持单/多 Sheet，自动开启筛选、首行冻结、适应列宽）"""
    if isinstance(df_dict_or_df, pd.DataFrame):
        sheets = {sheet_name: df_dict_or_df}
    else:
        sheets = df_dict_or_df

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for s_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=s_name)
            worksheet = writer.sheets[s_name]

            # 1. 开启 Excel 自动筛选
            worksheet.auto_filter.ref = worksheet.dimensions
            # 2. 冻结首行
            worksheet.freeze_panes = "A2"

            # 3. 自动适应列宽
            for col in worksheet.columns:
                max_len = max(get_display_width(cell.value) for cell in col)
                col_letter = get_column_letter(col[0].column)
                worksheet.column_dimensions[col_letter].width = max(
                    max_len + 5, 15
                )

    return output.getvalue()


def smart_round(x):
    """智能数值舍入"""
    if abs(x - round(x)) < 1e-6:
        return int(round(x))
    return round(x, 2)


# ==========================================
# 业务逻辑实现
# ==========================================

# ------------------------------------------
# 模式 1: Cable Reel 格式转换
# ------------------------------------------
if "Cable Reel" in app_mode:
    st.subheader("🔌 Cable Reel 欠料格式转换")
    file1 = st.file_uploader(
        "请上传 Cable Reel 原始文件 (.xlsx / .xls / .csv)",
        type=["xlsx", "xls", "csv"],
        key="uploader_cable",
    )

    if file1:
        try:
            df_raw = read_uploaded_file(file1, key_suffix="cable")

            # 定位关键列
            cable_col_idx = find_target_column(
                df_raw, ["cable", "reel", "线缆", "型号", "料号"]
            )
            remark_col_idx = find_target_column(
                df_raw, ["remark", "备注", "nozzle", "地刷"]
            )

            if cable_col_idx is None:
                cable_col_idx = 0
                st.warning("未自动识别线缆列，默认使用第 A 列 (索引 0)")
            if remark_col_idx is None:
                remark_col_idx = 1
                st.warning("未自动识别备注/地刷列，默认使用第 B 列 (索引 1)")

            meta_cols = {remark_col_idx, cable_col_idx}

            # 扫描日期表头行
            best_row_idx, max_date_count = None, 0
            for r in range(min(15, len(df_raw))):
                row_vals = df_raw.iloc[r].tolist()
                parsed_vals = [parse_header_date(v) for v in row_vals]
                date_count = sum(1 for p in parsed_vals if p is not None)
                if date_count > max_date_count:
                    max_date_count = date_count
                    best_row_idx = r

            col_date_map = {}
            if best_row_idx is not None and max_date_count > 0:
                header_row = df_raw.iloc[best_row_idx]
                for c in range(len(df_raw.columns)):
                    if c not in meta_cols:
                        parsed_d = parse_header_date(header_row.iloc[c])
                        if parsed_d:
                            col_date_map[c] = parsed_d

            # 单元格清洗与填充
            df_filled = df_raw.copy()
            for col in [remark_col_idx, cable_col_idx]:
                df_filled[col] = df_filled[col].apply(clean_cell).ffill().bfill()

            # 提取数据
            raw_records = []
            plan_keywords = ["计划", "plan", "計劃"]

            for r in range(len(df_raw)):
                row_values = df_raw.iloc[r].tolist()
                row_str = " ".join(str(v).strip().lower() for v in row_values)

                if not any(kw in row_str for kw in plan_keywords) or "计划库存" in row_str:
                    continue

                if r > 0:
                    floor_nozzle_val = (
                        clean_cell(df_filled.iloc[r - 1, cable_col_idx]) or ""
                    )
                    full_name_val = (
                        clean_cell(df_filled.iloc[r - 1, remark_col_idx]) or ""
                    )
                else:
                    floor_nozzle_val, full_name_val = "", ""

                if digits_only_nozzle and floor_nozzle_val:
                    nozzle_num = re.sub(r"\D", "", floor_nozzle_val)
                    if nozzle_num:
                        floor_nozzle_val = nozzle_num

                if not full_name_val and not floor_nozzle_val:
                    continue

                for c, raw_date in col_date_map.items():
                    if c < len(row_values):
                        num = safe_convert_number(row_values[c])
                        # 仅提取大于 0 的有效数值
                        if num is not None and num > 0:
                            monday_date = get_monday_date(raw_date)
                            raw_records.append(
                                {
                                    "Floor nozzle": floor_nozzle_val,
                                    "FullName": full_name_val,
                                    "Daily Date": raw_date,
                                    "Monday Date": monday_date,
                                    "Released": num,
                                }
                            )

            if raw_records:
                df_temp = pd.DataFrame(raw_records).astype(
                    {
                        "Floor nozzle": str,
                        "FullName": str,
                        "Daily Date": str,
                        "Monday Date": str,
                    }
                )
                df_temp["Released"] = pd.to_numeric(df_temp["Released"], errors="coerce")

                # --- 按周汇总 ---
                df_weekly = (
                    df_temp.groupby(
                        ["Floor nozzle", "FullName", "Monday Date"], as_index=False
                    )["Released"]
                    .sum()
                )
                df_weekly["Released"] = df_weekly["Released"].apply(smart_round)
                df_weekly = df_weekly[df_weekly["Released"] > 0]
                df_weekly = df_weekly.rename(columns={"Monday Date": "Demand Time"})
                df_weekly = df_weekly[["Floor nozzle", "FullName", "Demand Time", "Released"]]

                # --- 按天汇总 ---
                df_daily = (
                    df_temp.groupby(
                        ["Floor nozzle", "FullName", "Daily Date"], as_index=False
                    )["Released"]
                    .sum()
                )
                df_daily["Released"] = df_daily["Released"].apply(smart_round)
                df_daily = df_daily[df_daily["Released"] > 0]
                df_daily = df_daily.rename(columns={"Daily Date": "Demand Time"})
                df_daily = df_daily[["Floor nozzle", "FullName", "Demand Time", "Released"]]

                st.subheader("📋 转换结果预览 (Cable Reel 格式)")
                tab_weekly, tab_daily = st.tabs(["📅 按周汇总 (归集至周一)", "📆 按天汇总 (每日明细)"])

                with tab_weekly:
                    st.dataframe(df_weekly, use_container_width=True)

                with tab_daily:
                    st.dataframe(df_daily, use_container_width=True)

                excel_bytes = build_excel_bytes(
                    {"按周汇总": df_weekly, "按天汇总": df_daily}
                )
                now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="✅ 下载 Cable Reel 导出 Excel (带筛选)",
                    data=excel_bytes,
                    file_name=f"CR_欠料格式_{now_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("⚠️ 未检测到包含“计划”的有效数据行。")

        except Exception as e:
            st.error(f"❌ 处理出错: {str(e)}")
            with st.expander("点击查看完整报错堆栈"):
                st.code(traceback.format_exc())

# ------------------------------------------
# 模式 2: Floor Nozzle / SBD 格式转换
# ------------------------------------------
else:
    st.subheader("🧹 Floor Nozzle / SBD 欠料格式转换")

    file2 = st.file_uploader(
        "请上传 Floor Nozzle / SBD 原始文件 (.xlsx / .xls / .csv)",
        type=["xlsx", "xls", "csv"],
        key="uploader_chentou",
    )

    if file2:
        try:
            df_raw = read_uploaded_file(file2, key_suffix="chentou")

            best_row_idx, max_date_count = None, 0
            for r in range(min(15, len(df_raw))):
                row_vals = df_raw.iloc[r].tolist()
                parsed_vals = [parse_header_date(v) for v in row_vals]
                date_count = sum(1 for p in parsed_vals if p is not None)
                if date_count > max_date_count:
                    max_date_count = date_count
                    best_row_idx = r

            if best_row_idx is None or max_date_count == 0:
                st.error("❌ 未能在上传的文件中找到包含日期的表头列。")
                st.stop()

            header_row = df_raw.iloc[best_row_idx]
            col_date_map = {}
            for c in range(len(header_row)):
                d_parsed = parse_header_date(header_row.iloc[c])
                if d_parsed:
                    col_date_map[c] = d_parsed

            records = []
            plan_keywords = ["计划", "plan", "計劃"]

            for r in range(best_row_idx + 1, len(df_raw)):
                row_vals = [
                    str(x).strip()
                    if pd.notna(x) and str(x).strip().lower() not in ["nan", "none"]
                    else ""
                    for x in df_raw.iloc[r].tolist()
                ]
                row_str = " ".join(row_vals).lower()

                if not any(row_vals) or "计划库存" in row_str:
                    continue

                # 判定同行或上一行模式
                nozzle_same = row_vals[0]
                desc_same = row_vals[1]
                boxing_same = row_vals[2] if len(row_vals) > 2 else ""

                nozzle_prev = (
                    str(df_raw.iloc[r - 1, 0]).strip()
                    if r > 0 and pd.notna(df_raw.iloc[r - 1, 0])
                    else ""
                )
                desc_prev = (
                    str(df_raw.iloc[r - 1, 1]).strip()
                    if r > 0 and pd.notna(df_raw.iloc[r - 1, 1])
                    else ""
                )
                boxing_prev = (
                    str(df_raw.iloc[r - 1, 2]).strip()
                    if r > 0 and pd.notna(df_raw.iloc[r - 1, 2])
                    else ""
                )

                is_plan_row = any(kw in row_str for kw in plan_keywords)
                nozzle_val, desc_val, boxing_val = "", "", ""

                if nozzle_same and nozzle_same not in plan_keywords:
                    nozzle_val, desc_val, boxing_val = (
                        nozzle_same,
                        desc_same,
                        boxing_same,
                    )
                elif is_plan_row and r > 0 and nozzle_prev:
                    nozzle_val, desc_val, boxing_val = (
                        nozzle_prev,
                        desc_prev,
                        boxing_prev,
                    )

                if not nozzle_val or nozzle_val.lower() in ["nan", "none"]:
                    continue

                if digits_only_nozzle:
                    num_only = re.sub(r"\D", "", nozzle_val)
                    if num_only:
                        nozzle_val = num_only

                for c, date_str in col_date_map.items():
                    if c < len(row_vals):
                        num = safe_convert_number(row_vals[c])
                        # 仅提取大于 0 的有效数值
                        if num is not None and num > 0:
                            monday_date = get_monday_date(date_str)
                            records.append(
                                {
                                    "Floor nozzle": nozzle_val,
                                    "Description": desc_val,
                                    "分箱情况": boxing_val,
                                    "Daily Date": date_str,
                                    "Monday Date": monday_date,
                                    "Released": num,
                                }
                            )

            if records:
                df_out = pd.DataFrame(records).astype(
                    {
                        "Floor nozzle": str,
                        "Description": str,
                        "分箱情况": str,
                        "Daily Date": str,
                        "Monday Date": str,
                    }
                )
                df_out["Released"] = pd.to_numeric(df_out["Released"], errors="coerce")

                # --- 按周汇总 ---
                df_weekly = (
                    df_out.groupby(
                        ["Floor nozzle", "Description", "分箱情况", "Monday Date"],
                        as_index=False,
                    )["Released"]
                    .sum()
                )
                df_weekly["Released"] = df_weekly["Released"].apply(smart_round)
                df_weekly = df_weekly[df_weekly["Released"] > 0]
                df_weekly = df_weekly.rename(columns={"Monday Date": "Demand Time"})
                df_weekly = df_weekly.sort_values(
                    by=["Floor nozzle", "Demand Time"]
                ).reset_index(drop=True)

                # --- 按天汇总 ---
                df_daily = (
                    df_out.groupby(
                        ["Floor nozzle", "Description", "分箱情况", "Daily Date"],
                        as_index=False,
                    )["Released"]
                    .sum()
                )
                df_daily["Released"] = df_daily["Released"].apply(smart_round)
                df_daily = df_daily[df_daily["Released"] > 0]
                df_daily = df_daily.rename(columns={"Daily Date": "Demand Time"})
                df_daily = df_daily.sort_values(
                    by=["Floor nozzle", "Demand Time"]
                ).reset_index(drop=True)

                st.subheader("📋 转换结果预览（Floor Nozzle / SBD 格式）")
                tab_weekly, tab_daily = st.tabs(["📅 按周汇总 (归集至周一)", "📆 按天汇总 (每日明细)"])

                with tab_weekly:
                    st.dataframe(df_weekly, use_container_width=True)

                with tab_daily:
                    st.dataframe(df_daily, use_container_width=True)

                excel_bytes = build_excel_bytes(
                    {"按周汇总": df_weekly, "按天汇总": df_daily}
                )
                now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="✅ 下载 Floor Nozzle 导出 Excel (带筛选)",
                    data=excel_bytes,
                    file_name=f"Floor_Nozzle_明细汇总_{now_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("⚠️ 未能从文件中解析出有效的数量数据。")

        except Exception as e:
            st.error(f"❌ 处理出错: {str(e)}")
            with st.expander("点击查看报错堆栈"):
                st.code(traceback.format_exc())
