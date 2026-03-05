import streamlit as st
import pandas as pd
import numpy as np
from datetime import timedelta, date
import re

st.set_page_config(page_title="BillMatch - 个人对账助手", layout="wide")

st.title("💰 BillMatch - 个人对账助手")
st.markdown("上传 **记账软件导出文件** (钱迹) 和 **银行账单** (Excel)，自动找出漏记或错误的账单。")

# --- 侧边栏：文件上传 ---
st.sidebar.header("1. 上传文件")
qianji_file = st.sidebar.file_uploader("上传钱迹导出文件 (xlsx)", type=["xlsx", "xls"])
bill_files = st.sidebar.file_uploader("上传银行账单文件 (支持多选)", type=["xlsx", "xls"], accept_multiple_files=True)

# --- 侧边栏：配置 ---
st.sidebar.header("2. 对账设置")

# 日期范围选择器
today = date.today()
# 默认为当月（1号到今天）
default_start = today.replace(day=1)
date_range = st.sidebar.date_input("日期范围 (必选)", value=(default_start, today), help="只对账此范围内的交易，漏记和多余记账也只显示此范围内的数据。")

target_card = st.sidebar.text_input("目标卡号 (末四位)", value="8820", help="只对账包含此卡号的记录。留空则不对卡号进行强制过滤。")
days_tolerance = st.sidebar.slider("日期容差 (天)", 0, 7, 2, help="记账日期和账单日期允许相差的天数")

# --- 辅助函数 ---
def load_excel(file):
    try:
        # 尝试直接读取
        return pd.read_excel(file)
    except Exception:
        # 一些旧的 xls 文件需要特殊处理或跳过表头
        try:
            return pd.read_excel(file, header=1) # 尝试跳过第一行
        except:
            return pd.read_excel(file, engine='xlrd') # 尝试使用 xlrd 读取旧版 xls

def normalize_date(df, col_name):
    df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
    # 关键修复：归一化为午夜（去除时间部分）以进行精确的日期比较
    df[col_name] = df[col_name].dt.normalize()
    return df

def normalize_amount(df, col_name):
    # 如果是字符串，去除货币符号，转换为浮点数
    if df[col_name].dtype == object:
        df[col_name] = df[col_name].astype(str).str.replace('￥', '').str.replace(',', '', regex=False)
    df[col_name] = pd.to_numeric(df[col_name], errors='coerce').fillna(0.0)
    return df

def extract_card_tail(val):
    """从字符串中提取最后 4 位数字，忽略所有非数字字符。"""
    if pd.isna(val):
        return None
    
    # 处理浮点数 8820.0 的情况
    if isinstance(val, float):
        if val.is_integer():
            val = int(val)
            
    s = str(val)
    # 去除除数字以外的所有内容
    digits = re.sub(r'\D', '', s)
    if len(digits) >= 4:
        return digits[-4:]
    return None

# --- 主逻辑 ---
if qianji_file and bill_files:
    # 加载数据
    try:
        df_q = load_excel(qianji_file)
        
        # 加载并合并多个账单文件
        bill_dfs = []
        for f in bill_files:
            df_temp = load_excel(f)
            # 启发式：为每个文件查找标题行
            if "交易日期" not in df_temp.columns and len(df_temp) > 1:
                for i in range(5):
                    row_vals = df_temp.iloc[i].astype(str).values.tolist()
                    if any("交易日期" in s or "日期" in s for s in row_vals):
                        f.seek(0) 
                        df_temp = pd.read_excel(f, header=i+1)
                        break
            bill_dfs.append(df_temp)
        
        # 在合并前标准化列名以处理细微差异（例如空格）
        if bill_dfs:
            for i in range(len(bill_dfs)):
                bill_dfs[i].columns = bill_dfs[i].columns.astype(str).str.strip().str.replace('\n', '')
        
        if bill_dfs:
            df_b = pd.concat(bill_dfs, ignore_index=True)
        else:
            df_b = pd.DataFrame()

        st.success(f"成功加载! 钱迹: {len(df_q)} 条, 账单: {len(bill_files)} 份 (共 {len(df_b)} 条)")

        # 列选择
        st.subheader("3. 列映射配置")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**记账文件 (钱迹)**")
            q_cols = df_q.columns.tolist()
            q_date_idx = next((i for i, c in enumerate(q_cols) if "时间" in str(c) or "日期" in str(c)), 0)
            q_amt_idx = next((i for i, c in enumerate(q_cols) if "金额" in str(c)), 0)
            q_acc1_idx = next((i for i, c in enumerate(q_cols) if "账户" in str(c)), 0)
            q_acc2_idx = next((i for i, c in enumerate(q_cols) if "账户2" in str(c) or "转入" in str(c)), min(q_acc1_idx+1, len(q_cols)-1))
            q_desc_idx = next((i for i, c in enumerate(q_cols) if "备注" in str(c) or "说明" in str(c)), 0)

            q_date_col = st.selectbox("选择[时间]列", q_cols, index=q_date_idx, key='q_date')
            q_amt_col = st.selectbox("选择[金额]列", q_cols, index=q_amt_idx, key='q_amt')
            q_acc1_col = st.selectbox("选择[账户1]列", q_cols, index=q_acc1_idx, key='q_acc1', help="主账户/支出账户")
            q_acc2_col = st.selectbox("选择[账户2]列", q_cols, index=q_acc2_idx, key='q_acc2', help="可选：转入账户/对方账户。")
            q_desc_col = st.selectbox("选择[备注]列 (用于展示)", q_cols, index=q_desc_idx, key='q_desc')

        with col2:
            st.markdown("**银行账单**")
            b_cols = df_b.columns.tolist()
            b_date_idx = next((i for i, c in enumerate(b_cols) if "交易日期" in str(c) or "日期" in str(c)), 0)
            b_amt_idx = next((i for i, c in enumerate(b_cols) if "交易金额" in str(c) or "金额" in str(c) or "支出" in str(c)), 0)
            b_card_idx = next((i for i, c in enumerate(b_cols) if "卡号" in str(c) or "卡末四位" in str(c) or "卡" in str(c)), 0)
            b_desc_idx = next((i for i, c in enumerate(b_cols) if "描述" in str(c) or "摘要" in str(c) or "商户" in str(c)), 0)

            b_date_col = st.selectbox("选择[交易日期]列", b_cols, index=b_date_idx, key='b_date')
            b_amt_col = st.selectbox("选择[交易金额]列", b_cols, index=b_amt_idx, key='b_amt')
            b_card_col = st.selectbox("选择[卡末四位]列", b_cols, index=b_card_idx, key='b_card')
            b_desc_col = st.selectbox("选择[交易描述]列", b_cols, index=b_desc_idx, key='b_desc')

        if st.button("开始对账"):
            # 1. 预处理与标准化
            df_q_clean = df_q.copy()
            df_b_clean = df_b.copy()
            
            # 应用日期标准化（去除时间）
            normalize_date(df_q_clean, q_date_col)
            normalize_amount(df_q_clean, q_amt_col)
            
            # 从两个账户列中提取卡号末位
            df_q_clean['_card1'] = df_q_clean[q_acc1_col].apply(extract_card_tail)
            df_q_clean['_card2'] = df_q_clean[q_acc2_col].apply(extract_card_tail)
            
            normalize_date(df_b_clean, b_date_col)
            normalize_amount(df_b_clean, b_amt_col)
            df_b_clean['_b_card_clean'] = df_b_clean[b_card_col].apply(extract_card_tail)

            # 2. 按日期范围过滤（新功能）
            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_d, end_d = date_range
                
                # 过滤钱迹数据
                df_q_clean = df_q_clean[
                    (df_q_clean[q_date_col].dt.date >= start_d) & 
                    (df_q_clean[q_date_col].dt.date <= end_d)
                ]
                
                # 过滤账单数据
                df_b_clean = df_b_clean[
                    (df_b_clean[b_date_col].dt.date >= start_d) & 
                    (df_b_clean[b_date_col].dt.date <= end_d)
                ]
                
                st.info(f"📅 日期范围: {start_d} ~ {end_d} | 范围筛选后: 钱迹 {len(df_q_clean)} 条, 账单 {len(df_b_clean)} 条")
            else:
                st.warning("⚠️ 请选择完整的开始和结束日期")
                st.stop()

            # 3. 按目标卡号过滤（如果提供）
            if target_card:
                target_clean = extract_card_tail(target_card)
                if target_clean:
                    # 过滤钱迹数据：如果账户1或账户2匹配目标则保留
                    df_q_clean = df_q_clean[
                        (df_q_clean['_card1'] == target_clean) | 
                        (df_q_clean['_card2'] == target_clean)
                    ]
                    
                    df_b_clean = df_b_clean[df_b_clean['_b_card_clean'] == target_clean]
                    
                    st.success(f"💳 卡号过滤 **{target_clean}**: 最终待比对 钱迹 {len(df_q_clean)} 条, 账单 {len(df_b_clean)} 条")
                else:
                    st.warning("输入的卡号无法识别为数字，跳过过滤。")

            # 4. 匹配逻辑
            # 候选池：仅包含未匹配的钱迹条目
            df_q_pool = df_q_clean[df_q_clean[q_amt_col] != 0].copy()
            df_q_pool['matched'] = False
            
            df_b_clean['match_status'] = "未匹配 (漏记?)"
            df_b_clean['match_id'] = None
            
            matched_pairs = []

            for idx, row in df_b_clean.iterrows():
                b_date = row[b_date_col]
                b_amt = row[b_amt_col]
                
                if pd.isna(b_date) or b_amt == 0:
                    df_b_clean.at[idx, 'match_status'] = "忽略 (无效数据)"
                    continue

                # 在候选池中搜索
                candidates = df_q_pool[
                    (~df_q_pool['matched']) & 
                    (np.isclose(df_q_pool[q_amt_col].abs(), abs(b_amt), atol=0.01)) & 
                    (df_q_pool[q_date_col] >= b_date - timedelta(days=days_tolerance)) & 
                    (df_q_pool[q_date_col] <= b_date + timedelta(days=days_tolerance))
                ]
                
                if not candidates.empty:
                    # 选择最接近的日期
                    best_match = candidates.iloc[(candidates[q_date_col] - b_date).abs().argsort()[:1]]
                    match_idx = best_match.index[0]
                    
                    df_q_pool.at[match_idx, 'matched'] = True
                    df_b_clean.at[idx, 'match_status'] = "已匹配"
                    df_b_clean.at[idx, 'match_id'] = match_idx
                    
                    matched_pairs.append({
                        "Bill Date": b_date,
                        "Bill Amount": b_amt,
                        "Bill Desc": row[b_desc_col],
                        "QianJi Date": df_q_pool.at[match_idx, q_date_col],
                        "QianJi Desc": df_q_pool.at[match_idx, q_desc_col],
                        "QianJi Acc1": df_q_pool.at[match_idx, q_acc1_col],
                        "QianJi Acc2": df_q_pool.at[match_idx, q_acc2_col]
                    })
            
            # 结果
            unmatched_bills = df_b_clean[df_b_clean['match_status'] == "未匹配 (漏记?)"]
            unmatched_bills = unmatched_bills.sort_values(by=b_date_col, ascending=False)
            
            unmatched_qianji = df_q_pool[~df_q_pool['matched']]
            unmatched_qianji = unmatched_qianji.sort_values(by=q_date_col, ascending=False)
            
            st.divider() 
            
            t1, t2, t3 = st.tabs([f"🚨 漏记账单 ({len(unmatched_bills)})", f"✅ 已匹配 ({len(matched_pairs)})", f"❓ 多余记账 ({len(unmatched_qianji)})"])
            
            with t1:
                st.caption(f"此列表仅显示日期 **{date_range[0]} ~ {date_range[1]}** 内，且属于卡号 **{target_card}** 的漏记项。")
                st.dataframe(unmatched_bills[[b_date_col, b_card_col, b_amt_col, b_desc_col]])
                st.markdown(f"**漏记总额**: {unmatched_bills[b_amt_col].sum():.2f}")

            with t2:
                df_matched = pd.DataFrame(matched_pairs)
                if not df_matched.empty:
                    df_matched = df_matched.sort_values(by="Bill Date", ascending=False)
                st.dataframe(df_matched)

            with t3:
                st.caption(f"此列表仅显示日期 **{date_range[0]} ~ {date_range[1]}** 内，且属于卡号 **{target_card}** 的多余项。")
                st.dataframe(unmatched_qianji[[q_date_col, q_acc1_col, q_acc2_col, q_amt_col, q_desc_col]])

    except Exception as e:
        st.error(f"处理文件时出错: {e}")
else:
    st.info("👈 请在左侧上传两个 Excel 文件以开始。")