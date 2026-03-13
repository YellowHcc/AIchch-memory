#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银联数据比对工具 - V3 (适配实际数据结构)
功能：
1. 从银行流水摘要中提取门店名称和日期
2. 按门店+日期分组汇总后比对
3. 金额反查：根据差额定位未匹配流水
"""

import pandas as pd
import numpy as np
from datetime import datetime
from difflib import SequenceMatcher
import argparse
import sys
import re


def similarity(a, b):
    """计算两个字符串的相似度"""
    if pd.isna(a) or pd.isna(b):
        return 0
    return SequenceMatcher(None, str(a).strip(), str(b).strip()).ratio()


def find_best_match(store_name, store_list, threshold=0.6):
    """
    在门店列表中找到最佳匹配
    返回: (匹配的门店名, 相似度)
    """
    if pd.isna(store_name) or not store_list:
        return None, 0
    
    best_match = None
    best_score = 0
    
    for s in store_list:
        score = similarity(store_name, s)
        if score > best_score:
            best_score = score
            best_match = s
    
    if best_score >= threshold:
        return best_match, best_score
    return None, best_score


def standardize_date(date_val):
    """标准化日期格式为 YYYY-MM-DD"""
    if pd.isna(date_val):
        return None
    
    # 如果已经是datetime类型
    if isinstance(date_val, datetime):
        return date_val.strftime('%Y-%m-%d')
    
    # 尝试多种格式解析
    date_str = str(date_val).strip()
    
    # 处理纯数字格式如 "0311" 或 "20250311"
    if date_str.isdigit():
        if len(date_str) == 4:  # MMDD 格式，假设当年
            try:
                month = int(date_str[:2])
                day = int(date_str[2:])
                year = datetime.now().year
                dt = datetime(year, month, day)
                return dt.strftime('%Y-%m-%d')
            except:
                pass
        elif len(date_str) == 8:  # YYYYMMDD 格式
            try:
                dt = datetime.strptime(date_str, '%Y%m%d')
                return dt.strftime('%Y-%m-%d')
            except:
                pass
    
    formats = [
        '%Y-%m-%d',
        '%Y/%m/%d',
        '%Y%m%d',
        '%d/%m/%Y',
        '%m/%d/%Y',
        '%m-%d',
        '%m/%d',
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            if dt.year < 100:
                dt = dt.replace(year=dt.year + 2000)
            return dt.strftime('%Y-%m-%d')
        except:
            continue
    
    return date_str


def format_date_display(date_str):
    """将 YYYY-MM-DD 格式化为 MMDD 用于显示"""
    if not date_str or date_str == 'N/A':
        return 'N/A'
    try:
        dt = datetime.strptime(str(date_str), '%Y-%m-%d')
        return dt.strftime('%m%d')  # 返回如 0311
    except:
        return date_str


def extract_store_from_summary(summary):
    """从摘要中提取门店名称"""
    if pd.isna(summary):
        return None
    
    summary_str = str(summary).strip()
    
    # 模式1：银联入账：XXX店日期
    patterns = [
        r'银联入账：(.+?)(?:\d{4}|\d{2}|费|\+|-)',
        r'银联入账：(.+?)(?:费|\+|-)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, summary_str)
        if match:
            store = match.group(1).strip()
            store = store.rstrip('+').strip()
            # 过滤掉纯数字（提取失败的情况）
            if store and not store.isdigit():
                return store
    
    return None


def extract_date_from_summary(summary):
    """从摘要中提取日期（如0311）"""
    if pd.isna(summary):
        return None
    
    summary_str = str(summary).strip()
    
    # 提取 MMDD 模式（4位数字后跟-）
    match = re.search(r'(\d{4})-', summary_str)
    if match:
        date_mmdd = match.group(1)
        # 验证是有效的月日
        try:
            month = int(date_mmdd[:2])
            day = int(date_mmdd[2:])
            if 1 <= month <= 12 and 1 <= day <= 31:
                year = datetime.now().year
                dt = datetime(year, month, day)
                return dt.strftime('%Y-%m-%d')
        except:
            pass
    
    return None


def load_and_process_data(bill_path, flow_path):
    """
    加载并预处理数据（适配实际数据结构）
    """
    # 读取数据
    try:
        df_bill = pd.read_excel(bill_path)
        df_flow = pd.read_excel(flow_path)
    except Exception as e:
        print(f"读取文件失败: {e}")
        sys.exit(1)
    
    print(f"✓ 结算账单: {len(df_bill)} 行")
    print(f"✓ 银行流水: {len(df_flow)} 行")
    
    # 标准化列名
    df_bill.columns = [c.strip() for c in df_bill.columns]
    df_flow.columns = [c.strip() for c in df_flow.columns]
    
    # ========== 处理结算账单 ==========
    # 门店名称
    if 'POS门店名称' in df_bill.columns:
        df_bill['_store'] = df_bill['POS门店名称'].astype(str).str.strip()
    else:
        print("错误：结算账单缺少 'POS门店名称' 列")
        sys.exit(1)
    
    # 日期 - 使用结算日期（与银行流水到账日期对应）
    if '结算日期' in df_bill.columns:
        df_bill['_date'] = df_bill['结算日期'].apply(standardize_date)
    elif '营业日期' in df_bill.columns:
        df_bill['_date'] = df_bill['营业日期'].apply(standardize_date)
    else:
        print("错误：结算账单缺少日期列")
        sys.exit(1)
    
    # 金额
    if '交易金额' in df_bill.columns:
        df_bill['_amount'] = pd.to_numeric(df_bill['交易金额'], errors='coerce')
    else:
        print("错误：结算账单缺少 '交易金额' 列")
        sys.exit(1)
    
    # 手续费
    if '手续费' in df_bill.columns:
        df_bill['_fee'] = pd.to_numeric(df_bill['手续费'], errors='coerce').fillna(0)
    else:
        df_bill['_fee'] = 0
    
    # 结算日期（用于显示）
    if '结算日期' in df_bill.columns:
        df_bill['_settle_date'] = df_bill['结算日期'].apply(standardize_date)
    else:
        df_bill['_settle_date'] = df_bill['_date']
    
    # ========== 处理银行流水 ==========
    # 从摘要提取门店和日期
    df_flow['_extracted_store'] = df_flow['摘要'].apply(extract_store_from_summary)
    df_flow['_extracted_date'] = df_flow['摘要'].apply(extract_date_from_summary)
    
    # 金额
    if '交易金额' in df_flow.columns:
        df_flow['_amount'] = pd.to_numeric(df_flow['交易金额'], errors='coerce')
    else:
        print("错误：银行流水缺少 '交易金额' 列")
        sys.exit(1)
    
    # 从摘要提取手续费（如果有）
    def extract_fee(summary):
        if pd.isna(summary):
            return 0
        match = re.search(r'费([\d.]+)元?', str(summary))
        if match:
            try:
                return float(match.group(1))
            except:
                pass
        return 0
    
    df_flow['_fee'] = df_flow['摘要'].apply(extract_fee)
    
    # 保存原始数据用于后续分析
    df_flow['_summary'] = df_flow['摘要']
    
    return df_bill, df_flow


def reconcile_data_v3(df_bill, df_flow, amount_tolerance=0.01):
    """
    V3 比对逻辑：
    1. 建立门店映射（提取的门店名称 -> 结算单门店）
    2. 按（映射后门店+日期）分组汇总
    3. 金额比对
    4. 差额分析 - 找出未匹配的流水
    """
    results = []
    bill_stores = df_bill['_store'].unique().tolist()
    
    print("\n开始比对...")
    print("="*60)
    
    # 1. 建立门店映射
    print("\n【1. 门店映射】")
    store_mapping = {}
    flow_stores = df_flow['_extracted_store'].dropna().unique()
    
    for flow_store in flow_stores:
        matched_store, score = find_best_match(flow_store, bill_stores, threshold=0.5)
        if matched_store:
            store_mapping[flow_store] = matched_store
            print(f"  {flow_store} -> {matched_store} (相似度: {score:.1%})")
        else:
            print(f"  ⚠ {flow_store} -> 未找到匹配")
    
    print("="*60)
    
    # 2. 给流水添加匹配后的门店和日期
    df_flow['_matched_store'] = df_flow['_extracted_store'].map(store_mapping)
    df_flow['_matched_date'] = df_flow['_extracted_date']
    
    # 3. 按（匹配后门店+日期）分组汇总流水
    flow_valid = df_flow[df_flow['_matched_store'].notna() & df_flow['_matched_date'].notna()].copy()
    
    flow_grouped = flow_valid.groupby(['_matched_store', '_matched_date']).agg({
        '_amount': 'sum',      # 到账金额汇总
        '_fee': 'sum',         # 手续费汇总
        '_extracted_store': lambda x: ', '.join(x.unique()),  # 原始提取的门店名
        '摘要': lambda x: ' | '.join(x.tolist())  # 所有摘要
    }).reset_index()
    flow_grouped.columns = ['门店', '日期', '到账金额汇总', '手续费汇总', '原始门店名', '摘要汇总']
    
    # 4. 按（门店+日期）分组汇总结算单
    bill_grouped = df_bill.groupby(['_store', '_date']).agg({
        '_amount': 'sum',  # 交易金额汇总
        '_fee': 'sum',     # 手续费汇总
        '_settle_date': 'first'  # 结算日期（用于显示）
    }).reset_index()
    bill_grouped.columns = ['门店', '日期', '交易金额汇总', '手续费汇总', '结算日期']
    bill_grouped['净额汇总'] = bill_grouped['交易金额汇总'] - bill_grouped['手续费汇总']
    
    # 5. 逐组比对
    matched_bill_groups = set()
    unmatched_flow_details = []  # 记录未匹配的流水明细
    
    for idx, flow_group in flow_grouped.iterrows():
        store = flow_group['门店']
        date = flow_group['日期']
        flow_amount = flow_group['到账金额汇总']
        flow_fee = flow_group['手续费汇总']
        original_store = flow_group['原始门店名']
        summary = flow_group['摘要汇总']
        
        # 查找对应结算单
        bill_candidates = bill_grouped[
            (bill_grouped['门店'] == store) & 
            (bill_grouped['日期'] == date)
        ]
        
        if len(bill_candidates) == 0:
            # 无对应结算单
            results.append({
                '门店': store,
                '日期': format_date_display(date),
                '结算日期': 'N/A',
                '流水笔数': len(flow_valid[(flow_valid['_matched_store'] == store) & (flow_valid['_matched_date'] == date)]),
                '到账金额汇总': flow_amount,
                '流水手续费汇总': flow_fee,
                '交易金额汇总': 'N/A',
                '结算手续费汇总': 'N/A',
                '结算净额': 'N/A',
                '差额': 'N/A',
                '匹配状态': '✗ 无法匹配',
                '备注': '无对应结算单（门店或日期不匹配）',
            })
        else:
            bill_row = bill_candidates.iloc[0]
            bill_amount = bill_row['交易金额汇总']
            bill_fee = bill_row['手续费汇总']
            bill_net = bill_row['净额汇总']
            bill_settle_date = bill_row['结算日期']
            bill_key = (store, date)
            
            # 计算差额（到账金额 vs 结算净额）
            diff = flow_amount - bill_net
            
            # 判断是否匹配
            if abs(diff) <= amount_tolerance:
                # 完全匹配
                matched_bill_groups.add(bill_key)
                
                results.append({
                    '门店': store,
                    '日期': format_date_display(date),
                    '结算日期': format_date_display(bill_settle_date),
                    '流水笔数': len(flow_valid[(flow_valid['_matched_store'] == store) & (flow_valid['_matched_date'] == date)]),
                    '到账金额汇总': flow_amount,
                    '流水手续费汇总': flow_fee,
                    '交易金额汇总': bill_amount,
                    '结算手续费汇总': bill_fee,
                    '结算净额': bill_net,
                    '差额': diff,
                    '匹配状态': '✓ 已匹配',
                    '备注': '',
                })
            else:
                # 金额不匹配 - 标记但先不记录为未匹配
                matched_bill_groups.add(bill_key)  # 先标记为有结算单
                
                # 找出该组的所有流水明细
                flow_details = flow_valid[
                    (flow_valid['_matched_store'] == store) & 
                    (flow_valid['_matched_date'] == date)
                ].copy()
                
                if flow_amount < bill_net:
                    # 到账金额 < 结算净额：有结算金额找不到对应流水
                    status = '✗ 金额不符（结算多）'
                    note = f'结算净额比到账多 {abs(diff):.2f} 元，可能有未匹配流水'
                else:
                    # 到账金额 > 结算净额：有多余流水
                    status = '✗ 金额不符（流水多）'
                    note = f'到账比结算净额多 {abs(diff):.2f} 元，可能有多余流水'
                
                results.append({
                    '门店': store,
                    '日期': format_date_display(date),
                    '结算日期': format_date_display(bill_settle_date),
                    '流水笔数': len(flow_details),
                    '到账金额汇总': flow_amount,
                    '流水手续费汇总': flow_fee,
                    '交易金额汇总': bill_amount,
                    '结算手续费汇总': bill_fee,
                    '结算净额': bill_net,
                    '差额': diff,
                    '匹配状态': status,
                    '备注': note,
                })
                
                # 记录流水明细供后续分析
                for _, f in flow_details.iterrows():
                    unmatched_flow_details.append({
                        '门店组': store,
                        '日期组': format_date_display(date),
                        '匹配状态': status,
                        '摘要': f['_summary'],
                        '提取门店': f['_extracted_store'],
                        '到账金额': f['_amount'],
                        '提取手续费': f['_fee'],
                    })
    
    # 6. 找出结算单中未被匹配的记录
    unmatched_bill = []
    for idx, bill_row in bill_grouped.iterrows():
        bill_key = (bill_row['门店'], bill_row['日期'])
        if bill_key not in matched_bill_groups:
            unmatched_bill.append({
                '门店': bill_row['门店'],
                '日期': format_date_display(bill_row['日期']),
                '结算日期': format_date_display(bill_row['结算日期']),
                '交易金额汇总': bill_row['交易金额汇总'],
                '手续费汇总': bill_row['手续费汇总'],
                '净额汇总': bill_row['净额汇总'],
            })
    
    unmatched_bill_df = pd.DataFrame(unmatched_bill)
    unmatched_flow_df = pd.DataFrame(unmatched_flow_details)
    
    return pd.DataFrame(results), unmatched_bill_df, unmatched_flow_df, flow_valid


def generate_report_v3(match_result, unmatched_bill, unmatched_flow_details, flow_valid, output_path):
    """
    生成V3比对报告
    """
    # 统计信息
    total_groups = len(match_result)
    matched = len(match_result[match_result['匹配状态'] == '✓ 已匹配'])
    amount_mismatch_settlement_more = len(match_result[match_result['匹配状态'] == '✗ 金额不符（结算多）'])
    amount_mismatch_flow_more = len(match_result[match_result['匹配状态'] == '✗ 金额不符（流水多）'])
    unmatched = len(match_result[match_result['匹配状态'] == '✗ 无法匹配'])
    
    # 未提取到门店的流水
    no_store_extracted = len(flow_valid[flow_valid['_extracted_store'].isna()])
    
    print("\n" + "="*60)
    print("比对结果统计")
    print("="*60)
    print(f"流水总笔数: {len(flow_valid)}")
    print(f"分组数(门店+日期): {total_groups}")
    print(f"  ✓ 已匹配: {matched}")
    print(f"  ✗ 金额不符(结算多): {amount_mismatch_settlement_more}")
    print(f"  ✗ 金额不符(流水多): {amount_mismatch_flow_more}")
    print(f"  ✗ 无法匹配: {unmatched}")
    print(f"结算单未匹配组数: {len(unmatched_bill)}")
    if no_store_extracted > 0:
        print(f"⚠ 未提取到门店的流水: {no_store_extracted} 笔")
    
    # 计算金额汇总
    if len(match_result) > 0:
        matched_rows = match_result[match_result['匹配状态'] == '✓ 已匹配']
        if len(matched_rows) > 0:
            total_bill_matched = matched_rows['结算净额'].sum()
            total_flow_matched = matched_rows['到账金额汇总'].sum()
            print(f"\n已匹配结算净额: {total_bill_matched:,.2f}")
            print(f"已匹配到账金额: {total_flow_matched:,.2f}")
    print("="*60)
    
    # 导出到Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: 全部比对结果
        display_cols = ['门店', '日期', '结算日期', '流水笔数', '到账金额汇总', '流水手续费汇总', 
                       '交易金额汇总', '结算手续费汇总', '结算净额', '差额', '匹配状态', '备注']
        match_result_display = match_result[display_cols] if len(match_result) > 0 else pd.DataFrame(columns=display_cols)
        match_result_display.to_excel(writer, sheet_name='比对结果', index=False)
        
        # Sheet 2: 已匹配
        matched_rows = match_result[match_result['匹配状态'] == '✓ 已匹配']
        if len(matched_rows) > 0:
            matched_rows[display_cols].to_excel(writer, sheet_name='已匹配', index=False)
        
        # Sheet 3: 金额不符（结算多）- 可能有漏掉的流水
        amount_mismatch1 = match_result[match_result['匹配状态'] == '✗ 金额不符（结算多）']
        if len(amount_mismatch1) > 0:
            amount_mismatch1[display_cols].to_excel(writer, sheet_name='金额不符-结算多', index=False)
        
        # Sheet 4: 金额不符（流水多）- 可能有多余流水
        amount_mismatch2 = match_result[match_result['匹配状态'] == '✗ 金额不符（流水多）']
        if len(amount_mismatch2) > 0:
            amount_mismatch2[display_cols].to_excel(writer, sheet_name='金额不符-流水多', index=False)
        
        # Sheet 5: 无法匹配流水组
        unmatched_flow = match_result[match_result['匹配状态'] == '✗ 无法匹配']
        if len(unmatched_flow) > 0:
            unmatched_flow[display_cols].to_excel(writer, sheet_name='无法匹配流水', index=False)
        
        # Sheet 6: 未匹配结算单
        if len(unmatched_bill) > 0:
            unmatched_bill.to_excel(writer, sheet_name='未匹配结算单', index=False)
        
        # Sheet 7: 金额不符流水明细（用于分析差额）
        if len(unmatched_flow_details) > 0:
            unmatched_flow_details.to_excel(writer, sheet_name='金额不符流水明细', index=False)
        
        # Sheet 8: 所有流水明细（用于反查）
        flow_detail_cols = ['_extracted_store', '_extracted_date', '_amount', '_fee', '_summary']
        if all(col in flow_valid.columns for col in flow_detail_cols):
            flow_detail = flow_valid[flow_detail_cols].copy()
            flow_detail.columns = ['提取门店', '提取日期', '到账金额', '手续费', '摘要']
            flow_detail.to_excel(writer, sheet_name='所有流水明细', index=False)
        
        # Sheet 9: 汇总统计
        summary_data = {
            '项目': ['流水总笔数', '分组数', '已匹配', '金额不符(结算多)', '金额不符(流水多)', 
                    '无法匹配', '结算单未匹配组数'],
            '数量': [len(flow_valid), total_groups, matched, amount_mismatch_settlement_more,
                    amount_mismatch_flow_more, unmatched, len(unmatched_bill)],
        }
        summary = pd.DataFrame(summary_data)
        summary.to_excel(writer, sheet_name='汇总', index=False)
    
    print(f"\n✓ 报告已保存: {output_path}")


def main():
    parser = argparse.ArgumentParser(description='银联数据比对工具 V3 - 适配实际数据结构')
    parser.add_argument('bill_file', help='银联结算账单文件路径 (Excel)')
    parser.add_argument('flow_file', help='银行流水文件路径 (Excel)')
    parser.add_argument('-o', '--output', default='unionpay_reconcile_report_v3.xlsx', 
                        help='输出报告路径')
    parser.add_argument('--amount-tol', type=float, default=0.01, help='金额容差 (默认0.01)')
    
    args = parser.parse_args()
    
    print("="*60)
    print("银联数据比对工具 V3")
    print("匹配逻辑: 从摘要提取门店→分组汇总→金额比对")
    print("="*60)
    
    # 加载数据
    df_bill, df_flow = load_and_process_data(args.bill_file, args.flow_file)
    
    # 执行比对
    match_result, unmatched_bill, unmatched_flow_details, flow_valid = reconcile_data_v3(
        df_bill, df_flow,
        amount_tolerance=args.amount_tol
    )
    
    # 生成报告
    generate_report_v3(match_result, unmatched_bill, unmatched_flow_details, flow_valid, args.output)


if __name__ == '__main__':
    main()
