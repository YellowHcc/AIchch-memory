#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银联数据比对工具 - V4 (增加金额反查)
功能：
1. 从银行流水摘要中提取门店名称和日期
2. 按门店+日期分组汇总后比对
3. 金额反查：对差额，检查是否有无门店名的流水正好匹配差额
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


def find_best_match(store_name, store_list, threshold=0.5):
    """在门店列表中找到最佳匹配"""
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
    
    if isinstance(date_val, datetime):
        return date_val.strftime('%Y-%m-%d')
    
    date_str = str(date_val).strip()
    
    if date_str.isdigit():
        if len(date_str) == 4:
            try:
                month = int(date_str[:2])
                day = int(date_str[2:])
                year = datetime.now().year
                dt = datetime(year, month, day)
                return dt.strftime('%Y-%m-%d')
            except:
                pass
        elif len(date_str) == 8:
            try:
                dt = datetime.strptime(date_str, '%Y%m%d')
                return dt.strftime('%Y-%m-%d')
            except:
                pass
    
    formats = ['%Y-%m-%d', '%Y/%m/%d', '%Y%m%d', '%d/%m/%Y', '%m/%d/%Y', '%m-%d', '%m/%d']
    
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
    """将 YYYY-MM-DD 格式化为 MMDD"""
    if not date_str or date_str == 'N/A':
        return 'N/A'
    try:
        dt = datetime.strptime(str(date_str), '%Y-%m-%d')
        return dt.strftime('%m%d')
    except:
        return date_str


def extract_store_from_summary(summary):
    """从摘要中提取门店名称"""
    if pd.isna(summary):
        return None
    
    summary_str = str(summary).strip()
    
    patterns = [
        r'银联入账：(.+?)(?:\d{4}|\d{2}|费|\+|-)',
        r'银联入账：(.+?)(?:费|\+|-)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, summary_str)
        if match:
            store = match.group(1).strip().rstrip('+')
            if store and not store.isdigit():
                return store
    
    return None


def extract_date_from_summary(summary):
    """从摘要中提取日期"""
    if pd.isna(summary):
        return None
    
    summary_str = str(summary).strip()
    
    match = re.search(r'(\d{4})-', summary_str)
    if match:
        date_mmdd = match.group(1)
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


def extract_fee_from_summary(summary):
    """从摘要中提取手续费"""
    if pd.isna(summary):
        return 0
    match = re.search(r'费([\d.]+)元?', str(summary))
    if match:
        try:
            return float(match.group(1))
        except:
            pass
    return 0


def load_and_process_data(bill_path, flow_path):
    """加载并预处理数据"""
    try:
        df_bill = pd.read_excel(bill_path)
        df_flow = pd.read_excel(flow_path)
    except Exception as e:
        print(f"读取文件失败: {e}")
        sys.exit(1)
    
    print(f"✓ 结算账单: {len(df_bill)} 行")
    print(f"✓ 银行流水: {len(df_flow)} 行")
    
    df_bill.columns = [c.strip() for c in df_bill.columns]
    df_flow.columns = [c.strip() for c in df_flow.columns]
    
    # 结算账单
    df_bill['_store'] = df_bill['POS门店名称'].astype(str).str.strip()
    df_bill['_date'] = df_bill['结算日期'].apply(standardize_date)
    
    # 使用清分金额作为比对金额（扣除手续费和优惠后的实际到账金额）
    if '清分金额' in df_bill.columns:
        df_bill['_net_amount'] = pd.to_numeric(df_bill['清分金额'], errors='coerce')
    elif '清算金额' in df_bill.columns:
        df_bill['_net_amount'] = pd.to_numeric(df_bill['清算金额'], errors='coerce')
    else:
        # 兼容旧数据：交易金额 - 手续费
        df_bill['_net_amount'] = pd.to_numeric(df_bill['交易金额'], errors='coerce') - pd.to_numeric(df_bill['手续费'], errors='coerce').fillna(0)
    
    # 保留原始金额字段用于显示
    df_bill['_amount'] = pd.to_numeric(df_bill['交易金额'], errors='coerce')
    df_bill['_fee'] = pd.to_numeric(df_bill['手续费'], errors='coerce').fillna(0)
    
    # 银行流水
    df_flow['_extracted_store'] = df_flow['摘要'].apply(extract_store_from_summary)
    df_flow['_extracted_date'] = df_flow['摘要'].apply(extract_date_from_summary)
    df_flow['_amount'] = pd.to_numeric(df_flow['交易金额'], errors='coerce')
    df_flow['_fee'] = df_flow['摘要'].apply(extract_fee_from_summary)
    df_flow['_summary'] = df_flow['摘要']
    
    return df_bill, df_flow


def reconcile_data_v4(df_bill, df_flow, amount_tolerance=0.01):
    """
    V4 比对逻辑：
    1. 先按门店+日期匹配
    2. 对金额不符的，用无门店名的流水进行金额反查
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
    
    # 2. 给流水添加匹配后的门店和日期
    df_flow['_matched_store'] = df_flow['_extracted_store'].map(store_mapping)
    df_flow['_matched_date'] = df_flow['_extracted_date']
    
    # 标记无门店名的流水（用于后续金额反查）
    df_flow['_no_store'] = df_flow['_extracted_store'].isna()
    
    print("="*60)
    
    # 3. 按（匹配后门店+日期）分组汇总流水
    flow_with_store = df_flow[df_flow['_matched_store'].notna() & df_flow['_matched_date'].notna()].copy()
    
    flow_grouped = flow_with_store.groupby(['_matched_store', '_matched_date']).agg({
        '_amount': 'sum',
        '_fee': 'sum',
        '_extracted_store': lambda x: ', '.join(x.unique()),
    }).reset_index()
    flow_grouped.columns = ['门店', '日期', '到账金额汇总', '手续费汇总', '原始门店名']
    
    # 4. 按（门店+日期）分组汇总结算单
    bill_grouped = df_bill.groupby(['_store', '_date']).agg({
        '_amount': 'sum',
        '_fee': 'sum',
        '_net_amount': 'sum',
    }).reset_index()
    bill_grouped.columns = ['门店', '日期', '交易金额汇总', '手续费汇总', '净额汇总']
    
    # 5. 逐组比对
    matched_bill_groups = set()
    used_no_store_flows = set()  # 记录已被使用的无门店流水
    
    for idx, flow_group in flow_grouped.iterrows():
        store = flow_group['门店']
        date = flow_group['日期']
        flow_amount = flow_group['到账金额汇总']
        flow_fee = flow_group['手续费汇总']
        
        bill_candidates = bill_grouped[
            (bill_grouped['门店'] == store) & 
            (bill_grouped['日期'] == date)
        ]
        
        if len(bill_candidates) == 0:
            results.append({
                '门店': store,
                '日期': format_date_display(date),
                '流水笔数': len(flow_with_store[(flow_with_store['_matched_store'] == store) & (flow_with_store['_matched_date'] == date)]),
                '到账金额汇总': flow_amount,
                '交易金额汇总': 'N/A',
                '结算净额': 'N/A',
                '差额': 'N/A',
                '匹配状态': '✗ 无法匹配',
                '备注': '无对应结算单',
                '反查匹配': '',
            })
        else:
            bill_row = bill_candidates.iloc[0]
            bill_amount = bill_row['交易金额汇总']
            bill_fee = bill_row['手续费汇总']
            bill_net = bill_row['净额汇总']
            bill_key = (store, date)
            
            diff = flow_amount - bill_net
            
            # 检查是否完全匹配
            if abs(diff) <= amount_tolerance:
                matched_bill_groups.add(bill_key)
                results.append({
                    '门店': store,
                    '日期': format_date_display(date),
                    '流水笔数': len(flow_with_store[(flow_with_store['_matched_store'] == store) & (flow_with_store['_matched_date'] == date)]),
                    '到账金额汇总': flow_amount,
                    '交易金额汇总': bill_amount,
                    '结算净额': bill_net,
                    '差额': diff,
                    '匹配状态': '✓ 已匹配',
                    '备注': '',
                    '反查匹配': '',
                })
            else:
                # 金额不匹配 - 尝试金额反查
                # 查找无门店名的流水，金额是否等于差额
                matched_no_store = None
                
                if flow_amount < bill_net:
                    # 到账金额 < 结算净额，找无门店流水补差额
                    needed = bill_net - flow_amount
                    no_store_candidates = df_flow[
                        (df_flow['_no_store'] == True) & 
                        (~df_flow.index.isin(used_no_store_flows)) &
                        (abs(df_flow['_amount'] - needed) <= amount_tolerance)
                    ]
                    
                    if len(no_store_candidates) > 0:
                        matched_no_store = no_store_candidates.iloc[0]
                        used_no_store_flows.add(matched_no_store.name)
                        
                        # 重新计算
                        new_flow_amount = flow_amount + matched_no_store['_amount']
                        new_diff = new_flow_amount - bill_net
                        
                        if abs(new_diff) <= amount_tolerance:
                            matched_bill_groups.add(bill_key)
                            results.append({
                                '门店': store,
                                '日期': format_date_display(date),
                                '流水笔数': len(flow_with_store[(flow_with_store['_matched_store'] == store) & (flow_with_store['_matched_date'] == date)]) + 1,
                                '到账金额汇总': new_flow_amount,
                                '交易金额汇总': bill_amount,
                                '结算净额': bill_net,
                                '差额': new_diff,
                                '匹配状态': '✓ 已匹配(含反查)',
                                '备注': f'通过金额反查匹配无门店流水 {matched_no_store["_amount"]:.2f}',
                                '反查匹配': f'{matched_no_store["_summary"][:40]}...',
                            })
                            continue
                
                # 反查失败，记录为金额不符
                matched_bill_groups.add(bill_key)
                
                if flow_amount < bill_net:
                    status = '✗ 金额不符(结算多)'
                    note = f'结算净额比到账多 {abs(diff):.2f} 元'
                else:
                    status = '✗ 金额不符(流水多)'
                    note = f'到账比结算净额多 {abs(diff):.2f} 元'
                
                results.append({
                    '门店': store,
                    '日期': format_date_display(date),
                    '流水笔数': len(flow_with_store[(flow_with_store['_matched_store'] == store) & (flow_with_store['_matched_date'] == date)]),
                    '到账金额汇总': flow_amount,
                    '交易金额汇总': bill_amount,
                    '结算净额': bill_net,
                    '差额': diff,
                    '匹配状态': status,
                    '备注': note,
                    '反查匹配': '',
                })
    
    # 6. 第二阶段金额反查：对结算单未匹配的，用净额匹配无门店流水
    additional_matches = []
    
    for idx, bill_row in bill_grouped.iterrows():
        bill_key = (bill_row['门店'], bill_row['日期'])
        if bill_key not in matched_bill_groups:
            # 这笔结算单还没匹配到，尝试用净额匹配无门店流水
            bill_net = bill_row['净额汇总']
            
            # 找无门店流水，金额匹配
            matching_flows = df_flow[
                (df_flow['_no_store'] == True) & 
                (~df_flow.index.isin(used_no_store_flows)) &
                (abs(df_flow['_amount'] - bill_net) <= amount_tolerance)
            ]
            
            if len(matching_flows) > 0:
                mf = matching_flows.iloc[0]
                used_no_store_flows.add(mf.name)
                
                additional_matches.append({
                    '门店': bill_row['门店'],
                    '日期': format_date_display(bill_row['日期']),
                    '流水笔数': 1,
                    '到账金额汇总': mf['_amount'],
                    '交易金额汇总': bill_row['交易金额汇总'],
                    '结算净额': bill_net,
                    '差额': mf['_amount'] - bill_net,
                    '匹配状态': '✓ 已匹配(纯金额反查)',
                    '备注': f'结算单无对应门店流水，通过金额反查匹配',
                    '反查匹配': f'{mf["_summary"][:50]}...',
                })
                
                matched_bill_groups.add(bill_key)
    
    # 7. 第三阶段金额反查：按日期汇总无门店流水，匹配同日期未匹配的结算单
    date_grouped_matches = []
    
    # 按日期分组汇总剩余的无门店流水
    remaining_no_store = df_flow[
        (df_flow['_no_store'] == True) & 
        (~df_flow.index.isin(used_no_store_flows))
    ].copy()
    
    if len(remaining_no_store) > 0:
        # 按日期汇总
        no_store_by_date = remaining_no_store.groupby('_extracted_date').agg({
            '_amount': 'sum',
            '_summary': lambda x: ' | '.join(x.tolist())
        }).reset_index()
        
        # 尝试匹配同日期但未匹配的结算单
        for idx, bill_row in bill_grouped.iterrows():
            bill_key = (bill_row['门店'], bill_row['日期'])
            if bill_key not in matched_bill_groups:
                bill_net = bill_row['净额汇总']
                bill_date = bill_row['日期']
                
                # 找同日期的汇总流水
                date_match = no_store_by_date[
                    (no_store_by_date['_extracted_date'] == bill_date) &
                    (abs(no_store_by_date['_amount'] - bill_net) <= amount_tolerance)
                ]
                
                if len(date_match) > 0:
                    dm = date_match.iloc[0]
                    # 标记这些流水为已使用
                    used_in_match = remaining_no_store[
                        (remaining_no_store['_extracted_date'] == bill_date)
                    ]
                    for idx2 in used_in_match.index:
                        used_no_store_flows.add(idx2)
                    
                    date_grouped_matches.append({
                        '门店': bill_row['门店'],
                        '日期': format_date_display(bill_row['日期']),
                        '流水笔数': len(used_in_match),
                        '到账金额汇总': dm['_amount'],
                        '交易金额汇总': bill_row['交易金额汇总'],
                        '结算净额': bill_net,
                        '差额': dm['_amount'] - bill_net,
                        '匹配状态': '✓ 已匹配(日期汇总反查)',
                        '备注': f'按日期汇总{len(used_in_match)}笔无门店流水匹配',
                        '反查匹配': f'{dm["_summary"][:50]}...',
                    })
                    
                    matched_bill_groups.add(bill_key)
    
    # 合并所有结果
    if additional_matches or date_grouped_matches:
        results_df = pd.DataFrame(results)
        if additional_matches:
            results_df = pd.concat([results_df, pd.DataFrame(additional_matches)], ignore_index=True)
        if date_grouped_matches:
            results_df = pd.concat([results_df, pd.DataFrame(date_grouped_matches)], ignore_index=True)
        results = results_df
    
    # 8. 收集未匹配的无门店流水
    unmatched_no_store = df_flow[
        (df_flow['_no_store'] == True) & 
        (~df_flow.index.isin(used_no_store_flows))
    ][['_amount', '_fee', '_summary', '_extracted_date']].copy()
    unmatched_no_store.columns = ['到账金额', '手续费', '摘要', '提取日期']
    
    # 9. 找出结算单中未被匹配的记录
    unmatched_bill = []
    for idx, bill_row in bill_grouped.iterrows():
        bill_key = (bill_row['门店'], bill_row['日期'])
        if bill_key not in matched_bill_groups:
            unmatched_bill.append({
                '门店': bill_row['门店'],
                '日期': format_date_display(bill_row['日期']),
                '交易金额汇总': bill_row['交易金额汇总'],
                '手续费汇总': bill_row['手续费汇总'],
                '净额汇总': bill_row['净额汇总'],
            })
    
    unmatched_bill_df = pd.DataFrame(unmatched_bill)
    
    return pd.DataFrame(results), unmatched_bill_df, unmatched_no_store, df_flow


def generate_report_v4(match_result, unmatched_bill, unmatched_no_store, df_flow, output_path):
    """生成V4比对报告"""
    total_groups = len(match_result)
    matched = len(match_result[match_result['匹配状态'] == '✓ 已匹配'])
    matched_with_lookup = len(match_result[match_result['匹配状态'] == '✓ 已匹配(含反查)'])
    matched_pure_lookup = len(match_result[match_result['匹配状态'] == '✓ 已匹配(纯金额反查)'])
    matched_date_grouped = len(match_result[match_result['匹配状态'] == '✓ 已匹配(日期汇总反查)'])
    amount_mismatch = len(match_result[match_result['匹配状态'].str.contains('金额不符', na=False)])
    unmatched = len(match_result[match_result['匹配状态'] == '✗ 无法匹配'])
    
    print("\n" + "="*60)
    print("比对结果统计")
    print("="*60)
    print(f"流水总笔数: {len(df_flow)}")
    print(f"分组数(门店+日期): {total_groups}")
    print(f"  ✓ 已匹配: {matched}")
    print(f"  ✓ 已匹配(含金额反查): {matched_with_lookup}")
    print(f"  ✓ 已匹配(纯金额反查): {matched_pure_lookup}")
    print(f"  ✓ 已匹配(日期汇总反查): {matched_date_grouped}")
    print(f"  ✗ 金额不符: {amount_mismatch}")
    print(f"  ✗ 无法匹配: {unmatched}")
    print(f"未匹配无门店流水: {len(unmatched_no_store)} 笔")
    print(f"结算单未匹配组数: {len(unmatched_bill)}")
    
    if len(match_result) > 0:
        matched_rows = match_result[match_result['匹配状态'].str.contains('已匹配', na=False)]
        if len(matched_rows) > 0:
            total_bill_matched = matched_rows['结算净额'].sum()
            total_flow_matched = matched_rows['到账金额汇总'].sum()
            print(f"\n已匹配结算净额: {total_bill_matched:,.2f}")
            print(f"已匹配到账金额: {total_flow_matched:,.2f}")
    print("="*60)
    
    # 导出
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 全部比对结果
        display_cols = ['门店', '日期', '流水笔数', '到账金额汇总', '交易金额汇总', '结算净额', '差额', '匹配状态', '备注', '反查匹配']
        match_result_display = match_result[display_cols] if len(match_result) > 0 else pd.DataFrame(columns=display_cols)
        match_result_display.to_excel(writer, sheet_name='比对结果', index=False)
        
        # 已匹配
        matched_rows = match_result[match_result['匹配状态'].str.contains('已匹配', na=False)]
        if len(matched_rows) > 0:
            matched_rows[display_cols].to_excel(writer, sheet_name='已匹配', index=False)
        
        # 金额反查匹配
        lookup_matched = match_result[match_result['匹配状态'] == '✓ 已匹配(含反查)']
        if len(lookup_matched) > 0:
            lookup_matched[display_cols].to_excel(writer, sheet_name='金额反查匹配', index=False)
        
        # 纯金额反查匹配
        pure_lookup_matched = match_result[match_result['匹配状态'] == '✓ 已匹配(纯金额反查)']
        if len(pure_lookup_matched) > 0:
            pure_lookup_matched[display_cols].to_excel(writer, sheet_name='纯金额反查匹配', index=False)
        
        # 日期汇总反查匹配
        date_grouped_matched = match_result[match_result['匹配状态'] == '✓ 已匹配(日期汇总反查)']
        if len(date_grouped_matched) > 0:
            date_grouped_matched[display_cols].to_excel(writer, sheet_name='日期汇总反查匹配', index=False)
        
        # 金额不符
        amount_mismatch_rows = match_result[match_result['匹配状态'].str.contains('金额不符', na=False)]
        if len(amount_mismatch_rows) > 0:
            amount_mismatch_rows[display_cols].to_excel(writer, sheet_name='金额不符', index=False)
        
        # 无法匹配
        unmatched_flow = match_result[match_result['匹配状态'] == '✗ 无法匹配']
        if len(unmatched_flow) > 0:
            unmatched_flow[display_cols].to_excel(writer, sheet_name='无法匹配', index=False)
        
        # 未匹配结算单
        if len(unmatched_bill) > 0:
            unmatched_bill.to_excel(writer, sheet_name='未匹配结算单', index=False)
        
        # 未匹配无门店流水
        if len(unmatched_no_store) > 0:
            unmatched_no_store.to_excel(writer, sheet_name='未匹配无门店流水', index=False)
        
        # 所有流水明细
        flow_detail = df_flow[['_extracted_store', '_extracted_date', '_amount', '_fee', '_summary', '_no_store']].copy()
        flow_detail.columns = ['提取门店', '提取日期', '到账金额', '手续费', '摘要', '是否无门店']
        flow_detail.to_excel(writer, sheet_name='所有流水明细', index=False)
        
        # 汇总
        summary_data = {
            '项目': ['流水总笔数', '分组数', '已匹配', '已匹配(含反查)', '已匹配(纯金额反查)', '已匹配(日期汇总反查)', '金额不符', '无法匹配', '未匹配无门店流水', '结算单未匹配组数'],
            '数量': [len(df_flow), total_groups, matched, matched_with_lookup, matched_pure_lookup, matched_date_grouped,
                    amount_mismatch, unmatched, len(unmatched_no_store), len(unmatched_bill)],
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='汇总', index=False)
    
    print(f"\n✓ 报告已保存: {output_path}")


def main():
    parser = argparse.ArgumentParser(description='银联数据比对工具 V4 - 增加金额反查')
    parser.add_argument('bill_file', help='银联结算账单文件路径')
    parser.add_argument('flow_file', help='银行流水文件路径')
    parser.add_argument('-o', '--output', default='unionpay_reconcile_report_v4.xlsx', help='输出报告路径')
    parser.add_argument('--amount-tol', type=float, default=0.01, help='金额容差')
    
    args = parser.parse_args()
    
    print("="*60)
    print("银联数据比对工具 V4")
    print("匹配逻辑: 门店匹配 → 分组汇总 → 金额反查")
    print("="*60)
    
    df_bill, df_flow = load_and_process_data(args.bill_file, args.flow_file)
    
    match_result, unmatched_bill, unmatched_no_store, df_flow = reconcile_data_v4(
        df_bill, df_flow, amount_tolerance=args.amount_tol
    )
    
    generate_report_v4(match_result, unmatched_bill, unmatched_no_store, df_flow, args.output)


if __name__ == '__main__':
    main()
