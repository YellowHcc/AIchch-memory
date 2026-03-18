#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银联数据比对工具 - V5 (增加终端号反查)
功能：
1. 从银行流水摘要中提取门店名称和日期
2. 按门店+日期分组汇总后比对
3. 金额反查：对差额，检查是否有无门店名的流水正好匹配差额
4. 终端号反查：对于有多个终端号的门店，按终端号分组匹配无门店流水
   - 支持按终端号分组汇总匹配（多笔流水对应一个终端号）
   - 优先处理多终端号门店
"""

import pandas as pd
import numpy as np
from datetime import datetime
from difflib import SequenceMatcher
import argparse
import sys
import re
from itertools import combinations


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
    
    summary = str(summary).strip()
    
    # 模式1: "银联入账：XXX店/品牌X店MMdd-MMdd" - 有明确门店名
    match = re.search(r'银联入账：([^\d（\(]+?)(?:品牌|店).*?(?:\d{4}-\d{4}|\d{4})', summary)
    if match:
        store = match.group(1).strip()
        # 清理常见后缀
        store = re.sub(r'(品牌[一二三]店|品牌\d+店|一店|二店|三店)$', '', store)
        return store if store else None
    
    # 模式2: "银联入账：XXX费X元" - 检查提取的是否为纯数字（如0314-0314）
    match = re.search(r'银联入账：([^费]+?)(?:费\d+|=)', summary)
    if match:
        store = match.group(1).strip()
        # 排除纯数字或日期格式（如0314-0314）
        if re.match(r'^\d{4}-\d{4}$', store) or re.match(r'^\d+$', store):
            return None
        return store if store else None
    
    # 模式3: 其他，同样需要排除纯数字
    match = re.search(r'银联入账：(.+?)(?:（|\(|费|$)', summary)
    if match:
        store = match.group(1).strip()
        # 排除纯数字或日期格式
        if re.match(r'^\d{4}-\d{4}$', store) or re.match(r'^\d+$', store):
            return None
        return store if store else None
    
    return None


def extract_date_range_from_summary(summary):
    """从摘要中提取日期范围，返回(开始日期, 结束日期)"""
    if pd.isna(summary):
        return None, None
    
    summary = str(summary).strip()
    
    # 模式: "MMdd-MMdd" 表示日期范围
    match = re.search(r'(\d{4})-(\d{4})', summary)
    if match:
        start_str = match.group(1)
        end_str = match.group(2)
        try:
            start_month = int(start_str[:2])
            start_day = int(start_str[2:])
            end_month = int(end_str[:2])
            end_day = int(end_str[2:])
            year = datetime.now().year
            start_date = f"{year}-{start_month:02d}-{start_day:02d}"
            end_date = f"{year}-{end_month:02d}-{end_day:02d}"
            return start_date, end_date
        except:
            pass
    
    # 单个日期格式
    single_date = extract_date_from_summary(summary)
    return single_date, single_date


def extract_date_from_summary(summary):
    """从摘要中提取单个日期（兼容旧代码）"""
    if pd.isna(summary):
        return None
    
    summary = str(summary).strip()
    
    # 模式: "MMdd-MMdd" 或 "MMdd"
    patterns = [
        r'(\d{4})-(\d{4})',  # 0314-0314
        r'(=\d{4})',  # =0314
        r'[^=](=\d{4})',  # =0314 after non-=
    ]
    
    for pattern in patterns:
        match = re.search(pattern, summary)
        if match:
            date_str = match.group(1)
            try:
                if len(date_str) == 5 and date_str[0] == '=':  # =0314
                    date_str = date_str[1:]
                if len(date_str) == 4:
                    month = int(date_str[:2])
                    day = int(date_str[2:])
                    year = datetime.now().year
                    return f"{year}-{month:02d}-{day:02d}"
            except:
                pass
    
    return None


def find_matching_combination(target_amount, flows, tolerance=0.01):
    """
    从流水列表中找到金额组合等于目标金额的组合
    返回匹配的组合（列表）和剩余未匹配的流水索引
    """
    flow_indices = list(range(len(flows)))
    
    # 首先尝试单条匹配
    for i in flow_indices:
        if abs(flows[i] - target_amount) <= tolerance:
            return [i], [j for j in flow_indices if j != i]
    
    # 尝试两条组合
    for combo in combinations(flow_indices, 2):
        combo_sum = sum(flows[i] for i in combo)
        if abs(combo_sum - target_amount) <= tolerance:
            return list(combo), [j for j in flow_indices if j not in combo]
    
    # 尝试三条组合
    for combo in combinations(flow_indices, 3):
        combo_sum = sum(flows[i] for i in combo)
        if abs(combo_sum - target_amount) <= tolerance:
            return list(combo), [j for j in flow_indices if j not in combo]
    
    return None, flow_indices


def load_and_process_data(bill_file, flow_file):
    """加载并预处理数据"""
    print(f"\n读取结算账单: {bill_file}")
    df_bill = pd.read_excel(bill_file)
    print(f"✓ 结算账单: {len(df_bill)} 行")
    
    print(f"\n读取银行流水: {flow_file}")
    df_flow = pd.read_excel(flow_file)
    print(f"✓ 银行流水: {len(df_flow)} 行")
    
    return df_bill, df_flow


def reconcile_data_v5(df_bill, df_flow, amount_tolerance=0.01):
    """V5比对逻辑：门店+日期分组 → 三层金额反查 → 终端号反查"""
    
    # 1. 预处理结算单
    bill_amount_col = None
    if '清分金额' in df_bill.columns:
        bill_amount_col = '清分金额'
    elif '清算金额' in df_bill.columns:
        bill_amount_col = '清算金额'
    else:
        for col in df_bill.columns:
            if '金额' in col or '清算' in col:
                bill_amount_col = col
                break
    
    if bill_amount_col is None:
        raise ValueError("结算单中未找到金额字段")
    
    store_col = None
    if 'POS门店名称' in df_bill.columns:
        store_col = 'POS门店名称'
    elif '门店' in df_bill.columns:
        store_col = '门店'
    else:
        for col in df_bill.columns:
            if '门店' in col or '商户' in col:
                store_col = col
                break
    
    if store_col is None:
        raise ValueError("结算单中未找到门店字段")
    
    date_col = None
    for col in ['结算日期', '日期', '清算日期', '交易日期']:
        if col in df_bill.columns:
            date_col = col
            break
    
    if date_col is None:
        raise ValueError("结算单中未找到日期字段")
    
    fee_col = None
    for col in ['手续费', '清算手续费']:
        if col in df_bill.columns:
            fee_col = col
            break
    
    terminal_col = None
    for col in ['终端号', '终端编号', '设备号']:
        if col in df_bill.columns:
            terminal_col = col
            break
    
    if terminal_col:
        print(f"✓ 发现终端号字段: {terminal_col}")
    
    # 标准化结算单数据
    df_bill['_store'] = df_bill[store_col]
    df_bill['_date'] = df_bill[date_col].apply(standardize_date)
    df_bill['_net_amount'] = pd.to_numeric(df_bill[bill_amount_col], errors='coerce').fillna(0)
    df_bill['_fee'] = pd.to_numeric(df_bill[fee_col], errors='coerce').fillna(0) if fee_col else 0
    df_bill['_trans_amount'] = df_bill['_net_amount'] + df_bill['_fee']
    
    if terminal_col:
        df_bill['_terminal'] = df_bill[terminal_col]
    else:
        df_bill['_terminal'] = None
    
    # 2. 预处理流水
    flow_amount_col = None
    for col in ['到账金额', '交易金额', '金额', '入账金额']:
        if col in df_flow.columns:
            flow_amount_col = col
            break
    
    if flow_amount_col is None:
        raise ValueError("流水中未找到金额字段")
    
    flow_fee_col = None
    for col in ['手续费', '费用']:
        if col in df_flow.columns:
            flow_fee_col = col
            break
    
    flow_summary_col = None
    for col in ['摘要', '备注', '说明', '交易备注']:
        if col in df_flow.columns:
            flow_summary_col = col
            break
    
    df_flow['_amount'] = pd.to_numeric(df_flow[flow_amount_col], errors='coerce').fillna(0)
    df_flow['_fee'] = pd.to_numeric(df_flow[flow_fee_col], errors='coerce').fillna(0) if flow_fee_col else 0
    df_flow['_summary'] = df_flow[flow_summary_col] if flow_summary_col else ''
    
    # 提取门店和日期
    df_flow['_extracted_store'] = df_flow['_summary'].apply(extract_store_from_summary)
    df_flow['_extracted_date'] = df_flow['_summary'].apply(extract_date_from_summary)
    
    # 新增：提取日期范围（用于跨天匹配）
    df_flow[['_date_start', '_date_end']] = df_flow['_summary'].apply(
        lambda x: pd.Series(extract_date_range_from_summary(x))
    )
    
    # 标记无门店名的流水
    df_flow['_no_store'] = df_flow['_extracted_store'].isna()
    
    # 获取无门店流水的列表
    no_store_flows = df_flow[df_flow['_no_store'] == True].copy()
    
    # 3. 建立门店映射
    bill_stores = df_bill['_store'].unique()
    store_mapping = {}
    
    for flow_store in df_flow['_extracted_store'].dropna().unique():
        best_match, score = find_best_match(flow_store, bill_stores, threshold=0.5)
        if best_match:
            store_mapping[flow_store] = best_match
    
    print("\n" + "="*60)
    print("【1. 门店映射】")
    for k, v in store_mapping.items():
        print(f"  {k} -> {v}")
    
    # 应用映射
    df_flow['_mapped_store'] = df_flow['_extracted_store'].map(store_mapping)
    
    # 4. 按门店+日期分组
    flow_grouped = df_flow[df_flow['_no_store'] == False].groupby(['_mapped_store', '_extracted_date']).agg({
        '_amount': 'sum',
        '_fee': 'sum',
        '_summary': lambda x: ' | '.join(x.tolist())
    }).reset_index()
    flow_grouped.columns = ['门店', '日期', '到账金额汇总', '手续费汇总', '流水摘要']
    
    bill_grouped = df_bill.groupby(['_store', '_date']).agg({
        '_net_amount': 'sum',
        '_trans_amount': 'sum',
        '_fee': 'sum'
    }).reset_index()
    bill_grouped.columns = ['门店', '日期', '净额汇总', '交易金额汇总', '手续费汇总']
    
    # 5. V5新增：优先执行终端号反查（处理多终端号门店）
    terminal_matches = []
    used_no_store_indices = set()
    
    if terminal_col:
        print("\n" + "="*60)
        print("【2. 终端号反查】")
        
        # 找出有多个终端号的门店
        store_terminal_counts = df_bill.groupby('_store')['_terminal'].nunique()
        multi_terminal_stores = store_terminal_counts[store_terminal_counts > 1].index.tolist()
        
        print(f"发现 {len(multi_terminal_stores)} 个多终端号门店")
        
        # 获取按终端号分组的数据
        terminal_grouped = df_bill.groupby(['_store', '_date', '_terminal']).agg({
            '_net_amount': 'sum',
            '_trans_amount': 'sum',
            '_fee': 'sum'
        }).reset_index()
        
        # 对每个日期处理
        for date in no_store_flows['_extracted_date'].unique():
            if pd.isna(date):
                continue
                
            # 获取该日期的无门店流水
            date_no_store = no_store_flows[no_store_flows['_extracted_date'] == date].copy()
            
            # 获取该日期按终端号分组的结算单
            date_terminal = terminal_grouped[terminal_grouped['_date'] == date]
            
            # 优先处理多终端号门店
            for store in multi_terminal_stores:
                store_terminals = date_terminal[date_terminal['_store'] == store]
                
                for _, term_row in store_terminals.iterrows():
                    term_net = term_row['_net_amount']
                    terminal_id = term_row['_terminal']
                    
                    # 从剩余未匹配的无门店流水中查找组合
                    remaining_flows = date_no_store[~date_no_store.index.isin(used_no_store_indices)]
                    
                    if len(remaining_flows) == 0:
                        continue
                    
                    flow_amounts = remaining_flows['_amount'].tolist()
                    flow_indices_local = list(range(len(flow_amounts)))
                    
                    # 查找匹配的组合
                    matched_combo, _ = find_matching_combination(term_net, flow_amounts, amount_tolerance)
                    
                    if matched_combo:
                        # 获取实际的数据库索引
                        matched_db_indices = remaining_flows.iloc[list(matched_combo)].index.tolist()
                        matched_amount = sum(remaining_flows.loc[idx, '_amount'] for idx in matched_db_indices)
                        matched_summaries = ' | '.join(remaining_flows.loc[idx, '_summary'][:30] + '...' for idx in matched_db_indices)
                        
                        # 标记为已使用
                        for idx in matched_db_indices:
                            used_no_store_indices.add(idx)
                        
                        terminal_matches.append({
                            '门店': store,
                            '日期': format_date_display(date),
                            '流水笔数': len(matched_combo),
                            '到账金额汇总': matched_amount,
                            '交易金额汇总': term_row['_trans_amount'],
                            '结算净额': term_net,
                            '差额': matched_amount - term_net,
                            '匹配状态': '✓ 已匹配(终端号反查)',
                            '备注': f'终端号 {terminal_id}',
                            '反查匹配': matched_summaries,
                        })
                        
                        print(f"  ✓ {store} 终端:{terminal_id} = {term_net:.2f} ({len(matched_combo)}笔流水)")
            
            # 再处理单终端号门店（可能有多笔流水汇总匹配）
            single_terminal_stores = [s for s in bill_stores if s not in multi_terminal_stores]
            
            for store in single_terminal_stores:
                store_terminals = date_terminal[date_terminal['_store'] == store]
                
                if len(store_terminals) == 0:
                    continue
                
                # 单终端号门店可能有多笔流水汇总匹配
                for _, term_row in store_terminals.iterrows():
                    term_net = term_row['_net_amount']
                    terminal_id = term_row['_terminal']
                    
                    remaining_flows = date_no_store[~date_no_store.index.isin(used_no_store_indices)]
                    
                    if len(remaining_flows) == 0:
                        continue
                    
                    flow_amounts = remaining_flows['_amount'].tolist()
                    
                    matched_combo, _ = find_matching_combination(term_net, flow_amounts, amount_tolerance)
                    
                    if matched_combo:
                        matched_db_indices = remaining_flows.iloc[list(matched_combo)].index.tolist()
                        matched_amount = sum(remaining_flows.loc[idx, '_amount'] for idx in matched_db_indices)
                        matched_summaries = ' | '.join(remaining_flows.loc[idx, '_summary'][:30] + '...' for idx in matched_db_indices)
                        
                        for idx in matched_db_indices:
                            used_no_store_indices.add(idx)
                        
                        terminal_matches.append({
                            '门店': store,
                            '日期': format_date_display(date),
                            '流水笔数': len(matched_combo),
                            '到账金额汇总': matched_amount,
                            '交易金额汇总': term_row['_trans_amount'],
                            '结算净额': term_net,
                            '差额': matched_amount - term_net,
                            '匹配状态': '✓ 已匹配(终端号反查)',
                            '备注': f'终端号 {terminal_id}' if terminal_id else '',
                            '反查匹配': matched_summaries,
                        })
                        
                        print(f"  ✓ {store} = {term_net:.2f} ({len(matched_combo)}笔流水)")
    
    # 6. 常规比对（排除已匹配的终端号结算单），支持日期区间
    results = []
    matched_bill_groups = set()
    
    # 标记已匹配的结算单
    for match in terminal_matches:
        store = match['门店']
        date = standardize_date(match['日期'])
        matched_bill_groups.add((store, date))
    
    for idx, flow_row in flow_grouped.iterrows():
        flow_store = flow_row['门店']
        flow_date = flow_row['日期']
        flow_net = flow_row['到账金额汇总']
        
        # 获取该分组下的原始流水，检查是否有日期区间
        flow_group_detail = df_flow[
            (df_flow['_mapped_store'] == flow_store) & 
            (df_flow['_extracted_date'] == flow_date)
        ]
        
        # 检查是否是日期区间（跨天）
        date_start = flow_group_detail['_date_start'].iloc[0] if len(flow_group_detail) > 0 else flow_date
        date_end = flow_group_detail['_date_end'].iloc[0] if len(flow_group_detail) > 0 else flow_date
        
        is_date_range = date_start != date_end
        
        # 查找匹配 - 支持单日期或日期区间
        if is_date_range:
            # 日期区间：汇总结算单中该日期范围的金额
            match = bill_grouped[
                (bill_grouped['门店'] == flow_store) & 
                (bill_grouped['日期'] >= date_start) &
                (bill_grouped['日期'] <= date_end)
            ]
            if len(match) > 0:
                # 汇总多天的金额
                bill_net = match['净额汇总'].sum()
                bill_trans = match['交易金额汇总'].sum()
                bill_dates = f"{format_date_display(date_start)}-{format_date_display(date_end)}"
                matched_bill_keys = [(flow_store, d) for d in match['日期'].unique()]
            else:
                bill_net = 0
                bill_trans = 0
                bill_dates = format_date_display(flow_date)
                matched_bill_keys = []
        else:
            # 单日：原有逻辑
            match = bill_grouped[
                (bill_grouped['门店'] == flow_store) & 
                (bill_grouped['日期'] == flow_date)
            ]
            if len(match) > 0:
                bill_row = match.iloc[0]
                bill_net = bill_row['净额汇总']
                bill_trans = bill_row['交易金额汇总']
                bill_dates = format_date_display(flow_date)
                matched_bill_keys = [(flow_store, flow_date)]
            else:
                bill_net = 0
                bill_trans = 0
                bill_dates = format_date_display(flow_date)
                matched_bill_keys = []
        
        if bill_net == 0:
            results.append({
                '门店': flow_store,
                '日期': bill_dates,
                '流水笔数': len(flow_group_detail),
                '到账金额汇总': flow_net,
                '交易金额汇总': 0,
                '结算净额': 0,
                '差额': flow_net,
                '匹配状态': '✗ 无法匹配',
                '备注': '未找到对应结算单',
                '反查匹配': None,
            })
        else:
            diff = flow_net - bill_net
            
            if abs(diff) <= amount_tolerance:
                results.append({
                    '门店': flow_store,
                    '日期': bill_dates,
                    '流水笔数': len(flow_group_detail),
                    '到账金额汇总': flow_net,
                    '交易金额汇总': bill_trans,
                    '结算净额': bill_net,
                    '差额': diff,
                    '匹配状态': '✓ 已匹配' + ('(日期区间)' if is_date_range else ''),
                    '备注': f'{date_start}至{date_end}汇总' if is_date_range else '',
                    '反查匹配': None,
                })
                for key in matched_bill_keys:
                    matched_bill_groups.add(key)
            else:
                results.append({
                    '门店': flow_store,
                    '日期': bill_dates,
                    '流水笔数': len(flow_group_detail),
                    '到账金额汇总': flow_net,
                    '交易金额汇总': bill_trans,
                    '结算净额': bill_net,
                    '差额': diff,
                    '匹配状态': '✗ 金额不符(待反查)',
                    '备注': f'差额 {abs(diff):.2f} 元' + (f', 区间{date_start}至{date_end}' if is_date_range else ''),
                    '反查匹配': None,
                })
                for key in matched_bill_keys:
                    matched_bill_groups.add(key)
    
    # 7. 第一层金额反查：补足差额
    additional_matches = []
    
    for idx, result in enumerate(results):
        if '金额不符' in result['匹配状态']:
            flow_store = result['门店']
            flow_date_raw = standardize_date(result['日期'])
            diff = result['差额']
            
            no_store_candidates = df_flow[
                (df_flow['_no_store'] == True) & 
                (~df_flow.index.isin(used_no_store_indices)) &
                (df_flow['_extracted_date'] == flow_date_raw)
            ]
            
            if len(no_store_candidates) > 0:
                no_store_candidates['_diff_match'] = abs(no_store_candidates['_amount'] - abs(diff))
                best_candidate = no_store_candidates.loc[no_store_candidates['_diff_match'].idxmin()]
                
                if abs(best_candidate['_amount'] - abs(diff)) <= amount_tolerance:
                    used_no_store_indices.add(best_candidate.name)
                    
                    results[idx]['匹配状态'] = '✓ 已匹配(含反查)'
                    results[idx]['备注'] = f'门店流水+无门店流水补足差额'
                    results[idx]['反查匹配'] = f"{best_candidate['_summary'][:30]}... ({best_candidate['_amount']:.2f})"
    
    # 8. 第二层纯金额反查
    unmatched_bills = []
    for idx, bill_row in bill_grouped.iterrows():
        bill_key = (bill_row['门店'], bill_row['日期'])
        if bill_key not in matched_bill_groups:
            unmatched_bills.append(bill_row)
    
    for bill_row in unmatched_bills:
        bill_store = bill_row['门店']
        bill_date = bill_row['日期']
        bill_net = bill_row['净额汇总']
        
        no_store_candidates = df_flow[
            (df_flow['_no_store'] == True) & 
            (~df_flow.index.isin(used_no_store_indices)) &
            (df_flow['_extracted_date'] == bill_date)
        ]
        
        if len(no_store_candidates) > 0:
            exact_match = no_store_candidates[abs(no_store_candidates['_amount'] - bill_net) <= amount_tolerance]
            
            if len(exact_match) > 0:
                matched_flow = exact_match.iloc[0]
                used_no_store_indices.add(matched_flow.name)
                
                additional_matches.append({
                    '门店': bill_store,
                    '日期': format_date_display(bill_date),
                    '流水笔数': 1,
                    '到账金额汇总': matched_flow['_amount'],
                    '交易金额汇总': bill_row['交易金额汇总'],
                    '结算净额': bill_net,
                    '差额': matched_flow['_amount'] - bill_net,
                    '匹配状态': '✓ 已匹配(纯金额反查)',
                    '备注': '无门店名流水直接匹配结算单',
                    '反查匹配': f"{matched_flow['_summary'][:30]}...",
                })
                matched_bill_groups.add((bill_store, bill_date))
    
    # 9. 第三层日期汇总反查
    date_grouped_matches = []
    remaining_no_store = df_flow[
        (df_flow['_no_store'] == True) & 
        (~df_flow.index.isin(used_no_store_indices))
    ].copy()
    
    if len(remaining_no_store) > 0:
        no_store_by_date = remaining_no_store.groupby('_extracted_date').agg({
            '_amount': 'sum',
            '_summary': lambda x: ' | '.join(x.tolist())
        }).reset_index()
        
        for idx, bill_row in bill_grouped.iterrows():
            bill_key = (bill_row['门店'], bill_row['日期'])
            if bill_key not in matched_bill_groups:
                bill_net = bill_row['净额汇总']
                bill_date = bill_row['日期']
                
                date_match = no_store_by_date[
                    (no_store_by_date['_extracted_date'] == bill_date) &
                    (abs(no_store_by_date['_amount'] - bill_net) <= amount_tolerance)
                ]
                
                if len(date_match) > 0:
                    dm = date_match.iloc[0]
                    used_in_match = remaining_no_store[
                        (remaining_no_store['_extracted_date'] == bill_date)
                    ]
                    for idx2 in used_in_match.index:
                        used_no_store_indices.add(idx2)
                    
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
    all_results = []
    if results:
        all_results.extend(results)
    if terminal_matches:
        all_results.extend(terminal_matches)
    if additional_matches:
        all_results.extend(additional_matches)
    if date_grouped_matches:
        all_results.extend(date_grouped_matches)
    
    results_df = pd.DataFrame(all_results) if all_results else pd.DataFrame(columns=[
        '门店', '日期', '流水笔数', '到账金额汇总', '交易金额汇总', '结算净额', '差额', '匹配状态', '备注', '反查匹配'
    ])
    
    # 10. 收集未匹配的无门店流水
    unmatched_no_store = df_flow[
        (df_flow['_no_store'] == True) & 
        (~df_flow.index.isin(used_no_store_indices))
    ][['_amount', '_fee', '_summary', '_extracted_date']].copy()
    unmatched_no_store.columns = ['到账金额', '手续费', '摘要', '提取日期']
    
    # 11. 找出结算单中未被匹配的记录
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
    
    return results_df, unmatched_bill_df, unmatched_no_store, df_flow


def generate_report_v5(match_result, unmatched_bill, unmatched_no_store, df_flow, output_path):
    """生成V5比对报告"""
    total_groups = len(match_result)
    matched = len(match_result[match_result['匹配状态'] == '✓ 已匹配'])
    matched_with_lookup = len(match_result[match_result['匹配状态'] == '✓ 已匹配(含反查)'])
    matched_pure_lookup = len(match_result[match_result['匹配状态'] == '✓ 已匹配(纯金额反查)'])
    matched_date_grouped = len(match_result[match_result['匹配状态'] == '✓ 已匹配(日期汇总反查)'])
    matched_terminal = len(match_result[match_result['匹配状态'] == '✓ 已匹配(终端号反查)'])
    amount_mismatch = len(match_result[match_result['匹配状态'].str.contains('金额不符', na=False)])
    unmatched = len(match_result[match_result['匹配状态'] == '✗ 无法匹配'])
    
    print("\n" + "="*60)
    print("比对结果统计")
    print("="*60)
    print(f"流水总笔数: {len(df_flow)}")
    print(f"分组数: {total_groups}")
    print(f"  ✓ 已匹配: {matched}")
    print(f"  ✓ 已匹配(含反查): {matched_with_lookup}")
    print(f"  ✓ 已匹配(纯金额反查): {matched_pure_lookup}")
    print(f"  ✓ 已匹配(日期汇总反查): {matched_date_grouped}")
    print(f"  ✓ 已匹配(终端号反查): {matched_terminal}")
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
        display_cols = ['门店', '日期', '流水笔数', '到账金额汇总', '交易金额汇总', '结算净额', '差额', '匹配状态', '备注', '反查匹配']
        
        # 全部比对结果
        match_result_display = match_result[display_cols] if len(match_result) > 0 else pd.DataFrame(columns=display_cols)
        match_result_display.to_excel(writer, sheet_name='比对结果', index=False)
        
        # 已匹配
        for status, sheet_name in [
            ('✓ 已匹配', '已匹配'),
            ('✓ 已匹配(含反查)', '金额反查匹配'),
            ('✓ 已匹配(纯金额反查)', '纯金额反查匹配'),
            ('✓ 已匹配(日期汇总反查)', '日期汇总反查匹配'),
            ('✓ 已匹配(终端号反查)', '终端号反查匹配'),
        ]:
            rows = match_result[match_result['匹配状态'] == status]
            if len(rows) > 0:
                rows[display_cols].to_excel(writer, sheet_name=sheet_name, index=False)
        
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
            '项目': ['流水总笔数', '分组数', '已匹配', '已匹配(含反查)', '已匹配(纯金额反查)', 
                    '已匹配(日期汇总反查)', '已匹配(终端号反查)', '金额不符', '无法匹配', 
                    '未匹配无门店流水', '结算单未匹配组数'],
            '数量': [len(df_flow), total_groups, matched, matched_with_lookup, matched_pure_lookup, 
                    matched_date_grouped, matched_terminal, amount_mismatch, unmatched, 
                    len(unmatched_no_store), len(unmatched_bill)],
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='汇总', index=False)
    
    print(f"\n✓ 报告已保存: {output_path}")


def main():
    parser = argparse.ArgumentParser(description='银联数据比对工具 V5 - 增加终端号反查')
    parser.add_argument('bill_file', help='银联结算账单文件路径')
    parser.add_argument('flow_file', help='银行流水文件路径')
    parser.add_argument('-o', '--output', default='unionpay_reconcile_report_v5.xlsx', help='输出报告路径')
    parser.add_argument('--amount-tol', type=float, default=0.01, help='金额容差')
    
    args = parser.parse_args()
    
    print("="*60)
    print("银联数据比对工具 V5")
    print("匹配逻辑: 终端号反查 → 门店匹配 → 金额反查 → 日期汇总反查")
    print("="*60)
    
    df_bill, df_flow = load_and_process_data(args.bill_file, args.flow_file)
    
    match_result, unmatched_bill, unmatched_no_store, df_flow = reconcile_data_v5(
        df_bill, df_flow, amount_tolerance=args.amount_tol
    )
    
    generate_report_v5(match_result, unmatched_bill, unmatched_no_store, df_flow, args.output)


if __name__ == '__main__':
    main()
