#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银联数据比对工具
功能：比对银联结算账单和银行流水，确认所有入金都有结算明细
"""

import pandas as pd
import numpy as np
from datetime import datetime
from difflib import SequenceMatcher
import argparse
import sys


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
    """标准化日期格式"""
    if pd.isna(date_val):
        return None
    
    # 如果已经是datetime类型
    if isinstance(date_val, datetime):
        return date_val.strftime('%Y-%m-%d')
    
    # 尝试多种格式解析
    date_str = str(date_val).strip()
    formats = [
        '%Y-%m-%d',
        '%Y/%m/%d',
        '%Y%m%d',
        '%d/%m/%Y',
        '%m/%d/%Y',
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime('%Y-%m-%d')
        except:
            continue
    
    return date_str


def load_and_process_data(bill_path, flow_path, 
                          bill_store_col='门店', bill_date_col='日期', bill_amount_col='金额',
                          flow_store_col='门店', flow_date_col='日期', flow_amount_col='交易金额', flow_fee_col='手续费'):
    """
    加载并预处理数据
    """
    # 读取数据
    try:
        if bill_path.endswith('.csv'):
            df_bill = pd.read_csv(bill_path, encoding='utf-8')
        else:
            df_bill = pd.read_excel(bill_path)
            
        if flow_path.endswith('.csv'):
            df_flow = pd.read_csv(flow_path, encoding='utf-8')
        else:
            df_flow = pd.read_excel(flow_path)
    except Exception as e:
        print(f"读取文件失败: {e}")
        sys.exit(1)
    
    print(f"✓ 结算账单: {len(df_bill)} 行")
    print(f"✓ 银行流水: {len(df_flow)} 行")
    
    # 标准化列名（尝试自动识别）
    df_bill.columns = [c.strip() for c in df_bill.columns]
    df_flow.columns = [c.strip() for c in df_flow.columns]
    
    # 日期标准化
    df_bill['_date'] = df_bill[bill_date_col].apply(standardize_date)
    df_flow['_date'] = df_flow[flow_date_col].apply(standardize_date)
    
    # 金额转为数值
    df_bill['_amount'] = pd.to_numeric(df_bill[bill_amount_col], errors='coerce')
    df_flow['_amount'] = pd.to_numeric(df_flow[flow_amount_col], errors='coerce')
    
    if flow_fee_col in df_flow.columns:
        df_flow['_fee'] = pd.to_numeric(df_flow[flow_fee_col], errors='coerce').fillna(0)
    else:
        df_flow['_fee'] = 0
    
    # 门店名称清洗
    df_bill['_store'] = df_bill[bill_store_col].astype(str).str.strip()
    df_flow['_store'] = df_flow[flow_store_col].astype(str).str.strip()
    
    return df_bill, df_flow


def reconcile_data(df_bill, df_flow, 
                   date_tolerance=0,  # 日期容差天数
                   amount_tolerance=0.01,  # 金额容差
                   match_by_fee=True):
    """
    执行比对逻辑
    以银行流水为准，找出对应的结算明细
    """
    results = []
    matched_bill_indices = set()
    
    # 获取所有唯一门店名（用于模糊匹配）
    bill_stores = df_bill['_store'].unique().tolist()
    
    print("\n开始比对...")
    
    for idx, flow_row in df_flow.iterrows():
        flow_store = flow_row['_store']
        flow_date = flow_row['_date']
        flow_amount = flow_row['_amount']
        flow_fee = flow_row['_fee']
        flow_net = flow_amount - flow_fee
        
        # 1. 尝试门店模糊匹配
        matched_store, store_score = find_best_match(flow_store, bill_stores)
        
        # 2. 筛选候选记录
        candidates = df_bill.copy()
        
        if matched_store:
            candidates = candidates[candidates['_store'] == matched_store]
        
        # 日期匹配（允许容差）
        if date_tolerance == 0:
            candidates = candidates[candidates['_date'] == flow_date]
        
        # 金额匹配（原始金额或扣除手续费后的金额）
        if match_by_fee:
            # 尝试匹配净额（交易金额-手续费）
            amount_matches = (
                (candidates['_amount'] - flow_net).abs() <= amount_tolerance
            ) | (
                (candidates['_amount'] - flow_amount).abs() <= amount_tolerance
            )
        else:
            amount_matches = (candidates['_amount'] - flow_amount).abs() <= amount_tolerance
        
        candidates = candidates[amount_matches]
        
        # 记录结果
        if len(candidates) > 0:
            # 找到匹配
            matched_idx = candidates.index[0]
            matched_bill_indices.add(matched_idx)
            
            bill_row = candidates.iloc[0]
            status = '✓ 已匹配'
            
            results.append({
                '流水序号': idx + 1,
                '流水门店': flow_store,
                '匹配门店': bill_row['_store'],
                '门店相似度': f"{store_score:.2%}",
                '流水日期': flow_date,
                '结算日期': bill_row['_date'],
                '流水金额': flow_amount,
                '手续费': flow_fee,
                '结算金额': bill_row['_amount'],
                '差额': flow_amount - flow_fee - bill_row['_amount'],
                '匹配状态': status,
                '备注': ''
            })
        else:
            # 未找到匹配
            status = '✗ 未匹配'
            
            # 分析未匹配原因
            reason = []
            if matched_store is None:
                reason.append("门店未识别")
            
            results.append({
                '流水序号': idx + 1,
                '流水门店': flow_store,
                '匹配门店': matched_store if matched_store else 'N/A',
                '门店相似度': f"{store_score:.2%}" if matched_store else 'N/A',
                '流水日期': flow_date,
                '结算日期': 'N/A',
                '流水金额': flow_amount,
                '手续费': flow_fee,
                '结算金额': 'N/A',
                '差额': 'N/A',
                '匹配状态': status,
                '备注': '; '.join(reason) if reason else '金额/日期不匹配'
            })
    
    # 找出结算单中未被匹配的记录
    unmatched_bill = df_bill[~df_bill.index.isin(matched_bill_indices)]
    
    return pd.DataFrame(results), unmatched_bill


def generate_report(match_result, unmatched_bill, output_path):
    """
    生成比对报告
    """
    # 统计信息
    total_flow = len(match_result)
    matched = len(match_result[match_result['匹配状态'] == '✓ 已匹配'])
    unmatched = len(match_result[match_result['匹配状态'] == '✗ 未匹配'])
    
    print("\n" + "="*60)
    print("比对结果统计")
    print("="*60)
    print(f"银行流水总数: {total_flow}")
    print(f"已匹配: {matched} ({matched/total_flow*100:.1f}%)")
    print(f"未匹配: {unmatched} ({unmatched/total_flow*100:.1f}%)")
    print(f"结算单未匹配记录: {len(unmatched_bill)}")
    print("="*60)
    
    # 导出到Excel，多个sheet
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: 匹配结果
        match_result.to_excel(writer, sheet_name='比对结果', index=False)
        
        # Sheet 2: 未匹配流水
        unmatched_flow = match_result[match_result['匹配状态'] == '✗ 未匹配']
        if len(unmatched_flow) > 0:
            unmatched_flow.to_excel(writer, sheet_name='未匹配流水', index=False)
        
        # Sheet 3: 未匹配结算单
        if len(unmatched_bill) > 0:
            unmatched_bill.to_excel(writer, sheet_name='未匹配结算单', index=False)
        
        # Sheet 4: 汇总统计
        summary = pd.DataFrame({
            '项目': ['银行流水总数', '已匹配', '未匹配', '结算单未匹配'],
            '数量': [total_flow, matched, unmatched, len(unmatched_bill)],
            '占比': [
                '100%',
                f"{matched/total_flow*100:.1f}%",
                f"{unmatched/total_flow*100:.1f}%",
                'N/A'
            ]
        })
        summary.to_excel(writer, sheet_name='汇总', index=False)
    
    print(f"\n✓ 报告已保存: {output_path}")
    
    # 如果有未匹配项，给出警告
    if unmatched > 0:
        print(f"\n⚠️ 警告: 发现 {unmatched} 条银行流水未找到对应结算明细！")
    if len(unmatched_bill) > 0:
        print(f"⚠️ 警告: 发现 {len(unmatched_bill)} 条结算单记录未匹配到银行流水！")


def main():
    parser = argparse.ArgumentParser(description='银联数据比对工具')
    parser.add_argument('bill_file', help='银联结算账单文件路径 (Excel/CSV)')
    parser.add_argument('flow_file', help='银行流水文件路径 (Excel/CSV)')
    parser.add_argument('-o', '--output', default='unionpay_reconcile_report.xlsx', 
                        help='输出报告路径 (默认: unionpay_reconcile_report.xlsx)')
    parser.add_argument('--bill-store', default='门店', help='结算单门店列名')
    parser.add_argument('--bill-date', default='日期', help='结算单日期列名')
    parser.add_argument('--bill-amount', default='金额', help='结算单金额列名')
    parser.add_argument('--flow-store', default='门店', help='流水门店列名')
    parser.add_argument('--flow-date', default='日期', help='流水日期列名')
    parser.add_argument('--flow-amount', default='交易金额', help='流水金额列名')
    parser.add_argument('--flow-fee', default='手续费', help='流水手续费列名')
    parser.add_argument('--amount-tol', type=float, default=0.01, help='金额容差 (默认0.01)')
    
    args = parser.parse_args()
    
    print("="*60)
    print("银联数据比对工具")
    print("="*60)
    
    # 加载数据
    df_bill, df_flow = load_and_process_data(
        args.bill_file, args.flow_file,
        args.bill_store, args.bill_date, args.bill_amount,
        args.flow_store, args.flow_date, args.flow_amount, args.flow_fee
    )
    
    # 执行比对
    match_result, unmatched_bill = reconcile_data(
        df_bill, df_flow,
        amount_tolerance=args.amount_tol
    )
    
    # 生成报告
    generate_report(match_result, unmatched_bill, args.output)


if __name__ == '__main__':
    main()
