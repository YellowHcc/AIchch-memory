#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银联数据比对工具 - V2
功能：按门店+日期分组汇总后比对，解决多笔流水对应一笔结算的场景
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
            # 如果是两位数年份，需要补全
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
    
    # 标准化列名
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
    
    # 流水净额 = 交易金额 - 手续费
    df_flow['_net_amount'] = df_flow['_amount'] - df_flow['_fee']
    
    # 门店名称清洗
    df_bill['_store'] = df_bill[bill_store_col].astype(str).str.strip()
    df_flow['_store'] = df_flow[flow_store_col].astype(str).str.strip()
    
    return df_bill, df_flow


def reconcile_data_v2(df_bill, df_flow, amount_tolerance=0.01):
    """
    V2 比对逻辑：按门店+日期分组汇总后比对
    """
    results = []
    bill_stores = df_bill['_store'].unique().tolist()
    
    print("\n开始比对...")
    print("="*60)
    
    # 1. 建立门店映射（流水门店 -> 结算单门店）
    store_mapping = {}
    flow_stores = df_flow['_store'].unique()
    
    for flow_store in flow_stores:
        matched_store, score = find_best_match(flow_store, bill_stores, threshold=0.6)
        if matched_store:
            store_mapping[flow_store] = matched_store
            print(f"门店映射: {flow_store} -> {matched_store} (相似度: {score:.1%})")
    
    print("="*60)
    
    # 2. 给流水添加匹配后的门店名
    df_flow['_matched_store'] = df_flow['_store'].map(store_mapping)
    
    # 3. 按（匹配后门店+日期）分组汇总流水
    flow_grouped = df_flow.groupby(['_matched_store', '_date']).agg({
        '_amount': 'sum',      # 交易金额汇总
        '_fee': 'sum',         # 手续费汇总
        '_net_amount': 'sum',  # 净额汇总
        '_store': lambda x: ', '.join(x.unique())  # 原始门店名
    }).reset_index()
    flow_grouped.columns = ['门店', '日期', '交易金额汇总', '手续费汇总', '净额汇总', '原始门店名']
    
    # 4. 按（门店+日期）分组汇总结算单
    bill_grouped = df_bill.groupby(['_store', '_date']).agg({
        '_amount': 'sum'
    }).reset_index()
    bill_grouped.columns = ['门店', '日期', '结算金额汇总']
    
    # 5. 逐组比对
    matched_flow_indices = set()
    matched_bill_groups = set()
    
    for idx, flow_group in flow_grouped.iterrows():
        store = flow_group['门店']
        date = flow_group['日期']
        flow_net = flow_group['净额汇总']
        flow_gross = flow_group['交易金额汇总']
        flow_fee = flow_group['手续费汇总']
        original_stores = flow_group['原始门店名']
        
        # 查找对应结算单
        bill_candidates = bill_grouped[
            (bill_grouped['门店'] == store) & 
            (bill_grouped['日期'] == date)
        ]
        
        if len(bill_candidates) == 0:
            # 无对应结算单 - 可能是日期或门店不匹配
            results.append({
                '门店': store if pd.notna(store) else original_stores,
                '日期': format_date_display(date),
                '流水笔数': len(df_flow[(df_flow['_matched_store'] == store) & (df_flow['_date'] == date)]),
                '交易金额汇总': flow_gross,
                '手续费汇总': flow_fee,
                '净额汇总': flow_net,
                '结算金额汇总': 'N/A',
                '差额': 'N/A',
                '匹配状态': '✗ 无法匹配',
                '备注': '无对应结算单（门店或日期不匹配）',
                '_flow_store': store,
                '_flow_date': date,
                '_bill_amount': 0
            })
        else:
            bill_amount = bill_candidates.iloc[0]['结算金额汇总']
            bill_key = (store, date)
            
            # 计算差额
            diff = flow_net - bill_amount
            
            # 判断是否匹配（考虑容差）
            if abs(diff) <= amount_tolerance:
                # 完全匹配
                matched_flow_indices.add(idx)
                matched_bill_groups.add(bill_key)
                
                results.append({
                    '门店': store,
                    '日期': format_date_display(date),
                    '流水笔数': len(df_flow[(df_flow['_matched_store'] == store) & (df_flow['_date'] == date)]),
                    '交易金额汇总': flow_gross,
                    '手续费汇总': flow_fee,
                    '净额汇总': flow_net,
                    '结算金额汇总': bill_amount,
                    '差额': diff,
                    '匹配状态': '✓ 已匹配',
                    '备注': '',
                    '_flow_store': store,
                    '_flow_date': date,
                    '_bill_amount': bill_amount
                })
            else:
                # 金额不匹配 - 需要进一步分析
                if flow_net < bill_amount:
                    # 流水 < 结算单：有流水未匹配到
                    results.append({
                        '门店': store,
                        '日期': format_date_display(date),
                        '流水笔数': len(df_flow[(df_flow['_matched_store'] == store) & (df_flow['_date'] == date)]),
                        '交易金额汇总': flow_gross,
                        '手续费汇总': flow_fee,
                        '净额汇总': flow_net,
                        '结算金额汇总': bill_amount,
                        '差额': diff,
                        '匹配状态': '✗ 金额不符',
                        '备注': f'结算单金额大，可能缺少 {abs(diff):.2f} 元的流水',
                        '_flow_store': store,
                        '_flow_date': date,
                        '_bill_amount': bill_amount
                    })
                else:
                    # 流水 > 结算单：有多余流水或结算单金额有误
                    results.append({
                        '门店': store,
                        '日期': format_date_display(date),
                        '流水笔数': len(df_flow[(df_flow['_matched_store'] == store) & (df_flow['_date'] == date)]),
                        '交易金额汇总': flow_gross,
                        '手续费汇总': flow_fee,
                        '净额汇总': flow_net,
                        '结算金额汇总': bill_amount,
                        '差额': diff,
                        '匹配状态': '✗ 金额不符',
                        '备注': f'流水金额大，超出 {abs(diff):.2f} 元',
                        '_flow_store': store,
                        '_flow_date': date,
                        '_bill_amount': bill_amount
                    })
    
    # 6. 找出结算单中未被匹配的记录
    unmatched_bill = []
    for idx, bill_row in bill_grouped.iterrows():
        bill_key = (bill_row['门店'], bill_row['日期'])
        if bill_key not in matched_bill_groups:
            unmatched_bill.append({
                '门店': bill_row['门店'],
                '日期': format_date_display(bill_row['日期']),
                '结算金额汇总': bill_row['结算金额汇总'],
                '原始日期': bill_row['日期']
            })
    
    unmatched_bill_df = pd.DataFrame(unmatched_bill)
    
    return pd.DataFrame(results), unmatched_bill_df


def generate_report_v2(match_result, unmatched_bill, df_flow, output_path):
    """
    生成V2比对报告
    """
    # 统计信息
    total_flow_groups = len(match_result)
    matched = len(match_result[match_result['匹配状态'] == '✓ 已匹配'])
    unmatched = len(match_result[match_result['匹配状态'].str.contains('无法匹配|金额不符')])
    
    print("\n" + "="*60)
    print("比对结果统计")
    print("="*60)
    print(f"流水组数(按门店+日期): {total_flow_groups}")
    print(f"已匹配: {matched} ({matched/total_flow_groups*100:.1f}%)")
    print(f"未匹配/金额不符: {unmatched} ({unmatched/total_flow_groups*100:.1f}%)")
    print(f"结算单未匹配记录: {len(unmatched_bill)}")
    
    # 计算金额汇总
    if len(match_result) > 0:
        total_bill_matched = match_result[match_result['匹配状态'] == '✓ 已匹配']['结算金额汇总'].sum()
        total_flow_matched = match_result[match_result['匹配状态'] == '✓ 已匹配']['净额汇总'].sum()
        print(f"\n已匹配结算金额: {total_bill_matched:,.2f}")
        print(f"已匹配流水净额: {total_flow_matched:,.2f}")
    print("="*60)
    
    # 导出到Excel，多个sheet
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: 全部比对结果
        display_cols = ['门店', '日期', '流水笔数', '交易金额汇总', '手续费汇总', '净额汇总', '结算金额汇总', '差额', '匹配状态', '备注']
        match_result_display = match_result[display_cols] if len(match_result) > 0 else pd.DataFrame(columns=display_cols)
        match_result_display.to_excel(writer, sheet_name='比对结果', index=False)
        
        # Sheet 2: 未匹配流水组
        unmatched_flow = match_result[match_result['匹配状态'].str.contains('无法匹配')]
        if len(unmatched_flow) > 0:
            unmatched_flow[display_cols].to_excel(writer, sheet_name='未匹配流水组', index=False)
        
        # Sheet 3: 金额不符
        amount_mismatch = match_result[match_result['匹配状态'].str.contains('金额不符')]
        if len(amount_mismatch) > 0:
            amount_mismatch[display_cols].to_excel(writer, sheet_name='金额不符', index=False)
        
        # Sheet 4: 未匹配结算单
        if len(unmatched_bill) > 0:
            unmatched_bill.to_excel(writer, sheet_name='未匹配结算单', index=False)
        
        # Sheet 5: 汇总统计
        summary_data = {
            '项目': ['流水组数(门店+日期)', '已匹配', '未匹配', '金额不符', '结算单未匹配组数'],
            '数量': [total_flow_groups, matched, 
                    len(match_result[match_result['匹配状态'].str.contains('无法匹配')]),
                    len(amount_mismatch),
                    len(unmatched_bill)],
        }
        summary = pd.DataFrame(summary_data)
        summary.to_excel(writer, sheet_name='汇总', index=False)
    
    print(f"\n✓ 报告已保存: {output_path}")
    
    # 如果有未匹配项，给出警告
    if unmatched > 0:
        print(f"\n⚠️ 警告: 发现 {unmatched} 组流水存在问题！")
    if len(unmatched_bill) > 0:
        print(f"⚠️ 警告: 发现 {len(unmatched_bill)} 组结算单未匹配到流水！")


def main():
    parser = argparse.ArgumentParser(description='银联数据比对工具 V2 - 按门店+日期分组汇总比对')
    parser.add_argument('bill_file', help='银联结算账单文件路径 (Excel/CSV)')
    parser.add_argument('flow_file', help='银行流水文件路径 (Excel/CSV)')
    parser.add_argument('-o', '--output', default='unionpay_reconcile_report_v2.xlsx', 
                        help='输出报告路径 (默认: unionpay_reconcile_report_v2.xlsx)')
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
    print("银联数据比对工具 V2")
    print("匹配逻辑: 按门店+日期分组汇总后比对")
    print("="*60)
    
    # 加载数据
    df_bill, df_flow = load_and_process_data(
        args.bill_file, args.flow_file,
        args.bill_store, args.bill_date, args.bill_amount,
        args.flow_store, args.flow_date, args.flow_amount, args.flow_fee
    )
    
    # 执行比对
    match_result, unmatched_bill = reconcile_data_v2(
        df_bill, df_flow,
        amount_tolerance=args.amount_tol
    )
    
    # 生成报告
    generate_report_v2(match_result, unmatched_bill, df_flow, args.output)


if __name__ == '__main__':
    main()
