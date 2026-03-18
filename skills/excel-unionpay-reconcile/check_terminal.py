import pandas as pd

# 读取数据检查
settle_df = pd.read_excel('/root/.openclaw/media/inbound/结算账单_银联商务_20260316091946---eb8e67f9-216e-46fb-a662-ba87e0d04c88.xlsx')
flow_df = pd.read_excel('/root/.openclaw/media/inbound/银行到账20260316092009---7d9a23f9-8353-469f-bf27-6127d1d1a98c.xlsx')

# 检查那4笔无门店流水的详细信息
amounts = [10867.09, 6373.65, 245.54, 1968.18]
print('未匹配无门店流水（原始数据）:')
for amt in amounts:
    row = flow_df[flow_df['交易金额'] == amt]
    if not row.empty:
        print(f"\n金额 {amt}:")
        print(f"  交易时间: {row.iloc[0]['交易时间']}")
        print(f"  摘要: {row.iloc[0]['摘要']}")
        print(f"  银行交易流水号: {row.iloc[0]['银行交易流水号']}")

print("\n\n结算单中的终端号分组:")
terminal_groups = settle_df.groupby(['POS门店名称', '终端号'])['清分金额'].sum().reset_index()
print(terminal_groups.to_string(index=False))
