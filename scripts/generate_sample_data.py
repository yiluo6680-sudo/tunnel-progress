"""生成田头山隧道示例数据 V2 — 简化格式"""

import os, random, re
import pandas as pd
from datetime import datetime, timedelta

random.seed(42)

# === 桩号辅助函数 ===
def ch_to_m(ch):
    m = re.match(r'(?:[A-Z]*)?K?(\d+)\+(\d+(?:\.\d+)?)', ch.strip().upper())
    return int(m.group(1)) * 1000 + float(m.group(2)) if m else 0

def m_to_ch(m, prefix=''):
    return f"{prefix}K{int(m/1000)}+{m - int(m/1000)*1000:.1f}"

# === 1. 隧道参数表 ===
design = [
    ('田头山隧道', '左线', 'ZK85+147',  'ZK86+523.5', 'Ⅲ'),
    ('田头山隧道', '左线', 'ZK86+523.5', 'ZK87+022.5', 'Ⅳ'),
    ('田头山隧道', '左线', 'ZK87+022.5', 'ZK87+736.5', 'Ⅴ'),
    ('田头山隧道', '右线', 'YK84+790',  'YK86+310.5', 'Ⅲ'),
    ('田头山隧道', '右线', 'YK86+310.5', 'YK87+128.5', 'Ⅳ'),
    ('田头山隧道', '右线', 'YK87+128.5', 'YK87+781.6', 'Ⅴ'),
]
df_design = pd.DataFrame(design, columns=['隧道名称', '线别', '起始桩号', '终点桩号', '围岩级别'])
df_design.to_csv('data/design_baseline.csv', index=False, encoding='utf-8-sig')
print(f"隧道参数: {len(df_design)} 行")

# === 2. 施工进度表 ===
# 左线: 开挖初支到960m (约37%), 仰拱到560m, 二衬到480m
# 右线: 开挖初支到1210m (约40%), 仰拱到750m, 二衬到660m
start_date = datetime(2024, 7, 19)

progress_data = []
# (线别, 起点m, 终点m, 工序, 起始日期偏移天数)
left_base = ch_to_m('ZK85+147')
right_base = ch_to_m('YK84+790')

progress_spec = [
    # 左线
    ('左线', left_base, left_base + 960, '开挖初支', 0),
    ('左线', left_base, left_base + 560, '仰拱', 60),
    ('左线', left_base, left_base + 480, '二衬', 90),
    # 右线
    ('右线', right_base, right_base + 1210, '开挖初支', 5),
    ('右线', right_base, right_base + 750, '仰拱', 65),
    ('右线', right_base, right_base + 660, '二衬', 95),
]

for line, s, e, step, day_off in progress_spec:
    step_size = 30 if step == '开挖初支' else 20
    pos = s
    seq = 0
    while pos < e:
        seq += 1
        seg_end = min(pos + step_size, e)
        days = day_off + seq * random.randint(10, 25)
        comp = start_date + timedelta(days=days)
        progress_data.append({
            '线别': line,
            '起始桩号': m_to_ch(pos),
            '终点桩号': m_to_ch(seg_end),
            '工序名称': step,
            '完成日期': comp.strftime('%Y-%m-%d'),
        })
        pos = seg_end

df_progress = pd.DataFrame(progress_data)
df_progress.to_csv('data/construction_progress.csv', index=False, encoding='utf-8-sig')
print(f"施工进度: {len(df_progress)} 行")

# === 3. 计量数据表 ===
meas_data = []
# 左线: 前400m已计量审批, 400-450m申报中
# 右线: 前580m已计量审批, 580-630m申报中
meas_spec = [
    ('左线', left_base, left_base + 400, '第04期', '已审批', 95000),
    ('左线', left_base + 400, left_base + 450, '第05期', '申报中', 0),
    ('右线', right_base, right_base + 580, '第06期', '已审批', 95000),
    ('右线', right_base + 580, right_base + 630, '第07期', '申报中', 0),
]

for line, s, e, period, status, unit_price in meas_spec:
    seg_size = 25
    pos = s
    seq = 0
    while pos < e:
        seq += 1
        seg_end = min(pos + seg_size, e)
        length = seg_end - pos
        amount = length * unit_price * random.uniform(0.9, 1.1)
        meas_data.append({
            '线别': line,
            '起始桩号': m_to_ch(pos),
            '终点桩号': m_to_ch(seg_end),
            '已计量金额': round(amount, 2),
            '计量批次': period,
            '计量状态': status,
        })
        pos = seg_end

df_meas = pd.DataFrame(meas_data)
df_meas.to_csv('data/measurement_ledger.csv', index=False, encoding='utf-8-sig')
print(f"计量数据: {len(df_meas)} 行")
print("\n✅ 数据生成完毕!")
