"""核心逻辑：根据围岩级别动态判定进度计量状态"""

import pandas as pd
from config import STEPS_BY_ROCK, SEGMENT_LENGTH


def create_segment_grid(design_df):
    """将隧道按围岩切分为微区段(10m)网格"""
    segs = []
    for _, r in design_df.iterrows():
        line = r['线别']
        pos = r['起始桩号_m']
        end = r['终点桩号_m']
        rock = r['围岩级别']
        while pos < end:
            seg_end = min(pos + SEGMENT_LENGTH, end)
            segs.append({
                '线别': line, '起始桩号_m': pos, '终点桩号_m': seg_end,
                '围岩级别': rock,
                # 根据围岩级别决定该区段应有的工序
                '工序列表': STEPS_BY_ROCK.get(rock, STEPS_BY_ROCK['Ⅳ']),
            })
            pos = seg_end
    return pd.DataFrame(segs)


def determine_status(segments_df, progress_df, measurement_df):
    """核心判定：每个微区段 × 每道工序 → 综合状态

    状态定义:
      - 未施工: 进度台账中无此工序记录
      - 已完工: 工序已完成，但二衬完成后的计量无记录
      - 申报中: 计量已申报待批
      - 已计量: 计量已审批
      - 可计量: (核心) 工序已完成，计量条件已满足但尚未申报
    """
    if segments_df.empty:
        return pd.DataFrame()

    rows = []
    for _, seg in segments_df.iterrows():
        sl = seg['线别']
        ss = seg['起始桩号_m']
        se = seg['终点桩号_m']
        rock = seg['围岩级别']
        steps = seg['工序列表']

        for step in steps:
            # --- 查进度：该工序在此区段是否已完成 ---
            done = False
            comp_date = None
            if not progress_df.empty:
                mask = (
                    (progress_df['线别'] == sl) &
                    (progress_df['工序名称'] == step) &
                    (progress_df['起始桩号_m'] <= ss) &
                    (progress_df['终点桩号_m'] >= se)
                )
                matches = progress_df[mask]
                if not matches.empty:
                    done = True
                    comp_date = matches.iloc[0].get('完成日期', None)

            # --- 状态判定 ---
            if not done:
                status = '未施工'
                meas_status = ''
                meas_amount = 0.0
            else:
                # 只有二衬才查计量状态
                if step == '二衬' and not measurement_df.empty:
                    m_mask = (
                        (measurement_df['线别'] == sl) &
                        (measurement_df['起始桩号_m'] <= ss) &
                        (measurement_df['终点桩号_m'] >= se)
                    )
                    m_matches = measurement_df[m_mask]
                    if not m_matches.empty:
                        m_row = m_matches.iloc[0]
                        m_stat = m_row['计量状态']
                        if m_stat == '已审批':
                            status = '已计量'
                        elif m_stat == '申报中':
                            status = '申报中'
                        else:
                            status = '可计量'  # 未计量的二衬完成段
                        meas_status = m_stat
                        meas_amount = float(m_row.get('已计量金额', 0))
                    else:
                        status = '可计量'  # 核心：二衬完成了但台账无记录
                        meas_status = ''
                        meas_amount = 0.0
                else:
                    status = '已完工'
                    meas_status = ''
                    meas_amount = 0.0

            rows.append({
                '线别': sl,
                '起始桩号_m': ss,
                '终点桩号_m': se,
                '围岩级别': rock,
                '工序名称': step,
                '状态': status,
                '已计量金额': meas_amount,
                '可预估金额': 0.0,
                '完成日期': comp_date,
            })

    return pd.DataFrame(rows)


def estimate_eligible_amounts(status_df):
    """估算可计量区段的金额（基于已计量区段的平均单价推算）"""
    df = status_df.copy()

    # 计算各线别各围岩的综合单价（延米）
    meas_done = df[(df['状态'] == '已计量') & (df['工序名称'] == '二衬')]
    unit_prices = {}
    if not meas_done.empty:
        for (line, rock), grp in meas_done.groupby(['线别', '围岩级别']):
            total_len = (grp['终点桩号_m'] - grp['起始桩号_m']).sum()
            total_amt = grp['已计量金额'].sum()
            if total_len > 0:
                unit_prices[(line, rock)] = total_amt / total_len

    for idx, r in df.iterrows():
        if r['状态'] == '可计量':
            seg_len = r['终点桩号_m'] - r['起始桩号_m']
            price = unit_prices.get((r['线别'], r['围岩级别']), 80000)
            df.at[idx, '可预估金额'] = seg_len * price

    return df
