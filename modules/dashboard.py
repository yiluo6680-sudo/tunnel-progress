"""KPI 面板计算"""


def calculate_kpis(status_df, measurement_df):
    """计算顶部面板的三个核心KPI"""
    kpis = {}

    if status_df.empty:
        return {'进度_pct': 0, '已计量_总额': 0, '可计量_总额': 0, '可计量_区段': 0, '可计量_长度': 0}

    # 进度：二衬完成比例
    lining = status_df[status_df['工序名称'] == '二衬']
    if not lining.empty:
        total_len = (lining['终点桩号_m'] - lining['起始桩号_m']).sum()
        done_len = lining[lining['状态'].isin(['已计量', '可计量', '申报中', '已完工'])].apply(
            lambda r: r['终点桩号_m'] - r['起始桩号_m'], axis=1
        ).sum()
        kpis['进度_pct'] = round(done_len / total_len * 100, 1) if total_len else 0
    else:
        kpis['进度_pct'] = 0

    # 已计量总额（取较大值：从已计量区段求和 vs 计量台账直接求和）
    amt_from_status = lining[lining['状态'] == '已计量']['已计量金额'].sum()
    if not measurement_df.empty and '已计量金额' in measurement_df.columns:
        amt_from_meas = measurement_df[measurement_df['计量状态'] == '已审批']['已计量金额'].sum()
        kpis['已计量_总额'] = max(amt_from_status, amt_from_meas)
    else:
        kpis['已计量_总额'] = amt_from_status

    # 可计量
    eligible = status_df[status_df['状态'] == '可计量']
    kpis['可计量_总额'] = eligible['可预估金额'].sum()
    kpis['可计量_区段'] = len(eligible)
    kpis['可计量_长度'] = (eligible['终点桩号_m'] - eligible['起始桩号_m']).sum()

    return kpis
