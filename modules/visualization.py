"""可视化模块 V2 — 紧凑左右并排布局"""

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from config import ALL_STEPS, STATUS_COLORS, ROCK_COLORS


def create_tunnel_figure(status_df, design_df, tunnel_name="田头山隧道"):
    """生成隧道进度计量视图

    布局：每个线别一个独立图框，内部包含多行工序条
    - 围岩背景行（半透明色块）
    - 各工序进度行（灰色=未施工，蓝色=已完工）
    - 二衬计量状态行（绿=已计量，黄=可计量，橙=申报中）
    """
    if status_df.empty:
        return go.Figure()

    lines = sorted(status_df['线别'].unique(), reverse=True)
    n_lines = len(lines)

    # 每个线别独占一个子图
    fig = make_subplots(
        rows=n_lines, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.15,
        subplot_titles=[f"{ln}" for ln in lines],
    )

    # 全局桩号范围
    all_m = status_df[['起始桩号_m', '终点桩号_m']]
    x_min, x_max = all_m.min().min(), all_m.max().max()
    x_pad = (x_max - x_min) * 0.02

    for line_idx, line in enumerate(lines):
        row = line_idx + 1
        ldf = status_df[status_df['线别'] == line]
        ldesign = design_df[design_df['线别'] == line] if not design_df.empty else pd.DataFrame()

        # 该线别的工序列表
        steps_in_line = [s for s in ALL_STEPS if s in ldf['工序名称'].values]

        # 每个工序的 Y 坐标
        y_positions = {}
        y_labels = {}
        y_idx = 0

        # === 第0行：围岩背景 ===
        if not ldesign.empty:
            for _, seg in ldesign.iterrows():
                rock = seg['围岩级别']
                color = ROCK_COLORS.get(rock, 'rgba(200,200,200,0.3)')
                fig.add_trace(go.Scatter(
                    x=[seg['起始桩号_m'], seg['终点桩号_m'], seg['终点桩号_m'], seg['起始桩号_m']],
                    y=[0, 0, 1, 1],
                    fill='toself', fillcolor=color,
                    mode='none', showlegend=False, hoverinfo='skip',
                ), row=row, col=1)
            # 添加围岩标记文本
            for _, seg in ldesign.iterrows():
                mid = (seg['起始桩号_m'] + seg['终点桩号_m']) / 2
                fig.add_annotation(
                    x=mid, y=0.5,
                    text=seg['围岩级别'],
                    showarrow=False, font=dict(size=9, color='#555'),
                    xref='x', yref=f'y{row}',
                )
            y_positions['围岩'] = 0
            y_labels[0] = '围岩'
            y_idx = 1

        # === 各工序行 ===
        proc_y = y_idx
        for step in steps_in_line:
            sdf = ldf[ldf['工序名称'] == step].sort_values('起始桩号_m')
            for _, seg in sdf.iterrows():
                status = seg['状态']
                color = STATUS_COLORS.get(status, '#D3D3D3')
                seg_len = seg['终点桩号_m'] - seg['起始桩号_m']
                base = seg['起始桩号_m']
                label = status

                # 悬浮提示
                hov = (
                    f"<b>{line} {_fmt_chainage(seg['起始桩号_m'])} ~ {_fmt_chainage(seg['终点桩号_m'])}</b><br>"
                    f"工序: {step}<br>"
                    f"围岩: {seg['围岩级别']}级<br>"
                    f"状态: {label}<br>"
                    f"长度: {seg_len:.0f}m"
                )
                if seg.get('已完成金额', 0) > 0:
                    hov += f"<br>已计量: ¥{seg['已计量金额']:,.0f}"
                if seg.get('可预估金额', 0) > 0:
                    hov += f"<br>可预估: ¥{seg['可预估金额']:,.0f}"
                cd = seg.get('完成日期')
                if cd is not None and pd.notna(cd):
                    hov += f"<br>完成: {cd.strftime('%Y-%m-%d')}"

                fig.add_trace(go.Bar(
                    x=[seg_len], y=[proc_y],
                    base=base, orientation='h',
                    marker=dict(color=color, line=dict(width=0.2, color='white')),
                    hoverinfo='text', hovertext=hov,
                    showlegend=False, width=0.6,
                ), row=row, col=1)

            y_positions[step] = proc_y
            y_labels[proc_y] = step
            proc_y += 1

        # === 底部：计量状态行 ===
        meas_y = proc_y
        meas_df = ldf[ldf['工序名称'] == '二衬'].sort_values('起始桩号_m')
        for _, seg in meas_df.iterrows():
            status = seg['状态']
            color = STATUS_COLORS.get(status, '#D3D3D3')
            if status in ('可计量', '申报中', '已计量'):
                seg_len = seg['终点桩号_m'] - seg['起始桩号_m']
                base = seg['起始桩号_m']
                hov = (
                    f"<b>{line} {_fmt_chainage(seg['起始桩号_m'])} ~ {_fmt_chainage(seg['终点桩号_m'])}</b><br>"
                    f"计量状态: {status}<br>"
                )
                if seg.get('已计量金额', 0) > 0:
                    hov += f"已计量: ¥{seg['已计量金额']:,.0f}"
                if seg.get('可预估金额', 0) > 0:
                    hov += f"可预估: ¥{seg['可预估金额']:,.0f}"
                fig.add_trace(go.Bar(
                    x=[seg_len], y=[meas_y],
                    base=base, orientation='h',
                    marker=dict(color=color, line=dict(width=0.4, color='#333')),
                    hoverinfo='text', hovertext=hov,
                    showlegend=False, width=0.6,
                ), row=row, col=1)
        y_positions['计量'] = meas_y
        y_labels[meas_y] = '💰 计量'

        # === 设置Y轴 ===
        y_tickvals = sorted(y_labels.keys())
        y_ticktext = [y_labels[v] for v in y_tickvals]
        fig.update_yaxes(
            tickvals=y_tickvals, ticktext=y_ticktext,
            row=row, col=1,
            gridcolor='#f8f8f8',
            zeroline=False,
        )

        # 添加线别信息
        start_s = _fmt_chainage(x_min)
        end_s = _fmt_chainage(x_max)
        fig.update_xaxes(
            title_text='桩号' if row == n_lines else '',
            tickmode='array',
            tickvals=_gen_ticks(x_min, x_max),
            ticktext=[_fmt_chainage(v) for v in _gen_ticks(x_min, x_max)],
            gridcolor='#eee',
            range=[x_min - x_pad, x_max + x_pad],
            row=row, col=1,
        )

    # === 全局布局 ===
    fig.update_layout(
        title=dict(
            text=f"<b> {tunnel_name}</b>  施工进度 · 计量状态",
            font=dict(size=16), x=0.5,
        ),
        height=max(250, n_lines * 180),
        margin=dict(l=20, r=30, t=60, b=30),
        hovermode='closest',
        plot_bgcolor='white',
        barmode='overlay',
        bargap=0.3,
    )

    return fig


def _fmt_chainage(meters):
    km = int(meters / 1000)
    m_val = meters - km * 1000
    return f"K{km}+{m_val:.0f}"


def _gen_ticks(start, end, step=200):
    s = int(start / step) * step
    ticks = []
    while s <= end:
        ticks.append(s)
        s += step
    return ticks
