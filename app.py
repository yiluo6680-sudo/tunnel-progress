"""隧道工程施工进度及计量可视化系统 V2"""

import os, io
import pandas as pd
import streamlit as st
from config import STATUS_COLORS, STATUS_LABELS
from modules.data_loader import load_design_baseline, load_construction_progress, load_measurement_ledger
from modules.status_engine import create_segment_grid, determine_status, estimate_eligible_amounts
from modules.dashboard import calculate_kpis
from modules.visualization import create_tunnel_figure

st.set_page_config(page_title="隧道进度计量可视化", page_icon="🚇", layout="wide")

# === 初始化 ===
if 'status_df' not in st.session_state:
    st.session_state.status_df = None
if 'df_design' not in st.session_state:
    st.session_state.df_design = None
    st.session_state.kpis = None

DEMO_DIR = os.path.join(os.path.dirname(__file__), 'data')


# === 引擎函数 ===
def _run_engine(df_design, df_progress, df_meas):
    segs = create_segment_grid(df_design)
    status_df = determine_status(segs, df_progress, df_meas)
    status_df = estimate_eligible_amounts(status_df)
    kpis = calculate_kpis(status_df, df_meas)

    st.session_state.df_design = df_design
    st.session_state.df_progress = df_progress
    st.session_state.df_meas = df_meas
    st.session_state.status_df = status_df
    st.session_state.kpis = kpis


def _fch(m):
    return f"K{int(m/1000)}+{m%1000:.0f}"


# === 侧边栏 ===
with st.sidebar:
    st.title("🚇 隧道可视化")
    st.markdown("---")

    # 数据来源
    data_source = st.radio("数据来源", ["📥 加载示例数据", "📂 上传我的数据"])

    if data_source == "📥 加载示例数据":
        if st.button("✅ 加载田头山隧道示例", use_container_width=True, type="primary"):
            df_design = load_design_baseline(os.path.join(DEMO_DIR, 'design_baseline.csv'))
            df_progress = load_construction_progress(os.path.join(DEMO_DIR, 'construction_progress.csv'))
            df_meas = load_measurement_ledger(os.path.join(DEMO_DIR, 'measurement_ledger.csv'))
            _run_engine(df_design, df_progress, df_meas)
            st.rerun()
    else:
        design_file = st.file_uploader("① 隧道参数表", type=['xlsx', 'xls', 'csv'])
        progress_file = st.file_uploader("② 施工进度表", type=['xlsx', 'xls', 'csv'])
        meas_file = st.file_uploader("③ 计量数据表", type=['xlsx', 'xls', 'csv'])
        if design_file and progress_file and meas_file:
            df_design = load_design_baseline(design_file)
            df_progress = load_construction_progress(progress_file)
            df_meas = load_measurement_ledger(meas_file)
            _run_engine(df_design, df_progress, df_meas)
            st.rerun()

    # 数据模板下载
    st.markdown("---")
    with st.expander("📋 查看数据模板格式", expanded=False):
        st.markdown("**表1：隧道参数表**")
        st.code("隧道名称,线别,起始桩号,终点桩号,围岩级别\n田头山隧道,左线,ZK85+147,ZK86+523.5,Ⅲ")
        st.markdown("**表2：施工进度表**")
        st.code("线别,起始桩号,终点桩号,工序名称,完成日期\n左线,ZK85+147,ZK86+047,开挖初支,2024-08-20")
        st.markdown("**表3：计量数据表**")
        st.code("线别,起始桩号,终点桩号,已计量金额,计量批次,计量状态\n左线,ZK85+147,ZK85+447,500000,第01期,已审批")
        st.caption("工序名称可选: 开挖初支 / 仰拱 / 二衬")
        st.caption("计量状态可选: 已审批 / 申报中 / 未计量")

    # 图例
    st.markdown("---")
    st.subheader("🎨 图例")
    for status, color in STATUS_COLORS.items():
        label = STATUS_LABELS.get(status, status)
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:8px;margin:4px 0;">'
            f'<div style="width:24px;height:16px;background:{color};border-radius:2px;"></div>'
            f'<span>{label}</span></div>',
            unsafe_allow_html=True,
        )

# === 主界面 ===
if st.session_state.status_df is not None:
    status_df = st.session_state.status_df
    df_design = st.session_state.df_design
    kpis = st.session_state.kpis

    # --- KPI 面板 ---
    st.subheader("📊 总览")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("📈 总进度（二衬）", f"{kpis['进度_pct']}%")
    with c2:
        st.metric("💰 已计量总额", f"¥{kpis['已计量_总额']/1e4:.0f}万")
    with c3:
        st.metric("⭐ 可计量待申报", f"¥{kpis['可计量_总额']/1e4:.0f}万",
                  delta=f"{kpis['可计量_区段']}个区段")
    with c4:
        st.metric("📏 可计量长度", f"{kpis['可计量_长度']:.0f}m")
    with c5:
        # 进度对比
        line_progress = status_df[status_df['工序名称'] == '二衬'].groupby('线别').apply(
            lambda g: f"{sum(g['状态'].isin(['已计量','可计量','申报中','已完工'])*(g['终点桩号_m']-g['起始桩号_m']))/(g['终点桩号_m']-g['起始桩号_m']).sum()*100:.0f}%"
        )
        progress_str = " | ".join(line_progress)
        st.metric("🚇 各线进度", progress_str)

    # --- 过滤器 ---
    st.markdown("---")
    fcols = st.columns(3)
    with fcols[0]:
        lines = ['全部'] + sorted(status_df['线别'].unique())
        sel_line = st.selectbox("线别", lines, key='line_filter')
    with fcols[1]:
        rocks = ['全部'] + sorted(status_df['围岩级别'].unique())
        sel_rock = st.selectbox("围岩级别", rocks, key='rock_filter')
    with fcols[2]:
        st.selectbox("显示模式", ["综合视图", "只看计量状态"], key='view_mode')

    # --- 主图 ---
    st.markdown("---")
    plot_df = status_df.copy()
    if sel_line != '全部':
        plot_df = plot_df[plot_df['线别'] == sel_line]
    if sel_rock != '全部':
        plot_df = plot_df[plot_df['围岩级别'] == sel_rock]

    fig = create_tunnel_figure(plot_df, df_design)
    st.plotly_chart(fig, use_container_width=True)

    # --- 可计量区段明细 ---
    eligible = status_df[status_df['状态'] == '可计量']
    if not eligible.empty:
        st.markdown("---")
        st.subheader(f"⭐ 可计量区段明细 ({len(eligible)} 个区段)")

        disp = eligible.copy()
        disp['起始桩号'] = disp.apply(lambda r: _fch(r['起始桩号_m']), axis=1)
        disp['终点桩号'] = disp.apply(lambda r: _fch(r['终点桩号_m']), axis=1)
        disp['长度(m)'] = disp['终点桩号_m'] - disp['起始桩号_m']
        disp['预估金额(万元)'] = disp['可预估金额'] / 1e4
        cols = ['线别', '起始桩号', '终点桩号', '围岩级别', '长度(m)', '预估金额(万元)']
        st.dataframe(
            disp[cols].sort_values(['线别', '起始桩号']),
            use_container_width=True, hide_index=True,
            column_config={
                '预估金额(万元)': st.column_config.NumberColumn(format="¥%.1f万"),
                '长度(m)': st.column_config.NumberColumn(format="%.0f"),
            }
        )
        csv = disp[cols].to_csv(index=False, encoding='utf-8-sig')
        st.download_button("📥 导出可计量区段报表", csv, "可计量区段.csv", "text/csv")

    # --- 原始数据检查 ---
    with st.expander("📋 数据详情"):
        t1, t2, t3 = st.tabs(["隧道参数", "施工进度", "计量数据"])
        with t1:
            st.dataframe(df_design, use_container_width=True, hide_index=True)
        with t2:
            st.dataframe(st.session_state.get('df_progress', pd.DataFrame()), use_container_width=True, hide_index=True)
        with t3:
            st.dataframe(st.session_state.get('df_meas', pd.DataFrame()), use_container_width=True, hide_index=True)

else:
    # --- 欢迎页 ---
    st.markdown("""
    <div style="text-align:center;padding:60px 20px;">
        <h1 style="font-size:2.2rem;">🚇 隧道施工进度 · 计量可视化系统</h1>
        <p style="font-size:1.1rem;color:#666;margin:20px 0;">一眼看清：已开挖初支 / 已仰拱 / 已二衬 / 已计量 /  ⭐可计量</p>
        <p style="color:#999;">👈 左侧点击 <b>"加载示例数据"</b> 立即体验</p>
    </div>
    """, unsafe_allow_html=True)
