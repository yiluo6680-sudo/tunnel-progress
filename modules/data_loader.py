"""数据加载模块 — 支持灵活列名映射"""

import pandas as pd
import streamlit as st
from utils.chainage_utils import parse_chainage


def load_design_baseline(file) -> pd.DataFrame:
    """加载隧道参数表
    必需列: 线别, 起始桩号, 终点桩号, 围岩级别
    可选: 隧道名称
    """
    df = _read_file(file)
    df = _rename_col(df, {
        '隧道名称': '隧道名称', '线别': '线别',
        '起始桩号': '起始桩号', '终点桩号': '终点桩号',
        '围岩级别': '围岩级别', '围岩': '围岩级别',
    })
    _require_cols(df, ['线别', '起始桩号', '终点桩号', '围岩级别'], '隧道参数表')
    df['起始桩号_m'] = df['起始桩号'].apply(parse_chainage)
    df['终点桩号_m'] = df['终点桩号'].apply(parse_chainage)
    df['隧道名称'] = df.get('隧道名称', '隧道')
    return df


def load_construction_progress(file) -> pd.DataFrame:
    """加载施工进度表
    必需列: 线别, 起始桩号, 终点桩号, 工序名称, 完成日期
    """
    df = _read_file(file)
    df = _rename_col(df, {
        '线别': '线别', '起始桩号': '起始桩号', '终点桩号': '终点桩号',
        '工序名称': '工序名称', '工序': '工序名称',
        '完成日期': '完成日期',
    })
    _require_cols(df, ['线别', '起始桩号', '终点桩号', '工序名称', '完成日期'], '施工进度表')
    df['起始桩号_m'] = df['起始桩号'].apply(parse_chainage)
    df['终点桩号_m'] = df['终点桩号'].apply(parse_chainage)
    df['完成日期'] = pd.to_datetime(df['完成日期'], errors='coerce')
    return df


def load_measurement_ledger(file) -> pd.DataFrame:
    """加载计量数据表
    必需列: 线别, 起始桩号, 终点桩号, 计量状态
    可选: 已计量金额, 计量批次
    """
    df = _read_file(file)
    df = _rename_col(df, {
        '线别': '线别', '起始桩号': '起始桩号', '终点桩号': '终点桩号',
        '计量状态': '计量状态', '状态': '计量状态',
        '已计量金额': '已计量金额', '金额': '已计量金额',
        '计量批次': '计量批次', '批次': '计量批次', '期数': '计量批次',
    })
    _require_cols(df, ['线别', '起始桩号', '终点桩号', '计量状态'], '计量数据表')
    df['起始桩号_m'] = df['起始桩号'].apply(parse_chainage)
    df['终点桩号_m'] = df['终点桩号'].apply(parse_chainage)
    if '已计量金额' not in df.columns:
        df['已计量金额'] = 0.0
    if '计量批次' not in df.columns:
        df['计量批次'] = ''
    df['已计量金额'] = pd.to_numeric(df['已计量金额'], errors='coerce').fillna(0)
    status_map = {'已审批': '已审批', '已计量': '已审批', '审批通过': '已审批',
                  '申报中': '申报中', '审核中': '申报中',
                  '未计量': '未计量', '未申报': '未计量'}
    df['计量状态'] = df['计量状态'].map(status_map).fillna('未计量')
    return df


def _read_file(file):
    if isinstance(file, str):
        df = pd.read_csv(file) if file.endswith('.csv') else pd.read_excel(file, engine='openpyxl')
    else:
        df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file, engine='openpyxl')
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _rename_col(df, mapping):
    """智能列名映射（支持部分匹配）"""
    cols_lower = {c.lower(): c for c in df.columns}
    for target, source in mapping.items():
        if source in df.columns:
            continue
        # 尝试模糊匹配
        found = cols_lower.get(source.lower())
        if found:
            df = df.rename(columns={found: source})
        else:
            # 尝试在列名中包含关键词
            for c_lower, c_orig in cols_lower.items():
                if source.lower() in c_lower:
                    df = df.rename(columns={c_orig: source})
                    break
    return df


def _require_cols(df, required, table_name):
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"【{table_name}】缺少必需列: {', '.join(missing)}")
        st.info(f"现有列: {', '.join(df.columns)}")
        st.stop()
