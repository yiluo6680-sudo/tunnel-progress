"""
分项完工计量模板化生成工具
"""

import os, sys, re, io, glob
from datetime import datetime
import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from tools.generator_core import (
    scan_and_build_templates, load_template_index,
    generate_segment_files, ROCK_SUB_ITEMS,
    DEFAULT_TEMPLATE_SOURCE, DEFAULT_OUTPUT, TEMPLATE_DIR,
)

st.set_page_config(page_title="分项完工计量生成器", page_icon="📋", layout="wide")

for k in ["template_index", "parsed_segments"]:
    if k not in st.session_state:
        st.session_state[k] = None
if "template_source" not in st.session_state:
    st.session_state.template_source = DEFAULT_TEMPLATE_SOURCE


def _fmt(ch):
    """K87+492.00 → 87+492.00"""
    return ch.replace("K", "", 1) if ch.startswith("K") else ch


def _infer_item_name(filename):
    """从文件名推测工序名称，放在顶部避免NameError"""
    for kw in ["洞身开挖", "喷射砼支护", "锚杆支护", "钢筋网支护", "钢架支护",
               "衬砌钢筋加工及安装", "衬砌砼", "仰拱钢筋加工及安装",
               "仰拱砼", "仰拱回填", "超前小导管支护", "管棚",
               "衬砌防水层", "止水带", "明洞钢筋加工及安装", "明洞止水带",
               "超前小导管"]:
        if kw in filename:
            return kw
    name = os.path.splitext(filename)[0]
    name = re.sub(r'^WH3A4[\d\-]+', '', name)
    return name[:20] if name else filename

with st.sidebar:
    st.header("📂 模板源")
    ts = st.text_input("扫描目录", value=st.session_state.template_source,
        key="ts_inp", label_visibility="collapsed")
    st.session_state.template_source = ts
    if st.button("🔄 扫描模板", use_container_width=True):
        with st.spinner("扫描中..."):
            r = scan_and_build_templates(source_path=ts)
            if "error" in r:
                st.error(r["error"])
            else:
                st.success(f"✅ {r['templates_count']} 个模板文件")
                st.session_state.template_index = r.get("templates", {})
    idx = st.session_state.template_index
    if idx:
        st.caption(f"已扫描围岩: {list(idx.keys())}")

st.title("📋 分项完工计量 · 模板化批量生成")
st.caption("外环高速三期第4合同段 · 田头山隧道")

if not idx:
    st.info("👈 先在左侧输入路径，点击扫描")
    st.stop()

# ===================================================================
# ① 选围岩/线别
# ===================================================================
st.markdown("---")
st.header("① 选择围岩、线别")
c1, c2 = st.columns(2)
with c1:
    sel_rock = st.selectbox("围岩级别", list(idx.keys()))
with c2:
    ls = sorted(idx.get(sel_rock, {}).keys())
    sel_line = st.selectbox("线别", ls if ls else ["右线"])

items_data = idx.get(sel_rock, {}).get(sel_line, {})
all_item_names = list(items_data.keys())

# ===================================================================
# ② 选模板
# ===================================================================
st.markdown("---")
st.header("② 选择模板")

tpl_mode = st.radio("选模板方式", ["📦 批量勾选", "📂 从文件夹选择"], horizontal=True)

template_choice = {}

if tpl_mode == "📦 批量勾选":
    # ── 全选/取消全选按钮 ──
    c_all, c_none, _ = st.columns([1, 1, 4])
    with c_all:
        if st.button("✅ 全选", use_container_width=True):
            for item in all_item_names:
                st.session_state[f"ck_{item}"] = True
            st.rerun()
    with c_none:
        if st.button("⬜ 取消全选", use_container_width=True):
            for item in all_item_names:
                st.session_state[f"ck_{item}"] = False
            st.rerun()

    for item_name in all_item_names:
        versions = items_data.get(item_name, [])
        if not versions:
            continue
        n = len(versions)

        # 初始化状态（默认勾选）
        if f"ck_{item_name}" not in st.session_state:
            st.session_state[f"ck_{item_name}"] = True

        col_a, col_b = st.columns([1, 5])
        with col_a:
            checked = st.checkbox("", key=f"ck_{item_name}")
        with col_b:
            if not checked:
                st.markdown(f"<span style='color:#999'>~~{item_name}~~ （跳过）</span>",
                    unsafe_allow_html=True)
                continue
            st.markdown(f"<h4 style='margin:0'>🔹 {item_name}</h4>", unsafe_allow_html=True)
        if n == 1:
            v = versions[0]
            tpath = os.path.join(TEMPLATE_DIR, v["template_file"])
            template_choice[item_name] = tpath
            st.caption(f"来源: {v['source_seg']} | {v['mtime']}")
            st.caption(f"📄 {v.get('source_folder','')}")
        else:
            opts = [f"版{v['source_seg']} | {v['mtime']} | {v['size_kb']}KB" for v in versions]
            opt_map = {}
            for vi, v in enumerate(versions):
                opt_map[opts[vi]] = (os.path.join(TEMPLATE_DIR, v["template_file"]), v.get("source_folder",""))
            sel = st.selectbox("选模板版本", opts, index=n-1, key=f"tv_{item_name}")
            tpath, fname_full = opt_map[sel]
            template_choice[item_name] = tpath
            st.caption(f"📄 {fname_full}")

    st.caption(f"已选 {len(template_choice)} 个工序")

else:
    # ── 从文件夹选择 ──
    st.markdown("##### 选择模板文件（Excel格式）")
    st.caption("点击下方按钮，从文件夹中选取一个或多个 Excel 文件作为模板")

    uploaded_tpls = st.file_uploader(
        "选择模板文件",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="tpl_upload",
        label_visibility="collapsed",
    )

    if uploaded_tpls:
        # 保存到临时目录
        import tempfile
        tpl_dir = tempfile.mkdtemp(prefix="tpl_")
        st.success(f"已选择 {len(uploaded_tpls)} 个文件")

        for f in uploaded_tpls:
            # 保存到临时路径
            tmp_path = os.path.join(tpl_dir, f.name)
            with open(tmp_path, "wb") as fh:
                fh.write(f.getbuffer())

            item_name = _infer_item_name(f.name)
            template_choice[item_name] = tmp_path
            st.markdown(f"<h4 style='margin:0'>🔹 {item_name}</h4>", unsafe_allow_html=True)
            st.caption(f"📄 {f.name}")

        st.caption(f"已选 {len(template_choice)} 个模板文件")

        # 手动调整工序名映射
        with st.expander("✏️ 调整工序名称映射"):
            st.caption("如需修正自动识别的工序名，在此修改")
            corrected = {}
            for f in uploaded_tpls:
                orig_name = _infer_item_name(f.name)
                new_name = st.text_input(f"{f.name}", value=orig_name, key=f"rn_{f.name}")
                corrected[new_name] = template_choice.get(orig_name, "")
            if st.button("✅ 确认名称"):
                template_choice.clear()
                template_choice.update(corrected)
                st.success("已更新")

# ===================================================================
# ③ 输入编号表
# ===================================================================
st.markdown("---")
st.header("③ 输入编号表")

inp_mode = st.radio("输入方式", ["📄 上传文件", "⌨️ 直接粘贴"], horizontal=True)
st.code("""分项名称,分项工程编号
田头山隧道右线-洞身衬砌(K87+492.00～K87+675.00)(183.00m，S-Ⅴa)-衬砌砼,WH3A4-03-33-02-006
田头山隧道右线-洞身衬砌(K87+492.00～K87+675.00)(183.00m，S-Ⅴa)-仰拱砼,WH3A4-03-33-02-008""", language="csv")

df_raw = None

if inp_mode == "📄 上传文件":
    wh_file = st.file_uploader("选择编号表", type=["xlsx","xls","csv"], key="wh_upload")
    if wh_file is not None:
        try:
            df_raw = pd.read_excel(wh_file) if not wh_file.name.endswith(".csv") else pd.read_csv(wh_file)
        except Exception as e:
            st.error(f"读取失败: {e}")
else:
    paste_txt = st.text_area("粘贴数据（可从Excel直接复制两列）", height=150,
        placeholder="田头山隧道右线-洞身衬砌(K87+492.00～K87+675.00)(183.00m，S-Ⅴa)-衬砌砼,WH3A4-03-33-02-006\n田头山隧道右线-洞身衬砌(K87+492.00～K87+675.00)(183.00m，S-Ⅴa)-仰拱砼,WH3A4-03-33-02-008")
    if paste_txt.strip():
        try:
            import io
            # 自动识别分隔符和表头
            text = paste_txt.strip()
            # 检测分隔符（逗号、制表符、空格）
            if '\t' in text:
                sep = '\t'
            elif ',' in text:
                sep = ','
            else:
                sep = r'\s+'
            # 检测是否已有表头
            first_line = text.split('\n')[0].strip()
            has_header = '名称' in first_line or '编号' in first_line or '分项' in first_line
            df_raw = pd.read_csv(io.StringIO(text), sep=sep, header=0 if has_header else None)
            if not has_header:
                df_raw.columns = ['分项名称', '分项工程编号']
        except Exception as e:
            st.error(f"解析失败: {e}")

if df_raw is None:
    st.info("👆 上传或粘贴编号表后继续")
    st.stop()

name_col = code_col = None
for c in df_raw.columns:
    cl = c.strip().lower()
    if any(kw in cl for kw in ["分项名称","分项工程名称","项目名称","名称"]):
        name_col = c
    if any(kw in cl for kw in ["分项工程编号","分项编号","编号","编码","WH"]):
        code_col = c
if not name_col:
    name_col = st.selectbox("选分项名称列", df_raw.columns)
if not code_col:
    code_col = st.selectbox("选分项工程编号列", df_raw.columns)

ptn = re.compile(
    r'田头山隧道(左线|右线)'
    r'.*?洞身(?:开挖|衬砌)'
    r'\((?:[ZY])?K(\d+\+[\d.]+)～(?:[ZY])?K(\d+\+[\d.]+)\)'
    r'\(([\d.]+)m'
    r'[，,]\s*([SⅤⅣⅢa-zA-Z\-]+)\)'
    r'.*?-\s*([^-\s].*)$'
)

segments_info = {}
parse_err = []
for _, row in df_raw.iterrows():
    name = str(row.get(name_col, "")).strip()
    code = str(row.get(code_col, "")).strip()
    if not name or not code or code == "nan":
        continue
    m = ptn.search(name)
    if m:
        line, sc, ec, ln, rk_raw, si = m.groups()
        rk = rk_raw.replace("S-Va","S-Ⅴa").replace("S-Vb","S-Ⅴb").replace("S-Vc","S-Ⅴc")\
                   .replace("S-IVa","S-Ⅳa").replace("S-IVb","S-Ⅳb").replace("S-III","S-Ⅲ")
        sk = f"{line}|K{sc}|K{ec}"
        segments_info.setdefault(sk, {
            "line": line, "start": f"K{sc}", "end": f"K{ec}",
            "length": float(ln), "rock": rk, "items": {},
        })
        segments_info[sk]["items"][si] = code
    else:
        parse_err.append(name)

if not segments_info:
    st.error("未能解析出区段")
    with st.expander("未解析的名称"):
        for n in df_raw[name_col].head(10):
            st.code(n)
    st.stop()

st.success(f"解析到 {len(segments_info)} 个区段")
df_seg = pd.DataFrame([{
    "线别": s["line"], "起始": s["start"], "终点": s["end"],
    "长度(m)": s["length"], "围岩": s["rock"], "分项数": len(s["items"]),
} for s in segments_info.values()])
st.dataframe(df_seg, use_container_width=True, hide_index=True)

# ===================================================================
# ④ 输出 + 生成
# ===================================================================
st.markdown("---")
st.header("④ 输出设置")
st.warning(f"📂 文件将保存到: **{st.session_state.get('out_dir', DEFAULT_OUTPUT)}**")
st.caption("每个区段单独一个文件夹，命名与模板源一致")

out_col1, out_col2 = st.columns([3, 1])
with out_col1:
    output_dir = st.text_input("如需修改输出路径，在此粘贴", value=DEFAULT_OUTPUT, key="out_dir",
        help="正式保存到SMB请粘贴路径")
with out_col2:
    if st.button("🔍 打开输出目录", use_container_width=True):
        try:
            if os.name == 'nt':
                os.startfile(output_dir)
            elif sys.platform == 'darwin':
                os.system(f'open "{output_dir}"')
            else:
                os.system(f'xdg-open "{output_dir}"')
        except:
            pass

st.markdown("---")
st.header("⑤ 选择区段 → 生成")

sel_keys = st.multiselect(
    "勾选要生成的区段",
    list(segments_info.keys()),
    default=list(segments_info.keys()),
    format_func=lambda k: f"{segments_info[k]['start']}~{segments_info[k]['end']} ({segments_info[k]['rock']})",
)

if not sel_keys:
    st.warning("请至少选一个区段")
    st.stop()

# ── 检查是否有文件会被覆盖 ──
overwrite_files = []
for sk in sel_keys:
    seg = segments_info[sk]
    prefix = "ZK" if seg["line"] == "左线" else "K"
    folder = f"田头山隧道{seg['line']}-洞身衬砌({prefix}{_fmt(seg['start'])}～{prefix}{_fmt(seg['end'])})({seg['length']:.2f}m，{seg['rock']})"
    seg_dir = os.path.join(output_dir, folder)
    for si, code in seg["items"].items():
        if si in template_choice:
            wh = code
            fname = f"{wh}{folder}-{si}.xlsx"
            fpath = os.path.join(seg_dir, fname)
            if os.path.exists(fpath):
                overwrite_files.append(fpath)

if overwrite_files:
    st.warning(f"⚠️ **{len(overwrite_files)} 个文件已存在，将被覆盖**")
    with st.expander("查看将被覆盖的文件"):
        for f in overwrite_files[:20]:
            st.code(f"📄 {os.path.basename(f)}")
        if len(overwrite_files) > 20:
            st.caption(f"...还有 {len(overwrite_files)-20} 个")

    confirm_overwrite = st.checkbox("✅ 我确认覆盖已有文件", value=False)
    if not confirm_overwrite:
        st.info("☝️ 勾选确认后才可以生成")
        st.stop()

if st.button("⚡ 一键生成", type="primary", use_container_width=True):
    all_res = []
    pb = st.progress(0, text="生成中...")
    status = st.empty()

    for i, sk in enumerate(sel_keys):
        seg = segments_info[sk]
        status.info(f"[{i+1}/{len(sel_keys)}] {seg['start']}~{seg['end']}")
        valid = {si: code for si, code in seg["items"].items() if si in template_choice}
        valid_tpl = {si: template_choice[si] for si in valid}
        r = generate_segment_files(
            line=seg["line"], rock_grade=seg["rock"],
            start_chain=seg["start"], end_chain=seg["end"],
            length=seg["length"],
            sub_items=list(valid.keys()), wh_mapping=valid,
            template_choice=valid_tpl, output_dir=output_dir,
        )
        all_res.append(r)
        pb.progress((i + 1) / len(sel_keys))

    pb.empty(); status.empty()
    total_ok = sum(r["success_count"] for r in all_res)
    total_all = sum(len(r["results"]) for r in all_res)
    st.success(f"✅ {total_ok}/{total_all} 成功")

    for r in all_res:
        with st.expander(f"📁 {r['folder']}"):
            st.code(r["path"])
            for res in r["results"]:
                st.markdown(f"{res['status']} **{res['sub_item']}** {res.get('wh_number','')}")
                if res.get("reason"):
                    st.caption(f"  {res['reason']}")

    rows = [{"文件夹": r["folder"], "分项": res["sub_item"],
             "状态": res["status"], "WH编号": res.get("wh_number", "")}
            for r in all_res for res in r["results"]]
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False, engine='openpyxl')
    st.download_button("📥 下载生成报告", buf.getvalue(),
        f"计量生成报告_{datetime.now():%Y%m%d_%H%M%S}.xlsx")


st.caption(f"{datetime.now():%Y-%m-%d}")
