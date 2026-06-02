"""分项完工计量生成器 — Streamlit 多页面版"""

import os, sys, re, io
from datetime import datetime
import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from tools.generator_core import (
    scan_and_build_templates, load_template_index,
    generate_segment_files, ROCK_SUB_ITEMS,
    DEFAULT_TEMPLATE_SOURCE, DEFAULT_OUTPUT, TEMPLATE_DIR,
)

for k in ["template_index", "parsed_segments"]:
    if k not in st.session_state:
        st.session_state[k] = None
if "template_source" not in st.session_state:
    st.session_state.template_source = DEFAULT_TEMPLATE_SOURCE
if "ck_all" not in st.session_state:
    st.session_state.ck_all = True


def _fmt(ch):
    return ch.replace("K", "", 1) if ch.startswith("K") else ch


def _infer_name(filename):
    for kw in ["洞身开挖", "喷射砼支护", "锚杆支护", "钢筋网支护", "钢架支护",
               "衬砌钢筋加工及安装", "衬砌砼", "仰拱钢筋加工及安装",
               "仰拱砼", "仰拱回填", "超前小导管支护", "管棚"]:
        if kw in filename:
            return kw
    name = os.path.splitext(filename)[0]
    name = re.sub(r'^WH3A4[\d\-]+', '', name)
    return name[:20] if name else filename


st.title("📋 分项完工计量 · 模板化批量生成")
st.caption("外环高速三期第4合同段 · 田头山隧道")

# 模板源
with st.expander("📂 模板目录", expanded=not st.session_state.get("template_index")):
    # 第一行：直接输入
    col_inp, col_scan = st.columns([4, 1])
    with col_inp:
        ts = st.text_input("直接输入路径", value=st.session_state.template_source, key="ts_inp",
            placeholder="如 H:\正式版本 或 \\\\192.168.1.28\H\...")
    with col_scan:
        st.write("")
        if st.button("🔄 扫描", use_container_width=True):
            with st.spinner("扫描中..."):
                r = scan_and_build_templates(source_path=ts)
                if "error" in r:
                    st.error(r["error"])
                else:
                    st.success(f"✅ {r['templates_count']} 个模板")
                    st.session_state.template_index = r.get("templates", {})
    st.session_state.template_source = ts

    # 第二行：导航选择
    if ts and os.path.isdir(ts):
        st.markdown(f"**当前: {ts}**")
        # 获取子目录列表 + 上级目录
        nav_dirs = []
        parent = os.path.dirname(os.path.normpath(ts))
        if parent and parent != ts:
            nav_dirs.append((".. [返回上级]", parent))
        try:
            subs = sorted([d for d in os.listdir(ts)
                if os.path.isdir(os.path.join(ts, d)) and not d.startswith('.')
                and not d.startswith('~') and not d.startswith('$')])
            for d in subs[:30]:
                nav_dirs.append((f"📁 {d}", os.path.join(ts, d)))
        except:
            pass

        if nav_dirs:
            sel_nav = st.selectbox("导航选择", [d[0] for d in nav_dirs], key="nav_sel")
            col_go, col_cur = st.columns([1, 3])
            with col_go:
                if st.button("进入", use_container_width=True):
                    for label, path in nav_dirs:
                        if label == sel_nav:
                            st.session_state.template_source = path
                            st.rerun()
            with col_cur:
                st.caption("选择子目录后点「进入」")

idx = st.session_state.template_index
if not idx:
    st.info("👈 先设置模板扫描目录，点击扫描")
    st.stop()

st.caption(f"已扫描围岩: {list(idx.keys())}")

# ① 选围岩/线别
st.markdown("---")
st.subheader("① 选择围岩、线别")
c1, c2 = st.columns(2)
with c1:
    sel_rock = st.selectbox("围岩级别", list(idx.keys()), key="page_rock")
with c2:
    ls = sorted(idx.get(sel_rock, {}).keys())
    sel_line = st.selectbox("线别", ls if ls else ["右线"], key="page_line")

items_data = idx.get(sel_rock, {}).get(sel_line, {})
all_item_names = list(items_data.keys())

# ② 勾工序+选模板
st.markdown("---")
st.subheader("② 勾选工序 → 选模板版本")

tpl_mode = st.radio("选模板方式", ["📦 批量勾选", "📂 从文件夹选择"], horizontal=True, key="page_mode")
template_choice = {}

if tpl_mode == "📦 批量勾选":
    # 多选
    sel_items = st.multiselect("勾选本次要做的工序", all_item_names, default=all_item_names,
        key="page_sel_procs")
    st.caption(f"已选 {len(sel_items)} 个工序")

    for item_name in all_item_names:
        versions = items_data.get(item_name, [])
        if not versions:
            continue
        if item_name not in sel_items:
            st.markdown(f"<span style='color:#999'>~~{item_name}~~ （跳过）</span>", unsafe_allow_html=True)
            continue

        n = len(versions)
        st.markdown(f"**🔹 {item_name}**")
        if n == 1:
            v = versions[0]
            tpath = os.path.join(TEMPLATE_DIR, v["template_file"])
            template_choice[item_name] = tpath
            st.caption(f"来源: {v['source_seg']} | {v['mtime']}")
        else:
            opts = [f"版{v['source_seg']} | {v['mtime']} | {v['size_kb']}KB" for v in versions]
            opt_map = {opts[i]: os.path.join(TEMPLATE_DIR, versions[i]["template_file"]) for i in range(n)}
            sel = st.selectbox("选版本", opts, index=n-1, key=f"pg_tv_{item_name}")
            template_choice[item_name] = opt_map[sel]
else:
    uploaded_tpls = st.file_uploader("选择模板文件", type=["xlsx","xls","xlsm"],
        accept_multiple_files=True, key="pg_tpl_up")
    if uploaded_tpls:
        import tempfile
        tdir = tempfile.mkdtemp(prefix="tpl_")
        for f in uploaded_tpls:
            tmp = os.path.join(tdir, f.name)
            with open(tmp, "wb") as fh:
                fh.write(f.getbuffer())
            item_name = _infer_name(f.name)
            template_choice[item_name] = tmp
            st.markdown(f"**🔹 {item_name}**")
            st.caption(f"📄 {f.name}")

# ③ 上传编号表
st.markdown("---")
st.subheader("③ 输入编号表")

inp_mode = st.radio("输入方式", ["📄 上传文件", "⌨️ 直接粘贴"], horizontal=True, key="pg_inp")
st.code("""分项名称,分项工程编号
田头山隧道右线-洞身衬砌(K87+492.00～K87+675.00)(183.00m，S-Ⅴa)-衬砌砼,WH3A4-03-33-02-006""", language="csv")

df_raw = None
if inp_mode == "📄 上传文件":
    wh_file = st.file_uploader("选择编号表", type=["xlsx","xls","csv"], key="pg_wh")
    if wh_file is not None:
        try:
            df_raw = pd.read_excel(wh_file) if not wh_file.name.endswith(".csv") else pd.read_csv(wh_file)
        except Exception as e:
            st.error(f"读取失败: {e}")
else:
    paste_txt = st.text_area("粘贴CSV", height=120, key="pg_paste",
        placeholder="田头山隧道右线-洞身衬砌(K87+492.00～K87+675.00)(183.00m，S-Ⅴa)-衬砌砼,WH3A4-03-33-02-006")
    if paste_txt.strip():
        try:
            text = paste_txt.strip()
            sep = '\t' if '\t' in text else ',' if ',' in text else r'\s+'
            first_line = text.split('\n')[0].strip()
            has_header = '名称' in first_line or '编号' in first_line
            df_raw = pd.read_csv(io.StringIO(text), sep=sep, header=0 if has_header else None)
            if not has_header:
                df_raw.columns = ['分项名称', '分项工程编号']
        except Exception as e:
            st.error(f"解析失败: {e}")

if df_raw is None:
    st.info("👆 上传或粘贴编号表后继续")
    st.stop()

# 识别列
name_col = code_col = None
for c in df_raw.columns:
    cl = c.strip().lower()
    if any(kw in cl for kw in ["分项名称","分项工程名称","项目名称","名称"]):
        name_col = c
    if any(kw in cl for kw in ["分项工程编号","分项编号","编号","编码","WH"]):
        code_col = c
if not name_col:
    name_col = st.selectbox("选分项名称列", df_raw.columns, key="pg_nc")
if not code_col:
    code_col = st.selectbox("选分项工程编号列", df_raw.columns, key="pg_cc")

# 解析
ptn = re.compile(
    r'田头山隧道(左线|右线).*?洞身(?:开挖|衬砌)'
    r'\((?:[ZY])?K(\d+\+[\d.]+)～(?:[ZY])?K(\d+\+[\d.]+)\)'
    r'\(([\d.]+)m[，,]\s*([SⅤⅣⅢa-zA-Z\-]+)\)'
    r'.*?-\s*([^-\s].*)$'
)

segments_info = {}
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
        segments_info.setdefault(sk, {"line": line, "start": f"K{sc}", "end": f"K{ec}",
            "length": float(ln), "rock": rk, "items": {}})
        segments_info[sk]["items"][si] = code

if not segments_info:
    st.error("未能解析出区段")
    st.stop()

st.success(f"解析到 {len(segments_info)} 个区段")
df_seg = pd.DataFrame([{"线别": s["line"], "起始": s["start"], "终点": s["end"],
    "长度(m)": s["length"], "围岩": s["rock"], "分项数": len(s["items"])} for s in segments_info.values()])
st.dataframe(df_seg, use_container_width=True, hide_index=True)

# ④ 输出设置
st.markdown("---")
st.subheader("④ 输出设置")
st.caption("默认保存到模板源目录，可修改")
output_dir = st.text_input("输出目录", value=st.session_state.template_source, key="pg_out")

# ⑤ 模板映射检查
st.markdown("---")
st.subheader("⑤ 模板映射检查")

# 收集所有区段需要的分项
all_needed = set()
for sk, seg in segments_info.items():
    for si in seg["items"]:
        all_needed.add(si)

# 收集已选的模板
all_templates = set(template_choice.keys())

# 匹配情况
matched = all_needed & all_templates
unmatched_needed = all_needed - all_templates  # 区段要但没选模板
unused_templates = all_templates - all_needed  # 选了模板但区段没有

st.markdown(f"**区段需要的分项:** {len(all_needed)} 个")
st.markdown(f"**已选模板:** {len(all_templates)} 个")
st.markdown(f"**✅ 匹配成功:** {len(matched)} 个 — {'、'.join(matched) if matched else '无'}")

if unmatched_needed:
    st.warning(f"⚠️ **以下分项区段需要但未选模板:** {'、'.join(unmatched_needed)}")
if unused_templates:
    st.info(f"ℹ️ **以下已选模板在区段中未用到**（如需使用请在下方手动映射）: {'、'.join(unused_templates)}")

# 手动映射：允许把未匹配的模板指定给区段的分项
manual_map = {}
if unused_templates:
    st.markdown("##### 手动映射（把模板指定给某个分项）")
    for ut in unused_templates:
        target = st.selectbox(
            f"模板「{ut}」→ 用作",
            ["（跳过）"] + sorted(all_needed),
            key=f"mm_{ut}",
        )
        if target and target != "（跳过）":
            manual_map[target] = template_choice[ut]

# 合并映射：自动匹配 + 手动指定
final_tpl_map = {si: template_choice[si] for si in matched}
final_tpl_map.update(manual_map)

st.caption(f"最终可用映射: {len(final_tpl_map)} 个分项")

# ⑥ 选择区段 → 生成
st.markdown("---")
st.subheader("⑥ 选择区段 → 生成")

sel_keys = st.multiselect("勾选要生成的区段", list(segments_info.keys()),
    default=list(segments_info.keys()),
    format_func=lambda k: f"{segments_info[k]['start']}~{segments_info[k]['end']} ({segments_info[k]['rock']})",
    key="pg_segs")

if not sel_keys:
    st.warning("请至少选一个区段")
    st.stop()

# 覆盖检查
overwrite_count = 0
for sk in sel_keys:
    seg = segments_info[sk]
    prefix = "ZK" if seg["line"] == "左线" else "K"
    folder = f"田头山隧道{seg['line']}-洞身衬砌({prefix}{_fmt(seg['start'])}～{prefix}{_fmt(seg['end'])})({seg['length']:.2f}m，{seg['rock']})"
    seg_dir = os.path.join(output_dir, folder)
    for si, code in seg["items"].items():
        if si in final_tpl_map:
            fpath = os.path.join(seg_dir, f"{code}{folder}-{si}.xlsx")
            if os.path.exists(fpath):
                overwrite_count += 1

if overwrite_count > 0:
    st.warning(f"⚠️ {overwrite_count} 个文件已存在，将被覆盖")
    if not st.checkbox("✅ 确认覆盖已有文件", key="pg_confirm"):
        st.stop()

if st.button("⚡ 一键生成", type="primary", use_container_width=True, key="pg_gen"):
    all_res = []
    pb = st.progress(0, text="生成中...")
    status = st.empty()

    for i, sk in enumerate(sel_keys):
        seg = segments_info[sk]
        status.info(f"[{i+1}/{len(sel_keys)}] {seg['start']}~{seg['end']}")
        valid = {si: code for si, code in seg["items"].items() if si in final_tpl_map}
        valid_tpl = {si: final_tpl_map[si] for si in valid}
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
        with st.expander(f"📁 {r['folder']} ({r['success_count']}/{len(r['results'])})"):
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


