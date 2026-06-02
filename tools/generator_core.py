"""
分项完工计量 - 模板化批量生成 核心引擎 V2

原理：用已批复的Excel文件作模板，仅替换封面2个关键单元格
（分项名称 B12 + 编号 H15），其余全由公式级联自动计算。

核心改动 V2：
  - 分项编号由用户提供映射表（分项名称→WH编号），不自动生成
  - 支持同区段不同分项使用不同category编号
  - 更多分项类型（钢筋网、钢架、仰拱钢筋等）
"""

import os, re, shutil, json, sys
from datetime import datetime
from pathlib import Path
import openpyxl

# ── 常量 ──────────────────────────────────────────────────────────

CONTRACT_PREFIX = "WH3A4"
LINE_CODE = {"左线": "02", "右线": "03"}

# 各围岩级别全套分项（按工序顺序）
ROCK_SUB_ITEMS = {
    "S-Ⅴa": ["洞身开挖",
              "喷射砼支护", "锚杆支护", "钢筋网支护", "钢架支护",
              "衬砌钢筋加工及安装", "衬砌砼",
              "仰拱钢筋加工及安装", "仰拱砼", "仰拱回填",
              "超前小导管支护"],
    "S-Ⅴb": ["洞身开挖",
              "喷射砼支护", "锚杆支护", "钢筋网支护", "钢架支护",
              "衬砌钢筋加工及安装", "衬砌砼",
              "仰拱钢筋加工及安装", "仰拱砼", "仰拱回填",
              "超前小导管支护"],
    "S-Ⅴc": ["洞身开挖",
              "喷射砼支护", "锚杆支护", "钢筋网支护", "钢架支护",
              "衬砌钢筋加工及安装", "衬砌砼",
              "仰拱钢筋加工及安装", "仰拱砼", "仰拱回填",
              "超前小导管支护"],
    "S-Ⅳa": ["洞身开挖",
              "喷射砼支护", "锚杆支护", "钢筋网支护", "钢架支护",
              "衬砌钢筋加工及安装", "衬砌砼",
              "仰拱钢筋加工及安装", "仰拱砼", "仰拱回填"],
    "S-Ⅳb": ["洞身开挖",
              "喷射砼支护", "锚杆支护", "钢筋网支护", "钢架支护",
              "衬砌钢筋加工及安装", "衬砌砼",
              "仰拱钢筋加工及安装", "仰拱砼", "仰拱回填"],
    "S-Ⅲ":  ["洞身开挖",
              "喷射砼支护", "锚杆支护", "钢筋网支护", "钢架支护",
              "衬砌钢筋加工及安装", "衬砌砼"],
}

# 常用的编号分组方案（用户可自定义）
# 格式：{分组名: {"category": "XX", "sub_seg": "YY", "items": [分项名, ...]}}
DEFAULT_WH_SCHEMES = {
    "S-Ⅴa-右线-开挖类": {"category": "17", "sub_seg": "02", "items": ["洞身开挖"]},
    "S-Ⅴa-右线-衬砌类": {"category": "33", "sub_seg": "02",
        "items": ["喷射砼支护", "锚杆支护", "钢筋网支护", "钢架支护",
                  "衬砌钢筋加工及安装", "衬砌砼",
                  "仰拱钢筋加工及安装", "仰拱砼", "仰拱回填",
                  "超前小导管支护"]},
}

# ── 路径配置（默认值，可在界面修改） ─────────────────────────────
# Windows 用 H 盘映射路径，macOS 用 /Volumes/
if os.name == 'nt':  # Windows
    DEFAULT_TEMPLATE_SOURCE = r"H:\1深圳龙华片区\外环高速三期\合约\分项完工计量\正式版本"
else:  # macOS / Linux
    DEFAULT_TEMPLATE_SOURCE = "/Volumes/192.168.1.28/1深圳龙华片区/外环高速三期/合约/分项完工计量/正式版本"
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "templates")
TEMPLATE_INDEX = os.path.join(TEMPLATE_DIR, "template_index.json")
DEFAULT_OUTPUT = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "generated_test")


# ==================================================================
# 1. 模板管理（只读扫描，复制到本地）
# ==================================================================

def scan_and_build_templates(source_path=None):
    """从已批复文件扫描并建立模板库（只读，保留全部版本供用户选择）

    参数:
        source_path: 模板源目录，默认为 DEFAULT_TEMPLATE_SOURCE
    """
    if source_path is None:
        source_path = DEFAULT_TEMPLATE_SOURCE
    if not os.path.exists(source_path):
        return {"error": f"模板源目录不可访问: {source_path}"}

    os.makedirs(TEMPLATE_DIR, exist_ok=True)

    # 扫描所有文件，按 (线别, 围岩, 分项) 分组保留全部源文件
    all_sources = {}  # {(line, rock, item): [{"file": src_path, "source_seg": src_seg, "size": ...}]}

    for root, dirs, files in os.walk(source_path):
        if "草稿" in root or ".rar" in root:
            continue
        for f in files:
            if not f.endswith('.xlsx') or f.startswith('~$'):
                continue
            fp = os.path.join(root, f)
            rock = _extract_rock(os.path.basename(root) + f)
            if not rock:
                continue

            line = "右线" if "右线" in f or "右线" in root else \
                   "左线" if "左线" in f or "左线" in root else "右线"
            sub_item = _extract_sub_item(f)
            if not sub_item:
                continue

            key = (line, rock, sub_item)

            # 从文件夹名提取来源区段信息（用于显示给用户选择）
            folder_name = os.path.basename(root)
            # 提取桩号
            ch_match = re.search(r'K(\d+\+[\d.]*～K\d+\+[\d.]*)', folder_name)
            seg_info = ch_match.group(1) if ch_match else folder_name[:40]

            # 获取文件修改时间
            mtime = os.path.getmtime(fp)
            mtime_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d")

            all_sources.setdefault(key, []).append({
                "source_path": fp,
                "source_seg": seg_info,
                "source_folder": folder_name,
                "mtime": mtime_str,
                "size_kb": os.path.getsize(fp) // 1024,
            })

    # 复制全部源文件到本地模板目录，每个文件独立命名
    idx = {}
    template_file_map = {}  # 本地路径 → 源信息

    for (line, rock, sub_item), sources in all_sources.items():
        rk = rock.replace("/", "_")
        idx.setdefault(rk, {}).setdefault(line, {})[sub_item] = []

        for i, src in enumerate(sources):
            # 生成本地唯一文件名：线别_围岩_分项_序号.xlsx
            tname = f"{line}_{rock}_{sub_item}_{i+1}.xlsx"
            tpath = os.path.join(TEMPLATE_DIR, tname)
            try:
                shutil.copy2(src["source_path"], tpath)
                entry = {
                    "template_file": tname,
                    "source_seg": src["source_seg"],
                    "source_folder": src["source_folder"],
                    "mtime": src["mtime"],
                    "size_kb": src["size_kb"],
                    "id": i,
                }
                idx[rk][line][sub_item].append(entry)
                template_file_map[tpath] = src
            except Exception:
                pass

    with open(TEMPLATE_INDEX, 'w', encoding='utf-8') as f:
        json.dump(idx, f, ensure_ascii=False, indent=2)

    total_templates = sum(
        len(versions)
        for rock in idx.values()
        for line in rock.values()
        for versions in line.values()
    )

    return {
        "templates_count": total_templates,
        "templates": idx,
        "rock_grades": list(idx.keys()),
    }


def load_template_index():
    if not os.path.exists(TEMPLATE_INDEX):
        return {}
    with open(TEMPLATE_INDEX, 'r', encoding='utf-8') as f:
        return json.load(f)


def _extract_rock(text):
    for r in ["S-Ⅴa", "S-Ⅴb", "S-Ⅴc", "S-Ⅳa", "S-Ⅳb", "S-Ⅲ",
              "S-Va", "S-Vb", "S-Vc", "S-IVa", "S-IVb", "S-III",
              "Ⅴ级", "Ⅳ级", "Ⅲ级"]:
        if r in text:
            return r.replace("S-Va","S-Ⅴa").replace("S-Vb","S-Ⅴb").replace("S-Vc","S-Ⅴc")\
                     .replace("S-IVa","S-Ⅳa").replace("S-IVb","S-Ⅳb").replace("S-III","S-Ⅲ")
    return ""


def _extract_sub_item(text):
    items = ["衬砌钢筋加工及安装", "超前小导管支护", "衬砌防水层",
             "衬砌砼", "仰拱砼", "仰拱回填", "锚杆支护", "锚杆",
             "洞身开挖", "管棚", "止水带", "喷射砼支护",
             "钢筋网支护", "钢架支护", "仰拱钢筋加工及安装",
             "明洞钢筋加工及安装", "明洞止水带", "仰拱钢筋加工及安装",
             "超前小导管"]
    for item in items:
        if item in text:
            return item
    return ""


# ==================================================================
# 2. 生成引擎（接收用户提供的编号映射）
# ==================================================================

def _detect_prefix(line):
    """按线别返回桩号前缀 — 左线ZK，右线K"""
    return "ZK" if line == "左线" else "K"


def generate_segment_files(
    line, rock_grade, start_chain, end_chain, length,
    sub_items=None,
    wh_mapping=None,       # {子项名: 完整WH编号}
    template_choice=None,  # {子项名: 模板文件路径}  ← 用户选择的模板
    output_dir=None,
):
    """
    为一个区段生成分项完工计量文件

    参数:
      wh_mapping: { "洞身开挖": "WH3A4-03-17-02-001", ... }
      template_choice: { "洞身开挖": "/path/to/template.xlsx", ... }
                       用户从模板库中选择的具体模板文件
    """
    if sub_items is None:
        sub_items = ROCK_SUB_ITEMS.get(rock_grade, [])
    if output_dir is None:
        output_dir = DEFAULT_OUTPUT
    if wh_mapping is None:
        wh_mapping = {}
    if template_choice is None:
        template_choice = {}

    start_disp = _fmt_ch(start_chain)
    end_disp = _fmt_ch(end_chain)
    start_num = _ch_num(start_chain)
    end_num = _ch_num(end_chain)

    ch_prefix = _detect_prefix(line)

    folder_name = _folder_name(line, rock_grade, start_disp, end_disp, length, ch_prefix)
    seg_dir = os.path.join(output_dir, folder_name)
    os.makedirs(seg_dir, exist_ok=True)

    results = []
    for sub_item in sub_items:
        wh_number = wh_mapping.get(sub_item, "")
        if not wh_number:
            results.append({
                "sub_item": sub_item,
                "status": "⏭️ 跳过",
                "reason": "未设置编号"
            })
            continue

        # 使用用户选择的模板文件
        tpath = template_choice.get(sub_item, "")
        if not tpath or not os.path.exists(tpath):
            results.append({
                "sub_item": sub_item,
                "status": "⏭️ 跳过",
                "reason": "未选择模板"
            })
            continue

        fname = _file_name(line, rock_grade, sub_item,
                           start_disp, end_disp, length, wh_number, ch_prefix)
        fpath = os.path.join(seg_dir, fname)

        try:
            _gen_one(tpath, fpath, line, rock_grade, sub_item,
                     start_disp, end_disp, length, start_num, end_num, wh_number, ch_prefix)
            results.append({
                "sub_item": sub_item,
                "status": "✅ 成功",
                "wh_number": wh_number,
                "file": fname,
            })
        except Exception as e:
            results.append({
                "sub_item": sub_item,
                "status": "❌ 失败",
                "reason": str(e),
            })

    return {"folder": folder_name, "path": seg_dir,
            "results": results,
            "success_count": sum(1 for r in results if r["status"] == "✅ 成功")}


def _gen_one(template_path, output_path, line, rock, sub_item,
             start_disp, end_disp, length, start_num, end_num, wh_number, ch_prefix="K"):
    """复制模板 → 替换封面 + 全表替换桩号/编号

    用 zipfile 直接操作 xlsx 内部的 XML，纯 Python，跨平台兼容。
    """
    import shutil, os, re, zipfile
    seg_name = f"田头山隧道{line}-洞身衬砌({ch_prefix}{start_disp}～{ch_prefix}{end_disp})({length:.2f}m，{rock})"
    name_value = f"分项工程名称：{seg_name}-{sub_item}"

    shutil.copy2(template_path, output_path)
    with zipfile.ZipFile(output_path, 'r') as zin:
        all_files = {n: zin.read(n) for n in zin.namelist()}

    # 新桩号数值
    nsv = float(start_disp.split('+')[0]) * 1000 + float(start_disp.split('+')[1])
    nev = float(end_disp.split('+')[0]) * 1000 + float(end_disp.split('+')[1])
    new_s = str(int(nsv)) if nsv == int(nsv) else str(nsv)
    new_e = str(int(nev)) if nev == int(nev) else str(nev)

    # ── 找封面 sheet1 的旧桩号 ──
    old_ch = None
    old_wh = None
    s1 = all_files.get('xl/worksheets/sheet1.xml', b'').decode()
    m = re.search(r'<c[^>]*r="B12"[^>]*t="s"[^>]*?>.*?<v>(\d+)</v>', s1)
    if m:
        ss = all_files.get('xl/sharedStrings.xml', b'')
        if ss:
            import xml.etree.ElementTree as ET
            ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            root = ET.fromstring(ss)
            sis = root.findall(f'{ns}si')
            idx = int(m.group(1))
            if idx < len(sis):
                t = sis[idx].find(f'{ns}t')
                txt = t.text if t is not None else ''
                if not txt:
                    txt = ''.join(r.find(f'{ns}t').text or '' for r in sis[idx].findall(f'{ns}r') if r.find(f'{ns}t') is not None)
                cm = re.search(r'[ZY]?K\d+\+[\d.]+～[ZY]?K\d+\+[\d.]+', txt)
                if cm: old_ch = cm.group()
                wm = re.search(r'WH3A4[\d-]+', txt)
                if wm: old_wh = wm.group()

    new_ch_str = f"{ch_prefix}{start_disp}～{ch_prefix}{end_disp}"

    # ── 全表替换字符串 ──
    for name in all_files:
        if not name.endswith('.xml'): continue
        text = all_files[name].decode('utf-8')
        if old_ch: text = text.replace(old_ch, new_ch_str)
        if old_wh and old_wh != wh_number: text = text.replace(old_wh, wh_number)
        all_files[name] = text.encode('utf-8')

    # ── 按表头列替换数值 ──
    # 找"起点桩号""终点桩号"在 sharedStrings 中的索引
    ss_xml = all_files.get('xl/sharedStrings.xml', b'')
    start_idx = end_idx = None
    if ss_xml:
        import xml.etree.ElementTree as ET
        ns_s = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        root = ET.fromstring(ss_xml)
        for i, si in enumerate(root.findall(f'{ns_s}si')):
            t = si.find(f'{ns_s}t')
            txt = t.text if t is not None else ''
            if '起点桩号' in txt: start_idx = i
            if '终点桩号' in txt: end_idx = i

    # ── 找表头列，替换数值 ──
    found_cols = {}  # col_letter → 'start' or 'end'

    for name in all_files:
        if 'worksheets/sheet' not in name or not name.endswith('.xml'): continue
        text = all_files[name].decode('utf-8')

        # 方法1: 通过 shared strings 索引找
        for idx, typ in [(start_idx, 'start'), (end_idx, 'end')]:
            if idx is not None:
                for m in re.finditer(rf'<c[^>]*r="([A-Z]+)\d+"[^>]*t="s"[^>]*?><v>{idx}</v>', text):
                    found_cols[m.group(1)] = typ

        # 方法2: 直接在文本中搜表头文字
        for m in re.finditer(r'<t>(起[^\s<]*[桩位].*?桩号|起点桩号)</t>', text):
            found_cols[m.group(1)] = 'start'
        for m in re.finditer(r'<t>(终[^\s<]*[桩位].*?桩号|终点桩号)</t>', text):
            found_cols[m.group(1)] = 'end'

        # 替换已找到的列
        if found_cols:
            for cl, typ in found_cols.items():
                nv = new_s if typ == 'start' else new_e
                def _rn(m2, nv2=nv):
                    x = m2.group(0)
                    if 't="s"' in x: return x
                    return m2.group(1) + nv2 + m2.group(2)
                text = re.sub(rf'(<c[^>]*r="{cl}\d+"[^>]*?>.*?<v>)[\d.]+(</v>)', _rn, text)

        all_files[name] = text.encode('utf-8')

    # 方法3: 如果以上都没找到，硬编码 D列=起点，F列=终点（95%的模板都是这）
    if not found_cols:
        import sys as _sys
        _sys.stderr.write("⚠️ 未找到表头列，使用默认D/F列\n")
        for name in all_files:
            if 'worksheets/sheet' not in name or not name.endswith('.xml'): continue
            text = all_files[name].decode('utf-8')
            for cl, nv in [('D', new_s), ('F', new_e)]:
                def _rn3(m2, nv3=nv):
                    x = m2.group(0)
                    if 't="s"' in x: return x
                    return m2.group(1) + nv3 + m2.group(2)
                text = re.sub(rf'(<c[^>]*r="{cl}\d+"[^>]*?>.*?<v>)[\d.]+(</v>)', _rn3, text)
            all_files[name] = text.encode('utf-8')

    # ── 封面 B12/H15 改为 inlineStr ──
    s1_path = next((n for n in all_files if 'worksheets/sheet' in n and n.endswith('.xml')), None)
    if s1_path:
        xml = all_files[s1_path].decode('utf-8')
        for ref, nv in [('B12', name_value), ('H15', wh_number)]:
            ev = _xml_escape(nv)
            m = re.search(rf'<c[^>]*?r="{ref}"[^>]*?>.*?</c>', xml, re.DOTALL)
            if m:
                nc = m.group().replace('t="s"', 't="inlineStr"', 1)
                nc = re.sub(r'>.*?</c>', f'><is><t>{ev}</t></is></c>', nc, count=1, flags=re.DOTALL)
                xml = xml.replace(m.group(), nc, 1)
        all_files[s1_path] = xml.encode('utf-8')

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)


def _gen_one_zip(template_path, output_path, line, rock, sub_item,
                 start_disp, end_disp, length, start_num, end_num, wh_number, ch_prefix,
                 seg_name, name_value):
    """zipfile + ElementTree 方案（macOS / win32com 不可用时）"""
    import zipfile, re, xml.etree.ElementTree as ET
    old_chainage = None
    old_wh = None
    ss_texts = []
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    with zipfile.ZipFile(output_path, 'r') as zin:
        all_files = {n: zin.read(n) for n in zin.namelist()}

    ss_data = all_files.get('xl/sharedStrings.xml', b'')
    if ss_data:
        sr = ET.fromstring(ss_data)
        for si in sr.findall(f'{{{ns}}}si'):
            t = si.find(f'{{{ns}}}t')
            if t is not None and t.text:
                ss_texts.append(t.text)
            else:
                parts = [r.find(f'{{{ns}}}t').text or '' for r in si.findall(f'{{{ns}}}r') if r.find(f'{{{ns}}}t') is not None and r.find(f'{{{ns}}}t').text]
                ss_texts.append(''.join(parts))

    s1 = all_files.get('xl/worksheets/sheet1.xml', b'').decode()
    m = re.search(r'<c[^>]*r="B12"[^>]*t="s"[^>]*?>.*?<v>(\d+)</v>', s1)
    if m and ss_texts:
        old_b12 = ss_texts[int(m.group(1))]
        cm = re.search(r'([ZY]?K\d+\+[\d.]+～[ZY]?K\d+\+[\d.]+)', old_b12)
        if cm: old_chainage = cm.group(1)
        wm = re.search(r'WH3A4[\d-]+', old_b12)
        if wm: old_wh = wm.group(0)

    new_chainage = f"{ch_prefix}{start_disp}～{ch_prefix}{end_disp}"
    new_s = float(start_disp.split('+')[0]) * 1000 + float(start_disp.split('+')[1])
    new_e = float(end_disp.split('+')[0]) * 1000 + float(end_disp.split('+')[1])

    for name in all_files:
        if not name.endswith('.xml'): continue
        text = all_files[name].decode('utf-8')
        if old_chainage: text = text.replace(old_chainage, new_chainage)
        if old_wh and old_wh != wh_number: text = text.replace(old_wh, wh_number)
        all_files[name] = text.encode('utf-8')

    # ElementTree 解析，按表头列替换
    col_map = {}
    for name in all_files:
        if 'worksheets/sheet' not in name or not name.endswith('.xml'): continue
        try:
            root = ET.fromstring(all_files[name])
            sd = root.find(f'{{{ns}}}sheetData')
            if sd is None: continue
            for row in sd.findall(f'{{{ns}}}row'):
                for c in row.findall(f'{{{ns}}}c'):
                    ref = c.get('r', '')
                    v_el = c.find(f'{{{ns}}}v')
                    if v_el is not None and v_el.text and c.get('t') == 's' and v_el.text.isdigit():
                        idx = int(v_el.text)
                        if idx < len(ss_texts):
                            col_letter = re.match(r'([A-Z]+)', ref).group(1)
                            if '起点桩号' in ss_texts[idx]: col_map[col_letter] = 'start'
                            elif '终点桩号' in ss_texts[idx]: col_map[col_letter] = 'end'
        except: pass

    if col_map:
        for name in all_files:
            if 'worksheets/sheet' not in name or not name.endswith('.xml'): continue
            text = all_files[name].decode('utf-8')
            for cl, typ in col_map.items():
                nv = str(int(new_s)) if new_s == int(new_s) else str(new_s)
                nv = nv if typ == 'start' else str(int(new_e)) if new_e == int(new_e) else str(new_e)
                def _rpl(m, v=nv):
                    x = m.group(0)
                    if 't="s"' in x: return x  # 跳过公式和共享字符串
                    return m.group(1) + v + m.group(2)
                text = re.sub(rf'(<c[^>]*r="{cl}\d+"[^>]*?>.*?<v>)[\d.]+(</v>)', _rpl, text)
            all_files[name] = text.encode('utf-8')

    # 封面 inlineStr
    s1_path = next((n for n in all_files if 'worksheets/sheet' in n and n.endswith('.xml')), None)
    if s1_path:
        xml = all_files[s1_path].decode('utf-8')
        for ref, nv in [('B12', name_value), ('H15', wh_number)]:
            ev = _xml_escape(nv)
            m = re.search(rf'<c[^>]*?r="{ref}"[^>]*?>.*?</c>', xml, re.DOTALL)
            if m:
                nc = m.group().replace('t="s"', 't="inlineStr"', 1)
                nc = re.sub(r'>.*?</c>', f'><is><t>{ev}</t></is></c>', nc, count=1, flags=re.DOTALL)
                xml = xml.replace(m.group(), nc, 1)
        all_files[s1_path] = xml.encode('utf-8')

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)


def _parse_shared_strings(xml_bytes):
    """从 sharedStrings.xml 提取所有文本字符串"""
    import xml.etree.ElementTree as ET
    root = ET.fromstring(xml_bytes)
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    texts = []
    for si in root.findall(f'{ns}si'):
        # 普通文本
        t_node = si.find(f'{ns}t')
        if t_node is not None and t_node.text:
            texts.append(t_node.text)
        else:
            # 富文本（多个 <r> 子元素）
            parts = []
            for r_elem in si.findall(f'{ns}r'):
                t_elem = r_elem.find(f'{ns}t')
                if t_elem is not None and t_elem.text:
                    parts.append(t_elem.text)
            texts.append(''.join(parts))
    return texts


def _xml_escape(text):
    """XML 转义"""
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('"', '&quot;')
    text = text.replace("'", '&apos;')
    return text


# ==================================================================
# 3. 辅助函数
# ==================================================================

def _fmt_ch(ch):
    """桩号 → 87+492.00（去掉K/YK/ZK前缀）"""
    ch = re.sub(r'^[A-Z]+', '', ch.strip().upper())
    if '+' in ch:
        km, m = ch.split('+')
        try:
            return f"{int(km)}+{float(m):.2f}"
        except:
            return ch
    return ch


def _ch_num(ch):
    """桩号 → 数字 87492"""
    ch = re.sub(r'^[A-Z]+', '', ch.strip().upper())
    if '+' in ch:
        km, m = ch.split('+')
        try:
            return int(km) * 1000 + int(float(m))
        except:
            return 0
    return 0


def _folder_name(line, rock, start, end, length, ch_prefix="K"):
    return f"田头山隧道{line}-洞身衬砌({ch_prefix}{start}～{ch_prefix}{end})({length:.2f}m，{rock})"


def _file_name(line, rock, sub_item, start, end, length, wh_number, ch_prefix="K"):
    """文件名与模板一致：使用 ch_prefix（K/YK/ZK）"""
    seg = f"田头山隧道{line}-洞身衬砌({ch_prefix}{start}～{ch_prefix}{end})({length:.2f}m，{rock})"
    return f"{wh_number}{seg}-{sub_item}.xlsx"


# ==================================================================
# 4. 命令行测试
# ==================================================================
if __name__ == "__main__":
    print("=" * 60)
    print("  分项完工计量模板化批量生成 V2")
    print("=" * 60)

    r = scan_and_build_templates()
    print(f"\n模板: {r['templates_count']} 个, 围岩: {r['rock_grades']}")

    idx = load_template_index()
    wh = {
        "洞身开挖": "WH3A4-03-17-02-001",
        "喷射砼支护": "WH3A4-03-33-02-001",
        "锚杆支护": "WH3A4-03-33-02-002",
        "钢筋网支护": "WH3A4-03-33-02-003",
        "钢架支护": "WH3A4-03-33-02-004",
        "衬砌钢筋加工及安装": "WH3A4-03-33-02-005",
        "衬砌砼": "WH3A4-03-33-02-006",
        "仰拱钢筋加工及安装": "WH3A4-03-33-02-007",
        "仰拱砼": "WH3A4-03-33-02-008",
        "仰拱回填": "WH3A4-03-33-02-009",
        "超前小导管支护": "WH3A4-03-33-02-010",
    }

    result = generate_segment_files(
        line="右线", rock_grade="S-Ⅴa",
        start_chain="K87+492", end_chain="K87+675", length=183.0,
        sub_items=list(wh.keys()),
        wh_mapping=wh,
    )
    print(f"\n生成: {result['folder']}")
    for r2 in result["results"]:
        print(f"  {r2['status']} {r2['sub_item']} → {r2.get('wh_number','')}")
