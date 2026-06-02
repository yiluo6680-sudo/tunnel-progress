"""桩号解析与区间操作工具"""

import re
from config import CHAINAGE_PATTERN


def parse_chainage(ch: str) -> float:
    """将桩号字符串转为绝对米数值
    支持格式: K85+147, ZK85+147, YK85+147, K85+147.5, 85147
    """
    if not ch or not isinstance(ch, str):
        return 0.0

    ch = ch.strip().upper()

    # 尝试标准格式 Kxx+xxx
    m = re.match(CHAINAGE_PATTERN, ch)
    if m:
        km = int(m.group(1))
        m_val = float(m.group(2))
        return km * 1000 + m_val

    # 尝试纯数字
    try:
        return float(ch)
    except (ValueError, TypeError):
        return 0.0


def format_chainage(meters: float) -> str:
    """将绝对米数值转回桩号显示格式: K85+147"""
    km = int(meters / 1000)
    m_val = meters - km * 1000
    return f"K{km}+{m_val:.1f}"


def chainage_to_relative(chainage: float, start_of_tunnel: float) -> float:
    """将绝对桩号转为相对隧道的米数（从洞口起算）"""
    return chainage - start_of_tunnel


def relative_to_chainage(relative_m: float, start_of_tunnel: float) -> float:
    """将相对米数转回绝对桩号"""
    return start_of_tunnel + relative_m


def parse_chainage_pair(ch_pair: str):
    """解析桩号区间字符串 'ZK85+147~ZK87+736.5' → (start, end)"""
    parts = ch_pair.replace('—', '~').replace('～', '~').split('~')
    if len(parts) == 2:
        return parse_chainage(parts[0]), parse_chainage(parts[1])
    return None, None


def segments_overlap(s1, e1, s2, e2) -> bool:
    """判断两个桩号区间是否有重叠"""
    return max(s1, s2) < min(e1, e2)


def overlap_length(s1, e1, s2, e2) -> float:
    """返回两个区间的重叠长度（米）"""
    return max(0, min(e1, e2) - max(s1, s2))
