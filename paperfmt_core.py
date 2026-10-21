# -*- coding: utf-8 -*-
"""
paperfmt_core.py — 论文格式锁定与应用核心库
================================================

功能：
  lock  : 读取"论文范稿" .docx，提取格式档案（页面设置 / docDefaults / 全部样式 /
          角色映射 / 自动编号 / 表格边框），输出为可复用的 JSON 档案文件。
  apply : 将格式档案应用到"待排版论文" .docx —— 复制样式定义、识别段落角色
          （标题/各级标题/摘要/关键词/参考文献/题注/表格文字/正文）、
          清除与范稿冲突的直接格式，输出符合范稿格式的新 docx 与处理报告。
  verify: 从输出文件反提格式档案，与锁定档案逐项比对，输出核对结果。

仅依赖 python-docx（含 lxml）。支持 .docx，不支持旧版 .doc。
"""
import os
import re
import json
import datetime
from collections import Counter

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

ORDER_RPR = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps',
             'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow',
             'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish',
             'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern',
             'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect',
             'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs',
             'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath']

ORDER_PPR = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore',
             'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers',
             'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku',
             'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct',
             'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd',
             'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing',
             'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection',
             'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl',
             'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange']

ORDER_NUMPR = ['w:ilvl', 'w:numId']

ORDER_TBLPR = ['w:tblStyle', 'w:tblpPr', 'w:tblOverlap', 'w:bidiVisual',
               'w:tblStyleRowBandSize', 'w:tblStyleColBandSize', 'w:tblW',
               'w:jc', 'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders',
               'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook',
               'w:tblCaption', 'w:tblDescription']

# 清除"直接格式"时要删除的 pPr / rPr 子元素（让样式定义接管）
CLEAR_PPR = ['w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:spacing',
             'w:ind', 'w:contextualSpacing', 'w:jc', 'w:textAlignment',
             'w:outlineLvl', 'w:rPr', 'w:suppressLineNumbers']
CLEAR_RPR = ['w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps',
             'w:smallCaps', 'w:color', 'w:spacing', 'w:w', 'w:kern',
             'w:position', 'w:sz', 'w:szCs', 'w:u', 'w:highlight']

CN_NUM = '一二三四五六七八九十'
ABSTRACT_HEADS = {'摘要', '摘 要', '摘　要', 'abstract', 'summary'}
REF_HEADS = {'参考文献', 'references', 'bibliography', '参考书目'}
RE_KEYWORD = re.compile(r'^\s*(关键词|关键字|關鍵詞|key\s*words?|keywords)\s*[:：]?', re.I)
RE_CAPTION = re.compile(r'^\s*(图|表|图表|figure|fig\.?|table|scheme|chart)\s*[\d一二三四五六七八九十]', re.I)
RE_H1 = re.compile(r'^\s*[%s]+\s*[、．.]' % CN_NUM)
RE_H2 = re.compile(r'^\s*[（(][%s]+[)）]' % CN_NUM)
RE_H3 = re.compile(r'^\s*\d+\s*[、．.]')
RE_H3B = re.compile(r'^\s*\d+\.\d+(\.\d+)?\b')
RE_END_PUNCT = re.compile(r'[。！？；，、;,.!?:：]$')


def _norm_head(t):
    return re.sub(r'[\s　]+', '', (t or '')).lower()

SIZE_NAMES = {42: '初号', 36: '小初', 26: '一号', 24: '小一', 22: '二号',
              18: '小二', 16: '三号', 15: '小三', 14: '四号', 12: '小四',
              10.5: '五号', 9: '小五', 7.5: '六号', 6.5: '小六'}


# ---------------------------------------------------------------------------
# XML 基础工具
# ---------------------------------------------------------------------------

def _local(tag):
    return tag.split('}')[-1] if '}' in tag else tag


def _bare(name):
    """'w:firstLine' → 'firstLine'"""
    return name.split(':')[-1]


def child_el(parent, tag, order):
    """按 schema 顺序查找/创建子元素"""
    e = parent.find(qn(tag))
    if e is not None:
        return e
    e = OxmlElement(tag)
    idx = order.index(tag)
    for sib in parent:
        key = 'w:' + _local(sib.tag)
        if key in order and order.index(key) > idx:
            sib.addprevious(e)
            return e
    parent.append(e)
    return e


def ser(el):
    """序列化 XML 元素为字符串（含命名空间声明）"""
    if el is None:
        return None
    return etree.tostring(el, encoding='unicode')


def size_name(half_pt):
    if half_pt is None:
        return '—'
    pt = half_pt / 2.0
    return SIZE_NAMES.get(pt, '%gpt' % pt)


# ---------------------------------------------------------------------------
# 读取：字体 / 段落格式
# ---------------------------------------------------------------------------

def read_font(rPr):
    if rPr is None:
        return {}
    f = {}
    rF = rPr.find(qn('w:rFonts'))
    if rF is not None:
        for attr, key in (('w:ascii', 'ascii'), ('w:hAnsi', 'hAnsi'),
                          ('w:eastAsia', 'eastasia'), ('w:cs', 'cs')):
            v = rF.get(qn(attr))
            if v:
                f[key] = v
    sz = rPr.find(qn('w:sz'))
    if sz is not None and sz.get(qn('w:val')):
        f['size_half_pt'] = int(sz.get(qn('w:val')))
    for tag, key in (('w:b', 'bold'), ('w:bCs', 'bold_cs'),
                     ('w:i', 'italic'), ('w:iCs', 'italic_cs')):
        e = rPr.find(qn(tag))
        if e is not None:
            v = e.get(qn('w:val'))
            f[key] = (v not in ('0', 'false', 'off')) if v is not None else True
    c = rPr.find(qn('w:color'))
    if c is not None and c.get(qn('w:val')):
        f['color'] = c.get(qn('w:val'))
    return f


def read_para(pPr):
    if pPr is None:
        return {}
    out = {}
    jc = pPr.find(qn('w:jc'))
    if jc is not None and jc.get(qn('w:val')):
        out['jc'] = jc.get(qn('w:val'))
    sp = pPr.find(qn('w:spacing'))
    if sp is not None:
        d = {}
        for a in ('w:before', 'w:after', 'w:line', 'w:lineRule',
                  'w:beforeLines', 'w:afterLines'):
            v = sp.get(qn(a))
            if v is not None:
                d[_bare(a)] = v
        if d:
            out['spacing'] = d
    ind = pPr.find(qn('w:ind'))
    if ind is not None:
        d = {}
        for a in ('w:firstLine', 'w:firstLineChars', 'w:hanging',
                  'w:hangingChars', 'w:left', 'w:leftChars', 'w:right',
                  'w:rightChars'):
            v = ind.get(qn(a))
            if v is not None:
                d[_bare(a)] = v
        if d:
            out['ind'] = d
    ol = pPr.find(qn('w:outlineLvl'))
    if ol is not None and ol.get(qn('w:val')) is not None:
        out['outlineLvl'] = ol.get(qn('w:val'))
    for t, k in (('w:keepNext', 'keepNext'), ('w:keepLines', 'keepLines'),
                 ('w:pageBreakBefore', 'pageBreakBefore'),
                 ('w:widowControl', 'widowControl')):
        e = pPr.find(qn(t))
        if e is not None:
            v = e.get(qn('w:val'))
            out[k] = (v not in ('0', 'false', 'off')) if v is not None else True
    np = pPr.find(qn('w:numPr'))
    if np is not None:
        il = np.find(qn('w:ilvl'))
        ni = np.find(qn('w:numId'))
        out['numPr'] = {'ilvl': il.get(qn('w:val')) if il is not None else None,
                        'numId': ni.get(qn('w:val')) if ni is not None else None}
    return out


def read_style_def(style_el):
    """读取 w:style 元素 → {'font': {...}, 'para': {...}}"""
    out = {'font': {}, 'para': {}}
    if style_el is None:
        return out
    out['font'] = read_font(style_el.find(qn('w:rPr')))
    out['para'] = read_para(style_el.find(qn('w:pPr')))
    return out


def write_font_def(rPr, fd):
    if not fd:
        return
    rF = child_el(rPr, 'w:rFonts', ORDER_RPR)
    for k, a in (('ascii', 'w:ascii'), ('hAnsi', 'w:hAnsi'),
                 ('eastasia', 'w:eastAsia'), ('cs', 'w:cs')):
        if fd.get(k):
            rF.set(qn(a), fd[k])
    if 'size_half_pt' in fd:
        sz = child_el(rPr, 'w:sz', ORDER_RPR)
        sz.set(qn('w:val'), str(int(fd['size_half_pt'])))
        szcs = child_el(rPr, 'w:szCs', ORDER_RPR)
        szcs.set(qn('w:val'), str(int(fd['size_half_pt'])))
    for k, t in (('bold', 'w:b'), ('bold_cs', 'w:bCs'),
                 ('italic', 'w:i'), ('italic_cs', 'w:iCs')):
        if k in fd:
            e = child_el(rPr, t, ORDER_RPR)
            e.set(qn('w:val'), '1' if fd[k] else '0')
    if fd.get('color'):
        c = child_el(rPr, 'w:color', ORDER_RPR)
        c.set(qn('w:val'), fd['color'])


def write_para_def(pPr, pd):
    if not pd:
        return
    if pd.get('jc'):
        jc = child_el(pPr, 'w:jc', ORDER_PPR)
        jc.set(qn('w:val'), pd['jc'])
    sp = pd.get('spacing')
    if sp:
        el = child_el(pPr, 'w:spacing', ORDER_PPR)
        for k, v in sp.items():
            el.set(qn('w:' + k), str(v))
    ind = pd.get('ind')
    if ind:
        el = child_el(pPr, 'w:ind', ORDER_PPR)
        for k, v in ind.items():
            el.set(qn('w:' + k), str(v))
    if pd.get('outlineLvl') is not None:
        el = child_el(pPr, 'w:outlineLvl', ORDER_PPR)
        el.set(qn('w:val'), str(pd['outlineLvl']))
    for k, t in (('keepNext', 'w:keepNext'), ('keepLines', 'w:keepLines'),
                 ('pageBreakBefore', 'w:pageBreakBefore'),
                 ('widowControl', 'w:widowControl')):
        if k in pd:
            el = child_el(pPr, t, ORDER_PPR)
            el.set(qn('w:val'), '1' if pd[k] else '0')


def merge_def(base, over):
    """base 为样式定义，over 为直接格式，后者优先"""
    out = {'font': dict(base.get('font') or {}), 'para': {}}
    bp = dict(base.get('para') or {})
    for k, v in (over.get('para') or {}).items():
        if k in ('spacing', 'ind') and isinstance(v, dict) and isinstance(bp.get(k), dict):
            merged = dict(bp[k])
            merged.update(v)
            bp[k] = merged
        else:
            bp[k] = v
    out['para'] = bp
    out['font'].update(over.get('font') or {})
    return out


# ---------------------------------------------------------------------------
# 描述辅助（报告用）
# ---------------------------------------------------------------------------

def font_desc(fd):
    if not fd:
        return '—', '—', '—', '—'
    zh = fd.get('eastasia') or fd.get('ascii') or '—'
    en = fd.get('ascii') or fd.get('hAnsi') or '—'
    sz = size_name(fd.get('size_half_pt'))
    bold = '加粗' if fd.get('bold') else '常规'
    return zh, en, sz, bold


def spacing_desc(pd):
    sp = (pd or {}).get('spacing') or {}
    if 'line' not in sp:
        return '单倍（默认）'
    rule = sp.get('lineRule', 'auto')
    if rule == 'auto':
        return '%g 倍行距' % round(int(sp['line']) / 240.0, 2)
    return '固定值 %gpt' % round(int(sp['line']) / 20.0, 1)


def indent_desc(pd):
    ind = (pd or {}).get('ind') or {}
    parts = []
    if ind.get('firstLineChars'):
        parts.append('首行缩进 %g 字符' % (int(ind['firstLineChars']) / 100.0))
    elif ind.get('firstLine'):
        parts.append('首行缩进 %gpt' % round(int(ind['firstLine']) / 20.0, 1))
    if ind.get('hangingChars'):
        parts.append('悬挂缩进 %g 字符' % (int(ind['hangingChars']) / 100.0))
    if ind.get('left'):
        parts.append('左缩进 %gpt' % round(int(ind['left']) / 20.0, 1))
    return '；'.join(parts) if parts else '无'


ALIGN_CN = {'left': '左对齐', 'center': '居中', 'right': '右对齐',
            'both': '两端对齐', 'distribute': '分散对齐'}


def align_desc(pd):
    v = (pd or {}).get('jc')
    return ALIGN_CN.get(v, v) if v else '默认（左对齐）'


# ---------------------------------------------------------------------------
# lock：从范稿提取格式档案
# ---------------------------------------------------------------------------

def _direct_def(p):
    """段落自身携带的直接格式定义（含段内混合加粗的前缀标签识别）"""
    pPr = p._p.find(qn('w:pPr'))
    para = read_para(pPr)
    font = {}
    bold_prefix = None
    run_infos = []
    for r in p._p.iter(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        txt = ''.join(t.text or '' for t in r.iter(qn('w:t')))
        if not txt.strip():
            continue
        run_infos.append((txt, read_font(rPr)))
    if run_infos:
        font = dict(run_infos[0][1])
        bolds = [bool(fd.get('bold')) for _, fd in run_infos]
        if len(set(bolds)) > 1:
            # 段内混合加粗：若加粗部分是短前缀（≤15字，如“关键词：”），
            # 样式按不加粗合成，并记录加粗前缀，应用时还原标签加粗
            prefix = ''
            n = 0
            for txt, fd in run_infos:
                if fd.get('bold'):
                    prefix += txt
                    n += len(txt)
                else:
                    break
            if 0 < n <= 15:
                font['bold'] = False
                bold_prefix = prefix
            else:
                bold_chars = sum(len(t) for t, fd in run_infos if fd.get('bold'))
                total = sum(len(t) for t, _ in run_infos) or 1
                font['bold'] = bold_chars * 2 > total
    return {'font': font, 'para': para, 'bold_prefix': bold_prefix}


def _has_direct_fmt(p):
    d = _direct_def(p)
    return bool(d['font']) or bool(d['para'])


def _style_name(p):
    try:
        return p.style.name if p.style is not None else 'Normal'
    except Exception:
        return 'Normal'


def style_or_synth(p, synth_name, prof):
    """范稿段落使用的样式名；若为 Normal+直接格式则合成一个新样式"""
    name = _style_name(p)
    if name == 'Normal' and _has_direct_fmt(p):
        base = prof['styles'].get('Normal', {'font': {}, 'para': {}})
        d = _direct_def(p)
        eff = merge_def(base, d)
        if d.get('bold_prefix'):
            eff['_bold_prefix'] = d['bold_prefix']
        prof['styles'][synth_name] = eff
        syn = prof['meta'].setdefault('synthesized', [])
        if synth_name not in syn:
            syn.append(synth_name)
        return synth_name
    return name


def _looks_like_title(p, text, body_half_pt):
    if not text or len(text) > 60:
        return False
    if RE_END_PUNCT.search(text):
        return False
    pPr = p._p.find(qn('w:pPr'))
    pd = read_para(pPr)
    jc = pd.get('jc')
    if jc == 'center':
        return True
    name = _style_name(p).lower()
    if 'title' in name or '标题' in name:
        return True
    d = _direct_def(p)
    size = d['font'].get('size_half_pt')
    if d['font'].get('bold'):
        return True
    if size is not None and body_half_pt is not None and size >= body_half_pt + 4:
        return True
    return False


def _heading_level(p, text, prof):
    """返回标题级别 1/2/3，非标题返回 None"""
    name = _style_name(p).lower()
    lvl = None
    if 'heading' in name or '标题' in name:
        m = re.search(r'(\d)', name)
        lvl = int(m.group(1)) if m else 1
    if lvl is None:
        ol = read_para(p._p.find(qn('w:pPr'))).get('outlineLvl')
        if ol is None:
            sd = prof['styles'].get(_style_name(p), {})
            ol = (sd.get('para') or {}).get('outlineLvl')
        if ol is not None:
            lvl = int(ol) + 1
    if lvl is None:
        if RE_H1.match(text):
            lvl = 1
        elif RE_H2.match(text):
            lvl = 2
        elif RE_H3B.match(text) or RE_H3.match(text):
            lvl = 3
    if lvl is not None:
        lvl = max(1, min(3, lvl))
    return lvl


def _read_first_table(doc):
    """读取范稿第一个表格的边框与单元格文字格式"""
    tbls = list(doc.element.body.iter(qn('w:tbl')))
    if not tbls:
        return None
    tbl = tbls[0]
    out = {'borders': {}, 'cell_style': None}
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is not None:
        st = tblPr.find(qn('w:tblStyle'))
        if st is not None and st.get(qn('w:val')):
            out['style_id'] = st.get(qn('w:val'))
        bd = tblPr.find(qn('w:tblBorders'))
        if bd is not None:
            for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                e = bd.find(qn('w:' + side))
                if e is not None:
                    d = {}
                    for a in ('w:val', 'w:sz', 'w:space', 'w:color'):
                        v = e.get(qn(a))
                        if v is not None:
                            d[_bare(a)] = v
                    out['borders'][side] = d
    # 单元格文字格式（取第一个有文字单元格所在段落的样式/直接格式）
    prof_stub = {'styles': {}, 'meta': {'synthesized': []}}
    from docx.text.paragraph import Paragraph
    for p_el in tbl.iter(qn('w:p')):
        para = Paragraph(p_el, doc)
        if para.text.strip():
            out['cell_style'] = style_or_synth(para, '论文-表格文字', prof_stub)
            out['cell_style_def'] = prof_stub['styles'].get(out['cell_style'])
            break
    return out


def extract_profile(template_path):
    """从范稿 docx 提取格式档案（dict）"""
    template_path = str(template_path)
    doc = Document(template_path)
    prof = {'version': '1.0',
            'meta': {'template': os.path.basename(template_path),
                     'extracted_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                     'tool': 'paperfmt 1.0', 'synthesized': []},
            'page': {}, 'doc_defaults': {}, 'doc_defaults_xml': None,
            'styles': {}, 'styles_xml': [], 'roles': {}, 'table': None,
            'numbering_xml': None, 'notes': []}

    # --- 页面设置（首个节） ---
    sectPr = doc.sections[0]._sectPr
    pgSz = sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        for a in ('w:w', 'w:h', 'w:orient'):
            v = pgSz.get(qn(a))
            if v:
                prof['page'][_bare(a)] = v
    pgMar = sectPr.find(qn('w:pgMar'))
    if pgMar is not None:
        for a in ('w:top', 'w:right', 'w:bottom', 'w:left', 'w:header',
                  'w:footer', 'w:gutter'):
            v = pgMar.get(qn(a))
            if v:
                prof['page'][_bare(a)] = v

    # --- docDefaults ---
    styles_el = doc.styles.element
    dd = styles_el.find(qn('w:docDefaults'))
    if dd is not None:
        prof['doc_defaults_xml'] = ser(dd)
        rprd = dd.find(qn('w:rPrDefault'))
        if rprd is not None:
            prof['doc_defaults']['font'] = read_font(rprd.find(qn('w:rPr')))
        pprd = dd.find(qn('w:pPrDefault'))
        if pprd is not None:
            prof['doc_defaults']['para'] = read_para(pprd.find(qn('w:pPr')))

    # --- 全部样式 XML（供整体复制）+ 用到的样式定义 ---
    for st in styles_el.findall(qn('w:style')):
        prof['styles_xml'].append(ser(st))
    used = set()
    for p_el in doc.element.body.iter(qn('w:p')):
        from docx.text.paragraph import Paragraph
        p = Paragraph(p_el, doc)
        used.add(_style_name(p))
    if 'Normal' not in used:
        used.add('Normal')
    name_by_id = {}
    for st in styles_el.findall(qn('w:style')):
        nm = st.find(qn('w:name'))
        sid = st.get(qn('w:styleId'))
        if nm is not None and nm.get(qn('w:val')):
            name_by_id[sid] = nm.get(qn('w:val'))
    for nm in sorted(used):
        for st in styles_el.findall(qn('w:style')):
            n = st.find(qn('w:name'))
            if n is not None and n.get(qn('w:val')) == nm:
                prof['styles'][nm] = read_style_def(st)
                break

    # --- 角色识别 ---
    from docx.text.paragraph import Paragraph
    paras = [Paragraph(p_el, doc) for p_el in doc.element.body.iter(qn('w:p'))]
    body_half = (prof['styles'].get('Normal', {}).get('font') or {}).get('size_half_pt')
    roles = {'heading': {}}
    abs_head_style = kw_style = refs_head_style = None
    abs_body_style = ref_body_style = caption_style = title_style = None
    abs_idx = kw_idx = refs_idx = None
    first_done = False
    body_counter = Counter()
    for i, p in enumerate(paras):
        t = p.text.strip()
        if not t:
            continue
        if RE_KEYWORD.match(t) and kw_style is None:
            kw_style = style_or_synth(p, '论文-关键词', prof)
            kw_idx = i
            first_done = True
            continue
        if _norm_head(t) in ('摘要', 'abstract', 'summary') and abs_idx is None:
            abs_head_style = style_or_synth(p, '论文-摘要标题', prof)
            abs_idx = i
            first_done = True
            continue
        if _norm_head(t) in ('参考文献', 'references', 'bibliography', '参考书目') \
                and refs_idx is None:
            refs_head_style = style_or_synth(p, '论文-参考文献标题', prof)
            refs_idx = i
            first_done = True
            continue
        if not first_done and _looks_like_title(p, t, body_half):
            title_style = style_or_synth(p, '论文-标题', prof)
            first_done = True
            continue
        first_done = True
        lvl = _heading_level(p, t, prof)
        if lvl:
            sname = style_or_synth(p, '论文-%d级标题' % lvl, prof)
            roles['heading'].setdefault(str(lvl), sname)
            continue
        if RE_CAPTION.match(t) and caption_style is None:
            caption_style = style_or_synth(p, '论文-图表题注', prof)
            continue
        if refs_idx is not None and i > refs_idx and ref_body_style is None:
            ref_body_style = style_or_synth(p, '论文-参考文献条目', prof)
            body_counter[_style_name(p)] += 1
        if abs_idx is not None and i > abs_idx and (kw_idx is None or i < kw_idx) \
                and abs_body_style is None:
            abs_body_style = style_or_synth(p, '论文-摘要正文', prof)
            body_counter[_style_name(p)] += 1
        # 统计正文样式使用频次（标题/特殊段之外的普通段落）
        if lvl is None and refs_idx is None and abs_idx is None \
                and not RE_CAPTION.match(t):
            body_counter[_style_name(p)] += 1

    h = roles['heading']
    roles['body'] = body_counter.most_common(1)[0][0] if body_counter else 'Normal'
    if title_style:
        roles['title'] = title_style
    roles['abstract_heading'] = abs_head_style or h.get('2') or h.get('1') or 'Normal'
    if not abs_head_style:
        prof['notes'].append('范稿无“摘要”标题段，已回退为 %s' % roles['abstract_heading'])
    roles['abstract_body'] = abs_body_style or 'Normal'
    roles['keywords'] = kw_style or 'Normal'
    roles['references_heading'] = refs_head_style or h.get('1') or 'Normal'
    if not refs_head_style:
        prof['notes'].append('范稿无“参考文献”标题段，已回退为 %s' % roles['references_heading'])
    roles['reference_body'] = ref_body_style or 'Normal'
    roles['caption'] = caption_style or 'Normal'
    prof['roles'] = roles

    # --- 表格 ---
    prof['table'] = _read_first_table(doc)
    if prof['table'] and prof['table'].get('cell_style_def'):
        prof['styles'].setdefault(prof['table']['cell_style'],
                                  prof['table']['cell_style_def'])
        prof['meta'].setdefault('synthesized', []).append(prof['table']['cell_style'])

    # --- numbering.xml（自动编号定义，供 apply 导入） ---
    for part in doc.part.package.parts:
        if str(part.partname) == '/word/numbering.xml':
            try:
                prof['numbering_xml'] = part.blob.decode('utf-8')
            except Exception:
                pass
            break
    return prof


# ---------------------------------------------------------------------------
# apply：把档案应用到目标文档
# ---------------------------------------------------------------------------

def _remove_style_by(tstyles_el, style_id=None, name=None):
    for st in tstyles_el.findall(qn('w:style')):
        sid = st.get(qn('w:styleId'))
        nm = st.find(qn('w:name'))
        nmv = nm.get(qn('w:val')) if nm is not None else None
        if (style_id and sid == style_id) or (name and nmv == name):
            tstyles_el.remove(st)


def _style_exists(doc, name):
    for s in doc.styles:
        try:
            if s.name == name:
                return True
        except Exception:
            continue
    return False


def _apply_section(sectPr, page):
    pgSz = sectPr.find(qn('w:pgSz'))
    if pgSz is None:
        pgSz = OxmlElement('w:pgSz')
        sectPr.append(pgSz)
    for k in ('w', 'h', 'orient'):
        if page.get(k):
            pgSz.set(qn('w:' + k), page[k])
    pgMar = sectPr.find(qn('w:pgMar'))
    if pgMar is None:
        pgMar = OxmlElement('w:pgMar')
        sectPr.append(pgMar)
    for k in ('top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'):
        if page.get(k):
            pgMar.set(qn('w:' + k), page[k])


def _get_target_numbering(doc):
    """返回目标文档 numbering 根元素；不存在则尽力创建"""
    for part in doc.part.package.parts:
        if str(part.partname) == '/word/numbering.xml':
            return parse_xml(part.blob), part
    # 创建 numbering part
    try:
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        xml = ('<w:numbering xmlns:w="http://schemas.openxmlformats.org/'
               'wordprocessingml/2006/main"/>')
        part = Part(PackURI('/word/numbering.xml'),
                    'application/vnd.openxmlformats-officedocument.'
                    'wordprocessingml.numbering+xml',
                    xml.encode('utf-8'), doc.part.package)
        doc.part.relate_to(part, RT.NUMBERING)
        return parse_xml(xml), part
    except Exception:
        return None, None


def _import_numbering(doc, prof):
    """把范稿样式用到的自动编号定义导入目标文档，返回 {旧numId: 新numId}"""
    used = set()
    for xml in prof.get('styles_xml') or []:
        try:
            el = parse_xml(xml)
        except Exception:
            continue
        pPr = el.find(qn('w:pPr'))
        if pPr is None:
            continue
        np = pPr.find(qn('w:numPr'))
        if np is not None:
            ni = np.find(qn('w:numId'))
            if ni is not None and ni.get(qn('w:val')):
                used.add(ni.get(qn('w:val')))
    if not used or not prof.get('numbering_xml'):
        return {}
    tgt_num, tgt_part = _get_target_numbering(doc)
    if tgt_num is None:
        return {}
    try:
        src_num = parse_xml(prof['numbering_xml'].encode('utf-8'))
    except Exception:
        return {}

    def max_id(tag, attr):
        m = 0
        for e in tgt_num.findall(qn('w:' + tag)):
            v = e.get(qn('w:' + attr))
            if v and v.isdigit():
                m = max(m, int(v))
        return m

    next_abs = max_id('abstractNum', 'abstractNumId') + 1
    next_num = max_id('num', 'numId') + 1
    mapping = {}
    for nid in sorted(used):
        num_el = None
        for e in src_num.findall(qn('w:num')):
            if e.get(qn('w:numId')) == nid:
                num_el = e
                break
        if num_el is None:
            continue
        aid_el = num_el.find(qn('w:abstractNumId'))
        if aid_el is None:
            continue
        old_aid = aid_el.get(qn('w:val'))
        abs_el = None
        for e in src_num.findall(qn('w:abstractNum')):
            if e.get(qn('w:abstractNumId')) == old_aid:
                abs_el = e
                break
        if abs_el is None:
            continue
        # 目标中是否已有同名编号定义（按内容简单判重：跳过，直接复制）
        new_abs = etree.fromstring(etree.tostring(abs_el))
        new_abs.set(qn('w:abstractNumId'), str(next_abs))
        tgt_num.append(new_abs)
        new_num = etree.fromstring(etree.tostring(num_el))
        new_num.set(qn('w:numId'), str(next_num))
        aid2 = new_num.find(qn('w:abstractNumId'))
        if aid2 is not None:
            aid2.set(qn('w:val'), str(next_abs))
        tgt_num.append(new_num)
        mapping[nid] = str(next_num)
        next_abs += 1
        next_num += 1
    if tgt_part is not None and hasattr(tgt_part, '_blob'):
        try:
            tgt_part._blob = etree.tostring(tgt_num, xml_declaration=True,
                                            encoding='UTF-8', standalone=True)
        except Exception:
            pass
    return mapping


def _style_numpr(doc, style_name):
    """样式中携带的自动编号 (numId, ilvl)"""
    for s in doc.styles:
        try:
            if s.name != style_name:
                continue
        except Exception:
            continue
        pPr = s.element.find(qn('w:pPr'))
        if pPr is None:
            return None
        np = pPr.find(qn('w:numPr'))
        if np is None:
            return None
        ni = np.find(qn('w:numId'))
        il = np.find(qn('w:ilvl'))
        if ni is None or ni.get(qn('w:val')) in (None, '0'):
            return None
        return (ni.get(qn('w:val')), il.get(qn('w:val')) if il is not None else '0')
    return None


def _clear_para_direct(p, keep_layout=False):
    pPr = p._p.find(qn('w:pPr'))
    if pPr is None:
        return
    for tag in CLEAR_PPR:
        if keep_layout and tag in ('w:jc', 'w:ind', 'w:spacing'):
            continue
        e = pPr.find(qn(tag))
        if e is not None:
            pPr.remove(e)


def _clear_runs_direct(p, keep_bold=False):
    for r in p._p.iter(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            continue
        for tag in CLEAR_RPR:
            if keep_bold and tag in ('w:b', 'w:bCs'):
                continue
            e = rPr.find(qn(tag))
            if e is not None:
                rPr.remove(e)
        if len(rPr) == 0 and not rPr.attrib:
            r.remove(rPr)


def _para_has_image(p):
    for tag in ('w:drawing', 'w:pict', 'w:object'):
        if p._p.find('.//' + qn(tag)) is not None:
            return True
    return False


def _attach_numpr(p, numId, ilvl):
    pPr = p._p.get_or_add_pPr()
    old = pPr.find(qn('w:numPr'))
    if old is not None:
        pPr.remove(old)
    numPr = child_el(pPr, 'w:numPr', ORDER_PPR)
    e = child_el(numPr, 'w:ilvl', ORDER_NUMPR)
    e.set(qn('w:val'), str(ilvl))
    e = child_el(numPr, 'w:numId', ORDER_NUMPR)
    e.set(qn('w:val'), str(numId))


MANUAL_NUM_RES = [re.compile(r'^\s*[%s]+\s*[、．.]\s*' % CN_NUM),
                  re.compile(r'^\s*[（(][%s]+[)）]\s*' % CN_NUM),
                  re.compile(r'^\s*\d+\s*[、．.]\s*'),
                  re.compile(r'^\s*\d+\.\d+(\.\d+)?\s*'),
                  re.compile(r'^\s*\d+\.\s+')]


def _strip_manual_number(p):
    for r in p._p.iter(qn('w:r')):
        for t in r.iter(qn('w:t')):
            if t.text and t.text.strip():
                for rx in MANUAL_NUM_RES:
                    new, n = rx.subn('', t.text, count=1)
                    if n:
                        t.text = new
                        return True
                return False
    return False


def _set_bold(r, on):
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r.insert(0, rPr)
    for t in ('w:b', 'w:bCs'):
        e = child_el(rPr, t, ORDER_RPR)
        e.set(qn('w:val'), '1' if on else '0')


def _bold_prefix(p, prefix):
    """给段落开头与 prefix 匹配的文字恢复加粗（如“关键词：”标签）；
    前缀落在 run 中间时自动拆分 run。"""
    if not prefix or not p.text.startswith(prefix):
        return False
    remaining = len(prefix)
    for r in list(p._p.iter(qn('w:r'))):
        ts = list(r.iter(qn('w:t')))
        txt = ''.join(t.text or '' for t in ts)
        if not txt:
            continue
        if remaining <= 0:
            break
        if len(txt) <= remaining:
            _set_bold(r, True)
            remaining -= len(txt)
            continue
        # 前缀边界落在本 run 内部
        if len(ts) == 1:
            cut = remaining
            full = ts[0].text or ''
            ts[0].text = full[:cut]
            r2 = etree.fromstring(etree.tostring(r))
            t2 = list(r2.iter(qn('w:t')))[0]
            t2.text = full[cut:]
            _set_bold(r2, False)
            r.addnext(r2)
            _set_bold(r, True)
        else:
            _set_bold(r, True)  # 多 w:t 结构罕见，整 run 加粗
        remaining = 0
        break
    return True


def _resolve_style_name(doc, name, fallback):
    if name and _style_exists(doc, name):
        return name
    if _style_exists(doc, fallback):
        return fallback
    return 'Normal'


def _in_table(p_el):
    anc = p_el.getparent()
    while anc is not None:
        if anc.tag == qn('w:tbl'):
            return True
        anc = anc.getparent()
    return False


def apply_profile(prof, target_path, out_path, report_path=None):
    """把格式档案应用到目标 docx，返回报告 dict"""
    from docx.text.paragraph import Paragraph
    target_path = str(target_path)
    out_path = str(out_path)
    doc = Document(target_path)
    warnings = []
    rows = []
    stats = {}

    def bump(role):
        stats[role] = stats.get(role, 0) + 1

    # 1. 复制范稿全部样式 + docDefaults
    tstyles = doc.styles.element
    for xml in prof.get('styles_xml') or []:
        try:
            el = parse_xml(xml)
        except Exception:
            continue
        sid = el.get(qn('w:styleId'))
        nm_el = el.find(qn('w:name'))
        nm = nm_el.get(qn('w:val')) if nm_el is not None else None
        _remove_style_by(tstyles, style_id=sid)
        if nm:
            _remove_style_by(tstyles, name=nm)
        tstyles.append(el)
    if prof.get('doc_defaults_xml'):
        old_dd = tstyles.find(qn('w:docDefaults'))
        if old_dd is not None:
            tstyles.remove(old_dd)
        tstyles.insert(0, parse_xml(prof['doc_defaults_xml']))

    # 2. 合成样式（范稿中靠直接格式定义的段落角色）
    for name, sd in prof.get('styles', {}).items():
        if not _style_exists(doc, name):
            try:
                st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
                write_font_def(st.element.get_or_add_rPr(), sd.get('font'))
                write_para_def(st.element.get_or_add_pPr(), sd.get('para'))
            except Exception as e:
                warnings.append('合成样式 %s 失败：%s' % (name, e))

    # 3. 页面设置（所有节）
    for sec in doc.sections:
        _apply_section(sec._sectPr, prof.get('page') or {})

    # 4. 自动编号导入
    num_map = _import_numbering(doc, prof)
    if prof.get('numbering_xml') and not num_map:
        warnings.append('范稿使用了自动编号，但目标文档无法导入编号定义，标题编号可能缺失')

    roles = prof.get('roles') or {}
    body_style = roles.get('body', 'Normal')
    body_half = (prof.get('styles', {}).get(body_style, {}).get('font') or {}).get('size_half_pt')
    heading_styles = roles.get('heading', {})
    num_by_level = {}
    for lvl, sname in heading_styles.items():
        np = _style_numpr(doc, sname)
        if np:
            old_id, ilvl = np
            if old_id in num_map:
                num_by_level[lvl] = (num_map[old_id], ilvl)
            else:
                warnings.append('%d 级标题样式的自动编号未能导入' % int(lvl))

    table_info = prof.get('table') or {}
    cell_style = table_info.get('cell_style')

    # 5. 段落遍历：识别角色 → 套样式 → 清直接格式
    ctx = {'first_done': False, 'in_abstract': False, 'in_refs': False}
    idx = 0
    for p_el in doc.element.body.iter(qn('w:p')):
        p = Paragraph(p_el, doc)
        text = p.text.strip()
        idx += 1
        role = None
        note = ''
        lvl = None

        if _in_table(p_el):
            role = 'table_cell'
        elif not text:
            role = 'empty'
        else:
            if RE_KEYWORD.match(text):
                role = 'keywords'
                ctx['in_abstract'] = False
                ctx['first_done'] = True
            elif _norm_head(text) in ('摘要', 'abstract', 'summary'):
                role = 'abstract_heading'
                ctx['in_abstract'] = True
                ctx['first_done'] = True
            elif _norm_head(text) in ('参考文献', 'references', 'bibliography', '参考书目'):
                role = 'references_heading'
                ctx['in_refs'] = True
                ctx['first_done'] = True
            elif not ctx['first_done'] and _looks_like_title(p, text, body_half):
                role = 'title'
                ctx['first_done'] = True
            else:
                lvl = _heading_level(p, text, prof)
                if lvl:
                    role = 'heading_%d' % lvl
                elif ctx['in_refs']:
                    role = 'reference_body'
                elif ctx['in_abstract']:
                    role = 'abstract_body'
                elif RE_CAPTION.match(text):
                    role = 'caption'
                else:
                    role = 'body'
                ctx['first_done'] = True

        # 角色 → 样式名
        if role == 'table_cell':
            sname = _resolve_style_name(doc, cell_style, body_style)
            if not cell_style:
                note = '范稿无表格样式，表格文字按正文'
        elif role == 'empty':
            sname = body_style
        elif role == 'title':
            sname = _resolve_style_name(doc, roles.get('title'), body_style)
            if not roles.get('title'):
                note = '范稿无标题样式，按正文'
        elif role and role.startswith('heading_'):
            sname = _resolve_style_name(doc, heading_styles.get(str(lvl)), body_style)
            if not heading_styles.get(str(lvl)):
                note = '范稿无 %d 级标题样式，按正文' % lvl
        else:
            sname = _resolve_style_name(doc, roles.get(role), body_style)

        # 套样式
        try:
            p.style = doc.styles[sname]
        except Exception:
            sname = body_style
            try:
                p.style = doc.styles[body_style]
            except Exception:
                pass
            note = (note + '；' if note else '') + '样式套用失败，回退正文'

        # 清直接格式（图片段保留版面；表格加粗保留表头强调）
        keep_layout = _para_has_image(p)
        keep_bold = (role == 'table_cell')
        _clear_para_direct(p, keep_layout=keep_layout)
        _clear_runs_direct(p, keep_bold=keep_bold)

        # 还原段内标签加粗（如“关键词：”）
        sd = (prof.get('styles') or {}).get(sname) or {}
        if sd.get('_bold_prefix'):
            _bold_prefix(p, sd['_bold_prefix'])

        # 自动编号挂载 + 剥离手动编号
        if role and role.startswith('heading_') and str(lvl) in num_by_level:
            nid, nilvl = num_by_level[str(lvl)]
            _attach_numpr(p, nid, nilvl)
            if _strip_manual_number(p):
                note = (note + '；' if note else '') + '手动编号已剥离（改用自动编号）'

        if role != 'empty':
            bump(role)
        rows.append({'idx': idx, 'role': role, 'text': text[:18],
                     'style': sname, 'note': note})
        if keep_layout:
            note_img = '图片段落保留原版面（对齐/缩进未清除）'

    # 6. 表格边框
    if table_info.get('borders'):
        tblPr_order = ORDER_TBLPR
        for tbl in doc.element.body.iter(qn('w:tbl')):
            tblPr = tbl.find(qn('w:tblPr'))
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr')
                tbl.insert(0, tblPr)
            old = tblPr.find(qn('w:tblBorders'))
            if old is not None:
                tblPr.remove(old)
            bd = OxmlElement('w:tblBorders')
            for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                d = table_info['borders'].get(side)
                if d is None:
                    continue
                e = OxmlElement('w:' + side)
                for k, v in d.items():
                    e.set(qn('w:' + k), str(v))
                bd.append(e)
            tgt_idx = tblPr_order.index('w:tblBorders')
            inserted = False
            for sib in tblPr:
                key = 'w:' + _local(sib.tag)
                if key in tblPr_order and tblPr_order.index(key) > tgt_idx:
                    sib.addprevious(bd)
                    inserted = True
                    break
            if not inserted:
                tblPr.append(bd)

    # 页眉页脚说明
    if any(True for s in doc.sections for hf in (s.header, s.footer)
           if hf is not None and hf.paragraphs and any(pp.text.strip() for pp in hf.paragraphs)):
        warnings.append('页眉/页脚内容未纳入格式锁定（不同论文页眉内容不同，如需统一请告知）')
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or '.', exist_ok=True)
    doc.save(out_path)

    report = {'template': prof.get('meta', {}).get('template'),
              'target': os.path.basename(target_path),
              'output': os.path.basename(out_path),
              'generated_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
              'page': prof.get('page') or {},
              'roles': roles, 'styles': prof.get('styles') or {},
              'rows': rows, 'stats': stats, 'warnings': warnings,
              'notes': prof.get('notes') or []}
    if report_path:
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(render_report(report))
    return report


# ---------------------------------------------------------------------------
# 报告渲染
# ---------------------------------------------------------------------------

def _page_desc(page):
    def mm(k):
        v = page.get(k)
        return round(int(v) / 56.6929, 1) if v else None
    w, h = mm('w'), mm('h')
    paper = 'A4 (%.0f×%.0f mm)' % (w, h) if w and h else '—'
    m = {k: mm(k) for k in ('top', 'right', 'bottom', 'left')}
    return paper, '上 %g / 下 %g / 左 %g / 右 %g cm' % (
        m['top'] / 10, m['bottom'] / 10, m['left'] / 10, m['right'] / 10)


ROLE_CN = {'title': '标题', 'body': '正文', 'abstract_heading': '摘要标题',
           'abstract_body': '摘要正文', 'keywords': '关键词',
           'references_heading': '参考文献标题', 'reference_body': '参考文献条目',
           'caption': '图表题注', 'table_cell': '表格文字', 'empty': '空段'}
for _i in (1, 2, 3):
    ROLE_CN['heading_%d' % _i] = '%d 级标题' % _i


def render_report(rep):
    L = []
    L.append('# 论文格式规范报告')
    L.append('')
    L.append('- 格式范稿：%s' % rep.get('template'))
    L.append('- 待排版论文：%s' % rep.get('target'))
    L.append('- 输出文件：%s' % rep.get('output'))
    L.append('- 生成时间：%s' % rep.get('generated_at'))
    L.append('')
    paper, margin = _page_desc(rep.get('page') or {})
    L.append('## 一、页面设置（已应用到所有节）')
    L.append('')
    L.append('| 项目 | 值 |')
    L.append('| --- | --- |')
    L.append('| 纸张 | %s |' % paper)
    L.append('| 页边距 | %s |' % margin)
    L.append('')
    L.append('## 二、锁定的样式定义')
    L.append('')
    L.append('| 角色 | 样式名 | 中文字体 | 西文字体 | 字号 | 字重 | 对齐 | 行距 | 缩进 |')
    L.append('| --- | --- | --- | --- | --- | --- | --- | --- | --- |')
    roles = rep.get('roles') or {}
    styles = rep.get('styles') or {}

    def row(role_cn, sname):
        sd = styles.get(sname) or {}
        zh, en, sz, bold = font_desc(sd.get('font'))
        L.append('| %s | %s | %s | %s | %s | %s | %s | %s | %s |' % (
            role_cn, sname, zh, en, sz, bold,
            align_desc(sd.get('para')), spacing_desc(sd.get('para')),
            indent_desc(sd.get('para'))))

    if roles.get('title'):
        row('标题', roles['title'])
    for lvl in ('1', '2', '3'):
        if (roles.get('heading') or {}).get(lvl):
            row('%s 级标题' % lvl, roles['heading'][lvl])
    row('正文', roles.get('body', 'Normal'))
    for r in ('abstract_heading', 'abstract_body', 'keywords',
              'references_heading', 'reference_body', 'caption'):
        if roles.get(r):
            row(ROLE_CN[r], roles[r])
    if rep.get('table_style_note'):
        L.append('')
        L.append('> %s' % rep['table_style_note'])
    L.append('')
    L.append('## 三、段落处理明细')
    L.append('')
    L.append('| # | 角色 | 内容（前 18 字） | 应用样式 | 备注 |')
    L.append('| --- | --- | --- | --- | --- |')
    for r in rep.get('rows') or []:
        txt = (r.get('text') or '').replace('|', '／') or '（空段）'
        L.append('| %d | %s | %s | %s | %s |' % (
            r['idx'], ROLE_CN.get(r['role'], r['role']), txt,
            r.get('style', ''), r.get('note', '') or '—'))
    L.append('')
    L.append('## 四、统计')
    L.append('')
    st = rep.get('stats') or {}
    L.append('；'.join('%s %d 段' % (ROLE_CN.get(k, k), v) for k, v in st.items())
             or '（无段落）')
    L.append('')
    warns = list(rep.get('warnings') or []) + list(rep.get('notes') or [])
    L.append('## 五、警告与说明')
    L.append('')
    if warns:
        for w in warns:
            L.append('- %s' % w)
    else:
        L.append('- 无')
    L.append('')
    return '\n'.join(L)


# ---------------------------------------------------------------------------
# verify：核对输出是否符合锁定的格式
# ---------------------------------------------------------------------------

def verify_profile(prof, docx_path):
    out = extract_profile(docx_path)
    checks = []

    def cmp(path, a, b):
        checks.append({'item': path, 'expect': a, 'actual': b,
                       'pass': a == b})

    for k in ('w', 'h', 'orient', 'top', 'right', 'bottom', 'left', 'gutter'):
        cmp('页面.' + k, (prof.get('page') or {}).get(k), (out.get('page') or {}).get(k))
    ddf = (prof.get('doc_defaults') or {}).get('font') or {}
    if ddf:
        for k in ('eastasia', 'ascii', 'size_half_pt'):
            cmp('默认字体.' + k, ddf.get(k),
                ((out.get('doc_defaults') or {}).get('font') or {}).get(k))
    names = set()
    roles = prof.get('roles') or {}
    if roles.get('title'):
        names.add(roles['title'])
    names.add(roles.get('body', 'Normal'))
    for v in (roles.get('heading') or {}).values():
        names.add(v)
    for r in ('abstract_heading', 'abstract_body', 'keywords',
              'references_heading', 'reference_body', 'caption'):
        if roles.get(r):
            names.add(roles[r])
    for n in sorted(names):
        a = (prof.get('styles') or {}).get(n) or {}
        b = (out.get('styles') or {}).get(n) or {}
        for k in ('eastasia', 'ascii', 'size_half_pt', 'bold'):
            cmp('%s.字体.%s' % (n, k), (a.get('font') or {}).get(k),
                (b.get('font') or {}).get(k))
        cmp('%s.对齐' % n, (a.get('para') or {}).get('jc'),
            (b.get('para') or {}).get('jc'))
        cmp('%s.行距' % n, ((a.get('para') or {}).get('spacing') or {}).get('line'),
            ((b.get('para') or {}).get('spacing') or {}).get('line'))
        for k in ('firstLineChars', 'firstLine', 'hanging', 'left'):
            cmp('%s.缩进.%s' % (n, k), ((a.get('para') or {}).get('ind') or {}).get(k),
                ((b.get('para') or {}).get('ind') or {}).get(k))
    passed = all(c['pass'] for c in checks)
    return {'checks': checks, 'passed': passed,
            'failed': [c for c in checks if not c['pass']]}


def render_verify(vr):
    L = ['# 格式核对结果', '',
         '总体结论：%s' % ('✅ 全部通过' if vr['passed'] else '❌ 存在不一致'), '',
         '| 项目 | 范稿锁定值 | 输出实际值 | 结果 |',
         '| --- | --- | --- | --- |']
    for c in vr['checks']:
        L.append('| %s | %s | %s | %s |' % (
            c['item'], c['expect'] if c['expect'] is not None else '—',
            c['actual'] if c['actual'] is not None else '—',
            '✅' if c['pass'] else '❌'))
    return '\n'.join(L)
