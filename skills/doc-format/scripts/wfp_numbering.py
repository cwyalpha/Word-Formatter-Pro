"""Convert DOCX paragraph list labels to text without rebuilding the document.

Only XML parts containing converted paragraphs are rewritten. All other ZIP
members, including media, relationships, styles and numbering definitions, keep
their original contents. No Office installation is required.
"""
from __future__ import annotations

import argparse
from copy import deepcopy
from dataclasses import dataclass, field
import json
import os
from pathlib import Path
import posixpath
import re
import tempfile
import unicodedata
from zipfile import BadZipFile, ZipFile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
PLACEHOLDER = re.compile(r"%([1-9])")
NUMBER_FIELD = re.compile(r"\b(?:LISTNUM|AUTONUM|AUTONUMLGL|AUTONUMOUT)\b", re.I)
PPR_ORDER = """pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr
    suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap
    overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd
    snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc
    textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr
    sectPr pPrChange""".split()


class NumberingConversionError(ValueError):
    """The input cannot be converted reliably; the output has not been written."""


@dataclass
class ConversionReport:
    converted: int = 0
    parts: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _tag(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def _child(element, ns, name):
    return None if element is None else element.find(_tag(ns, name))


def _val(element, ns, name, default=None):
    item = _child(element, ns, name)
    return default if item is None else item.get(_tag(ns, "val"), default)


def _int(value, default=0):
    return default if value is None else int(value)


def _on(element, ns, name):
    item = _child(element, ns, name)
    return item is not None and item.get(_tag(ns, "val"), "1") not in ("0", "false", "off")


def _insert_property(parent, child, ns):
    """Keep newly added pPr children in schema order."""
    name = etree.QName(child).localname
    rank = PPR_ORDER.index(name) if name in PPR_ORDER else len(PPR_ORDER)
    for index, sibling in enumerate(parent):
        other = etree.QName(sibling).localname
        if other in PPR_ORDER and PPR_ORDER.index(other) > rank:
            parent.insert(index, child)
            return
    parent.append(child)


def _merge_properties(elements, ns, name):
    result = etree.Element(_tag(ns, name))
    for element in elements:
        if element is None:
            continue
        for item in element:
            old = result.find(item.tag)
            if old is None:
                result.append(deepcopy(item))
            else:
                # Property attributes (notably partial rFonts and ind) inherit
                # individually; complex children of the last property replace.
                if item.tag == _tag(ns, "rFonts"):
                    # An explicit list font must override a theme font from
                    # docDefaults/styles. Keeping asciiTheme next to a newly
                    # supplied ascii="Arial" makes Office use the theme font,
                    # changing label width and the following tab position.
                    for literal, theme in (("ascii", "asciiTheme"), ("hAnsi", "hAnsiTheme"),
                                           ("eastAsia", "eastAsiaTheme"), ("cs", "cstheme")):
                        if _tag(ns, literal) in item.attrib and _tag(ns, theme) not in item.attrib:
                            old.attrib.pop(_tag(ns, theme), None)
                old.attrib.update(item.attrib)
                if len(item):
                    old[:] = [deepcopy(child) for child in item]
    return result


def _effective_indent(properties, ns):
    result = etree.Element(_tag(ns, "ind"))
    for ppr in properties:
        ind = _child(ppr, ns, "ind")
        if ind is None:
            continue
        # Character-unit indentation takes precedence over twips even if the
        # twip value appears later in the inheritance. Switch the opposing
        # character-unit variant only when a new character-unit value is set.
        for active, opposing in (("firstLine", "hanging"), ("hanging", "firstLine")):
            if (_tag(ns, active) in ind.attrib
                    or (_tag(ns, active + "Chars") in ind.attrib
                        and _int(ind.get(_tag(ns, active + "Chars"))) != 0)):
                result.attrib.pop(_tag(ns, opposing), None)
            # A zero character-unit override cancels an opposing character
            # indent, but does not cancel a hanging/first-line twip indent.
            # WPS list paragraphs commonly contain firstLineChars="0" beside
            # a list-level hanging indent; discarding hanging loses layout.
            if _tag(ns, active + "Chars") in ind.attrib:
                result.attrib.pop(_tag(ns, opposing + "Chars"), None)
        result.attrib.update(ind.attrib)
    return result


def _effective_tabs(properties, ns):
    stops = {}
    for ppr in properties:
        tabs = _child(ppr, ns, "tabs")
        if tabs is None:
            continue
        for tab in tabs.findall(_tag(ns, "tab")):
            position = tab.get(_tag(ns, "pos"))
            stops[position] = deepcopy(tab)
    result = etree.Element(_tag(ns, "tabs"))
    for position in sorted(stops, key=lambda p: _int(p)):
        tab = stops[position]
        if tab.get(_tag(ns, "val")) == "num":
            # A numbering tab is no longer active once numId=0. Retain its
            # position as an ordinary left tab for the inserted text run.
            tab.set(_tag(ns, "val"), "left")
        result.append(tab)
    return result


def _cjk_label_fills_hanging_indent(label, rpr, definition, cascade, settings, ns):
    """Recognize the one-em CJK boundary where a literal tab adds a new gap.

    A list suffix can finish at the hanging-indent boundary. An ordinary run
    tab at that same boundary advances to the NEXT stop instead. Only omit
    that redundant tab for full-width CJK text of known size with an exact
    match. This is deliberately not a general font-width/layout estimator.
    """
    if _val(definition, ns, "lvlJc", "left") != "left":
        return False
    if not label or not all(unicodedata.east_asian_width(c) in ("W", "F")
                            and unicodedata.category(c) in ("Lo", "Nl", "Nd", "Po", "Ps", "Pe")
                            for c in label):
        return False
    # Custom stops at/after the boundary, character indents and unusual
    # shaping need a renderer. Earlier left stops cannot affect this tab.
    ind = _effective_indent(cascade, ns)
    left = _int(ind.get(_tag(ns, "start"), ind.get(_tag(ns, "left"))))
    if any(tab.get(_tag(ns, "val")) != "left"
           or _int(tab.get(_tag(ns, "pos"))) >= left
           for tab in _effective_tabs(cascade, ns)):
        return False
    ppr = _merge_properties(cascade, ns, "pPr")
    if (_on(ppr, ns, "bidi") or _on(ppr, ns, "mirrorIndents")
            or _child(ppr, ns, "textDirection") is not None):
        return False
    compat = _child(settings, ns, "compat")
    if any(_on(compat, ns, name) for name in
           ("doNotUseIndentAsNumberingTabStop", "noTabHangInd")):
        return False
    if any(_int(value) for name, value in ind.attrib.items()
           if etree.QName(name).localname.endswith("Chars")):
        return False
    if (_child(rpr, ns, "rStyle") is not None
            or _val(rpr, ns, "spacing", "0") != "0"
            or _val(rpr, ns, "w", "100") != "100"
            or _val(rpr, ns, "vertAlign", "baseline") != "baseline"
            or any(_on(rpr, ns, name) for name in ("rtl", "cs", "vanish"))
            or _child(rpr, ns, "fitText") is not None):
        return False
    size = _int(_val(rpr, ns, "sz"))  # half-points; one em = size * 10 twips
    hanging = _int(ind.get(_tag(ns, "hanging")))
    return size > 0 and hanging > 0 and hanging == len(label) * size * 10


def _roman(number):
    if not 0 < number < 4000:
        return str(number)
    result = ""
    for value, text in [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]:
        count, number = divmod(number, value)
        result += text * count
    return result


def _letters(number):
    if number < 1:
        return str(number)
    result = ""
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def _chinese(number, legal=False, traditional=False):
    """Chinese cardinal numbers, including zeros between 4-digit groups."""
    digits = "零壹貳參肆伍陸柒捌玖" if traditional else (
        "零壹贰叁肆伍陆柒捌玖" if legal else "〇一二三四五六七八九")
    units = ("", "拾", "佰", "仟") if legal or traditional else ("", "十", "百", "千")
    groups = ("", "萬", "億", "兆") if traditional else ("", "万", "亿", "兆")
    if number == 0:
        return digits[0]
    if number < 0 or number >= 10 ** 16:
        raise NumberingConversionError(f"中文编号数值超出支持范围: {number}")
    chunks = []
    while number:
        number, value = divmod(number, 10000)
        chunks.append(value)
    result = ""
    zero_pending = False
    for index in range(len(chunks) - 1, -1, -1):
        value = chunks[index]
        if not value:
            zero_pending = bool(result)
            continue
        if result and (zero_pending or value < 1000):
            result += digits[0]
        text, inner_zero = "", False
        for power in range(3, -1, -1):
            digit = value // (10 ** power) % 10
            if digit:
                if inner_zero:
                    text += digits[0]
                text += digits[digit] + units[power]
                inner_zero = False
            elif text:
                inner_zero = True
        result += text + groups[index]
        zero_pending = False
    if not legal and not traditional and result.startswith("一十"):
        result = result[1:]
    return result


def _japanese(number, legal=False):
    """Japanese unit-based numerals, without intervening zero characters.

    Unlike Chinese counting, 101 is 百一 and 1001 is 千一. A coefficient
    of one is omitted for ten/hundred/thousand, but retained for 万/億.
    This range has been checked against WPS at signed 32-bit boundaries.
    """
    if not 0 <= number <= 2147483647:
        raise NumberingConversionError(f"日式编号数值超出已验证范围: {number}")
    digits = "〇壱弐参四五六七八九" if legal else "〇一二三四五六七八九"
    units = ("", "拾", "百", "阡") if legal else ("", "十", "百", "千")
    groups = ("", "萬", "億") if legal else ("", "万", "億")
    if number == 0:
        return digits[0]
    chunks = []
    while number:
        number, chunk = divmod(number, 10000)
        chunks.append(chunk)
    result = ""
    for index in range(len(chunks)-1, -1, -1):
        value = chunks[index]
        if not value:
            continue
        for power in range(3, -1, -1):
            digit = value // (10 ** power) % 10
            if digit:
                result += (digits[digit] if digit != 1 or power == 0 else "") + units[power]
        result += groups[index]
    return result


def format_number(value: int, fmt: str) -> str:
    """Render the supported Word numFmt values; never guess unknown formats."""
    if fmt == "decimal":
        return str(value)
    if fmt == "decimalZero":
        return f"{value:02d}"
    if fmt in ("decimalFullWidth", "decimalFullWidth2"):
        return str(value).translate(str.maketrans("0123456789", "０１２３４５６７８９"))
    if fmt in ("upperLetter", "lowerLetter"):
        if value < 1:
            raise NumberingConversionError("字母编号从 0 开始时各办公软件显示不一致")
        text = _letters(value)
        return text if fmt == "upperLetter" else text.lower()
    if fmt in ("upperRoman", "lowerRoman"):
        if not 1 <= value <= 999:
            raise NumberingConversionError("罗马数字编号暂只支持 1–999，超出时各办公软件显示不一致")
        text = _roman(value)
        return text if fmt == "upperRoman" else text.lower()
    if fmt in ("chineseCounting", "chineseCountingThousand"):
        return _chinese(value)
    if fmt == "chineseLegalSimplified":
        return _chinese(value, legal=True)
    if fmt in ("japaneseCounting", "japaneseLegal"):
        return _japanese(value, legal=fmt == "japaneseLegal")
    if fmt in ("ideographTraditional", "ideographZodiac"):
        symbols = {"ideographTraditional": "甲乙丙丁戊己庚辛壬癸",
                   "ideographZodiac": "子丑寅卯辰巳午未申酉戌亥"}[fmt]
        return symbols[value - 1] if 1 <= value <= len(symbols) else str(value)
    if fmt == "decimalEnclosedCircle":
        if 1 <= value <= 20:
            return chr(0x2460 + value - 1)
        return str(value)
    if fmt == "decimalEnclosedParen":
        return chr(0x2474 + value - 1) if 1 <= value <= 20 else str(value)
    if fmt == "decimalEnclosedFullstop":
        return chr(0x2488 + value - 1) if 1 <= value <= 20 else str(value)
    if fmt == "ordinal":
        suffix = "th" if 10 <= value % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(value % 10, "th")
        return str(value) + suffix
    if fmt == "hex":
        return format(value, "X")
    if fmt == "none":
        return ""
    raise NumberingConversionError(f"暂不支持编号格式 {fmt!r}")


class _Definitions:
    def __init__(self, styles, numbering, ns):
        self.ns = ns
        self.settings = None
        self.has_character_grid = False
        self.styles = {} if styles is None else {
            e.get(_tag(ns, "styleId")): e for e in styles.findall(_tag(ns, "style"))}
        self.default_style = next((name for name, style in self.styles.items()
                                   if style.get(_tag(ns, "type")) == "paragraph"
                                   and style.get(_tag(ns, "default")) in ("1", "true", "on")), None)
        self.defaults = _child(styles, ns, "docDefaults")
        self.nums = {} if numbering is None else {
            e.get(_tag(ns, "numId")): e for e in numbering.findall(_tag(ns, "num"))}
        self.abstracts = {} if numbering is None else {
            e.get(_tag(ns, "abstractNumId")): e for e in numbering.findall(_tag(ns, "abstractNum"))}
        self._level_cache = {}
        self._chain_cache = {}

    def style_chain(self, name):
        name = name or self.default_style
        original_name = name
        if name in self._chain_cache:
            return self._chain_cache[name]
        chain, seen = [], set()
        while name and name in self.styles:
            if name in seen:
                raise NumberingConversionError(f"样式继承存在循环: {name}")
            seen.add(name)
            style = self.styles[name]
            chain.append(style)
            name = _val(style, self.ns, "basedOn")
        chain.reverse()
        self._chain_cache[original_name] = chain
        return chain

    def paragraph_properties(self, paragraph):
        ns = self.ns
        direct = _child(paragraph, ns, "pPr")
        name = _val(direct, ns, "pStyle") or self.default_style
        chain = self.style_chain(name)
        default = _child(_child(self.defaults, ns, "pPrDefault"), ns, "pPr")
        inherited = [default] + [_child(style, ns, "pPr") for style in chain]
        number, inherited_level = None, None
        for properties in inherited + [direct]:
            numpr = _child(properties, ns, "numPr")
            if numpr is not None:
                number = _val(numpr, ns, "numId", number)
        for properties in inherited:
            inherited_level = _val(_child(properties, ns, "numPr"), ns, "ilvl", inherited_level)
        if number in (None, "0"):
            return None
        # Prefer pStyle bindings; fall back to a style's ilvl for documents
        # produced by Word/WPS which use it without a corresponding binding.
        level = _val(_child(direct, ns, "numPr"), ns, "ilvl")
        levels = self.levels(number)
        if level is None:
            ids = [style.get(_tag(ns, "styleId")) for style in reversed(chain)]
            for style_id in ids:
                match = next((i for i, definition in levels.items()
                              if _val(definition, ns, "pStyle") == style_id), None)
                if match is not None:
                    level = match
                    break
        level = _int(inherited_level if level is None else level)
        if not 0 <= level <= 8 or level not in levels:
            raise NumberingConversionError(f"编号 {number} 缺少有效的第 {level + 1} 级定义")
        return number, level, levels, inherited, chain

    def levels(self, number, seen=()):
        if number in self._level_cache:
            return self._level_cache[number]
        if number in seen:
            raise NumberingConversionError(f"编号样式链接存在循环: {number}")
        ns = self.ns
        instance = self.nums.get(number)
        if instance is None:
            raise NumberingConversionError(f"找不到编号实例 numId={number}")
        abstract_id = _val(instance, ns, "abstractNumId")
        abstract = self.abstracts.get(abstract_id)
        if abstract is None:
            raise NumberingConversionError(f"找不到编号定义 abstractNumId={abstract_id}")
        linked_style = _val(abstract, ns, "numStyleLink")
        if linked_style:
            linked_number = None
            for style in self.style_chain(linked_style):
                numpr = _child(_child(style, ns, "pPr"), ns, "numPr")
                linked_number = _val(numpr, ns, "numId", linked_number)
            if linked_number is None:
                raise NumberingConversionError(f"编号样式链接没有 numId: {linked_style}")
            levels = {i: deepcopy(e) for i, e in self.levels(linked_number, seen + (number,)).items()}
        else:
            levels = {}
        for definition in abstract.findall(_tag(ns, "lvl")):
            levels[_int(definition.get(_tag(ns, "ilvl")))] = deepcopy(definition)
        for override in instance.findall(_tag(ns, "lvlOverride")):
            index = _int(override.get(_tag(ns, "ilvl")))
            definition = _child(override, ns, "lvl")
            if definition is not None:
                # lvl is a complete overridden level, not a delta to merge.
                levels[index] = deepcopy(definition)
            start = _child(override, ns, "startOverride")
            if start is not None and index in levels:
                old = _child(levels[index], ns, "start")
                if old is None:
                    old = etree.SubElement(levels[index], _tag(ns, "start"))
                old.set(_tag(ns, "val"), start.get(_tag(ns, "val"), "1"))
        self._level_cache[number] = levels
        return levels

    def restarts_after_section(self, number):
        instance = self.nums.get(number)
        abstract = self.abstracts.get(_val(instance, self.ns, "abstractNumId"))
        return abstract is not None and abstract.get(_tag(W15, "restartNumberingAfterBreak")) in ("1", "true", "on")

    def layout_cascade(self, paragraph, inherited, chain):
        """Resolve list layout where each numPr occurs in the inheritance.

        A new numPr (including an ilvl-only override) reapplies that level's
        layout at that point. Properties directly on the same style/paragraph
        then override it. Moving all list properties ahead of all styles loses
        this distinction, e.g. a direct level switch would keep an old indent.
        """
        ns = self.ns
        result, number, level = [], None, 0
        layers = inherited + [_child(paragraph, ns, "pPr")]
        for index, properties in enumerate(layers):
            numpr = _child(properties, ns, "numPr")
            if numpr is not None:
                number = _val(numpr, ns, "numId", number)
                explicit = _val(numpr, ns, "ilvl")
                if explicit is not None:
                    level = _int(explicit)
                if number not in (None, "0"):
                    levels = self.levels(number)
                    if explicit is None and 0 < index <= len(chain):
                        style_id = chain[index - 1].get(_tag(ns, "styleId"))
                        level = next((i for i,e in levels.items() if _val(e, ns, "pStyle") == style_id), level)
                    if level not in levels:
                        raise NumberingConversionError(f"编号 {number} 缺少缩进层级 {level + 1}")
                    result.append(_child(levels[level], ns, "pPr"))
            result.append(properties)
        return result


class _Counters:
    def __init__(self, ns):
        self.ns = ns
        self.values = {}
        self.pending = {}

    def section_break(self, definitions):
        for number in list(self.values):
            if definitions.restarts_after_section(number):
                self.values.pop(number, None)
                self.pending.pop(number, None)

    def advance(self, number, index, levels):
        ns = self.ns
        values = self.values.setdefault(number, {})
        pending = self.pending.setdefault(number, set())
        # Jumping directly to a child uses the parents' initial numbers too.
        # They have now appeared implicitly, so their next explicit occurrence
        # increments them rather than repeating the same initial number.
        for parent in range(index):
            if parent in levels and (parent not in values or parent in pending):
                values[parent] = _int(_val(levels[parent], ns, "start", "1"))
                pending.discard(parent)
        if index not in values or index in pending:
            values[index] = _int(_val(levels[index], ns, "start", "1"))
            pending.discard(index)
        else:
            values[index] += 1
        for deeper, definition in levels.items():
            if deeper <= index:
                continue
            restart = _val(definition, ns, "lvlRestart")
            # Ancestor changes also invalidate descendants through the
            # intervening level; an explicit zero disables this completely.
            trigger = deeper - 1 if restart is None else _int(restart) - 1
            if restart != "0" and index <= trigger:
                pending.add(deeper)
        return {i: (_int(_val(e, ns, "start", "1")) if i not in values or i in pending else values[i])
                for i, e in levels.items()}


def _label(levels, index, values, ns):
    definition = levels[index]
    if _child(definition, ns, "lvlPicBulletId") is not None:
        raise NumberingConversionError("图片项目符号无法转换为普通文字")
    fmt = _val(definition, ns, "numFmt", "decimal")
    pattern = _val(definition, ns, "lvlText", "" if fmt == "none" else "%" + str(index + 1) + ".")
    if fmt == "bullet":
        if PLACEHOLDER.search(pattern):
            raise NumberingConversionError("项目符号模板含 %n，Word 与 WPS 的显示规则不同")
        return pattern
    # Validate even when lvlText is literal and has no placeholder.
    format_number(values[index], fmt)
    legal = _on(definition, ns, "isLgl")

    def substitute(match):
        level = int(match.group(1)) - 1
        if level not in levels:
            raise NumberingConversionError(f"编号模板引用缺失的第 {level + 1} 级")
        return format_number(values[level], "decimal" if legal else _val(levels[level], ns, "numFmt", "decimal"))

    return PLACEHOLDER.sub(substitute, pattern)


def _story_groups(root, ns):
    """Each part is a story; isolated text boxes have their own counters.

    All footnotes in one part share counters, as do all endnotes. The same
    numId in the main story does not affect the counters in these parts.
    """
    groups = {}
    for paragraph in root.iter(_tag(ns, "p")):
        ancestors = list(paragraph.iterancestors())
        if any(a.tag in (_tag(ns, "del"), _tag(ns, "moveFrom")) for a in ancestors):
            continue
        note = next((a for a in ancestors if a.tag in (_tag(ns, "footnote"), _tag(ns, "endnote"))), None)
        if note is not None and (note.get(_tag(ns, "type"), "normal") != "normal"
                                 or _int(note.get(_tag(ns, "id"))) < 0):
            continue
        owner = next((a for a in ancestors if a.tag == _tag(ns, "txbxContent")), root)
        groups.setdefault(owner, []).append(paragraph)
    return groups.values()


def _convert_paragraph(paragraph, definition, label, inherited, chain, definitions, preserve_suffix):
    ns = definitions.ns
    ppr = _child(paragraph, ns, "pPr")
    if ppr is None:
        ppr = etree.Element(_tag(ns, "pPr"))
        paragraph.insert(0, ppr)
    # Freeze the effective layout BEFORE replacing numPr. In particular,
    # ilvl-only overrides reapply list indentation after the inherited style.
    cascade = definitions.layout_cascade(paragraph, inherited, chain)
    for name in ("ind", "tabs"):
        direct = _child(ppr, ns, name)
        if any(_child(layer, ns, name) is not None for layer in cascade):
            if name == "ind":
                effective = _effective_indent(cascade, ns)
                if direct is None:
                    _insert_property(ppr, effective, ns)
                else:
                    direct.attrib.clear()
                    direct.attrib.update(effective.attrib)
            else:
                effective = _effective_tabs(cascade, ns)
                if direct is None:
                    _insert_property(ppr, effective, ns)
                else:
                    direct[:] = list(effective)
    numpr = _child(ppr, ns, "numPr")
    if numpr is None:
        numpr = etree.Element(_tag(ns, "numPr"))
        _insert_property(ppr, numpr, ns)
    # Explicit suppression also overrides numbering inherited from styles.
    numpr[:] = []
    etree.SubElement(numpr, _tag(ns, "numId")).set(_tag(ns, "val"), "0")
    if not label:
        return
    run = etree.Element(_tag(ns, "r"))
    default_rpr = _child(_child(definitions.defaults, ns, "rPrDefault"), ns, "rPr")
    rpr = _merge_properties([default_rpr] + [_child(s, ns, "rPr") for s in chain]
                            + [_child(ppr, ns, "rPr"), _child(definition, ns, "rPr")], ns, "rPr")
    if len(rpr):
        run.append(rpr)
    # The default conversion joins the label directly to the original runs.
    # Do not strip the body's own spaces/tabs: only omit the generated suffix.
    suffix = _val(definition, ns, "suff", "tab") if preserve_suffix else "nothing"
    if (suffix == "tab" and not definitions.has_character_grid
            and _cjk_label_fills_hanging_indent(label, rpr, definition, cascade,
                                               definitions.settings, ns)):
        suffix = "nothing"
    text = etree.SubElement(run, _tag(ns, "t"))
    text.text = label + (" " if suffix == "space" else "")
    text.set(XML_SPACE, "preserve")
    if suffix == "tab":
        etree.SubElement(run, _tag(ns, "tab"))
    paragraph.insert(paragraph.index(ppr) + 1, run)


def _read_xml(archive, name):
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    return etree.fromstring(archive.read(name), parser)


def _related_parts(archive, main):
    directory, filename = posixpath.split(main)
    rels = posixpath.join(directory, "_rels", filename + ".rels")
    result = {}
    if rels in archive.namelist():
        for rel in _read_xml(archive, rels):
            if rel.get("TargetMode") == "External":
                continue
            kind = rel.get("Type", "").rsplit("/", 1)[-1]
            target = rel.get("Target", "")
            path = posixpath.normpath(posixpath.join(directory, target)) if not target.startswith("/") else target[1:]
            if path.startswith("../"):
                raise NumberingConversionError("文档关系指向包外路径")
            if kind in ("styles", "numbering", "settings"):
                result[kind] = path
    return result


def convert_docx_numbering(input_path, output_path=None, *, overwrite=False, strict=True,
                           preserve_suffix=False):
    """Write a converted copy and return a ConversionReport.

    strict=True refuses unsupported numbering rather than silently changing it
    to decimal. With strict=False supported paragraphs are converted and every
    unsupported paragraph is reported and left intact. Source files are never
    overwritten; running on a converted document does not add labels twice.
    By default labels directly precede the original content, without a new
    tab/space suffix. Paragraph indents and original whitespace are retained,
    but body text can move closer to the label. preserve_suffix=True enables
    the previous spacing behavior, including the CJK boundary correction.
    """
    source = Path(input_path).resolve()
    destination = Path(output_path).resolve() if output_path else source.with_name(source.stem + "_numbering_text.docx")
    if source == destination:
        raise ValueError("输出路径必须与原文档不同")
    if destination.exists() and not overwrite:
        raise FileExistsError(f"输出文件已存在: {destination}")
    report = ConversionReport()
    replacements = {}
    with ZipFile(source) as archive:
        names = archive.namelist()
        if len(set(names)) != len(names):
            raise NumberingConversionError("DOCX 包中存在重复文件名")
        types = _read_xml(archive, "[Content_Types].xml")
        overrides = [(e.get("PartName", "").lstrip("/"), e.get("ContentType", "")) for e in types]
        mains = [name for name, kind in overrides if kind.endswith("wordprocessingml.document.main+xml")]
        if len(mains) != 1:
            raise NumberingConversionError("未找到唯一的 DOCX 主文档")
        main = mains[0]
        root = _read_xml(archive, main)
        ns = etree.QName(root).namespace
        if ns not in (W, "http://purl.oclc.org/ooxml/wordprocessingml/main"):
            raise NumberingConversionError("主文档使用了未知的 WordprocessingML 命名空间")
        related = _related_parts(archive, main)
        definitions = _Definitions(*[
            _read_xml(archive, related[key]) if key in related else None for key in ("styles", "numbering")], ns)
        if "settings" in related:
            definitions.settings = _read_xml(archive, related["settings"])
        definitions.has_character_grid = any(
            grid.get(_tag(ns, "type")) in ("linesAndChars", "snapToChars")
            or _int(grid.get(_tag(ns, "charSpace"))) != 0
            for grid in root.iter(_tag(ns, "docGrid")))
        story_types = (".header+xml", ".footer+xml", ".footnotes+xml", ".endnotes+xml", ".comments+xml")
        parts = [main] + [name for name, kind in overrides
                          if kind.startswith("application/vnd.openxmlformats-officedocument.wordprocessingml")
                          and kind.endswith(story_types)]
        for part in parts:
            tree = root if part == main else _read_xml(archive, part)
            plans = []
            # Numbering fields are a separate mechanism, requiring a field
            # evaluator. Do not turn unrelated PAGE/TOC/REF fields into text.
            instructions = " ".join((e.text or "") for e in tree.iter(_tag(ns, "instrText")))
            instructions += " " + " ".join(e.get(_tag(ns, "instr"), "") for e in tree.iter(_tag(ns, "fldSimple")))
            if NUMBER_FIELD.search(instructions):
                report.warnings.append(f"{part}: 包含暂不支持的 LISTNUM/AUTONUM 编号字段")
            for group in _story_groups(tree, ns):
                counters = _Counters(ns)
                group_plans, families, family_seen, family_last, ambiguous = [], {}, {}, {}, set()
                for index, paragraph in enumerate(group, 1):
                    try:
                        info = definitions.paragraph_properties(paragraph)
                        if info is None:
                            continue
                        number, level, levels, inherited, chain = info
                        family = _val(definitions.nums[number], ns, "abstractNumId")
                        seen = family_seen.setdefault(family, set())
                        if number in seen and family_last.get(family) != number:
                            ambiguous.add(family)
                        seen.add(number)
                        family_last[family] = number
                        families[id(paragraph)] = family
                        if any(a.tag == _tag(MC, "AlternateContent") for a in paragraph.iterancestors()):
                            raise NumberingConversionError("编号位于兼容性备选文本框内，暂不支持")
                        numpr = _child(_child(paragraph, ns, "pPr"), ns, "numPr")
                        if (_child(numpr, ns, "numberingChange") is not None or _child(numpr, ns, "ins") is not None
                                or _child(_child(paragraph, ns, "pPr"), ns, "pPrChange") is not None):
                            raise NumberingConversionError("编号存在修订，请先接受修订")
                        values = counters.advance(number, level, levels)
                        label = _label(levels, level, values, ns)
                        if _val(levels[level], ns, "suff", "tab") not in ("tab", "space", "nothing"):
                            raise NumberingConversionError("未知的编号后分隔符")
                        group_plans.append((paragraph, levels[level], label, inherited, chain))
                    except (NumberingConversionError, ValueError) as exc:
                        report.warnings.append(f"{part} 段落 {index}: {exc}")
                    finally:
                        if _child(_child(paragraph, ns, "pPr"), ns, "sectPr") is not None:
                            counters.section_break(definitions)
                for family in sorted(ambiguous):
                    report.warnings.append(f"{part}: 多个 numId 共享 abstractNumId={family} 并交错使用，"
                                           "WPS 与独立列表计数规则不同；该组编号未转换")
                plans.extend(p for p in group_plans if families[id(p[0])] not in ambiguous)
            if plans:
                for paragraph, definition, label, inherited, chain in plans:
                    _convert_paragraph(paragraph, definition, label, inherited, chain, definitions,
                                       preserve_suffix)
                replacements[part] = etree.tostring(tree, encoding="UTF-8", xml_declaration=True,
                                                    standalone=archive.read(part).lstrip().split(b"?>", 1)[0].find(b'standalone="yes"') >= 0)
                report.parts.append(part)
                report.converted += len(plans)
        if report.warnings and strict:
            raise NumberingConversionError("转换未写出，以下编号需要确认：\n" + "\n".join(report.warnings))
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(dir=destination.parent, suffix=".docx", delete=False) as handle:
                temporary = Path(handle.name)
            with ZipFile(temporary, "w") as output:
                output.comment = archive.comment
                for info in archive.infolist():
                    data = replacements[info.filename] if info.filename in replacements else archive.read(info)
                    output.writestr(info, data)
            os.replace(temporary, destination)
        finally:
            if temporary is not None and temporary.exists():
                temporary.unlink()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description="独立 DOCX 自动编号转文字测试工具（包含表格、页眉页脚及文本框）")
    parser.add_argument("input", help="输入 DOCX 文件")
    parser.add_argument("-o", "--output", help="输出 DOCX 文件；默认在原文件名后加 _numbering_text")
    parser.add_argument("--overwrite", action="store_true", help="允许覆盖已有输出文件；不能覆盖原文档")
    parser.add_argument("--allow-partial", action="store_true", help="保留不支持的编号并报告；默认遇到不支持的编号不写出文件")
    parser.add_argument("--preserve-suffix", action="store_true", help="保留编号后的制表符/空格以维持间距；默认编号直接连接原文")
    args = parser.parse_args(argv)
    try:
        report = convert_docx_numbering(args.input, args.output, overwrite=args.overwrite,
                                       strict=not args.allow_partial, preserve_suffix=args.preserve_suffix)
    except (ValueError, OSError, etree.XMLSyntaxError, BadZipFile, KeyError) as exc:
        parser.exit(1, f"转换失败: {exc}\n")
    print(json.dumps(report.__dict__, ensure_ascii=False, indent=2))
    return 2 if report.warnings else 0


if __name__ == "__main__":
    raise SystemExit(main())
