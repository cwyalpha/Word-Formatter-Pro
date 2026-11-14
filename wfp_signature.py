"""Conservative signature detection, date parsing and cross-platform font metrics.

This module has no GUI or Office dependency. It never creates or removes paragraphs.
"""
from datetime import date
from contextlib import contextmanager
from copy import deepcopy
from functools import lru_cache
import os
import math
from pathlib import Path
import re
import sys
import struct
import unicodedata

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run

try:
    from fontTools.ttLib import TTLibError
except ImportError:
    class TTLibError(Exception):
        pass

_FONT_ERRORS = (ImportError, OSError, ValueError, KeyError, TTLibError)


_DIGITS = dict(zip('零〇○Ｏ一二三四五六七八九', (0, 0, 0, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9)))
_DATE_PART = r'[0-9零〇○Ｏ一二三四五六七八九十]+'
_DATE_CN = re.compile(r'^(?P<year>' + _DATE_PART + r')\s*年\s*(?P<month>' + _DATE_PART
                      + r')\s*月\s*(?P<day>' + _DATE_PART + r')\s*日$')
_DATE_NUMERIC = re.compile(r'^(?P<year>[0-9]{4})(?P<sep>[-/])(?P<month>[0-9]{1,2})'
                           r'(?P=sep)(?P<day>[0-9]{1,2})$')
_ORG_SUFFIX = re.compile(r'(?:人民政府|委员会|人民法院|人民检察院|研究院|研究所|办公室|办事处|'
                         r'工作组|领导小组|指挥部|管理局|总公司|分公司|公司|学校|大学|学院|医院|'
                         r'协会|学会|中心|党委|党组|党支部|党总支|总工会|工会|团委|'
                         r'政府|省委|市委|县委|区委|乡委|镇委|街道办|机关|[部厅局处科室院所站社厂])$')
_REVISIONS = ('ins', 'del', 'moveFrom', 'moveTo', 'pPrChange', 'rPrChange',
              'moveFromRangeStart', 'moveToRangeStart', 'numberingChange')
_UNSAFE = ('drawing', 'pict', 'object', 'fldChar', 'instrText', 'fldSimple', 'sdt',
           'footnoteReference', 'endnoteReference', 'sym', 'altChunk') + _REVISIONS
_TAIL_NOTE = re.compile(r'^[（(](?:此件|本文|本件|联系人|联系电话|联系方式|公开|不予公开|主动公开|依申请公开)')
_TAIL_COPY = re.compile(r'^(?:抄送|抄报|印发机关|印发日期|印送)[：:]')


def _chinese_number(value, year=False):
    if value.isascii() and value.isdigit():
        if year and len(value) != 4:
            raise ValueError('year must have four digits')
        return int(value)
    if year:
        if len(value) != 4 or any(ch not in _DIGITS for ch in value):
            raise ValueError('year must have four digits')
        return int(''.join(str(_DIGITS[ch]) for ch in value))
    if '十' in value:
        if value.count('十') != 1:
            raise ValueError('invalid Chinese number')
        left, right = value.split('十')
        if (left and (len(left) != 1 or left not in _DIGITS or _DIGITS[left] == 0)
                or right and (len(right) != 1 or right not in _DIGITS or _DIGITS[right] == 0)):
            raise ValueError('invalid Chinese number')
        return (_DIGITS[left] if left else 1) * 10 + (_DIGITS[right] if right else 0)
    if not 1 <= len(value) <= 2 or any(ch not in _DIGITS for ch in value):
        raise ValueError('invalid Chinese number')
    return int(''.join(str(_DIGITS[ch]) for ch in value))


def parse_signature_date(text):
    """Return a real calendar date, or None; never infer a missing year/month/day."""
    value = unicodedata.normalize('NFKC', text.strip())
    match = _DATE_CN.fullmatch(value) or _DATE_NUMERIC.fullmatch(value)
    if match is None:
        return None
    try:
        return date(_chinese_number(match['year'], year=True), _chinese_number(match['month']),
                    _chinese_number(match['day']))
    except (ValueError, KeyError, OverflowError):
        return None


def _font_key(name):
    return ''.join(ch for ch in name.casefold() if ch.isalnum())


@lru_cache(maxsize=1)
def _font_candidates():
    """Read font locations, including per-user fonts; do not run desktop utilities."""
    candidates = {}
    folders = []
    if sys.platform.startswith('win'):
        windows_fonts = Path(os.environ.get('WINDIR', r'C:\Windows')) / 'Fonts'
        folders = [windows_fonts, Path(os.environ.get('LOCALAPPDATA', str(Path.home())))
                   / 'Microsoft' / 'Windows' / 'Fonts']
        try:
            import winreg
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(hive, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts') as handle:
                        index = 0
                        while True:
                            try:
                                label, filename, _ = winreg.EnumValue(handle, index)
                            except OSError:
                                break
                            index += 1
                            if not isinstance(filename, str):
                                continue
                            path = Path(os.path.expandvars(filename))
                            if not path.is_absolute():
                                path = windows_fonts / path
                            if path.is_file():
                                candidates.setdefault(str(path), set()).add(_font_key(label))
                except OSError:
                    continue
        except ImportError:
            pass
    elif sys.platform == 'darwin':
        folders = [Path('/System/Library/Fonts'), Path('/Library/Fonts'), Path.home() / 'Library/Fonts']
    else:
        folders = [Path('/usr/share/fonts'), Path('/usr/local/share/fonts'), Path.home() / '.fonts',
                   Path(os.environ.get('XDG_DATA_HOME', str(Path.home() / '.local/share'))) / 'fonts']
        folders += [Path(folder) / 'fonts' for folder in
                    os.environ.get('XDG_DATA_DIRS', '/usr/local/share:/usr/share').split(os.pathsep) if folder]
    for folder in folders:
        try:
            for path in folder.rglob('*'):
                if path.suffix.lower() in ('.ttf', '.otf', '.ttc', '.otc') and path.is_file():
                    candidates.setdefault(str(path), set())
        except OSError:
            continue
    return tuple((path, tuple(labels)) for path, labels in candidates.items())


@lru_cache(maxsize=1024)
def _font_faces(path):
    if Path(path).suffix.lower() in ('.ttc', '.otc'):
        with open(path, 'rb') as stream:
            header = stream.read(12)
        if len(header) != 12:
            raise ValueError('incomplete font collection')
        tag, _, count = struct.unpack('>4sII', header)
        if tag != b'ttcf' or not 1 <= count <= 1024:
            raise ValueError('invalid font collection')
    else:
        count = 1
    faces = []
    for index in range(count):
        with _open_font(path, index) as font:
            names = set()
            for record in font['name'].names:
                if record.nameID in (1, 4, 6, 16):
                    try:
                        names.add(_font_key(record.toUnicode()))
                    except (UnicodeError, LookupError):
                        continue
            flags = font['head'].macStyle
            faces.append((index, frozenset(names), bool(flags & 1), bool(flags & 2)))
    return tuple(faces)


@contextmanager
def _open_font(path, index):
    from fontTools.ttLib import TTFont
    # Also close the input when TTFont rejects a corrupt file during construction.
    with open(path, 'rb') as stream:
        font = TTFont(stream, fontNumber=index, lazy=True)
        try:
            yield font
        finally:
            font.close()


@lru_cache(maxsize=128)
def _find_font(name, bold, italic):
    key = _font_key(name)
    candidates = sorted(_font_candidates(), key=lambda item: not any(key in label for label in item[1]))
    best, score = None, 3
    for path, _ in candidates:
        try:
            for index, names, face_bold, face_italic in _font_faces(path):
                if key not in names:
                    continue
                difference = int(face_bold != bold) + int(face_italic != italic)
                if difference < score:
                    best, score = (path, index), difference
                if score == 0:
                    return best
        except _FONT_ERRORS:
            continue
    return best


@lru_cache(maxsize=64)
def _font_advances(path, index):
    with _open_font(path, index) as font:
        cmap = font.getBestCmap() or {}
        metrics = font['hmtx'].metrics
        units = font['head'].unitsPerEm
        return {code: metrics[glyph][0] / units for code, glyph in cmap.items() if glyph in metrics}


class FontMetrics:
    def __init__(self, log):
        self.log = log
        self.warned = set()

    def measure(self, text, name, size, bold=False, italic=False):
        advances = None
        try:
            found = _find_font(name, bold, italic)
            if found:
                advances = _font_advances(*found)
        except _FONT_ERRORS:
            pass
        if advances is None or any(ord(ch) not in advances for ch in text):
            if name not in self.warned:
                self.warned.add(name)
                self.log('落款测宽：字体“{}”不可用或缺少字形，使用全角／半角宽度估算；'
                         '字体替换后的最终位置请在 Office/WPS 检查。'.format(name))
        total = 0.0
        for ch in text:
            if advances is not None and ord(ch) in advances:
                width = advances[ord(ch)]
            elif unicodedata.combining(ch):
                width = 0.0
            else:
                width = 1.0 if unicodedata.east_asian_width(ch) in ('W', 'F', 'A') else 0.5
            total += width * size
        return total


class SignatureFormatter:
    """Detect on original paragraphs; apply only to a confirmed, safe pair."""
    def __init__(self, processor):
        self.processor = processor
        self.config = processor.config
        self.log = processor._log
        self.metrics = FontMetrics(self.log)

    def text(self, para):
        return self.processor._paragraph_text(para)

    @staticmethod
    def has(para, tags):
        return any(para._p.find('.//' + qn('w:' + tag)) is not None for tag in tags)

    def blank(self, block):
        return (isinstance(block, Paragraph) and not self.text(block).strip()
                and not self.has(block, _UNSAFE + ('br', 'cr', 'sectPr', 'bookmarkStart', 'bookmarkEnd',
                                                  'commentRangeStart', 'commentRangeEnd')))

    def organization(self, para):
        text = self.text(para).strip()
        pieces = re.split(r'\s+', text)
        if len(pieces) > 1 and sum(bool(_ORG_SUFFIX.search(piece)) for piece in pieces) > 1:
            return False
        return (2 <= len(text) <= 60 and bool(_ORG_SUFFIX.search(text))
                and not re.search(r'[。；;：:！？!?\n\t、，,]|[（(](?:公章|印章|盖章|签章)', text)
                and not text.startswith(('附件', '抄送', '抄报', '印发', '主送'))
                and not re.match(r'^(?:请|由|联系|详见|根据|按照|本次|有关|如需|负责|参加|截至)', text)
                and not re.match(r'^(?:[一二三四五六七八九十]+[、.]|[（(]|[0-9]+[.、])', text))

    def tail(self, para):
        text = self.text(para).strip()
        if self.has(para, _UNSAFE):
            return False
        return bool((_TAIL_NOTE.match(text) and text.endswith(('）', ')')))
                    or _TAIL_COPY.match(text)
                    or re.fullmatch(r'.{2,60}\s+[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日印发', text))

    def safe(self, para):
        if self.has(para, _UNSAFE + ('br', 'cr', 'sectPr', 'numPr')):
            return False
        text = self.text(para)
        if '\n' in text or '\t' in text or (text.strip() and text != text.rstrip()):
            return False
        if self.has(para, ('bidi', 'textDirection')):
            return False
        for run in self.processor._iter_inline_runs(para):
            for tag in ('spacing', 'w', 'position', 'vertAlign', 'rtl'):
                element = run._r.find('./' + qn('w:rPr') + '/' + qn('w:' + tag))
                if element is not None:
                    value = element.get(qn('w:val'))
                    if (tag == 'w' and value == '100') or (tag in ('spacing', 'position') and value == '0'):
                        continue
                    return False
        return True

    def classify(self, blocks, attachment_roles, excluded=()):
        """Return a map keyed by paragraph XML, including roles to protect from body rules."""
        boundaries = [i for i, block in enumerate(blocks) if isinstance(block, Paragraph)
                      and attachment_roles.get(block._p, {}).get('kind') == 'body']
        starts = [0] + boundaries
        ends = boundaries + [len(blocks)]
        roles = {}
        for section_index, (start, end) in enumerate(zip(starts, ends)):
            if section_index and not self.config.get('signature_include_attachments', False):
                continue
            label = '主文' if section_index == 0 else '附件{}'.format(section_index)
            last = end - 1
            while last >= start:
                block = blocks[last]
                if self.blank(block) or isinstance(block, Paragraph) and self.tail(block):
                    last -= 1
                    continue
                break
            if last < start or not isinstance(blocks[last], Paragraph):
                continue
            date_para = blocks[last]
            parsed = parse_signature_date(self.text(date_para))
            if parsed is None:
                if self.has(date_para, _UNSAFE):
                    self.log('落款识别：{}末尾含图片、字段或修订，不跨过该区域寻找落款。'.format(label))
                continue
            previous = last - 1
            while previous >= start and self.blank(blocks[previous]):
                previous -= 1
            if previous < start or not isinstance(blocks[previous], Paragraph):
                self.log('落款识别：{}末尾日期缺少独立机关署名，跳过专项调整。'.format(label))
                continue
            unit = blocks[previous]
            if not self.organization(unit) or unit._p in excluded or unit._p in attachment_roles:
                self.log('落款识别：{}末尾日期的前一段无法确认为机关署名，跳过专项调整。'.format(label))
                continue
            before = previous - 1
            while before >= start and self.blank(blocks[before]):
                before -= 1
            if before >= start and isinstance(blocks[before], Paragraph) and self.organization(blocks[before]):
                self.log('落款识别：{}疑似联合署名，首版保留原排法。'.format(label))
                continue
            if not all(self.safe(blocks[i]) if isinstance(blocks[i], Paragraph) else False
                       for i in range(previous, last + 1)):
                self.log('落款识别：{}署名与日期区域含图片、字段、修订、分页、制表符、尾随空白或复杂布局，跳过专项调整。'.format(label))
                continue
            # A picture/anchor immediately before a pair can be its stamp. Do not reposition that pair.
            if before >= start and isinstance(blocks[before], Paragraph) and self.has(blocks[before], _UNSAFE):
                self.log('落款识别：{}署名前邻接图片、字段或修订，跳过专项调整。'.format(label))
                continue
            info = {'unit': unit, 'date': date_para, 'parsed': parsed, 'label': label,
                    'blank_before': (previous > start and isinstance(blocks[previous - 1], Paragraph)
                                     and not self.text(blocks[previous - 1]).strip()
                                     and not self.has(blocks[previous - 1], _UNSAFE + ('br', 'cr', 'sectPr'))),
                    'between': blocks[previous + 1:last], 'index': previous + 1}
            roles[unit._p] = {'info': info, 'kind': 'unit'}
            roles[date_para._p] = {'info': info, 'kind': 'date'}
            self.log('落款识别：{}第{}段，“{}”／“{}”。'.format(label, previous + 1,
                     self.text(unit).strip(), self.text(date_para).strip()))
        return roles

    @staticmethod
    def _effective_flag(para, run, key):
        if run is not None and getattr(run.font, key) is not None:
            return getattr(run.font, key)
        for style in (run.style if run is not None else None, para.style):
            seen = set()
            while style is not None and style.style_id not in seen:
                seen.add(style.style_id)
                value = getattr(style.font, key)
                if value is not None:
                    return value
                style = style.base_style
        return False

    def _measure_paragraph(self, para, replacement=None, trim=True):
        size = int(float(self.config.get('body_size', 16)) * 2) / 2
        body_font = self.config.get('body_font', '仿宋_GB2312')
        english_font = self.config.get('english_font', body_font) if self.config.get('use_custom_english_font') else body_font
        if replacement is not None:
            segments = [(text, Run(node.getparent(), para))
                        for node, text in self._replacement_values(para, replacement)]
        else:
            segments = [(self.processor._inline_run_text(run), run) for run in self.processor._iter_inline_runs(para)]
        # Strip only outside padding while retaining formatting changes within the text.
        if replacement is None and trim:
            combined = ''.join(text for text, _ in segments)
            begin, end = len(combined) - len(combined.lstrip()), len(combined.rstrip())
            offset = 0
            trimmed = []
            for text, run in segments:
                trimmed.append((text[max(0, begin - offset):max(0, min(len(text), end - offset))], run))
                offset += len(text)
            segments = trimmed
        total = 0.0
        for text, run in segments:
            for ch in text:
                font = english_font if ch.isascii() else body_font
                total += self.metrics.measure(ch, font, size, self._effective_flag(para, run, 'bold'),
                                              self._effective_flag(para, run, 'italic'))
        return total

    @staticmethod
    def _replacement_values(para, new_text):
        nodes = list(para._p.iter(qn('w:t')))
        remaining = new_text
        for index, node in enumerate(nodes):
            length = len(node.text or '')
            value = remaining if index == len(nodes) - 1 else remaining[:length]
            remaining = '' if index == len(nodes) - 1 else remaining[length:]
            yield node, value

    @classmethod
    def _replace_text(cls, para, new_text):
        """Keep the original text nodes, wrappers and all range markers."""
        for node, value in cls._replacement_values(para, new_text):
            node.text = value
            node.set(qn('xml:space'), 'preserve')

    @staticmethod
    def _set_before(para, value, unit='pt'):
        spacing = para._p.get_or_add_pPr().get_or_add_spacing()
        spacing.set(qn('w:beforeAutospacing'), '0')
        spacing.set(qn('w:before'), str(int(round(value * 20))) if unit == 'pt' else '0')
        spacing.set(qn('w:beforeLines'), str(int(round(value * 100))) if unit == 'lines' else '0')
        p_pr = para._p.get_or_add_pPr()
        contextual = p_pr.find(qn('w:contextualSpacing'))
        if contextual is None:
            contextual = OxmlElement('w:contextualSpacing')
            p_pr.insert_element_before(contextual, 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
                                       'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
                                       'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange')
        contextual.set(qn('w:val'), '0')

    @staticmethod
    def _disable_run_grid(run):
        properties = run._r.get_or_add_rPr()
        snap = properties.find(qn('w:snapToGrid'))
        if snap is None:
            snap = OxmlElement('w:snapToGrid')
            properties.insert_element_before(snap, 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing',
                                             'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight',
                                             'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign',
                                             'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout',
                                             'w:specVanish', 'w:oMath', 'w:rPrChange')
        snap.set(qn('w:val'), '0')

    @staticmethod
    def _disable_automatic_spacing(para):
        # East Asian auto-spacing otherwise adds quarter-em gaps around date digits,
        # making the displayed width differ from the font's advance widths.
        order = ('autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing',
                 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
                 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange')
        p_pr = para._p.get_or_add_pPr()
        for tag in ('autoSpaceDE', 'autoSpaceDN', 'adjustRightInd', 'snapToGrid'):
            element = p_pr.find(qn('w:' + tag))
            if element is None:
                element = OxmlElement('w:' + tag)
                p_pr.insert_element_before(element, *('w:' + successor for successor in order[order.index(tag) + 1:]))
            element.set(qn('w:val'), '0')

    @staticmethod
    def _section_for_paragraph(doc, para):
        # sectPr describes the section ending at that paragraph; body sectPr is last.
        sections = list(doc.sections)
        section_index = 0
        for element in doc.element.body:
            if element is para._p:
                break
            if element.tag == qn('w:p') and element.find('./' + qn('w:pPr') + '/' + qn('w:sectPr')) is not None:
                section_index += 1
        return sections[min(section_index, len(sections) - 1)]

    def _usable_width(self, doc, para):
        section = self._section_for_paragraph(doc, para)
        cols = section._sectPr.find(qn('w:cols'))
        if cols is not None and (int(cols.get(qn('w:num'), '1')) > 1 or len(cols.findall(qn('w:col'))) > 1):
            return None
        direction = section._sectPr.find(qn('w:textDirection'))
        if direction is not None and direction.get(qn('w:val')) not in (None, 'lrTb'):
            return None
        width = section.page_width.pt if section.page_width else 595.28
        left = section.left_margin.pt if section.left_margin else 0
        right = section.right_margin.pt if section.right_margin else 0
        gutter = section.gutter.pt if section.gutter else 0
        return width - left - right - gutter

    @classmethod
    def _indent_character_size(cls, doc, para=None):
        # Word/WPS character indents use the document's default paragraph font size,
        # not the run or paragraph-mark size of the particular signature paragraph.
        styles = {element.get(qn('w:styleId')): element
                  for element in doc.styles.element.findall(qn('w:style'))}
        current = next((element for element in styles.values()
                        if element.get(qn('w:type')) == 'paragraph'
                        and element.get(qn('w:default')) in ('1', 'true', 'on')), styles.get('Normal'))
        visited = set()
        while current is not None and current not in visited:
            visited.add(current)
            size = current.find('./' + qn('w:rPr') + '/' + qn('w:sz'))
            if size is not None:
                break
            parent = current.find(qn('w:basedOn'))
            current = styles.get(parent.get(qn('w:val'))) if parent is not None else None
        else:
            size = doc.styles.element.find('./' + qn('w:docDefaults') + '/' + qn('w:rPrDefault')
                                            + '/' + qn('w:rPr') + '/' + qn('w:sz'))
        try:
            points = float(size.get(qn('w:val'))) / 2 if size is not None else 0
            if para is not None:
                grid = cls._section_for_paragraph(doc, para)._sectPr.find(qn('w:docGrid'))
                if grid is not None and grid.get(qn('w:type')) in ('linesAndChars', 'snapToChars'):
                    points += int(grid.get(qn('w:charSpace'), '0')) / 4096
            return points if math.isfinite(points) and points > 0 else None
        except (ValueError, TypeError):
            return None

    def format(self, doc, info, apply_color):
        unit, date_para = info['unit'], info['date']
        normalized = '{:04d}年{}月{}日'.format(info['parsed'].year, info['parsed'].month, info['parsed'].day)
        original_date = self.text(date_para)
        date_text = normalized if self.config.get('signature_normalize_date', False) else original_date.strip()
        # OOXML font sizes are stored in half-points, just like _set_run_font.
        size = int(float(self.config.get('body_size', 16)) * 2) / 2
        if size <= 0:
            self.log('落款排版：正文字号无效，保留落款。')
            return False
        em = self.metrics.measure('国', self.config.get('body_font', '仿宋_GB2312'), size)
        unit_width = self._measure_paragraph(unit)
        date_width = self._measure_paragraph(date_para, date_text if self.config.get('signature_normalize_date') else None)
        unit_full_width = self._measure_paragraph(unit, trim=False)
        date_full_width = self._measure_paragraph(date_para, date_text if self.config.get('signature_normalize_date') else None,
                                                 trim=False)
        width = self._usable_width(doc, unit)
        layout = self.config.get('signature_layout', 'sealed')
        if layout == 'sealed':
            date_right = 4 * em
            unit_right = date_right + (date_width - unit_width) / 2
        else:
            unit_right = 2 * em
            date_right = unit_right + unit_width - 2 * em - date_width
            shift = max(0, 2 * em - date_right)
            unit_right += shift
            date_right += shift
        if (width is None or min(unit_right, date_right) < 0
                or unit_full_width + unit_right > width or date_full_width + date_right > width):
            self.log('落款排版：{}落款超出版心或为多栏布局，保留，不缩小字号。'.format(info['label']))
            return False
        indent_character_size = self._indent_character_size(doc, unit)
        if indent_character_size is None:
            self.log('落款排版：无法确定文档字符缩进基准，保留落款。')
            return False
        # Preserve the original text, including leading spaces, unless normalization was selected.
        if self.config.get('signature_normalize_date', False):
            self._replace_text(date_para, normalized)
        for para, right in ((unit, unit_right), (date_para, date_right)):
            self._disable_automatic_spacing(para)
            self.processor._apply_font_to_runs(para, self.config.get('body_font', '仿宋_GB2312'), size,
                                              set_color=apply_color)
            for run in self.processor._iter_inline_runs(para):
                self._disable_run_grid(run)
            p_pr = para._p.get_or_add_pPr()
            # Explicit zeros override inherited first-line/character indents.
            ind = p_pr.get_or_add_ind()
            for key in ('left', 'right', 'start', 'end', 'leftChars', 'rightChars', 'startChars',
                        'endChars', 'firstLine', 'firstLineChars', 'hanging', 'hangingChars'):
                ind.attrib.pop(qn('w:' + key), None)
            para.paragraph_format.left_indent = Pt(0)
            para.paragraph_format.first_line_indent = Pt(0)
            para.paragraph_format.right_indent = Pt(right)
            # Character indents inherited from styles override twip values in Word/WPS.
            # Shadow both legacy and start/end character attributes, using the document
            # font size for their unit while retaining twip values for other readers.
            right_chars = str(int(round(right / indent_character_size * 100)))
            for key, value in (('leftChars', '0'), ('startChars', '0'), ('rightChars', right_chars),
                               ('endChars', right_chars), ('firstLineChars', '0'), ('hangingChars', '0')):
                ind.set(qn('w:' + key), value)
            mark = Run(OxmlElement('w:r'), para)
            self.processor._set_run_font(mark, self.config.get('body_font', '仿宋_GB2312'), size)
            self._disable_run_grid(mark)
            mark_properties = p_pr.find(qn('w:rPr'))
            if mark_properties is None:
                mark_properties = OxmlElement('w:rPr')
                p_pr.insert_element_before(mark_properties, 'w:sectPr', 'w:pPrChange')
            for property_name in ('rFonts', 'snapToGrid', 'sz'):
                for old in list(mark_properties.findall(qn('w:' + property_name))):
                    mark_properties.remove(old)
                mark_properties.append(deepcopy(mark._r.rPr.find(qn('w:' + property_name))))
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            spacing = p_pr.get_or_add_spacing()
            spacing.set(qn('w:afterAutospacing'), '0')
            spacing.set(qn('w:after'), '0')
            spacing.set(qn('w:afterLines'), '0')
            self.processor._apply_line_spacing(para, 'line_spacing', 'line_spacing_unit',
                                               'line_spacing_multiple', 28)
            para.paragraph_format.keep_together = True
        unit.paragraph_format.keep_with_next = True
        self._set_before(date_para, 0)
        # Retain explicit page breaks. They take priority over keeping the pair on one page.
        for blank in info['between']:
            blank.paragraph_format.keep_with_next = True
        if info['between']:
            self.log('落款间距：署名与日期之间已有空段，保留原有段落。')
        mode = self.config.get('signature_before_spacing_mode', 'auto')
        if mode == 'custom':
            self._set_before(unit, float(self.config.get('signature_before_spacing', 0)),
                             self.config.get('signature_before_spacing_unit', 'pt'))
        elif mode == 'auto' and layout == 'unsealed':
            if info['blank_before']:
                self.log('落款间距：署名前已有空段，保留并不追加30磅段前间距。')
            else:
                self._set_before(unit, 30)
        if self.text(date_para) != original_date and self.config.get('signature_normalize_date', False):
            self.log('落款日期：{} → {}。'.format(original_date.strip(), normalized))
        self.log('落款排版：{}采用{}，署名右缩进{:.2f}磅，日期右缩进{:.2f}磅。'.format(
            info['label'], '盖章公文文字布局' if layout == 'sealed' else '不盖章布局', unit_right, date_right))
        return True
