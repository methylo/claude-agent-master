# -*- coding: utf-8 -*-
"""slide-design 조판 스크립트

기획표(JSON) -> 16:9 pptx. 장 유형 6종(표지·요점·표·흐름·비교·마무리)만 조판한다.
사용법: python3 scripts/build_slides.py plan.json out.pptx

색·글꼴은 references/theme.md 의 "설정값" 표를 읽는다. 표를 못 읽으면 아래 기본값을 쓴다.
글자 넘침 점검: 상자마다 글자 수 상한을 계산해 넘으면 경고를 출력한다(파일은 그대로 만든다).
"""
import json
import os
import re
import subprocess
import sys

# ------------------------------------------------------------ python-pptx 준비
try:
    import pptx  # noqa: F401
except ImportError:
    print("python-pptx 가 없어 설치한다: pip install python-pptx")
    cmds = [[sys.executable, "-m", "pip", "install", "-q", "python-pptx"],
            [sys.executable, "-m", "pip", "install", "-q", "--user", "python-pptx"],
            [sys.executable, "-m", "pip", "install", "-q", "--break-system-packages", "python-pptx"]]
    for c in cmds:
        if subprocess.call(c) == 0:
            break
    import site
    import importlib
    try:
        sys.path.append(site.getusersitepackages())
    except Exception:
        pass
    importlib.invalidate_caches()
    import pptx  # noqa: F401  설치 후 재시도. 여기서도 실패하면 오류로 멈춘다

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ------------------------------------------------------------ 기본값 (theme.md 가 없을 때)
FONT = "맑은 고딕"          # 기본 글꼴
INK_HEX = "#1F1D1B"         # 먹색: 제목·본문, 표지 패널·마무리 장 배경
ACC_HEX = "#C9654D"         # 강조색: 번호·키·선·강조 문장
BOX_HEX = "#F7F2EC"         # 상자색: 흐름 상자·표 강조 행·비교 머리

THEME_MD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "references", "theme.md")


def _read_theme():
    """theme.md 의 | 이름 | 값 | 표에서 INK·ACC·BOX·FONT 를 읽는다."""
    vals = {}
    try:
        with open(THEME_MD, encoding="utf-8") as f:
            for line in f:
                m = re.match(r"^\|\s*`?(INK|ACC|BOX|FONT)`?\s*\|\s*`?([^|`]+?)`?\s*\|", line)
                if m:
                    vals[m.group(1)] = m.group(2).strip()
    except (IOError, OSError):
        pass
    return vals


def _rgb(hexstr, fallback):
    s = (hexstr or "").strip().lstrip("#")
    if not re.match(r"^[0-9A-Fa-f]{6}$", s):
        s = fallback.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def _mix(a, b, t):
    """색 a 와 b 를 t(0~1) 비율로 섞는다. t=0 이면 a."""
    return RGBColor(*[int(round(a[i] * (1 - t) + b[i] * t)) for i in range(3)])


_tv = _read_theme()
FONT = _tv.get("FONT", FONT) or FONT
INK = _rgb(_tv.get("INK"), INK_HEX)
ACC = _rgb(_tv.get("ACC"), ACC_HEX)
BOX = _rgb(_tv.get("BOX"), BOX_HEX)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0, 0, 0)
# 나머지 색은 세 색에서 계산한다
ACC2 = _mix(ACC, BLACK, 0.15)        # 강조 2단
MUT = _mix(INK, WHITE, 0.50)         # 보조 글자
LINE = _mix(INK, WHITE, 0.80)        # 구분선
BOX2 = _mix(BOX, INK, 0.08)          # 2단 상자
PANEL_SUB = _mix(BOX, INK, 0.15)     # 어두운 바탕 위 보조 글자

# ------------------------------------------------------------ 판형 · 여백 · 글자 크기
W, H = 13.333, 7.5
ML = MR = 0.9
CW = W - ML - MR                      # 본문 폭 11.53in
Y_LABEL, Y_TITLE, Y_BODY = 0.50, 0.80, 2.05
Y_RULE, Y_MSG, Y_FOOT = 5.92, 6.08, 6.92

SZ_TITLE = 32      # 본문 장 제목
SZ_MSG = 17        # 핵심 문장
SZ_BODY = 17       # 본문
SZ_SMALL = 12      # 라벨·하단 문구

TYPES = ("표지", "요점", "표", "흐름", "비교", "마무리")

WARN = []


# ------------------------------------------------------------ 글자 수 점검
def _units(text):
    """글자 폭 환산: 한글·전각 1, 공백 0.3, 그 밖(영문·숫자·기호) 0.55."""
    u = 0.0
    for c in str(text):
        if c == " ":
            u += 0.3
        elif ord(c) >= 0x2E80:
            u += 1.0
        else:
            u += 0.55
    return u


def _cap(width_in, size_pt, lines=1):
    """너비(인치)·글자 크기에서 들어가는 한글 글자 수."""
    per_line = width_in / (size_pt / 72.0) * 0.96
    return int(per_line * lines)


def check(where, text, limit):
    if text is None:
        return
    n = _units(text)
    if n > limit + 0.01:
        WARN.append("%s: %d자 > 상한 %d자 — \"%s\"" % (where, int(round(n)), limit, str(text)[:30]))


# ------------------------------------------------------------ 그리기 유틸
class Deck(object):
    def __init__(self, meta):
        self.prs = Presentation()
        self.prs.slide_width = Inches(W)
        self.prs.slide_height = Inches(H)
        self.blank = self.prs.slide_layouts[6]
        self.title_line = meta.get("title_line", "")
        self.footer = meta.get("footer", "")
        self.total = 0
        self._set_theme_fonts()

    def _set_theme_fonts(self):
        """테마 기본 글꼴을 FONT 로 바꿔 나중에 친 글자도 같은 글꼴이 되게 한다."""
        from lxml import etree
        part = self.prs.slide_masters[0].part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        root = etree.fromstring(part.blob)
        for tag in ("majorFont", "minorFont"):
            node = root.find(".//" + qn("a:" + tag))
            if node is None:
                continue
            for child_tag in ("a:latin", "a:ea", "a:cs"):
                e = node.find(qn(child_tag))
                if e is None:
                    e = node.makeelement(qn(child_tag), {})
                    node.insert(0, e)
                e.set("typeface", FONT)
            for f in node.findall(qn("a:font")):
                if f.get("script") in ("Hang", "Hani"):
                    f.set("typeface", FONT)
        part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def new(self):
        return self.prs.slides.add_slide(self.blank)

    def tb(self, sl, x, y, w, h, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
        s = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = s.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = anchor
        tf.paragraphs[0].alignment = align
        return tf

    @staticmethod
    def p(tf, first=False, before=0, after=0, line=1.3, align=None):
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        para.space_before, para.space_after, para.line_spacing = Pt(before), Pt(after), line
        if align is not None:
            para.alignment = align
        return para

    @staticmethod
    def r(para, text, size=SZ_BODY, color=INK, bold=False):
        run = para.add_run()
        run.text = str(text)
        f = run.font
        f.name, f.size, f.bold = FONT, Pt(size), bold
        f.color.rgb = color
        rPr = run._r.get_or_add_rPr()
        for tag in ("a:ea", "a:cs"):
            e = rPr.find(qn(tag))
            if e is None:
                e = rPr.makeelement(qn(tag), {})
                rPr.append(e)
            e.set("typeface", FONT)
        return run

    def rect(self, sl, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
        s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        s.fill.solid()
        s.fill.fore_color.rgb = fill
        s.line.fill.background()
        s.shadow.inherit = False
        st = s._element.find(qn("p:style"))   # 기본 도형 스타일(그림자·테두리 참조)을 지운다
        if st is not None:
            s._element.remove(st)
        return s

    def line(self, sl, x, y, w, color=LINE, width=0.75):
        return self.rect(sl, x, y, w, Pt(width) / 914400.0, color)

    def chrome(self, sl, num, title, message):
        """본문 장 공통: 왼쪽 위 번호와 제목 줄, 제목, 아래 핵심 문장, 하단 문구, 쪽 번호."""
        t = self.tb(sl, ML, Y_LABEL, CW, 0.3)
        para = self.p(t, True)
        self.r(para, "%02d" % num, SZ_SMALL, ACC, True)
        if self.title_line:
            self.r(para, "   " + self.title_line, SZ_SMALL, MUT)
        t = self.tb(sl, ML, Y_TITLE, CW, 0.8)
        self.r(self.p(t, True, line=1.15), title, SZ_TITLE, INK, True)
        check("%02d장 제목" % num, title, _cap(CW, SZ_TITLE))
        if message:
            self.line(sl, ML, Y_RULE, CW)
            self.rect(sl, ML, Y_MSG + 0.05, 0.08, 0.32, ACC)
            t = self.tb(sl, ML + 0.25, Y_MSG, CW - 0.25, 0.45)
            self.r(self.p(t, True), message, SZ_MSG, INK, True)
            check("%02d장 핵심 문장" % num, message, _cap(CW - 0.25, SZ_MSG))
        self.foot(sl, num)

    def foot(self, sl, num=None, dark=False):
        col = PANEL_SUB if dark else MUT
        if self.footer:
            t = self.tb(sl, ML, Y_FOOT, CW * 0.7, 0.3)
            self.r(self.p(t, True), self.footer, 11, col)
        if num is not None:
            t = self.tb(sl, W - MR - 1.6, Y_FOOT, 1.6, 0.3, align=PP_ALIGN.RIGHT)
            self.r(self.p(t, True, align=PP_ALIGN.RIGHT), "%02d / %d" % (num, self.total), 11, col)

    # -------------------------------------------------------- 장 유형 6종
    def cover(self, num, d):
        """표지 — 왼쪽 흰 바탕 제목, 오른쪽 먹색 패널에 요점 2~3줄."""
        s = self.new()
        px = 8.6
        lw = px - ML - 0.5
        self.rect(s, px, 0, W - px, H, INK)
        self.line(s, ML, 0.95, 0.62, ACC, 3.5)
        if d.get("kicker"):
            t = self.tb(s, ML, 1.25, lw, 0.4)
            self.r(self.p(t, True), d["kicker"], 15, ACC)
            check("%02d장 표지 윗줄" % num, d["kicker"], _cap(lw, 15))
        lines = _title_lines(d.get("title", ""))
        t = self.tb(s, ML, 2.1, lw, 2.6)
        for i, ln in enumerate(lines):
            self.r(self.p(t, first=(i == 0), line=1.18), ln, 44, INK, True)
            check("%02d장 표지 제목 %d줄" % (num, i + 1), ln, _cap(lw, 44))
        if len(lines) > 2:
            WARN.append("%02d장 표지 제목: %d줄 > 상한 2줄" % (num, len(lines)))
        t = self.tb(s, ML, 5.25, lw, 0.9)
        if d.get("message"):
            self.r(self.p(t, True, after=4), d["message"], 16, INK)
            check("%02d장 표지 핵심 문장" % num, d["message"], _cap(lw, 16, 2))
        items = _list(d.get("items"))[:3]
        if len(_list(d.get("items"))) > 3:
            WARN.append("%02d장 표지 오른쪽 항목: %d개 > 상한 3개" % (num, len(_list(d.get("items")))))
        pw = W - px - 1.0
        t = self.tb(s, px + 0.55, 1.6, pw, 4.6)
        for i, it in enumerate(items):
            small, big = (it[0], it[1]) if isinstance(it, (list, tuple)) and len(it) > 1 else (None, _s(it))
            if small:
                self.r(self.p(t, first=(i == 0), after=2), small, 15, PANEL_SUB)
                check("%02d장 표지 오른쪽 %d 작은 줄" % (num, i + 1), small, _cap(pw, 15))
            self.r(self.p(t, first=(i == 0 and not small), after=22, line=1.2), big, 24,
                   ACC if i == 0 else WHITE, True)
            check("%02d장 표지 오른쪽 %d" % (num, i + 1), big, _cap(pw, 24, 2))
        self.foot(s)
        return s

    def points(self, num, d):
        """요점 — 키·값 3~4행."""
        s = self.new()
        self.chrome(s, num, d.get("title", ""), d.get("message"))
        rows = [_kv(x) for x in _list(d.get("items"))]
        _count(num, "요점 항목", rows, 4)
        rows = rows[:5]
        n = max(len(rows), 1)
        kw = 2.7
        rh = min(0.86, 3.7 / n - 0.1)
        yy = Y_BODY + 0.05
        for i, (k, v) in enumerate(rows):
            if k:
                t = self.tb(s, ML, yy, kw - 0.2, rh, anchor=MSO_ANCHOR.MIDDLE)
                self.r(self.p(t, True, line=1.15), k, SZ_BODY - 1, ACC, True)
                check("%02d장 요점 %d 키" % (num, i + 1), k, _cap(kw - 0.2, SZ_BODY - 1, 2))
            vx = ML + (kw if k else 0)
            vw = CW - (kw if k else 0)
            t = self.tb(s, vx, yy, vw, rh, anchor=MSO_ANCHOR.MIDDLE)
            self.r(self.p(t, True, line=1.2), v, SZ_BODY, INK)
            check("%02d장 요점 %d" % (num, i + 1), v, _cap(vw, SZ_BODY, 2 if rh >= 0.7 else 1))
            if i < n - 1:
                self.line(s, ML, yy + rh + 0.05, CW, LINE, 0.5)
            yy += rh + 0.1
        return s

    def table(self, num, d):
        """표 — 머리 1행 + 본문 2~6행. 열 폭은 글자 길이 비율로 정한다."""
        s = self.new()
        self.chrome(s, num, d.get("title", ""), d.get("message"))
        headers = [_s(h) for h in _list(d.get("headers"))]
        rows = [[_s(c) for c in _list(r)] for r in _list(d.get("rows"))]
        _count(num, "표 행", rows, 6)
        rows = rows[:7]
        nc = max([len(headers)] + [len(r) for r in rows] + [1])
        headers += [""] * (nc - len(headers))
        rows = [r + [""] * (nc - len(r)) for r in rows]
        if nc > 5:
            WARN.append("%02d장 표 열: %d개 > 상한 5개" % (num, nc))
        n = max(len(rows), 1)
        row_h = min(0.78, 3.25 / n)
        size = 16 if row_h >= 0.6 else 14
        # 열 폭: 각 열이 한 줄에 들어가는 폭을 먼저 잡고, 본문 폭을 넘으면 가장 긴 열부터 줄인다
        need = []
        for j in range(nc):
            m = max([_units(headers[j])] + [_units(r[j]) for r in rows] + [3.0])
            need.append(m * size / 72.0 / 0.96 + 0.35)
        if sum(need) <= CW:
            extra = (CW - sum(need)) / nc
            widths = [x + extra for x in need]
        else:
            lo, hi_w = 0.0, max(need)
            for _ in range(40):
                mid = (lo + hi_w) / 2
                if sum(min(x, mid) for x in need) > CW:
                    hi_w = mid
                else:
                    lo = mid
            widths = [min(x, lo) for x in need]
            widths[-1] += CW - sum(widths)
        hi = d.get("highlight_row")
        yy, xx = Y_BODY - 0.05, ML
        for j, h in enumerate(headers):
            t = self.tb(s, xx, yy, widths[j] - 0.15, 0.3)
            self.r(self.p(t, True), h, 13, MUT, True)
            check("%02d장 표 머리 %d" % (num, j + 1), h, _cap(widths[j] - 0.15, 13))
            xx += widths[j]
        yy += 0.36
        self.line(s, ML, yy, CW, ACC, 1.2)
        yy += 0.1
        lines = 2 if row_h >= 0.7 else 1
        for i, row in enumerate(rows):
            on = isinstance(hi, int) and hi == i + 1
            if on:
                self.rect(s, ML - 0.12, yy + 0.02, CW + 0.24, row_h - 0.06, BOX)
            xx = ML
            for j, cell in enumerate(row):
                t = self.tb(s, xx, yy, widths[j] - 0.2, row_h, anchor=MSO_ANCHOR.MIDDLE)
                first = (j == 0)
                self.r(self.p(t, True, line=1.2), cell, size, ACC2 if (on and first) else INK, first)
                check("%02d장 표 %d행 %d열" % (num, i + 1, j + 1), cell, _cap(widths[j] - 0.2, size, lines))
                xx += widths[j]
            if i < n - 1:
                self.line(s, ML, yy + row_h, CW, LINE, 0.5)
            yy += row_h
        return s

    def flow(self, num, d):
        """흐름 — 가로 단계 3~5칸. 단계마다 번호·이름·설명."""
        s = self.new()
        self.chrome(s, num, d.get("title", ""), d.get("message"))
        steps = [_step(x) for x in _list(d.get("steps") or d.get("items"))]
        _count(num, "흐름 단계", steps, 5)
        steps = steps[:6]
        n = max(len(steps), 1)
        gap = 0.34 if n <= 5 else 0.24
        bw = (CW - gap * (n - 1)) / n
        top, bh = Y_BODY - 0.05, 3.55
        f_name = 18 if n <= 4 else 16
        f_desc = 14 if n <= 4 else 13
        for i, (name, desc) in enumerate(steps):
            x = ML + (bw + gap) * i
            last = (i == n - 1)
            self.rect(s, x, top, bw, bh, BOX2 if last else BOX)
            e = self.rect(s, x + 0.25, top + 0.28, 0.42, 0.42, ACC, MSO_SHAPE.OVAL)
            tf = e.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            pp = tf.paragraphs[0]
            pp.alignment = PP_ALIGN.CENTER
            self.r(pp, str(i + 1), 13, WHITE, True)
            iw = bw - 0.45
            t = self.tb(s, x + 0.25, top + 0.9, iw, 0.85)
            self.r(self.p(t, True, line=1.15), name, f_name, ACC2 if last else INK, True)
            check("%02d장 흐름 %d단계 이름" % (num, i + 1), name, _cap(iw, f_name, 2))
            t = self.tb(s, x + 0.25, top + 1.8, iw, bh - 2.0)
            for j, ln in enumerate(desc[:3]):
                self.r(self.p(t, first=(j == 0), after=6, line=1.25), ln, f_desc, INK)
                check("%02d장 흐름 %d단계 설명 %d" % (num, i + 1, j + 1), ln, _cap(iw, f_desc, 2))
            if len(desc) > 3:
                WARN.append("%02d장 흐름 %d단계 설명: %d줄 > 상한 3줄" % (num, i + 1, len(desc)))
            if not last:
                t = self.tb(s, x + bw, top + bh / 2 - 0.15, gap, 0.3, align=PP_ALIGN.CENTER)
                self.r(self.p(t, True, align=PP_ALIGN.CENTER), "▶", 10, MUT)
        return s

    def compare(self, num, d):
        """비교 — 왼쪽(전·계획)과 오른쪽(후·실제) 두 열, 3~4행."""
        s = self.new()
        self.chrome(s, num, d.get("title", ""), d.get("message"))
        pairs = [_pair(x) for x in _list(d.get("pairs") or d.get("items"))]
        _count(num, "비교 행", pairs, 4)
        pairs = pairs[:5]
        has_key = any(p[0] for p in pairs)
        kw = 2.4 if has_key else 0.0
        cw2 = (CW - kw) / 2
        left, right = d.get("left", "전"), d.get("right", "후")
        yy = Y_BODY - 0.08
        self.rect(s, ML + kw, yy, cw2 - 0.1, 0.46, BOX)
        self.rect(s, ML + kw + cw2, yy, cw2, 0.46, ACC)
        for x, txt, col in ((ML + kw, left, MUT), (ML + kw + cw2, right, WHITE)):
            t = self.tb(s, x + 0.25, yy, cw2 - 0.4, 0.46, anchor=MSO_ANCHOR.MIDDLE)
            self.r(self.p(t, True), txt, 15, col, True)
            check("%02d장 비교 머리" % num, txt, _cap(cw2 - 0.4, 15))
        yy += 0.6
        n = max(len(pairs), 1)
        rh = min(0.82, 3.05 / n)
        for i, (k, a, b) in enumerate(pairs):
            if k:
                t = self.tb(s, ML, yy, kw - 0.2, rh, anchor=MSO_ANCHOR.MIDDLE)
                self.r(self.p(t, True, line=1.15), k, 16, INK, True)
                check("%02d장 비교 %d행 항목" % (num, i + 1), k, _cap(kw - 0.2, 16, 2))
            for x, txt, col, tag in ((ML + kw, a, MUT, "왼쪽"), (ML + kw + cw2, b, ACC2, "오른쪽")):
                t = self.tb(s, x + 0.25, yy, cw2 - 0.45, rh, anchor=MSO_ANCHOR.MIDDLE)
                self.r(self.p(t, True, line=1.2), txt, 17, col, tag == "오른쪽")
                check("%02d장 비교 %d행 %s" % (num, i + 1, tag), txt, _cap(cw2 - 0.45, 17, 2 if rh >= 0.7 else 1))
            if i < n - 1:
                self.line(s, ML, yy + rh + 0.03, CW, LINE, 0.5)
            yy += rh + 0.06
        return s

    def closing(self, num, d):
        """마무리 — 먹색 바탕. 큰 제목, 핵심 문장, 결정·다음 할 일 2~4줄."""
        s = self.new()
        self.rect(s, 0, 0, W, H, INK)
        self.line(s, ML, 1.2, 0.62, ACC, 3.5)
        title = d.get("title", "")
        t = self.tb(s, ML, 1.5, CW, 1.0)
        self.r(self.p(t, True, line=1.15), title, 44, WHITE, True)
        check("%02d장 마무리 제목" % num, title, _cap(CW, 44))
        y = 2.6
        if d.get("message"):
            t = self.tb(s, ML, y, CW, 0.5)
            self.r(self.p(t, True), d["message"], 20, ACC, True)
            check("%02d장 마무리 핵심 문장" % num, d["message"], _cap(CW, 20))
        items = [_s(x) for x in _list(d.get("items"))]
        _count(num, "마무리 항목", items, 4)
        t = self.tb(s, ML, y + 0.9, CW, 2.7)
        for i, ln in enumerate(items[:5]):
            para = self.p(t, first=(i == 0), after=10, line=1.25)
            self.r(para, "—  ", 17, PANEL_SUB)
            self.r(para, ln, 18, WHITE)
            check("%02d장 마무리 %d" % (num, i + 1), ln, _cap(CW - 0.4, 18))
        self.foot(s, dark=True)
        return s


# ------------------------------------------------------------ 입력 정리
def _s(x):
    if isinstance(x, (list, tuple)):
        return " ".join(_s(i) for i in x)
    if isinstance(x, dict):
        return " ".join(_s(v) for v in x.values())
    return "" if x is None else str(x)


def _list(x):
    if x is None:
        return []
    return list(x) if isinstance(x, (list, tuple)) else [x]


def _kv(x):
    """요점 항목: "키: 값" 문자열, [키, 값], {"key","value"} 모두 받는다."""
    if isinstance(x, dict):
        return _s(x.get("key", "")), _s(x.get("value", x.get("text", "")))
    if isinstance(x, (list, tuple)):
        return (_s(x[0]), _s(x[1])) if len(x) > 1 else ("", _s(x[0]))
    m = re.match(r"^\s*([^:：]{1,12})\s*[:：]\s*(.+)$", _s(x))
    return (m.group(1), m.group(2)) if m else ("", _s(x))


def _step(x):
    """흐름 단계: {"name","desc"}, [이름, 설명...], "이름: 설명" 모두 받는다."""
    if isinstance(x, dict):
        desc = x.get("desc", x.get("items", []))
        return _s(x.get("name", x.get("title", ""))), [_s(i) for i in _list(desc)]
    if isinstance(x, (list, tuple)):
        rest = list(x[1:])
        if len(rest) == 1 and isinstance(rest[0], (list, tuple)):
            rest = list(rest[0])
        return _s(x[0]), [_s(i) for i in rest]
    k, v = _kv(x)
    return (k, [v]) if k else (v, [])


def _pair(x):
    """비교 행: [항목, 왼쪽, 오른쪽] 또는 [왼쪽, 오른쪽] 또는 {"item","left","right"}."""
    if isinstance(x, dict):
        return _s(x.get("item", "")), _s(x.get("left", "")), _s(x.get("right", ""))
    x = _list(x)
    if len(x) >= 3:
        return _s(x[0]), _s(x[1]), _s(x[2])
    if len(x) == 2:
        return "", _s(x[0]), _s(x[1])
    return "", _s(x[0]) if x else "", ""


def _title_lines(title):
    """표지 제목을 2줄로 나눈다. 줄바꿈(\\n)이 있으면 그대로, 없으면 가운데에 가까운 띄어쓰기에서 나눈다."""
    title = _s(title)
    if "\n" in title:
        return [t.strip() for t in title.split("\n") if t.strip()]
    if _units(title) <= 10:
        return [title]
    words = title.split(" ")
    best, bi = None, 1
    for i in range(1, len(words)):
        a, b = " ".join(words[:i]), " ".join(words[i:])
        diff = max(_units(a), _units(b))
        if best is None or diff < best:
            best, bi = diff, i
    return [" ".join(words[:bi]), " ".join(words[bi:])] if len(words) > 1 else [title]


def _count(num, what, seq, limit):
    if len(seq) > limit:
        WARN.append("%02d장 %s: %d개 > 상한 %d개" % (num, what, len(seq), limit))


# ------------------------------------------------------------ 실행
def build(plan, out):
    meta = plan.get("meta", {}) if isinstance(plan, dict) else {}
    slides = plan.get("slides", []) if isinstance(plan, dict) else plan
    if not slides:
        sys.exit("오류: 기획표에 장(slides)이 없다")
    bad = [(i + 1, sd.get("type")) for i, sd in enumerate(slides) if sd.get("type") not in TYPES]
    if bad:
        sys.exit("오류: 장 유형은 %s 중 하나여야 한다. 맞지 않는 장: %s" % ("·".join(TYPES), bad))
    if not 5 <= len(slides) <= 8:
        WARN.append("장 수: %d장 (기준 5~8장)" % len(slides))
    deck = Deck(meta)
    deck.total = len(slides)
    fn = {"표지": deck.cover, "요점": deck.points, "표": deck.table,
          "흐름": deck.flow, "비교": deck.compare, "마무리": deck.closing}
    for i, sd in enumerate(slides):
        fn[sd["type"]](i + 1, sd)
    deck.prs.save(out)
    print("생성: %s (%d장, 글꼴 %s)" % (out, len(slides), FONT))
    if WARN:
        print("넘침 경고 %d건" % len(WARN))
        for w in WARN:
            print("- " + w)
    else:
        print("넘침 경고 0건")


def main():
    if len(sys.argv) != 3:
        sys.exit("사용법: python3 scripts/build_slides.py plan.json out.pptx")
    with open(sys.argv[1], encoding="utf-8") as f:
        plan = json.load(f)
    build(plan, sys.argv[2])


if __name__ == "__main__":
    main()
