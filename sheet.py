#!/usr/bin/env python3
"""ARC spreadsheet export.

Builds docs/arc-build.xlsx (and docs/prices.csv) from the checker's results on every run, so the copy
published on GitHub Pages always carries the latest prices. Python standard library only: the xlsx is
written as plain SpreadsheetML (zip + XML), with cached values for every formula so phone previews show
numbers without recalculating.
"""

import csv
import io
import re
import zipfile
from datetime import date, datetime, timezone
from xml.sax.saxutils import escape
from zoneinfo import ZoneInfo

UK = ZoneInfo("Europe/London")
EPOCH = date(1899, 12, 30)
_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")

GBP = '"£"#,##0.00'
DELTA = '+"£"#,##0.00;-"£"#,##0.00;"£"0.00'
LEFT = '"£"#,##0.00;[Red]-"£"#,##0.00;"£"0.00'
DATE = 'd mmm yyyy'
DAY = 'ddd d mmm'

INK = "FF111827"
MUTED = "FF6B7280"
FAINT = "FF9CA3AF"
HEAD_BG = "FF1F2937"
BEST_BG = "FFDCFCE7"
GIFT_BG = "FFFEF3C7"
SOFT_BG = "FFF3F4F6"
LINK = "FF1D4ED8"
GOOD = "FF15803D"
BAD = "FFB91C1C"
WARN = "FFB45309"


# ---------------------------------------------------------------- tiny xlsx writer

def col_letter(c):
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def ref(r, c):
    return "%s%d" % (col_letter(c), r)


def clean(s):
    return _BAD.sub("", str(s))


class Styles:
    def __init__(self):
        self.fonts = [("Arial", 10, False, False, False, INK)]
        self.fills = [None, "gray125"]
        self.borders = [None]
        self.numfmts = {}
        self.xfs = [(0, 0, 0, 0, None)]
        self.cache = {}

    def get(self, bold=False, italic=False, size=10, color=INK, underline=False, fill=None, fmt=None,
            border=None, h=None, v="center", wrap=False, indent=0):
        key = (bold, italic, size, color, underline, fill, fmt, border, h, v, wrap, indent)
        if key in self.cache:
            return self.cache[key]
        font = ("Arial", size, bold, italic, underline, color)
        if font not in self.fonts:
            self.fonts.append(font)
        fi = 0
        if fill:
            if fill not in self.fills:
                self.fills.append(fill)
            fi = self.fills.index(fill)
        bi = 0
        if border:
            if border not in self.borders:
                self.borders.append(border)
            bi = self.borders.index(border)
        ni = 0
        if fmt:
            if fmt not in self.numfmts:
                self.numfmts[fmt] = 164 + len(self.numfmts)
            ni = self.numfmts[fmt]
        xf = (self.fonts.index(font), fi, bi, ni, (h, v, wrap, indent))
        self.xfs.append(xf)
        idx = len(self.xfs) - 1
        self.cache[key] = idx
        return idx

    def xml(self):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
        if self.numfmts:
            out.append('<numFmts count="%d">' % len(self.numfmts))
            for code, i in self.numfmts.items():
                out.append('<numFmt numFmtId="%d" formatCode="%s"/>' % (i, escape(code, {'"': "&quot;"})))
            out.append("</numFmts>")
        out.append('<fonts count="%d">' % len(self.fonts))
        for name, size, b, i, u, color in self.fonts:
            out.append("<font>%s%s%s<sz val=\"%s\"/><color rgb=\"%s\"/><name val=\"%s\"/><family val=\"2\"/></font>" % (
                "<b/>" if b else "", "<i/>" if i else "", "<u/>" if u else "", size, color, name))
        out.append("</fonts>")
        out.append('<fills count="%d">' % len(self.fills))
        for f in self.fills:
            if f is None:
                out.append('<fill><patternFill patternType="none"/></fill>')
            elif f == "gray125":
                out.append('<fill><patternFill patternType="gray125"/></fill>')
            else:
                out.append('<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % f)
        out.append("</fills>")
        out.append('<borders count="%d">' % len(self.borders))
        for b in self.borders:
            if b is None:
                out.append("<border><left/><right/><top/><bottom/><diagonal/></border>")
            else:
                side, color = b
                parts = []
                for s in ("left", "right", "top", "bottom"):
                    if s in side:
                        parts.append('<%s style="thin"><color rgb="%s"/></%s>' % (s, color, s))
                    else:
                        parts.append("<%s/>" % s)
                out.append("<border>%s<diagonal/></border>" % "".join(parts))
        out.append("</borders>")
        out.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
        out.append('<cellXfs count="%d">' % len(self.xfs))
        for fo, fi, bi, ni, al in self.xfs:
            attrs = 'numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"' % (ni, fo, fi, bi)
            if ni:
                attrs += ' applyNumberFormat="1"'
            if fo:
                attrs += ' applyFont="1"'
            if fi:
                attrs += ' applyFill="1"'
            if bi:
                attrs += ' applyBorder="1"'
            if al:
                h, v, wrap, indent = al
                a = []
                if h:
                    a.append('horizontal="%s"' % h)
                if v:
                    a.append('vertical="%s"' % v)
                if wrap:
                    a.append('wrapText="1"')
                if indent:
                    a.append('indent="%d"' % indent)
                out.append('<xf %s applyAlignment="1"><alignment %s/></xf>' % (attrs, " ".join(a)))
            else:
                out.append("<xf %s/>" % attrs)
        out.append("</cellXfs>")
        out.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
        out.append("</styleSheet>")
        return "".join(out)


class Sheet:
    def __init__(self, name, tab=None):
        self.name = name[:31]
        self.tab = tab
        self.cells = {}
        self.widths = {}
        self.heights = {}
        self.merges = []
        self.links = []
        self.freeze = None
        self.filter = None
        self.landscape = True

    def set(self, r, c, value=None, style=0, formula=None):
        self.cells[(r, c)] = (value, style, formula)

    def link(self, r, c, url, text, style):
        self.set(r, c, text, style)
        if url and re.match(r"^https://", url):
            self.links.append((ref(r, c), url))

    def width(self, widths):
        for c, w in widths.items():
            self.widths[c] = w

    def xml(self):
        rows = {}
        for (r, c), cell in self.cells.items():
            rows.setdefault(r, {})[c] = cell
        maxr = max(rows) if rows else 1
        maxc = max((c for (_, c) in self.cells), default=1)
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
        out.append("<sheetPr>%s<pageSetUpPr fitToPage=\"1\"/></sheetPr>" % (
            '<tabColor rgb="%s"/>' % self.tab if self.tab else ""))
        out.append('<dimension ref="A1:%s"/>' % ref(maxr, maxc))
        if self.freeze:
            fr, fc = self.freeze
            attrs = []
            if fc > 1:
                attrs.append('xSplit="%d"' % (fc - 1))
            if fr > 1:
                attrs.append('ySplit="%d"' % (fr - 1))
            pane = "bottomRight" if (fr > 1 and fc > 1) else ("bottomLeft" if fr > 1 else "topRight")
            out.append('<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane %s topLeftCell="%s" activePane="%s" state="frozen"/>'
                       '<selection pane="%s" activeCell="%s" sqref="%s"/></sheetView></sheetViews>' % (
                           " ".join(attrs), ref(fr, fc), pane, pane, ref(fr, fc), ref(fr, fc)))
        else:
            out.append('<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>')
        out.append('<sheetFormatPr defaultRowHeight="16"/>')
        if self.widths:
            out.append("<cols>")
            for c in sorted(self.widths):
                out.append('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (c, c, self.widths[c]))
            out.append("</cols>")
        out.append("<sheetData>")
        for r in sorted(rows):
            ht = self.heights.get(r)
            out.append('<row r="%d"%s>' % (r, ' ht="%s" customHeight="1"' % ht if ht else ""))
            for c in sorted(rows[r]):
                value, style, formula = rows[r][c]
                a = ref(r, c)
                s = ' s="%d"' % style if style else ""
                if formula is not None:
                    f = "<f>%s</f>" % escape(clean(formula))
                    if isinstance(value, bool):
                        out.append('<c r="%s"%s t="b">%s<v>%d</v></c>' % (a, s, f, 1 if value else 0))
                    elif isinstance(value, (int, float)):
                        out.append('<c r="%s"%s>%s<v>%s</v></c>' % (a, s, f, repr(round(float(value), 6))))
                    elif value is None:
                        out.append('<c r="%s"%s>%s</c>' % (a, s, f))
                    else:
                        out.append('<c r="%s"%s t="str">%s<v>%s</v></c>' % (a, s, f, escape(clean(value))))
                elif value is None or value == "":
                    if style:
                        out.append('<c r="%s"%s/>' % (a, s))
                elif isinstance(value, bool):
                    out.append('<c r="%s"%s t="b"><v>%d</v></c>' % (a, s, 1 if value else 0))
                elif isinstance(value, (int, float)):
                    out.append('<c r="%s"%s><v>%s</v></c>' % (a, s, repr(round(float(value), 6))))
                else:
                    txt = clean(value)
                    sp = ' xml:space="preserve"' if txt != txt.strip() or "\n" in txt else ""
                    out.append('<c r="%s"%s t="inlineStr"><is><t%s>%s</t></is></c>' % (a, s, sp, escape(txt)))
            out.append("</row>")
        out.append("</sheetData>")
        if self.filter:
            out.append('<autoFilter ref="%s"/>' % self.filter)
        if self.merges:
            out.append('<mergeCells count="%d">' % len(self.merges))
            for m in self.merges:
                out.append('<mergeCell ref="%s"/>' % m)
            out.append("</mergeCells>")
        if self.links:
            out.append("<hyperlinks>")
            for i, (a, _) in enumerate(self.links, 1):
                out.append('<hyperlink ref="%s" r:id="rId%d"/>' % (a, i))
            out.append("</hyperlinks>")
        out.append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        out.append('<pageSetup paperSize="9" orientation="%s" fitToWidth="1" fitToHeight="0"/>' % (
            "landscape" if self.landscape else "portrait"))
        out.append("</worksheet>")
        return "".join(out)

    def rels(self):
        if not self.links:
            return None
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for i, (_, url) in enumerate(self.links, 1):
            out.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                       'Target="%s" TargetMode="External"/>' % (i, escape(url, {'"': "&quot;"})))
        out.append("</Relationships>")
        return "".join(out)


class Book:
    def __init__(self, title):
        self.title = title
        self.sheets = []
        self.st = Styles()

    def sheet(self, name, tab=None):
        s = Sheet(name, tab)
        self.sheets.append(s)
        return s

    def save(self, path, modified):
        names = []
        for s in self.sheets:
            base, n = s.name, 2
            while s.name in names:
                s.name = "%s %d" % (base, n)
                n += 1
            names.append(s.name)
        ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
              '<Default Extension="xml" ContentType="application/xml"/>',
              '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
              '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
              '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
              '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>']
        for i in range(1, len(self.sheets) + 1):
            ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
        ct.append("</Types>")
        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                "</Relationships>")
        wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
              '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
              '<bookViews><workbookView activeTab="0"/></bookViews><sheets>']
        for i, s in enumerate(self.sheets, 1):
            wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(s.name, {'"': "&quot;"}), i, i))
        wb.append("</sheets>")
        dn = [(i, s) for i, s in enumerate(self.sheets) if s.filter]
        if dn:
            wb.append("<definedNames>")
            for i, s in dn:
                a, b = s.filter.split(":")
                absr = "'%s'!$%s$%s:$%s$%s" % (s.name.replace("'", "''"), re.sub(r"\d", "", a), re.sub(r"\D", "", a),
                                                 re.sub(r"\d", "", b), re.sub(r"\D", "", b))
                wb.append('<definedName name="_xlnm._FilterDatabase" localSheetId="%d" hidden="1">%s</definedName>' % (i, escape(absr)))
            wb.append("</definedNames>")
        wb.append('<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        wbrels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for i in range(1, len(self.sheets) + 1):
            wbrels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, i))
        wbrels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (len(self.sheets) + 1))
        wbrels.append("</Relationships>")
        stamp = modified.strftime("%Y-%m-%dT%H:%M:%SZ")
        core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                "<dc:title>%s</dc:title><dc:creator>ARC price checker</dc:creator>"
                '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
                '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
                "</cp:coreProperties>") % (escape(self.title), stamp, stamp)
        app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
               "<Application>ARC price checker</Application></Properties>")
        styles = self.st.xml()
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", "".join(ct))
            z.writestr("_rels/.rels", rels)
            z.writestr("docProps/core.xml", core)
            z.writestr("docProps/app.xml", app)
            z.writestr("xl/workbook.xml", "".join(wb))
            z.writestr("xl/_rels/workbook.xml.rels", "".join(wbrels))
            z.writestr("xl/styles.xml", styles)
            for i, s in enumerate(self.sheets, 1):
                z.writestr("xl/worksheets/sheet%d.xml" % i, s.xml())
                r = s.rels()
                if r:
                    z.writestr("xl/worksheets/_rels/sheet%d.xml.rels" % i, r)
        tmp = path + ".tmp"
        with open(tmp, "wb") as f:
            f.write(buf.getvalue())
        import os
        os.replace(tmp, path)


# ---------------------------------------------------------------- helpers

def xl_date(d):
    if isinstance(d, str):
        try:
            d = date.fromisoformat(d[:10])
        except ValueError:
            return None
    return (d - EPOCH).days if d else None


def uk_stamp(iso):
    try:
        dt = datetime.strptime(iso, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc).astimezone(UK)
        return dt.strftime("%a %-d %b %Y, %H:%M UK")
    except (TypeError, ValueError):
        return iso or ""


def short_date(s):
    try:
        return date.fromisoformat(s[:10]).strftime("%-d %b")
    except (TypeError, ValueError):
        return s or ""


def gbp(x):
    return "£{:,.2f}".format(x) if x is not None else "—"


def stock_text(o):
    if not o:
        return ""
    if o.get("in_stock"):
        return "In stock"
    if o.get("in_stock") is False:
        s = (o.get("stock") or "").lower()
        return "Pre-order" if "pre" in s else "Out of stock"
    return "Unknown"


def source_text(o):
    src = o.get("source")
    if src == "pricespy":
        return "PriceSpy UK" + (" (last good)" if o.get("stale") else "")
    if src == "direct":
        return "Shop page (live)"
    if src == "verified":
        return "Checked by sweep %s%s" % (short_date(o.get("checked_on")), " (old)" if o.get("stale") else "")
    return src or ""


SIGNAL_TEXT = {
    "soldout": "Sold out at every vetted shop",
    "stock_unknown": "Stock unknown",
    "new_low": "Lowest price yet",
    "near_low": "Near its lowest",
    "below_sheet": "5%+ under plan",
    "not_live": "Price from the daily sweep",
    "ni_check": "NI delivery cost to confirm",
    "no_price": "No price found",
}


# ---------------------------------------------------------------- workbook

def build_workbook(data, hist, cfg, site):
    st = None
    book = Book("ARC build")
    S = book.st
    st = {
        "title": S.get(bold=True, size=16),
        "sub": S.get(color=MUTED, size=9),
        "h": S.get(bold=True, color="FFFFFFFF", fill=HEAD_BG, h="left"),
        "hr": S.get(bold=True, color="FFFFFFFF", fill=HEAD_BG, h="center"),
        "hc": S.get(bold=True, color="FFFFFFFF", fill=HEAD_BG, h="center", wrap=True),
        "sec": S.get(bold=True, size=12),
        "txt": S.get(),
        "wrap": S.get(wrap=True),
        "bold": S.get(bold=True),
        "muted": S.get(color=MUTED),
        "mwrap": S.get(color=MUTED, wrap=True),
        "d_good": S.get(fmt=DELTA, color=GOOD),
        "d_bad": S.get(fmt=DELTA, color=BAD),
        "d_flat": S.get(fmt=DELTA, color=MUTED),
        "d_good_b": S.get(fmt=DELTA, color=GOOD, bold=True),
        "d_bad_b": S.get(fmt=DELTA, color=BAD, bold=True),
        "d_flat_b": S.get(fmt=DELTA, color=MUTED, bold=True),
        "shop": S.get(indent=1),
        "faint": S.get(color=FAINT, italic=True),
        "gbp": S.get(fmt=GBP),
        "gbp_b": S.get(fmt=GBP, bold=True),
        "gbp_best": S.get(fmt=GBP, bold=True, color=GOOD, fill=BEST_BG),
        "gbp_oos": S.get(fmt=GBP, italic=True, color=FAINT),
        "gbp_gift": S.get(fmt=GBP, fill=GIFT_BG),
        "delta": S.get(fmt=DELTA),
        "delta_b": S.get(fmt=DELTA, bold=True),
        "left_b": S.get(fmt=LEFT, bold=True, size=12),
        "kpi_l": S.get(color=MUTED),
        "kpi_v": S.get(fmt=GBP, bold=True, size=12),
        "kpi_in": S.get(fmt=GBP, bold=True, size=12, color=LINK),
        "date": S.get(fmt=DATE),
        "day": S.get(fmt=DAY),
        "link": S.get(color=LINK, underline=True),
        "good": S.get(color=GOOD, bold=True),
        "bad": S.get(color=BAD, bold=True),
        "warn": S.get(color=WARN),
        "tot_l": S.get(bold=True, border=("top", "FF9CA3AF")),
        "tot_g": S.get(bold=True, fmt=GBP, border=("top", "FF9CA3AF")),
        "tot_d": S.get(bold=True, fmt=DELTA, border=("top", "FF9CA3AF")),
        "tot_e": S.get(border=("top", "FF9CA3AF")),
        "gift": S.get(fill=GIFT_BG),
        "gift_b": S.get(fill=GIFT_BG, bold=True),
        "center": S.get(h="center"),
        "bought": S.get(h="center", bold=True, color=GOOD),
    }

    def dsty(v, bold=False):
        if v is None or abs(v) < 0.005:
            return st["d_flat_b" if bold else "d_flat"]
        return st[("d_good" if v < 0 else "d_bad") + ("_b" if bold else "")]

    def lsty(v, size=10):
        return S.get(fmt=LEFT, bold=True, size=size, color=GOOD if (v or 0) >= 0 else BAD)

    parts = data.get("parts", [])
    totals = data.get("totals", {})
    settings = data.get("settings", {})
    base = data.get("baseline") or {}
    updated = uk_stamp(data.get("generated"))
    xlsx_url = site + "arc-build.xlsx"

    # ------------------------------------------------ Build
    b = book.sheet("Build", tab="FF1F2937")
    b.width({1: 15, 2: 46, 3: 11, 4: 17, 5: 12, 6: 11, 7: 10, 8: 11, 9: 19, 10: 24, 11: 34, 12: 8, 13: 8})
    b.set(1, 1, data.get("title_long") or "ARC build", st["title"])
    b.heights[1] = 26
    b.set(2, 1, "Live UK prices (inc VAT) from shops that deliver to Northern Ireland · updated %s · rebuilt about 4× a day, "
                "re-verified by the 07:46 sweep · newest copy: %s" % (updated, xlsx_url), st["sub"])

    hdr = 12
    first = hdr + 1
    last = first + len(parts) - 1
    tot = last + 1
    budget = settings.get("budget")

    # KPI block (rows 4-10)
    lines = totals.get("delivery_lines", [])
    d_first = tot + 4
    d_last = d_first + max(len(lines), 1) - 1
    kpis = [
        ("Parts total at today's best prices", "=C%d" % tot, totals.get("build_total"), st["kpi_v"]),
        ("NI delivery, one order per shop (estimate)", "=SUM(E%d:E%d)" % (d_first, d_last), totals.get("delivery"), st["kpi_v"]),
        ("Total delivered", "=C4+C5", (totals.get("build_total") or 0) + (totals.get("delivery") or 0), st["kpi_v"]),
        ("Budget (incl. Windows) · edit me", None, budget, st["kpi_in"]),
        ("Left in budget after delivery", "=C7-C6" if budget else None,
         (budget - (totals.get("build_total") or 0) - (totals.get("delivery") or 0)) if budget else None,
         lsty((budget or 0) - (totals.get("build_total") or 0) - (totals.get("delivery") or 0), 12)),
        ("Parts vs the %s (%s)" % (base.get("label", "plan"), gbp(base.get("parts_total"))), "=C%d-F%d" % (tot, tot),
         totals.get("delta_vs_baseline"), S.get(fmt=DELTA, bold=True, size=12,
                                                color=GOOD if (totals.get("delta_vs_baseline") or 0) < 0 else BAD if (totals.get("delta_vs_baseline") or 0) > 0 else MUTED)),
        ("Paid so far · still to buy", None, "%s paid · %d of %d parts left" % (gbp(totals.get("paid") or 0),
                                                                               totals.get("parts_left", 0), totals.get("parts_total", 0)), st["bold"]),
    ]
    for i, (label, f, v, sty) in enumerate(kpis):
        r = 4 + i
        b.set(r, 1, label, st["kpi_l"])
        b.merges.append("A%d:B%d" % (r, r))
        if f:
            b.set(r, 3, v, sty, formula=f[1:])
        else:
            b.set(r, 3, v, sty)
    b.set(11, 1, "Parts — cheapest in-stock offer at a vetted UK shop, ranked on item + NI delivery", st["sec"])
    b.set(11, 9, "Blue = you can edit · green = cheaper than plan · red = dearer · amber rows = Phase 1 (case + fans)", st["sub"])

    heads = ["Slot", "Part", "Price", "Shop", "Stock", "Plan %s" % short_date(base.get("date")), "vs plan",
             "vs yesterday", "Lowest seen", "Next best", "Notes", "Bought", "Link"]
    for c, h in enumerate(heads, 1):
        b.set(hdr, c, h, st["hr"] if c in (3, 6, 7, 8) else st["h"])
    phase_ids = set((settings.get("phase1") or {}).get("parts", []))
    row_of = {}
    for i, p in enumerate(parts):
        r = first + i
        row_of[p["id"]] = r
        best = p.get("best") or {}
        gift = p["id"] in phase_ids
        if p.get("bought"):
            price, shop, stock = p.get("paid"), p.get("paid_retailer") or "", "Bought %s" % short_date(p.get("bought_on"))
        else:
            price, shop, stock = best.get("price"), best.get("display") or "", stock_text(best)
        b.set(r, 1, p.get("slot"), st["gift_b"] if gift else st["bold"])
        b.set(r, 2, p.get("name"), st["gift"] if gift else st["txt"])
        oos = best and best.get("in_stock") is False and not p.get("bought")
        b.set(r, 3, price, st["gbp_oos"] if oos else (st["gbp_gift"] if gift else st["gbp"]))
        b.set(r, 4, shop, st["shop"])
        b.set(r, 5, stock, st["good"] if stock == "In stock" else st["bought"] if p.get("bought") else st["warn"])
        bp = (p.get("baseline") or {}).get("price")
        b.set(r, 6, bp, st["gbp"])
        if price is not None and bp is not None:
            b.set(r, 7, round(price - bp, 2), dsty(price - bp), formula="C%d-F%d" % (r, r))
        ch = (p.get("change") or {}).get("vs_yesterday")
        b.set(r, 8, ch, dsty(ch))
        lo = p.get("low")
        b.set(r, 9, "%s · %s" % (gbp(lo["price"]), short_date(lo["date"])) if lo else "", st["muted"])
        nxt = [o for o in p.get("offers", []) if o.get("shop") != best.get("shop") and o.get("price") is not None]
        nb = nxt[0] if nxt else None
        b.set(r, 10, "%s %s%s" % (nb.get("display"), gbp(nb.get("price")), "" if nb.get("in_stock") else " (%s)" % stock_text(nb).lower())
              if nb else "", st["muted"])
        notes = [SIGNAL_TEXT.get(s, s) for s in p.get("signals", []) if s in SIGNAL_TEXT]
        if p.get("note"):
            notes.append(p["note"])
        if gift:
            notes.insert(0, "Phase 1")
        b.set(r, 11, " · ".join(notes), st["mwrap"])
        b.set(r, 12, "Y" if p.get("bought") else "", st["bought"])
        url = best.get("url") if not p.get("bought") else None
        if url:
            b.link(r, 13, url, "Buy", st["link"])
    b.set(tot, 1, "", st["tot_e"])
    b.set(tot, 2, "Parts total", st["tot_l"])
    b.set(tot, 3, totals.get("build_total"), st["tot_g"], formula="SUM(C%d:C%d)" % (first, last))
    for c in (4, 5):
        b.set(tot, c, "", st["tot_e"])
    b.set(tot, 6, base.get("parts_total"), st["tot_g"], formula="SUM(F%d:F%d)" % (first, last))
    dv = totals.get("delta_vs_baseline")
    b.set(tot, 7, dv, S.get(bold=True, fmt=DELTA, border=("top", "FF9CA3AF"),
                            color=GOOD if (dv or 0) < 0 else BAD if (dv or 0) > 0 else MUTED), formula="C%d-F%d" % (tot, tot))
    for c in range(8, 14):
        b.set(tot, c, "", st["tot_e"])

    # Delivery plan
    b.set(tot + 2, 1, "Buying plan — one order per shop", st["sec"])
    for c, h in enumerate(["Shop", "Parts", "", "Items", "NI delivery", "Basis"], 1):
        if h:
            b.set(tot + 3, c, h, st["hr"] if c in (4, 5) else st["h"])
        else:
            b.set(tot + 3, c, "", st["h"])
    for i, l in enumerate(lines or [{"shop": "—", "parts": 0, "cost": 0, "basis": ""}]):
        r = d_first + i
        b.set(r, 1, l.get("shop"), st["bold"])
        names = [p["slot"] for p in parts if not p.get("bought") and (p.get("best") or {}).get("display") == l.get("shop")]
        b.set(r, 2, ", ".join(names), st["mwrap"])
        items = sum((p.get("best") or {}).get("price") or 0 for p in parts if not p.get("bought") and (p.get("best") or {}).get("display") == l.get("shop"))
        b.set(r, 4, round(items, 2), st["gbp"], formula='SUMIFS(C%d:C%d,D%d:D%d,A%d,L%d:L%d,"<>Y")' % (first, last, first, last, r, first, last))
        b.set(r, 5, l.get("cost") or 0, st["gbp"])
        basis = {"verified": "Shop's published rate", "assumed": "Free / at checkout (assumed)", "estimate": "Estimate — shown at checkout",
                 "unverified": "Not confirmed yet (estimate)"}.get(l.get("basis"), l.get("basis") or "")
        if l.get("note"):
            basis = l["note"]
        b.set(r, 6, basis, st["muted"])

    # Phase 1
    ph = data.get("phase1") or {}
    if ph:
        r0 = d_last + 2
        b.set(r0, 1, (ph.get("label") or "Phase 1") + " — first purchase", st["sec"])
        r = r0 + 1
        cells = []
        for it in ph.get("items", []):
            pr = row_of.get(it["id"])
            b.set(r, 1, it.get("slot"), st["gift_b"])
            b.set(r, 2, next((p["name"] for p in parts if p["id"] == it["id"]), ""), st["gift"])
            b.set(r, 3, it.get("price"), st["gbp_gift"], formula="C%d" % pr if pr else None)
            cells.append("C%d" % r)
            r += 1
        b.set(r, 2, "Phase 1 spend (items)", st["tot_l"])
        b.set(r, 3, ph.get("cost"), st["tot_g"], formula="SUM(%s)" % ",".join(cells) if cells else None)
        b.set(r + 1, 2, "Phase 1 budget · edit me", st["kpi_l"])
        b.set(r + 1, 3, ph.get("budget"), S.get(fmt=GBP, bold=True, color=LINK))
        b.set(r + 2, 2, "Left in Phase 1 budget (before delivery)", st["kpi_l"])
        b.set(r + 2, 3, ph.get("left"), lsty(ph.get("left")), formula="C%d-C%d" % (r + 1, r))
        if ph.get("order_by"):
            b.set(r + 3, 2, "Order by (allows for NI delivery)", st["kpi_l"])
            b.set(r + 3, 3, xl_date(ph["order_by"]), st["date"])
        if ph.get("tip"):
            b.set(r + 4, 2, ph["tip"], st["mwrap"])
            b.heights[r + 4] = 30

    # ------------------------------------------------ Prices by shop
    m = book.sheet("Prices by shop", tab="FF15803D")
    shops = {}
    for p in parts:
        for o in p.get("offers", []):
            if o.get("price") is None:
                continue
            k = o.get("display") or o.get("shop")
            shops.setdefault(k, set()).add(p["id"])
    order = sorted(shops, key=lambda k: (-len(shops[k]), k))
    m.set(1, 1, "Price at every vetted UK shop (inc VAT, before delivery)", st["title"])
    m.heights[1] = 26
    m.set(2, 1, "Green = today's pick (cheapest in stock once NI delivery is added; amber italic = pick is out of stock everywhere vetted) · grey italic = out of stock, pre-order or stock unknown · "
                "blank = not sold there or not vetted · updated %s" % updated, st["sub"])
    m.width({1: 15, 2: 40})
    m.set(4, 1, "Slot", st["h"])
    m.set(4, 2, "Part", st["h"])
    for j, k in enumerate(order):
        m.set(4, 3 + j, k, st["hc"])
        m.widths[3 + j] = 13
    m.heights[4] = 30
    m.set(4, 3 + len(order), "Lowest listed", st["hr"])
    m.widths[3 + len(order)] = 12
    m.freeze = (5, 3)
    for i, p in enumerate(parts):
        r = 5 + i
        m.set(r, 1, p.get("slot"), st["bold"])
        m.set(r, 2, p.get("name"), st["txt"])
        best = p.get("best") or {}
        by = {}
        for o in p.get("offers", []):
            k = o.get("display") or o.get("shop")
            if o.get("price") is None:
                continue
            if k not in by or (o.get("in_stock") and not by[k].get("in_stock")) or (
                    bool(o.get("in_stock")) == bool(by[k].get("in_stock")) and o["price"] < by[k]["price"]):
                by[k] = o
        for j, k in enumerate(order):
            o = by.get(k)
            if not o:
                continue
            is_best = best and k == best.get("display") and o.get("price") == best.get("price")
            if is_best and not o.get("in_stock"):
                sty = S.get(fmt=GBP, italic=True, bold=True, color=WARN, fill=BEST_BG)
            else:
                sty = st["gbp_best"] if is_best else (st["gbp_oos"] if not o.get("in_stock") else st["gbp"])
            url = o.get("url")
            if url and re.match(r"^https://", url):
                m.links.append((ref(r, 3 + j), url))
            m.set(r, 3 + j, o["price"], sty)
        c0, c1 = col_letter(3), col_letter(2 + len(order))
        cheapest = min((o["price"] for o in by.values()), default=None)
        if order:
            m.set(r, 3 + len(order), cheapest, st["gbp_b"], formula="IFERROR(MIN(%s%d:%s%d),\"\")" % (c0, r, c1, r))
        if p.get("bought"):
            m.set(r, 3 + len(order), p.get("paid"), st["gbp_b"])
    r = 5 + len(parts) + 1
    m.set(r, 2, "NI delivery for a single order", st["muted"])
    rets = data.get("retailers", {})
    disp = {v.get("display") or k: v for k, v in rets.items()}
    for j, k in enumerate(order):
        v = disp.get(k, {})
        if v.get("ni") == "no":
            txt = "No NI"
        elif v.get("ni_first") is not None:
            txt = "free" if not v["ni_first"] else gbp(v["ni_first"])
        elif v.get("ni") == "assumed":
            txt = "at checkout"
        else:
            txt = "est. " + gbp(settings.get("ni_default_estimate", 15))
        m.set(r, 3 + j, txt, S.get(color=MUTED, h="center", size=9))

    # ------------------------------------------------ All offers
    a = book.sheet("All offers", tab="FF2563EB")
    a.set(1, 1, "Every offer found this run", st["title"])
    a.heights[1] = 26
    a.set(2, 1, "Vetted UK shops only (marketplaces, used stock and shops that don't deliver to NI are left out) · Landed = item + NI delivery "
                "for a single-item order · updated %s" % updated, st["sub"])
    cols = ["Slot", "Part", "Shop", "Listing", "Price", "NI delivery", "Landed", "Stock", "Source", "Pick?", "Link"]
    a.width({1: 15, 2: 30, 3: 17, 4: 54, 5: 10, 6: 11, 7: 10, 8: 12, 9: 24, 10: 7, 11: 7})
    for c, h in enumerate(cols, 1):
        a.set(4, c, h, st["hr"] if c in (5, 6, 7) else st["h"])
    r = 5
    for p in parts:
        best = p.get("best") or {}
        for o in p.get("offers", []):
            is_best = best and o.get("shop") == best.get("shop") and o.get("price") == best.get("price")
            a.set(r, 1, p.get("slot"), st["bold"] if is_best else st["txt"])
            a.set(r, 2, p.get("name"), st["muted"])
            a.set(r, 3, o.get("display") or o.get("shop"), st["bold"] if is_best else st["txt"])
            a.set(r, 4, (o.get("name") or "")[:120], st["muted"])
            a.set(r, 5, o.get("price"), st["gbp_best"] if is_best else (st["gbp"] if o.get("in_stock") else st["gbp_oos"]))
            dl = o.get("delivery") if o.get("delivery") is not None else (0 if o.get("landed") is not None else None)
            a.set(r, 6, dl, st["gbp"])
            if o.get("price") is not None and dl is not None:
                a.set(r, 7, o.get("landed"), st["gbp"], formula="E%d+F%d" % (r, r))
            a.set(r, 8, stock_text(o), st["good"] if o.get("in_stock") else st["warn"])
            a.set(r, 9, source_text(o), st["muted"])
            a.set(r, 10, "✓" if is_best else "", st["good"])
            if o.get("url"):
                a.link(r, 11, o["url"], "Open", st["link"])
            r += 1
    a.filter = "A4:K%d" % max(r - 1, 5)
    a.freeze = (5, 1)

    # ------------------------------------------------ Price history
    h = book.sheet("Price history", tab="FF7C3AED")
    h.set(1, 1, "Best price per day", st["title"])
    h.heights[1] = 26
    h.set(2, 1, "Cheapest in-stock vetted offer each UK day (item price). The build total only counts days with a price for every part.", st["sub"])
    series = (hist or {}).get("parts", {})
    days = sorted({e[0] for p in parts for e in series.get(p["id"], [])})
    h.width({1: 14})
    h.set(4, 1, "Day", st["h"])
    for j, p in enumerate(parts):
        h.set(4, 2 + j, p.get("slot"), st["hc"])
        h.widths[2 + j] = 12
    h.set(4, 2 + len(parts), "Build total", st["hc"])
    h.widths[2 + len(parts)] = 13
    h.heights[4] = 30
    h.freeze = (5, 2)
    for i, dday in enumerate(days):
        r = 5 + i
        h.set(r, 1, xl_date(dday), st["day"])
        vals = []
        for j, p in enumerate(parts):
            pts = [e for e in series.get(p["id"], []) if e[0] == dday]
            v = pts[-1][1] if pts else None
            vals.append(v)
            h.set(r, 2 + j, v, st["gbp"] if (pts and pts[-1][3]) else st["gbp_oos"])
        c0, c1 = col_letter(2), col_letter(1 + len(parts))
        full = all(v is not None for v in vals)
        h.set(r, 2 + len(parts), round(sum(vals), 2) if full else None, st["gbp_b"],
              formula='IF(COUNT(%s%d:%s%d)=%d,SUM(%s%d:%s%d),"")' % (c0, r, c1, r, len(parts), c0, r, c1, r))
    refs = (cfg.get("history_refs") or [])
    r = 5 + len(days) + 1
    if refs:
        h.set(r, 1, "Earlier plans", st["sec"])
        for k, x in enumerate(refs):
            h.set(r + 1 + k, 1, xl_date(x.get("date")), st["day"])
            h.set(r + 1 + k, 2, x.get("label"), st["txt"])
            h.set(r + 1 + k, 2 + len(parts), x.get("total"), st["gbp_b"])

    # ------------------------------------------------ Swaps & upgrades
    s = book.sheet("Swaps & upgrades", tab="FFB45309")
    s.set(1, 1, "Swap ideas and optional upgrades", st["title"])
    s.heights[1] = 26
    s.set(2, 1, "Swap ideas need your OK (Approve / Not for me on the live page); they must fit the case, match or beat the spec and save £10+ "
                "(colour swaps excepted). Upgrades are tracked for reference only.", st["sub"])
    s.width({1: 14, 2: 44, 3: 11, 4: 17, 5: 12, 6: 12, 7: 9, 8: 70})
    heads = ["Part", "Idea", "Price", "Shop", "Stock", "Saving", "Fits", "Why"]
    s.set(4, 1, "Swap ideas waiting for your OK", st["sec"])
    for c, hh in enumerate(heads, 1):
        s.set(5, c, hh, st["hr"] if c in (3, 6) else st["h"])
    r = 6
    pend = [x for x in data.get("suggestions", []) if x.get("status") == "pending"]
    slot = {p["id"]: p.get("slot") for p in parts}
    for x in pend:
        bb = x.get("best") or {}
        s.set(r, 1, slot.get(x.get("part"), x.get("part")), st["bold"])
        s.set(r, 2, x.get("name"), st["wrap"])
        s.set(r, 3, bb.get("price"), st["gbp"])
        s.set(r, 4, bb.get("display"), st["txt"])
        s.set(r, 5, stock_text(bb), st["good"] if bb.get("in_stock") else st["warn"])
        s.set(r, 6, x.get("saving"), lsty(x.get("saving")))
        s.set(r, 7, "Yes" if x.get("fit_ok") else "No: " + ", ".join(x.get("fit_problems") or []), st["good"] if x.get("fit_ok") else st["bad"])
        s.set(r, 8, x.get("why"), st["mwrap"])
        r += 1
    if not pend:
        s.set(r, 2, "None right now — the daily sweep adds ideas when something cheaper passes every rule.", st["muted"])
        r += 1
    r += 1
    s.set(r, 1, "Optional upgrades and alternatives (difference vs today's pick)", st["sec"])
    r += 1
    for c, hh in enumerate(["Part", "Option", "Price", "Shop", "Stock", "Difference", "Fits", "Why"], 1):
        s.set(r, c, hh, st["hr"] if c in (3, 6) else st["h"])
    r += 1
    for x in data.get("upgrades", []):
        bb = x.get("best") or {}
        s.set(r, 1, slot.get(x.get("part"), x.get("part")), st["bold"])
        s.set(r, 2, x.get("name"), st["wrap"])
        s.set(r, 3, bb.get("price"), st["gbp"] if bb.get("in_stock") else st["gbp_oos"])
        s.set(r, 4, bb.get("display"), st["txt"])
        s.set(r, 5, stock_text(bb), st["good"] if bb.get("in_stock") else st["warn"])
        s.set(r, 6, x.get("extra"), dsty(x.get("extra")))
        s.set(r, 7, "Yes" if x.get("fit_ok") else ("No: " + ", ".join(x.get("fit_problems") or []) if x.get("fit_ok") is False else ""),
              st["good"] if x.get("fit_ok") else st["bad"])
        s.set(r, 8, x.get("why"), st["mwrap"])
        if bb.get("url"):
            s.links.append((ref(r, 4), bb["url"]))
        r += 1

    # ------------------------------------------------ Fit check
    f = book.sheet("Fit check", tab="FF0891B2")
    lim = data.get("fit_limits", {})
    f.set(1, 1, "Fit check — %s" % lim.get("case", "case"), st["title"])
    f.heights[1] = 26
    f.set(2, 1, "Every part against the case and board limits, recalculated each run. Swap ideas must pass the same checks.", st["sub"])
    f.width({1: 34, 2: 14, 3: 22, 4: 12, 5: 70})
    for c, hh in enumerate(["Check", "Value", "Limit", "Status", "Note"], 1):
        f.set(4, c, hh, st["h"])
    r = 5
    for x in data.get("fit", []):
        unit = x.get("unit") or ""
        lim_txt = ("≥ " if x.get("kind") == "min" else "" if x.get("kind") == "in" else "≤ ") + str(x.get("limit")) + (" " + unit if unit else "")
        f.set(r, 1, x.get("check"), st["bold"])
        f.set(r, 2, "%s%s" % (x.get("value"), " " + unit if unit else ""), st["txt"])
        f.set(r, 3, lim_txt, st["txt"])
        f.set(r, 4, "✓ OK" if x.get("ok") else "✕ Problem", st["good"] if x.get("ok") else st["bad"])
        f.set(r, 5, x.get("note") or "", st["mwrap"])
        r += 1
    if lim.get("notes"):
        f.set(r + 1, 1, lim["notes"], st["mwrap"])
        f.merges.append("A%d:E%d" % (r + 1, r + 1))
        f.heights[r + 1] = 60

    # ------------------------------------------------ How it works
    n = book.sheet("How it updates", tab="FF6B7280")
    n.width({1: 26, 2: 100})
    n.set(1, 1, "How this spreadsheet stays current", st["title"])
    n.heights[1] = 26
    rows = [
        ("Newest copy", xlsx_url, xlsx_url),
        ("Live page", site, site),
        ("Live CSV (for Google Sheets)", '=IMPORTDATA("%sprices.csv") in any Google Sheet keeps a live copy' % site, site + "prices.csv"),
        ("Updated", updated, None),
        ("Price checks", "GitHub Actions re-reads PriceSpy UK (every UK shop's price, stock and delivery) and the shops' own pages about 4× a day "
                         "(07:07, 11:07, 15:07, 19:07 UK) and after every change to the build, then rebuilds this file.", None),
        ("Daily sweep", "A Claude scheduled task at 07:46 UK re-verifies shops that block automated checks (Overclockers UK, Scan, Watercooling UK, "
                        "CCL), vets new shops' NI delivery, applies your page buttons and proposes swaps.", None),
        ("Rules", "UK shops that deliver to Northern Ireland only · new stock only (no marketplaces, used or refurbished; Amazon only when sold by "
                  "Amazon) · best price = item + NI delivery for a single order · all-black parts · a different model is only ever a suggestion "
                  "until you approve it, and must fit the case, match or beat the spec and save £10+.", None),
        ("Editing", "Edit the blue cells (budgets) freely; everything else is rebuilt from the live data, so download a fresh copy rather than "
                    "keeping edits here. Mark parts bought with the live page's Bought it button.", None),
    ]
    r = 3
    for k, v, url in rows:
        n.set(r, 1, k, st["bold"])
        if url and v == url:
            n.link(r, 2, url, v, st["link"])
        else:
            n.set(r, 2, v, st["wrap"])
            if url:
                n.links.append((ref(r, 2), url))
        lines_n = max(1, -(-len(str(v)) // 100))
        n.heights[r] = 15 * lines_n + 3 if lines_n > 1 else None
        r += 1
    plan = cfg.get("history_refs") or []
    changes = cfg.get("changes") or []
    if plan or changes:
        r += 1
        n.set(r, 1, "How the build got here", st["sec"])
        r += 1
        for x in plan:
            n.set(r, 1, short_date(x.get("date")) + " " + x.get("date", "")[:4], st["muted"])
            n.set(r, 2, "%s — %s" % (x.get("label"), gbp(x.get("total"))), st["txt"])
            r += 1
        for c in changes:
            n.set(r, 2, "• " + c, st["wrap"])
            lines_n = max(1, -(-len(c) // 100))
            n.heights[r] = 15 * lines_n + 3 if lines_n > 1 else None
            r += 1
    for k in list(n.heights):
        if n.heights[k] is None:
            del n.heights[k]
    return book


def export_csv(data, path):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["slot", "part", "price_gbp", "shop", "stock", "ni_delivery_gbp", "landed_gbp", "plan_gbp", "bought", "url", "updated"])
    for p in data.get("parts", []):
        b = p.get("best") or {}
        w.writerow([p.get("slot"), p.get("name"),
                    p.get("paid") if p.get("bought") else b.get("price"),
                    p.get("paid_retailer") if p.get("bought") else b.get("display"),
                    "bought" if p.get("bought") else stock_text(b),
                    b.get("delivery") if b.get("delivery") is not None else (0 if b.get("landed") is not None else None),
                    b.get("landed"), (p.get("baseline") or {}).get("price"),
                    "yes" if p.get("bought") else "", b.get("url") or "", data.get("generated")])
    t = data.get("totals", {})
    w.writerow(["TOTAL", "Parts total", t.get("build_total"), "", "", t.get("delivery"),
                round((t.get("build_total") or 0) + (t.get("delivery") or 0), 2), (data.get("baseline") or {}).get("parts_total"),
                "", "", data.get("generated")])
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as f:
        f.write(buf.getvalue())
    import os
    os.replace(tmp, path)


def export(data, hist, cfg, xlsx_path, csv_path, site):
    gen = data.get("generated")
    try:
        modified = datetime.strptime(gen, "%Y-%m-%dT%H:%M:%SZ")
    except (TypeError, ValueError):
        modified = datetime.now(timezone.utc).replace(tzinfo=None)
    build_workbook(data, hist, cfg, site).save(xlsx_path, modified)
    export_csv(data, csv_path)
