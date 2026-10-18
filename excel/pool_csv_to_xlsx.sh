#!/usr/bin/env bash
# ==============================================================================
#  pool_csv_to_xlsx.sh
#    jboss_pool_collector.sh が出力した CSV を Excel ブック (.xlsx) に変換します。
#
#  ・フォント : Meiryo UI (全セル・グラフ・印刷ヘッダ)
#  ・配色     : モノトーン (黒～グレー)。枯渇=黒 / 待機=濃グレー / 高負荷=淡グレー
#  ・セル設定 : 数値/日時の表示形式、罫線、縞模様、見出し固定、オートフィルタ、
#               列幅、印刷設定 (横向き・横1ページ・見出し行の繰り返し)
#  ・シート   : サマリ / グラフ / 時間別推移 / 枯渇イベント / 時系列 /
#               待機スレッド / 待機スタック / スレッドダンプ / プール設定 / 項目説明
#
#  動作要件 : RHEL 9.x (9.8) 標準の bash と python3 (3.6 以降, 標準ライブラリのみ)
#             ※ pip / openpyxl / Excel 本体は不要です
#
#  使い方   : bash pool_csv_to_xlsx.sh [-o 出力.xlsx] [--max-rows N] [--no-charts] CSVフォルダ
# ==============================================================================
set -euo pipefail

usage() {
    cat <<'USAGE'
使い方: bash pool_csv_to_xlsx.sh [オプション] CSVフォルダ

  jboss_pool_collector.sh の出力フォルダ (pool_timeseries.csv 等) から Excel を作成します。

オプション
  -o, --output FILE    出力ファイル (既定: <CSVフォルダ>/<フォルダ名>.xlsx)
      --max-rows N     時系列シートの最大行数 (既定: 1048575 = Excel の上限)
      --no-charts      グラフシートを作成しない
  -h, --help           このヘルプを表示

例
  bash pool_csv_to_xlsx.sh ./pool_report_task-1234_20261018_100000
  bash pool_csv_to_xlsx.sh -o /tmp/report.xlsx ./pool_report_task-1234_20261018_100000
USAGE
}

OUTPUT=""
MAX_ROWS=1048575
CHARTS=1
INPUT=""
while [ $# -gt 0 ]; do
    case "$1" in
        -o|--output)  [ $# -ge 2 ] || { echo "エラー: $1 には値が必要です" >&2; exit 1; }; OUTPUT=$2; shift 2 ;;
        --max-rows)   [ $# -ge 2 ] || { echo "エラー: $1 には値が必要です" >&2; exit 1; }; MAX_ROWS=$2; shift 2 ;;
        --no-charts)  CHARTS=0; shift ;;
        -h|--help)    usage; exit 0 ;;
        -*)           echo "エラー: 不明なオプションです: $1" >&2; usage >&2; exit 1 ;;
        *)            INPUT=$1; shift ;;
    esac
done
[ -n "$INPUT" ] || { usage >&2; exit 1; }
[ -d "$INPUT" ] || { echo "エラー: フォルダがありません: $INPUT" >&2; exit 1; }
[ -f "$INPUT/pool_timeseries.csv" ] || { echo "エラー: $INPUT に pool_timeseries.csv がありません" >&2; exit 1; }
case "$MAX_ROWS" in ''|*[!0-9]*) echo "エラー: --max-rows は整数で指定してください" >&2; exit 1 ;; esac
[ "$MAX_ROWS" -ge 1 ] && [ "$MAX_ROWS" -le 1048575 ] || { echo "エラー: --max-rows は 1～1048575 で指定してください" >&2; exit 1; }
INPUT=$(cd "$INPUT" && pwd)
if [ -z "$OUTPUT" ]; then OUTPUT="$INPUT/$(basename "$INPUT").xlsx"; fi

PYTHON=""
for p in python3 /usr/libexec/platform-python python3.12 python3.11 python3.9; do
    if command -v "$p" >/dev/null 2>&1; then PYTHON=$p; break; fi
done
[ -n "$PYTHON" ] || { echo "エラー: python3 が見つかりません (sudo dnf install python3)" >&2; exit 1; }
"$PYTHON" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 6) else 1)' ||
    { echo "エラー: python 3.6 以降が必要です" >&2; exit 1; }

export PYTHONIOENCODING=utf-8
exec "$PYTHON" - "$INPUT" "$OUTPUT" "$MAX_ROWS" "$CHARTS" <<'PYEOF'
# -*- coding: utf-8 -*-
# CSV → xlsx 変換 (標準ライブラリのみ。Office Open XML を直接生成)
import csv
import datetime
import math
import os
import re
import sys
import zipfile

IN_DIR, OUT_PATH, MAX_ROWS, CHARTS = sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4] == '1'
csv.field_size_limit(2147483647)

FONT = 'Meiryo UI'
NUM_RE = re.compile(r'^-?\d+(\.\d+)?([eE][-+]?\d+)?$')
DT_RE = re.compile(r'^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$')
CTRL_RE = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')
EPOCH = datetime.date(1899, 12, 30)


# ------------------------------------------------------------------------------
# 入力
# ------------------------------------------------------------------------------
def csv_path(name):
    return os.path.join(IN_DIR, name)


def read_csv(name):
    p = csv_path(name)
    if not os.path.isfile(p):
        return None
    with open(p, encoding='utf-8-sig', newline='') as f:
        return [r for r in csv.reader(f)]


META = {}
STATUS = {}
UI = {}
_cols = read_csv('columns.csv') or []
for r in _cols[1:]:
    if len(r) < 6:
        continue
    f, k, label, unit, fmt, desc = r[:6]
    if f == 'ui':
        UI[k] = label
    elif f == 'status':
        STATUS[k] = (label, desc)
    else:
        META[(f, k)] = (label, unit, fmt, desc)


def ui(key):
    return UI.get(key, key)


def meta(f, k):
    return META.get((f, k), (k, '', 'text', ''))


INFO = {}
INFO_ROWS = []
_info = read_csv('collection_info.csv') or []
for r in _info[1:]:
    if len(r) >= 3:
        INFO[r[0]] = r[2]
        INFO_ROWS.append(r)
try:
    WARN_PCT = float(INFO.get('warn_pct', '80'))
except ValueError:
    WARN_PCT = 80.0


# ------------------------------------------------------------------------------
# スタイル (固定テーブル。PowerShell 版と同一の定義)
# ------------------------------------------------------------------------------
FONTS = [  # (サイズ, 太字, 斜体, 色)
    (10, 0, 0, '000000'), (10, 1, 0, '000000'), (10, 1, 0, 'FFFFFF'), (16, 1, 0, '000000'),
    (9, 0, 0, '595959'), (11, 1, 0, '000000'), (10, 0, 1, '808080')]
FILLS = ['262626', 'F2F2F2', '000000', '595959', 'BFBFBF', 'D9D9D9']  # fillId = 2 + index
BORDERS = [None, ('thin', 'BFBFBF'), ('thin', '808080'), ('bottom', '000000')]
NUMFMTS = [(164, '#,##0.0'), (165, '0.0"%"'), (166, 'yyyy/mm/dd hh:mm:ss'), (167, 'yyyy/mm/dd hh:mm')]
# 名前, font, fill, border, numFmt, 横位置, 縦位置, 折返し
XF_DEFS = [
    ('default', 0, 0, 0, 0, '', '', 0),
    ('title', 3, 0, 0, 0, '', 'center', 0),
    ('subtitle', 4, 0, 0, 0, '', 'center', 0),
    ('section', 5, 0, 3, 0, '', 'center', 0),
    ('hdr', 2, 2, 2, 0, 'center', 'center', 1),
    ('lbl', 1, 3, 1, 0, 'left', 'center', 1),
    ('text', 0, 0, 1, 0, '', 'center', 0),
    ('text_z', 0, 3, 1, 0, '', 'center', 0),
    ('int', 0, 0, 1, 3, '', 'center', 0),
    ('int_z', 0, 3, 1, 3, '', 'center', 0),
    ('dec1', 0, 0, 1, 164, '', 'center', 0),
    ('dec1_z', 0, 3, 1, 164, '', 'center', 0),
    ('pct', 0, 0, 1, 165, '', 'center', 0),
    ('pct_z', 0, 3, 1, 165, '', 'center', 0),
    ('pctb', 1, 0, 1, 165, '', 'center', 0),
    ('pctb_z', 1, 3, 1, 165, '', 'center', 0),
    ('dt', 0, 0, 1, 166, 'center', 'center', 0),
    ('dt_z', 0, 3, 1, 166, 'center', 'center', 0),
    ('dtm', 0, 0, 1, 167, 'center', 'center', 0),
    ('dtm_z', 0, 3, 1, 167, 'center', 'center', 0),
    ('wrap', 0, 0, 1, 0, 'left', 'top', 1),
    ('wrap_z', 0, 3, 1, 0, 'left', 'top', 1),
    ('st_OK', 0, 0, 1, 0, 'center', 'center', 0),
    ('st_OK_z', 0, 3, 1, 0, 'center', 'center', 0),
    ('st_HIGH', 1, 6, 1, 0, 'center', 'center', 0),
    ('st_WAIT', 2, 5, 1, 0, 'center', 'center', 0),
    ('st_EXHAUSTED', 2, 4, 1, 0, 'center', 'center', 0),
    ('st_GRAY', 6, 0, 1, 0, 'center', 'center', 0),
    ('st_GRAY_z', 6, 3, 1, 0, 'center', 'center', 0),
    ('note', 6, 0, 0, 0, '', 'center', 0),
    ('textb', 1, 0, 1, 0, '', 'center', 0),
    ('textb_z', 1, 3, 1, 0, '', 'center', 0),
    ('textw', 0, 0, 1, 0, 'left', 'center', 1),
]
XF = dict((d[0], i) for i, d in enumerate(XF_DEFS))


def styles_xml():
    o = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
         '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    o.append('<numFmts count="%d">' % len(NUMFMTS))
    for i, code in NUMFMTS:
        o.append('<numFmt numFmtId="%d" formatCode="%s"/>' % (i, esc_attr(code)))
    o.append('</numFmts><fonts count="%d">' % len(FONTS))
    for sz, b, it, col in FONTS:
        o.append('<font>%s%s<sz val="%d"/><color rgb="FF%s"/><name val="%s"/><family val="3"/><charset val="128"/></font>'
                 % ('<b/>' if b else '', '<i/>' if it else '', sz, col, FONT))
    o.append('</fonts><fills count="%d"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
             % (len(FILLS) + 2))
    for col in FILLS:
        o.append('<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/><bgColor indexed="64"/></patternFill></fill>' % col)
    o.append('</fills><borders count="%d">' % len(BORDERS))
    for b in BORDERS:
        if b is None:
            o.append('<border><left/><right/><top/><bottom/><diagonal/></border>')
        elif b[0] == 'bottom':
            o.append('<border><left/><right/><top/><bottom style="medium"><color rgb="FF%s"/></bottom><diagonal/></border>' % b[1])
        else:
            s = '<%s style="thin"><color rgb="FF' + b[1] + '"/></%s>'
            o.append('<border>' + ''.join(s % (e, e) for e in ('left', 'right', 'top', 'bottom')) + '<diagonal/></border>')
    o.append('</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
    o.append('<cellXfs count="%d">' % len(XF_DEFS))
    for name, font, fill, border, nf, h, v, wrap in XF_DEFS:
        a = '<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"' % (nf, font, fill, border)
        a += ' applyNumberFormat="1"' if nf else ''
        a += ' applyFont="1"' if font else ''
        a += ' applyFill="1"' if fill else ''
        a += ' applyBorder="1"' if border else ''
        if h or v or wrap:
            al = '<alignment'
            al += ' horizontal="%s"' % h if h else ''
            al += ' vertical="%s"' % v if v else ''
            al += ' wrapText="1"' if wrap else ''
            o.append(a + ' applyAlignment="1">' + al + '/></xf>')
        else:
            o.append(a + '/>')
    o.append('</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
    o.append('<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>')
    return ''.join(o)


# ------------------------------------------------------------------------------
# セル生成
# ------------------------------------------------------------------------------
def esc(s):
    s = CTRL_RE.sub('', s.replace('\r', ''))
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def esc_attr(s):
    return esc(s).replace('"', '&quot;')


def col_letter(i):
    s = ''
    i += 1
    while i > 0:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def dispw(s):
    w = 0
    for ch in s:
        w += 2 if ord(ch) > 0x2E80 else 1
    return w


def c_str(ref, text, style):
    if text == '':
        return '<c r="%s" s="%d"/>' % (ref, XF[style])
    text = text[:32000]
    sp = ' xml:space="preserve"' if (text != text.strip() or '\n' in text) else ''
    return '<c r="%s" s="%d" t="inlineStr"><is><t%s>%s</t></is></c>' % (ref, XF[style], sp, esc(text))


def c_num(ref, num, style):
    return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, XF[style], num)


def c_empty(ref, style):
    return '<c r="%s" s="%d"/>' % (ref, XF[style])


def dt_serial(v):
    m = DT_RE.match(v)
    if not m:
        return None, False
    y, mo, d, hh, mi = (int(m.group(i)) for i in range(1, 6))
    ss = int(m.group(6)) if m.group(6) else 0
    try:
        days = (datetime.date(y, mo, d) - EPOCH).days
    except ValueError:
        return None, False
    return '%.8f' % (days + (hh * 3600 + mi * 60 + ss) / 86400.0), m.group(6) is not None


def status_style(code, zebra):
    if code in ('HIGH', 'WAIT', 'EXHAUSTED'):
        return 'st_' + code
    if code in ('NOSTATS', 'ERROR'):
        return 'st_GRAY_z' if zebra else 'st_GRAY'
    return 'st_OK_z' if zebra else 'st_OK'


def value_cell(ref, v, kind, zebra, bold=False):
    """kind: text/int/dec1/pct/datetime/status/wrap"""
    z = '_z' if zebra else ''
    if v == '':
        base = {'int': 'int', 'dec1': 'dec1', 'pct': 'pct', 'datetime': 'dt', 'wrap': 'wrap'}.get(kind, 'text')
        return c_empty(ref, base + z)
    if kind in ('int', 'dec1', 'pct'):
        if NUM_RE.match(v):
            if kind == 'pct' and bold:
                return c_num(ref, v, 'pctb' + z)
            return c_num(ref, v, kind + z)
        return c_str(ref, v, 'text' + z)
    if kind == 'datetime':
        n, has_sec = dt_serial(v)
        if n is None:
            return c_str(ref, v, 'text' + z)
        return c_num(ref, n, ('dt' if has_sec else 'dtm') + z)
    if kind == 'status':
        return c_str(ref, STATUS.get(v, (v, ''))[0], status_style(v, zebra))
    if kind == 'wrap':
        return c_str(ref, v, 'wrap' + z)
    return c_str(ref, v, ('textb' if bold else 'text') + z)


def value_width(v, kind):
    if kind == 'datetime':
        return 19 if len(v) > 16 else 16
    if kind == 'status':
        return dispw(STATUS.get(v, (v, ''))[0]) + 2
    if kind in ('int', 'dec1', 'pct') and NUM_RE.match(v):
        ip = v.lstrip('-').split('.')[0]
        return len(v) + (len(ip) - 1) // 3 + (2 if kind == 'pct' else 0)
    return dispw(v.split('\n')[0])


# ------------------------------------------------------------------------------
# 表 (ヘッダ 1 行 + データ行)
# ------------------------------------------------------------------------------
class Table(object):
    def __init__(self, fname, header, front=(), back=(), skip=(), fixed=None):
        self.fname = fname
        self.header = header
        idx = dict((h, i) for i, h in enumerate(header))
        order = [h for h in front if h in idx]
        order += [h for h in header if h not in order and h not in back and h not in skip]
        order += [h for h in back if h in idx]
        self.cols = []
        for h in order:
            label, unit, fmt, desc = meta(fname, h)
            self.cols.append({'key': h, 'src': idx[h], 'label': label, 'unit': unit, 'kind': fmt or 'text',
                              'fixed': (fixed or {}).get(h)})
        self.ncol = len(self.cols)
        self.last_col = col_letter(self.ncol - 1)

    def letter(self, key):
        for i, c in enumerate(self.cols):
            if c['key'] == key:
                return col_letter(i)
        return None

    def compute_widths(self, sample_rows):
        self.widths = []
        hdr_lines = 1
        for c in self.cols:
            lw = dispw(c['label'] + (' (%s)' % c['unit'] if c['unit'] else ''))
            if c['fixed']:
                w = c['fixed']
            else:
                vw = 0
                for r in sample_rows:
                    if c['src'] < len(r) and r[c['src']] != '':
                        vw = max(vw, value_width(r[c['src']], c['kind']))
                # 見出しは折り返すため、値の幅を優先 (見出しは最大 3 行程度に収める)
                w = max(vw + 2, (lw + 2) // 3 + 2, 8)
                w = min(w, 50)
            lines = int(math.ceil(lw / float(max(w - 1, 1)))) + int(math.ceil(len(c['key']) / float(max(w - 1, 1))))
            hdr_lines = max(hdr_lines, lines)
            self.widths.append(w)
        self.header_height = min(15 * hdr_lines + 3, 120)

    def header_row(self, r):
        cells = []
        for i, c in enumerate(self.cols):
            t = c['label'] + (' (%s)' % c['unit'] if c['unit'] else '') + '\n' + c['key']
            cells.append(c_str(col_letter(i) + str(r), t, 'hdr'))
        return '<row r="%d" ht="%s" customHeight="1">%s</row>' % (r, self.header_height, ''.join(cells))

    def data_row(self, r, row, zebra):
        cells = []
        ht = None
        for i, c in enumerate(self.cols):
            v = row[c['src']] if c['src'] < len(row) else ''
            bold = False
            if c['kind'] == 'pct' and c['key'].startswith('usage_pct') and NUM_RE.match(v) and float(v) >= WARN_PCT:
                bold = True
            if c['key'] == 'leak_suspect' and v == 'yes':
                bold = True
            if c['kind'] == 'wrap' and v:
                lines = v.count('\n') + 1
                for line in v.split('\n'):
                    lines += int(dispw(line) / max(self.widths[i] - 1, 1))
                ht = max(ht or 0, min(409, 13 * lines + 4))
            cells.append(value_cell(col_letter(i) + str(r), v, c['kind'], zebra, bold))
        h = ' ht="%s" customHeight="1"' % ht if ht else ''
        return '<row r="%d"%s>%s</row>' % (r, h, ''.join(cells))


# ------------------------------------------------------------------------------
# シート XML
# ------------------------------------------------------------------------------
class Sheet(object):
    def __init__(self, name, ui_key):
        self.name = name
        self.ui_key = ui_key
        self.charts = []        # (anchor, chart_xml)
        self.filter_ref = None
        self.print_titles = False
        self.title_cols = 0


SHEETS = []
WB = {'zip': None}


def sheet_ref(name):
    return "'" + name.replace("'", "''") + "'"


def write_sheet(sheet, rows_iter, widths, freeze=None, autofilter=None, merges=None, grid=True,
                selected=False, landscape=True, default_width=None):
    # 横幅の狭いシートは横 1 ページに収め、広いシートは等倍で固定列を各ページに繰り返す
    fit = widths is None or sum(widths) <= 260
    if not fit and freeze:
        sheet.title_cols = freeze[0]
    idx = len(SHEETS) + 1
    sheet.index = idx
    SHEETS.append(sheet)
    zf = WB['zip']
    with zf.open('xl/worksheets/sheet%d.xml' % idx, 'w', force_zip64=True) as fh:
        def w(s):
            fh.write(s.encode('utf-8'))
        w('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        w('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        w('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' if fit else '<sheetPr/>')
        sv = '<sheetView workbookViewId="0"%s%s>' % ('' if grid else ' showGridLines="0"', ' tabSelected="1"' if selected else '')
        if freeze and (freeze[0] or freeze[1]):
            xs, ys = freeze
            tl = col_letter(xs) + str(ys + 1)
            if xs and ys:
                sv += ('<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                       '<selection pane="topRight"/><selection pane="bottomLeft"/>'
                       '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>' % (xs, ys, tl, tl, tl))
            elif ys:
                sv += ('<pane ySplit="%d" topLeftCell="%s" activePane="bottomLeft" state="frozen"/>'
                       '<selection pane="bottomLeft" activeCell="%s" sqref="%s"/>' % (ys, tl, tl, tl))
            else:
                sv += ('<pane xSplit="%d" topLeftCell="%s" activePane="topRight" state="frozen"/>'
                       '<selection pane="topRight" activeCell="%s" sqref="%s"/>' % (xs, tl, tl, tl))
        else:
            sv += '<selection activeCell="A1" sqref="A1"/>'
        w('<sheetViews>' + sv + '</sheetView></sheetViews>')
        w('<sheetFormatPr %sdefaultRowHeight="15"/>' % ('defaultColWidth="%s" ' % default_width if default_width else ''))
        if widths:
            w('<cols>' + ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, wd)
                                 for i, wd in enumerate(widths)) + '</cols>')
        w('<sheetData>')
        for row in rows_iter:
            w(row)
        w('</sheetData>')
        if autofilter:
            w('<autoFilter ref="%s"/>' % autofilter)
            sheet.filter_ref = autofilter
        if merges:
            w('<mergeCells count="%d">%s</mergeCells>' % (len(merges), ''.join('<mergeCell ref="%s"/>' % m for m in merges)))
        w('<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
        w('<pageSetup paperSize="9" orientation="%s"%s/>' % ('landscape' if landscape else 'portrait',
                                                             ' fitToWidth="1" fitToHeight="0"' if fit else ''))
        w('<headerFooter><oddHeader>&amp;L&amp;"Meiryo UI,Regular"&amp;9&amp;A</oddHeader>'
          '<oddFooter>&amp;C&amp;"Meiryo UI,Regular"&amp;9&amp;P / &amp;N</oddFooter></headerFooter>')
        if sheet.charts:
            w('<drawing r:id="rId1"/>')
        w('</worksheet>')


def data_sheet(name_key, fname, rows, front=(), back=(), skip=(), fixed=None, freeze_cols=1):
    """CSV 1 ファイルを表形式のシートにする (rows はヘッダ込み)"""
    if not rows:
        return None
    t = Table(fname, rows[0], front, back, skip, fixed)
    data = rows[1:]
    t.compute_widths(data[:500])
    sheet = Sheet(ui(name_key), name_key)

    def gen():
        yield t.header_row(1)
        if not data:
            yield '<row r="2">%s</row>' % c_str('A2', ui('no_data'), 'note')
        for i, r in enumerate(data):
            yield t.data_row(i + 2, r, i % 2 == 1)
    last = max(len(data) + 1, 2)
    sheet.print_titles = True
    write_sheet(sheet, gen(), t.widths, freeze=(freeze_cols, 1), autofilter='A1:%s%d' % (t.last_col, last))
    return sheet


# ------------------------------------------------------------------------------
# グラフ (DrawingML)
# ------------------------------------------------------------------------------
def rpr(sz, bold, color='000000'):
    return ('<a:defRPr sz="%d" b="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/></a:defRPr>' % (sz, bold, color, FONT, FONT, FONT))


def txpr(sz=900, bold=0, color='404040', rot=None):
    bp = '<a:bodyPr rot="%d" vert="horz"/>' % rot if rot is not None else '<a:bodyPr/>'
    return '<c:txPr>%s<a:lstStyle/><a:p><a:pPr>%s</a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>' % (bp, rpr(sz, bold, color))


def rich_title(text, sz, rot=None):
    bp = '<a:bodyPr rot="%d" vert="horz"/>' % rot if rot is not None else '<a:bodyPr/>'
    return ('<c:title><c:tx><c:rich>%s<a:lstStyle/><a:p><a:pPr>%s</a:pPr><a:r><a:rPr lang="ja-JP" sz="%d" b="1">'
            '<a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="%s"/><a:ea typeface="%s"/>'
            '<a:cs typeface="%s"/></a:rPr><a:t>%s</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
            % (bp, rpr(sz, 1), sz, FONT, FONT, FONT, esc(text)))


def ln(color, w=9525, dash='solid'):
    return '<a:ln w="%d" cap="rnd"><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:prstDash val="%s"/><a:round/></a:ln>' % (w, color, dash)


GRID = '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
AXLINE = '<c:spPr><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr>'


def val_axis(ax_id, cross_id, pos, title, fmt, grid, min0, cross_between, rot=None, vmin=None, vmax=None, major=None):
    sc = '<c:orientation val="minMax"/>'
    if vmax is not None:
        sc += '<c:max val="%s"/>' % vmax
    if vmin is not None:
        sc += '<c:min val="%s"/>' % vmin
    elif min0:
        sc += '<c:min val="0"/>'
    o = '<c:valAx><c:axId val="%d"/><c:scaling>%s</c:scaling><c:delete val="0"/>' % (ax_id, sc)
    o += '<c:axPos val="%s"/>%s' % (pos, GRID if grid else '')
    if title:
        o += rich_title(title, 900, -5400000 if pos == 'l' else None)
    o += '<c:numFmt formatCode="%s" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>' % esc_attr(fmt)
    o += '<c:tickLblPos val="low"/>%s%s<c:crossAx val="%d"/><c:crosses val="autoZero"/><c:crossBetween val="%s"/>' % (
        AXLINE, txpr(800, 0, '404040', rot), cross_id, cross_between)
    if major is not None:
        o += '<c:majorUnit val="%s"/>' % major
    return o + '</c:valAx>'


def chart_space(title, plot, legend=True):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><c:chart>%s<c:autoTitleDeleted val="0"/>'
            '<c:plotArea><c:layout/>%s<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>%s'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
            '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>'
            '%s</c:chartSpace>' % (rich_title(title, 1100), plot,
                                   '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>%s</c:legend>' % txpr(900) if legend else '',
                                   txpr(900)))


def scatter_chart(title, xref, series, xfmt, ytitle, xmin, xmax):
    """series: (名前, Y 範囲, 色, 線幅EMU, 線種, マーカー)"""
    s = '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>'
    for i, (name, yref, color, wd, dash, marker) in enumerate(series):
        s += '<c:ser><c:idx val="%d"/><c:order val="%d"/><c:tx><c:v>%s</c:v></c:tx><c:spPr>%s</c:spPr>' % (i, i, esc(name), ln(color, wd, dash))
        if marker == 'none':
            s += '<c:marker><c:symbol val="none"/></c:marker>'
        else:
            s += ('<c:marker><c:symbol val="%s"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</c:spPr></c:marker>'
                  % (marker, color, ln(color, 9525)))
        s += '<c:xVal><c:numRef><c:f>%s</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>%s</c:f></c:numRef></c:yVal><c:smooth val="0"/></c:ser>' % (
            esc(xref), esc(yref))
    s += '<c:axId val="5001"/><c:axId val="5002"/></c:scatterChart>'
    s += val_axis(5001, 5002, 'b', None, xfmt, False, False, 'midCat', -2700000, xmin, xmax)
    s += val_axis(5002, 5001, 'l', ytitle, '#,##0', True, True, 'midCat')
    return chart_space(title, s)


def bar_chart(title, catref, series, catfmt, ytitle, ymax):
    """series: (名前, 値範囲, 色)"""
    s = '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
    for i, (name, vref, color) in enumerate(series):
        s += ('<c:ser><c:idx val="%d"/><c:order val="%d"/><c:tx><c:v>%s</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
              '<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/><c:cat><c:numRef><c:f>%s</c:f></c:numRef></c:cat>'
              '<c:val><c:numRef><c:f>%s</c:f></c:numRef></c:val></c:ser>' % (i, i, esc(name), color, esc(catref), esc(vref)))
    s += '<c:gapWidth val="60"/><c:axId val="6001"/><c:axId val="6002"/></c:barChart>'
    s += ('<c:catAx><c:axId val="6001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>'
          '<c:numFmt formatCode="%s" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>'
          '%s%s<c:crossAx val="6002"/><c:crosses val="autoZero"/><c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
          '<c:noMultiLvlLbl val="0"/></c:catAx>' % (esc_attr(catfmt), AXLINE, txpr(800, 0, '404040', -2700000)))
    # 件数が少ない場合に目盛りが小数 (0, 0.2 ...) にならないよう固定
    small = ymax < 5
    s += val_axis(6002, 6001, 'l', ytitle, '#,##0', True, True, 'between', None, None, '5' if small else None, '1' if small else None)
    return chart_space(title, s)


def drawing_xml(charts):
    o = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
         'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">']
    for i, (anchor, _) in enumerate(charts):
        c1, r1, c2, r2 = anchor
        o.append('<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row>'
                 '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row>'
                 '<xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="%d" name="Chart %d"/>'
                 '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                 '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                 '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId%d"/>'
                 '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>'
                 % (c1, r1, c2, r2, i + 2, i + 1, i + 1))
    o.append('</xdr:wsDr>')
    return ''.join(o)


# ------------------------------------------------------------------------------
# 各シート
# ------------------------------------------------------------------------------
def build_summary(summary_rows, truncated):
    sheet = Sheet(ui('sheet_summary'), 'sheet_summary')
    hdr = summary_rows[0] if summary_rows else []
    data = summary_rows[1:] if summary_rows else []
    idx = dict((h, i) for i, h in enumerate(hdr))
    nds = len(data)
    span = max(nds, 3)
    widths = [34]
    for d in data:
        wd = 26
        for key in ('datasource', 'jndi_name'):
            if key in idx and idx[key] < len(d):
                wd = max(wd, min(44, dispw(d[idx[key]]) + 3))
        widths.append(wd)
    while len(widths) < span + 1:
        widths.append(26)
    last = col_letter(span)
    rows = []
    merges = []
    rows.append('<row r="1" ht="30" customHeight="1">%s</row>' % c_str('A1', ui('report_title'), 'title'))
    gen = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
    sub = '%s: %s    %s: %s' % (ui('generated'), gen, ui('source_dir'), IN_DIR)
    rows.append('<row r="2">%s</row>' % c_str('A2', sub, 'subtitle'))
    r = 3
    if truncated:
        rows.append('<row r="3">%s</row>' % c_str('A3', ui('truncated'), 'note'))
    r = 4
    # データソース別サマリ
    rows.append('<row r="%d" ht="21" customHeight="1">%s</row>' % (r, ''.join(
        [c_str('A%d' % r, ui('sec_summary'), 'section')] + [c_empty(col_letter(i) + str(r), 'section') for i in range(1, span + 1)])))
    r += 1
    cells = [c_str('A%d' % r, ui('col_item'), 'hdr')]
    for j, d in enumerate(data):
        cells.append(c_str(col_letter(j + 1) + str(r), d[idx['datasource']], 'hdr'))
    rows.append('<row r="%d" ht="21" customHeight="1">%s</row>' % (r, ''.join(cells)))
    r += 1
    for key in hdr:
        if key in ('datasource', 'verdict_code'):
            continue
        label, unit, fmt, desc = meta('summary', key)
        cells = [c_str('A%d' % r, label + (' (%s)' % unit if unit else ''), 'lbl')]
        ht = None
        for j, d in enumerate(data):
            v = d[idx[key]] if idx[key] < len(d) else ''
            ref = col_letter(j + 1) + str(r)
            if key == 'verdict':
                cells.append(c_str(ref, v, status_style(d[idx['verdict_code']], False)))
            elif key == 'comment':
                cells.append(c_str(ref, v, 'wrap'))
                lines = 0
                for part in v.split('\n'):
                    lines += int(math.ceil(dispw(part) / float(widths[j + 1] - 2))) or 1
                ht = max(ht or 0, min(409, 15 * lines + 6))
            else:
                bold = (key.startswith('usage_pct') and NUM_RE.match(v) is not None and float(v) >= WARN_PCT) or \
                       (key == 'leak_suspect' and v == 'yes')
                cells.append(value_cell(ref, v, fmt or 'text', False, bold))
        rows.append('<row r="%d"%s>%s</row>' % (r, ' ht="%s" customHeight="1"' % ht if ht else '', ''.join(cells)))
        r += 1
    # 収集条件・環境
    r += 1
    rows.append('<row r="%d" ht="21" customHeight="1">%s</row>' % (r, ''.join(
        [c_str('A%d' % r, ui('sec_info'), 'section')] + [c_empty(col_letter(i) + str(r), 'section') for i in range(1, span + 1)])))
    r += 1
    merged_w = sum(widths[1:span + 1])
    for row in INFO_ROWS:
        cells = [c_str('A%d' % r, row[1], 'lbl'), c_str('B%d' % r, row[2], 'textw')]
        cells += [c_empty(col_letter(i) + str(r), 'textw') for i in range(2, span + 1)]
        lines = int(math.ceil(dispw(row[2]) / float(merged_w - 2))) if row[2] else 1
        ht = ' ht="%s" customHeight="1"' % (15 * lines + 3) if lines > 1 else ''
        rows.append('<row r="%d"%s>%s</row>' % (r, ht, ''.join(cells)))
        merges.append('B%d:%s%d' % (r, last, r))
        r += 1
    # 状態の凡例
    r += 1
    rows.append('<row r="%d" ht="21" customHeight="1">%s</row>' % (r, ''.join(
        [c_str('A%d' % r, ui('sec_legend'), 'section')] + [c_empty(col_letter(i) + str(r), 'section') for i in range(1, span + 1)])))
    r += 1
    for code in ('EXHAUSTED', 'WAIT', 'HIGH', 'OK', 'NOSTATS', 'ERROR'):
        if code not in STATUS:
            continue
        label, desc = STATUS[code]
        cells = [c_str('A%d' % r, label + ' (' + code + ')', status_style(code, False)), c_str('B%d' % r, desc, 'textw')]
        cells += [c_empty(col_letter(i) + str(r), 'textw') for i in range(2, span + 1)]
        rows.append('<row r="%d">%s</row>' % (r, ''.join(cells)))
        merges.append('B%d:%s%d' % (r, last, r))
        r += 1
    write_sheet(sheet, rows, widths, grid=False, selected=True, merges=merges, landscape=False)
    return sheet


def build_charts(ts_sheet_name, ts_ranges, ts_letters, hourly_sheet_name, hourly_ranges, h_letters):
    sheet = Sheet(ui('sheet_charts'), 'sheet_charts')
    rows = []
    r0 = 0
    tsn = sheet_ref(ts_sheet_name)
    hn = sheet_ref(hourly_sheet_name) if hourly_sheet_name else None
    for key, label in ts_ranges['order']:
        a, b = ts_ranges[key]
        rows.append('<row r="%d" ht="21" customHeight="1">%s</row>' % (r0 + 1, c_str('A%d' % (r0 + 1), '■ ' + label, 'section')))

        def rng(letter, s=a, e=b):
            return '%s!$%s$%d:$%s$%d' % (tsn, letter, s, letter, e)
        series = [
            (ui('ser_inuse'), rng(ts_letters['InUseCount']), '000000', 28575, 'solid', 'none'),
            (ui('ser_active'), rng(ts_letters['ActiveCount']), '7F7F7F', 19050, 'solid', 'none'),
            (ui('ser_idle'), rng(ts_letters['IdleCount']), 'BFBFBF', 12700, 'solid', 'none'),
            (ui('ser_maxpool'), rng(ts_letters['max_pool_size']), '000000', 15875, 'dash', 'none'),
        ]
        xmin, xmax = ts_ranges.get(('x', key), (None, None))
        sheet.charts.append(((0, r0 + 1, 16, r0 + 20),
                             scatter_chart(ui('chart_ts') + ' - ' + label, rng(ts_letters['timestamp']), series,
                                           'mm/dd hh:mm', ui('axis_conn'), xmin, xmax)))
        if hn and key in hourly_ranges:
            ha, hb = hourly_ranges[key]

            def hr(letter, s=ha, e=hb):
                return '%s!$%s$%d:$%s$%d' % (hn, letter, s, letter, e)
            hs = [
                (ui('ser_inuse_max'), hr(h_letters['inuse_max']), '000000', 22225, 'solid', 'circle'),
                (ui('ser_inuse_avg'), hr(h_letters['inuse_avg']), '7F7F7F', 19050, 'solid', 'square'),
                (ui('ser_inuse_min'), hr(h_letters['inuse_min']), 'BFBFBF', 12700, 'sysDash', 'none'),
                (ui('ser_maxpool'), hr(h_letters['max_pool_size']), '000000', 15875, 'dash', 'none'),
            ]
            hxmin, hxmax = hourly_ranges.get(('x', key), (None, None))
            sheet.charts.append(((0, r0 + 21, 8, r0 + 38),
                                 scatter_chart(ui('chart_hourly') + ' - ' + label, hr(h_letters['bucket_start']), hs,
                                               'mm/dd hh:mm', ui('axis_conn'), hxmin, hxmax)))
            bs = [
                (ui('ser_wait'), hr(h_letters['d_WaitCount']), '808080'),
                (ui('ser_bf'), hr(h_letters['d_BlockingFailureCount']), '000000'),
            ]
            sheet.charts.append(((8, r0 + 21, 16, r0 + 38),
                                 bar_chart(ui('chart_hourly_err') + ' - ' + label, hr(h_letters['bucket_start']), bs,
                                           'mm/dd hh:mm', ui('axis_count'), hourly_ranges.get(('y', key), 0.0))))
        r0 += 40
    write_sheet(sheet, rows, None, grid=False, default_width=9)
    return sheet


def ds_key(row, i_ds, i_ty):
    return (row[i_ds] if i_ds < len(row) else '') + '\x1f' + (row[i_ty] if i_ty < len(row) else '')


def build_timeseries():
    """時系列: データソース単位にまとめて出力 (メモリ節約のためデータソースごとに再読込)"""
    path = csv_path('pool_timeseries.csv')
    with open(path, encoding='utf-8-sig', newline='') as f:
        rd = csv.reader(f)
        header = next(rd)
        idx = dict((h, i) for i, h in enumerate(header))
        i_ds, i_ty = idx.get('datasource', 0), idx.get('ds_type', 0)
        i_ts = idx.get('timestamp', 0)
        order, counts, sample, stamps = [], {}, [], {}
        for row in rd:
            if not row:
                continue
            k = ds_key(row, i_ds, i_ty)
            if k not in counts:
                counts[k] = 0
                order.append(k)
                stamps[k] = []
            counts[k] += 1
            stamps[k].append(dt_serial(row[i_ts])[0] if i_ts < len(row) else None)
            if len(sample) < 500:
                sample.append(row)
    t = Table('timeseries', header, front=('timestamp', 'datasource'), back=('epoch', 'elapsed_sec', 'sample_no', 'host'))
    t.compute_widths(sample)
    nds = max(len(order), 1)
    per_ds = MAX_ROWS // nds
    truncated = any(c > per_ds for c in counts.values())
    ranges = {'order': []}
    r = 2
    for k in order:
        n = min(counts[k], per_ds)
        if n > 0:
            ranges[k] = (r, r + n - 1)
            ranges['order'].append((k, k.replace('\x1f', ' (') + ')'))
            xs = [float(x) for x in stamps[k][:n] if x is not None]
            if xs:
                lo, hi = min(xs), max(xs)
                if hi - lo < 1e-9:
                    lo, hi = lo - 5 / 1440.0, hi + 5 / 1440.0
                ranges['x', k] = ('%.8f' % lo, '%.8f' % hi)
            r += n
    last_row = max(r - 1, 2)

    def gen():
        yield t.header_row(1)
        rn = 2
        for k in order:
            written = 0
            with open(path, encoding='utf-8-sig', newline='') as f:
                rd = csv.reader(f)
                next(rd)
                for row in rd:
                    if not row or ds_key(row, i_ds, i_ty) != k:
                        continue
                    if written >= per_ds:
                        break
                    yield t.data_row(rn, row, rn % 2 == 1)
                    rn += 1
                    written += 1
    sheet = Sheet(ui('sheet_timeseries'), 'sheet_timeseries')
    sheet.print_titles = True
    sheet.pending = (gen, t, last_row)
    letters = dict((key, t.letter(key)) for key in ('timestamp', 'InUseCount', 'ActiveCount', 'IdleCount', 'max_pool_size'))
    return sheet, ranges, letters, truncated


def build_hourly_info(rows):
    if not rows:
        return {}, {}
    hdr = rows[0]
    t = Table('hourly', hdr, front=('bucket_start', 'bucket_end', 'datasource'))
    idx = dict((h, i) for i, h in enumerate(hdr))
    ranges = {}
    xs, ymax = {}, {}
    try:
        half = float(INFO.get('bucket_min', '60')) / 1440.0 / 2
    except ValueError:
        half = 60 / 1440.0 / 2
    for i, row in enumerate(rows[1:]):
        k = ds_key(row, idx.get('datasource', 0), idx.get('ds_type', 0))
        if k not in ranges:
            ranges[k] = [i + 2, i + 2]
            xs[k] = []
            ymax[k] = 0.0
        ranges[k][1] = i + 2
        v = dt_serial(row[idx['bucket_start']])[0] if 'bucket_start' in idx else None
        if v is not None:
            xs[k].append(float(v))
        for key in ('d_WaitCount', 'd_BlockingFailureCount'):
            if key in idx and idx[key] < len(row) and NUM_RE.match(row[idx[key]]):
                ymax[k] = max(ymax[k], float(row[idx[key]]))
    for k in list(ranges.keys()):
        if xs[k]:
            ranges['x', k] = ('%.8f' % (min(xs[k]) - half), '%.8f' % (max(xs[k]) + half))
        ranges['y', k] = ymax[k]
    letters = dict((key, t.letter(key)) for key in ('bucket_start', 'inuse_max', 'inuse_avg', 'inuse_min', 'max_pool_size',
                                                    'd_WaitCount', 'd_BlockingFailureCount'))
    return dict((k, tuple(v) if isinstance(v, list) else v) for k, v in ranges.items()), letters


def build_columns_sheet():
    sheet = Sheet(ui('sheet_columns'), 'sheet_columns')
    fmap = {'summary': 'sheet_summary', 'info': 'sheet_summary', 'hourly': 'sheet_hourly', 'events': 'sheet_events',
            'timeseries': 'sheet_timeseries', 'queued': 'sheet_queued', 'qstack': 'sheet_qstack',
            'threaddump': 'sheet_threaddump', 'config': 'sheet_config', 'status': 'sec_legend'}
    widths = [16, 30, 36, 9, 90]
    rows = ['<row r="1" ht="21" customHeight="1">%s</row>' % ''.join(
        c_str(col_letter(i) + '1', ui(k), 'hdr') for i, k in enumerate(('col_sheet', 'col_item', 'col_key', 'col_unit', 'col_desc')))]
    r = 2
    for row in _cols[1:]:
        if len(row) < 6 or row[0] == 'ui' or row[0] not in fmap:
            continue
        z = r % 2 == 1
        zs = '_z' if z else ''
        lab_style = status_style(row[1], z) if row[0] == 'status' else 'text' + zs
        cells = [c_str('A%d' % r, ui(fmap[row[0]]), 'text' + zs), c_str('B%d' % r, row[2], lab_style),
                 c_str('C%d' % r, row[1], 'text' + zs), c_str('D%d' % r, row[3], 'text' + zs),
                 c_str('E%d' % r, row[5], 'text' + zs)]
        rows.append('<row r="%d">%s</row>' % (r, ''.join(cells)))
        r += 1
    sheet.print_titles = True
    write_sheet(sheet, rows, widths, freeze=(0, 1), autofilter='A1:E%d' % max(r - 1, 2))
    return sheet


# ------------------------------------------------------------------------------
# パッケージ (zip)
# ------------------------------------------------------------------------------
def main():
    tmp = OUT_PATH + '.tmp'
    zf = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    WB['zip'] = zf
    summary = read_csv('pool_summary.csv')
    hourly = read_csv('pool_hourly.csv')
    events = read_csv('pool_events.csv')
    queued = read_csv('pool_queued_threads.csv')
    qstack = read_csv('pool_queued_stacks.csv')
    tdump = read_csv('pool_threaddump_threads.csv')
    config = read_csv('pool_config.csv')

    ts_sheet, ts_ranges, ts_letters, truncated = build_timeseries()
    h_ranges, h_letters = build_hourly_info(hourly)

    # 書き込み順 = シート順
    build_summary(summary, truncated)
    has_ts = len(ts_ranges['order']) > 0
    if CHARTS and has_ts:
        build_charts(ui('sheet_timeseries'), ts_ranges, ts_letters,
                     ui('sheet_hourly') if hourly and len(hourly) > 1 else None, h_ranges, h_letters)
    data_sheet('sheet_hourly', 'hourly', hourly, front=('bucket_start', 'bucket_end', 'datasource'), freeze_cols=3)
    data_sheet('sheet_events', 'events', events, front=('event_no', 'datasource'), skip=('severity_label',), freeze_cols=2)
    gen, t, last_row = ts_sheet.pending
    write_sheet(ts_sheet, gen(), t.widths, freeze=(2, 1), autofilter='A1:%s%d' % (t.last_col, last_row))
    if queued and len(queued) > 1:
        data_sheet('sheet_queued', 'queued', queued, front=('timestamp', 'datasource'), back=('epoch', 'sample_no', 'host'), freeze_cols=2)
    if qstack and len(qstack) > 1:
        data_sheet('sheet_qstack', 'qstack', qstack, fixed={'stack': 110}, freeze_cols=2)
    if tdump and len(tdump) > 1:
        data_sheet('sheet_threaddump', 'threaddump', tdump, freeze_cols=1)
    data_sheet('sheet_config', 'config', config, fixed={'value': 60}, freeze_cols=4)
    build_columns_sheet()

    # グラフ / 描画
    chart_no = 0
    for s in SHEETS:
        if not s.charts:
            continue
        d = len([x for x in SHEETS if x.charts and x.index <= s.index])
        zf.writestr('xl/drawings/drawing%d.xml' % d, drawing_xml(s.charts))
        rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for i, (_, xml) in enumerate(s.charts):
            chart_no += 1
            zf.writestr('xl/charts/chart%d.xml' % chart_no, xml)
            rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
                        'Target="../charts/chart%d.xml"/>' % (i + 1, chart_no))
        rels.append('</Relationships>')
        zf.writestr('xl/drawings/_rels/drawing%d.xml.rels' % d, ''.join(rels))
        zf.writestr('xl/worksheets/_rels/sheet%d.xml.rels' % s.index,
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                    'Target="../drawings/drawing%d.xml"/></Relationships>' % d)
        s.drawing_no = d

    # workbook
    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15600" activeTab="0"/></bookViews><sheets>']
    for s in SHEETS:
        wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (esc_attr(s.name), s.index, s.index))
    wb.append('</sheets>')
    names = []
    for s in SHEETS:
        if s.filter_ref:
            a, b = s.filter_ref.split(':')
            m1, m2 = re.match(r'([A-Z]+)(\d+)', a), re.match(r'([A-Z]+)(\d+)', b)
            names.append('<definedName name="_xlnm._FilterDatabase" localSheetId="%d" hidden="1">%s!$%s$%s:$%s$%s</definedName>'
                         % (s.index - 1, esc(sheet_ref(s.name)), m1.group(1), m1.group(2), m2.group(1), m2.group(2)))
        if s.print_titles:
            ref = esc(sheet_ref(s.name))
            cols = '%s!$A:$%s,' % (ref, col_letter(s.title_cols - 1)) if s.title_cols else ''
            names.append('<definedName name="_xlnm.Print_Titles" localSheetId="%d">%s%s!$1:$1</definedName>'
                         % (s.index - 1, cols, ref))
    if names:
        wb.append('<definedNames>' + ''.join(names) + '</definedNames>')
    wb.append('<calcPr calcId="191029"/></workbook>')
    zf.writestr('xl/workbook.xml', ''.join(wb))

    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for s in SHEETS:
        rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                    'Target="worksheets/sheet%d.xml"/>' % (s.index, s.index))
    rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                'Target="styles.xml"/></Relationships>' % (len(SHEETS) + 1))
    zf.writestr('xl/_rels/workbook.xml.rels', ''.join(rels))
    zf.writestr('xl/styles.xml', styles_xml())

    now = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    zf.writestr('docProps/core.xml',
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                '<dc:title>%s</dc:title><dc:creator>pool_csv_to_xlsx</dc:creator>'
                '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
                '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified></cp:coreProperties>'
                % (esc(ui('report_title')), now, now))
    zf.writestr('docProps/app.xml',
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
                'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                '<Application>Microsoft Excel</Application></Properties>')
    zf.writestr('_rels/.rels',
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                '</Relationships>')
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>']
    for s in SHEETS:
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % s.index)
        if s.charts:
            ct.append('<Override PartName="/xl/drawings/drawing%d.xml" '
                      'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' % s.drawing_no)
    for i in range(1, chart_no + 1):
        ct.append('<Override PartName="/xl/charts/chart%d.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' % i)
    ct.append('</Types>')
    zf.writestr('[Content_Types].xml', ''.join(ct))
    zf.close()
    os.replace(tmp, OUT_PATH)
    print('%s: %s' % (ui('msg_created'), OUT_PATH))
    print('  %s: %s' % (ui('msg_sheets'), ' / '.join(s.name for s in SHEETS)))
    if truncated:
        print('  %s (--max-rows)' % ui('truncated'))


try:
    main()
except Exception as e:  # noqa
    try:
        os.remove(OUT_PATH + '.tmp')
    except OSError:
        pass
    sys.stderr.write('エラー: %s: %s\n' % (type(e).__name__, e))
    raise
PYEOF
