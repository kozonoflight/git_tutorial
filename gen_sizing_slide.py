#!/usr/bin/env python3
"""システム構成(2/3) サーバサイジングのスライドを再作成する。

元スライドは注記が「30日/月換算・8時間/日換算」だけだったため、
29,232 / 30 = 974 にしかならず、ピーク日1,949・ピーク時609に辿り着けない。
日集中率・時間集中率を列と注記に明示し、同時アクセスと同時セッションを分けて示す。
"""

from math import ceil

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

FONT = "メイリオ"

NAVY = RGBColor(0x1B, 0x3A, 0x4B)
BLUE = RGBColor(0x3D, 0x5A, 0x80)
BLUE_DK = RGBColor(0x2B, 0x40, 0x60)
BLUE_LT = RGBColor(0xEA, 0xF0, 0xF6)
TEAL_DK = RGBColor(0x1D, 0x7A, 0x70)
TEAL_LT = RGBColor(0xE0, 0xF1, 0xEE)
RED = RGBColor(0xC0, 0x39, 0x2B)
GRAY = RGBColor(0x5C, 0x67, 0x70)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ROW_ALT = RGBColor(0xF4, 0xF7, 0xFB)
HILITE = RGBColor(0xFF, 0xF4, 0xD6)

PEAK_MONTH = 29232
DAYS = 30
DAY_F = 2.0
HOURS = 8
HOUR_F = 2.5
PROC_SEC = 2.0
BURST = 3.0
SAFE = 1.5
DWELL_MIN = 10.0
MIN_SESS = 3


def metrics(growth):
    month = PEAK_MONTH * growth
    avg_day = month / DAYS
    peak_day = avg_day * DAY_F
    peak_h = peak_day / HOURS * HOUR_F
    per_min = peak_h / 60
    per_sec = peak_h / 3600
    access = ceil(per_min * DWELL_MIN * SAFE - 1e-9)
    session = max(MIN_SESS, ceil(per_sec * PROC_SEC * BURST * SAFE - 1e-9))
    return month, avg_day, peak_day, peak_h, access, session


def style_cell(cell, text, size=11, color=NAVY, fill=None, align=PP_ALIGN.RIGHT, bold=False):
    cell.margin_left = Inches(0.05)
    cell.margin_right = Inches(0.05)
    cell.margin_top = Inches(0.02)
    cell.margin_bottom = Inches(0.02)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill is None:
        cell.fill.background()
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    tf = cell.text_frame
    tf.word_wrap = True
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = FONT
        run.font.size = Pt(size if i == 0 else max(7.5, size - 2.5))
        run.font.color.rgb = color
        run.font.bold = bold


def add_text(slide, x, y, w, h, text, size=11, color=NAVY, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_top = 0
    tf.margin_right = 0
    tf.margin_bottom = 0
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = FONT
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.bold = bold
    return box


def add_highlight(slide, x, y, w, h):
    from pptx.enum.shapes import MSO_SHAPE

    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.background()
    shp.line.color.rgb = RED
    shp.line.width = Pt(2.0)
    shp.shadow.inherit = False
    return shp


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

add_text(slide, 0.45, 0.22, 12.4, 0.5, "4.対応内容説明 - システム構成（2/3）", size=26, color=NAVY)
from pptx.enum.shapes import MSO_SHAPE

bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.45), Inches(0.82), Inches(12.45), Inches(0.03))
bar.fill.solid()
bar.fill.fore_color.rgb = BLUE
bar.line.fill.background()
bar.shadow.inherit = False

add_text(
    slide,
    0.5,
    0.98,
    12.4,
    0.3,
    "②　サーバサイジングの要素となる同時アクセス数・同時セッション数を、注文実績を根拠として想定しました。",
    size=13,
    color=NAVY,
)

# ---- 主表 ----
labels = ["現状", "1.2倍", "1.5倍", "2倍"]
growths = [1.0, 1.2, 1.5, 2.0]
heads = [
    "成長\n倍率",
    "ピーク月\n（件/月）",
    "日平均\n（件/日）",
    "ピーク日\n（件/日）",
    "ピーク時\n（件/時）",
    "同時アクセス\n（閲覧人数）",
    "同時セッション\n（同時処理）",
]
widths = [1.05, 1.62, 1.62, 1.72, 1.72, 2.36, 2.36]

tbl_x, tbl_y = 0.5, 1.38
rows, cols = 5, 7
gt = slide.shapes.add_table(rows, cols, Inches(tbl_x), Inches(tbl_y), Inches(sum(widths)), Inches(2.02)).table
for i, w in enumerate(widths):
    gt.columns[i].width = Inches(w)
gt.rows[0].height = Inches(0.62)
for r in range(1, rows):
    gt.rows[r].height = Inches(0.35)

for c, h in enumerate(heads):
    fill = TEAL_DK if c == 6 else BLUE
    style_cell(gt.cell(0, c), h, size=10.5, color=WHITE, fill=fill, align=PP_ALIGN.CENTER)

for r, (lab, g) in enumerate(zip(labels, growths), start=1):
    month, avg_day, peak_day, peak_h, access, session = metrics(g)
    base = HILITE if g == 2.0 else (ROW_ALT if r % 2 == 0 else WHITE)
    vals = [
        (lab, NAVY, PP_ALIGN.LEFT, base, 11),
        (f"{month:,.0f}", NAVY, PP_ALIGN.RIGHT, base, 11),
        (f"{avg_day:,.0f}", GRAY, PP_ALIGN.RIGHT, base, 11),
        (f"{peak_day:,.0f}", NAVY, PP_ALIGN.RIGHT, base, 11),
        (f"{peak_h:,.0f}", NAVY, PP_ALIGN.RIGHT, base, 11),
        (f"{access:,}", BLUE_DK, PP_ALIGN.RIGHT, base, 12),
        (f"{session}", TEAL_DK, PP_ALIGN.RIGHT, TEAL_LT if g != 2.0 else RGBColor(0xC8, 0xEB, 0xDF), 13),
    ]
    for c, (v, col, al, fl, sz) in enumerate(vals):
        style_cell(gt.cell(r, c), v, size=sz, color=col, fill=fl, align=al)

add_highlight(slide, tbl_x, tbl_y + 0.62 + 0.35 * 3, sum(widths), 0.35)

notes = (
    "※　営業日30日/月、営業8時間/日。ピーク日＝日平均×日集中率2.0（締切日への集中）、ピーク時＝ピーク日÷8時間×時間集中率2.5（ピーク時間帯）。集中率は仮定値。\n"
    "※　同時アクセス＝ピーク時到着率（件/分）×滞在10分×安全率1.5。画面を開いている人数で、Web接続数・メモリの根拠。\n"
    "※　同時セッション＝ピーク時到着率（件/秒）×当該処理2秒×秒バースト3.0×安全率1.5（下限3）。同一処理を同時に実行する件数で、在庫引当・注文確定の排他の根拠。"
)
add_text(slide, 0.5, 3.48, 12.45, 0.9, notes, size=9.5, color=GRAY)

add_text(
    slide,
    0.5,
    4.42,
    12.45,
    0.55,
    "③　上記①②の結果より、リリース当初は同時セッション数＝4をこなす能力のサーバを導入し、リリース後にレスポンス状況を監視して\n　　スペックアップしていくことをご提案します。2倍成長（年58,464件）まで同時セッションは4で収まります。",
    size=13,
    color=NAVY,
)
add_text(slide, 0.5, 5.32, 12.45, 0.3, "④　以下にサイジング結果を提示します。", size=13, color=NAVY)

# ---- サイジング表 ----
s_widths = [1.55, 1.42, 1.42, 1.45, 1.45, 1.45, 1.45]
s_x, s_y = 0.5, 5.72
st = slide.shapes.add_table(3, 7, Inches(s_x), Inches(s_y), Inches(sum(s_widths)), Inches(1.28)).table
for i, w in enumerate(s_widths):
    st.columns[i].width = Inches(w)
st.rows[0].height = Inches(0.36)
st.rows[1].height = Inches(0.32)
st.rows[2].height = Inches(0.36)

st.cell(0, 0).merge(st.cell(1, 0))
st.cell(0, 1).merge(st.cell(0, 2))
st.cell(0, 3).merge(st.cell(0, 4))
st.cell(0, 5).merge(st.cell(0, 6))

style_cell(st.cell(0, 0), "サーバ", size=11, color=WHITE, fill=BLUE, align=PP_ALIGN.CENTER)
style_cell(st.cell(0, 1), "必要性能（ピーク時・2倍成長）", size=10.5, color=WHITE, fill=BLUE, align=PP_ALIGN.CENTER)
style_cell(st.cell(0, 3), "プラン1（標準）", size=11, color=WHITE, fill=BLUE, align=PP_ALIGN.CENTER)
style_cell(st.cell(0, 5), "プラン2（拡張）", size=11, color=WHITE, fill=BLUE, align=PP_ALIGN.CENTER)
style_cell(st.cell(1, 1), "同時セッション", size=10, color=WHITE, fill=BLUE_DK, align=PP_ALIGN.CENTER)
style_cell(st.cell(1, 2), "同時アクセス", size=10, color=WHITE, fill=BLUE_DK, align=PP_ALIGN.CENTER)
style_cell(st.cell(1, 3), "SPEC", size=10, color=WHITE, fill=BLUE_DK, align=PP_ALIGN.CENTER)
style_cell(st.cell(1, 4), "セッション数", size=10, color=WHITE, fill=BLUE_DK, align=PP_ALIGN.CENTER)
style_cell(st.cell(1, 5), "SPEC", size=10, color=WHITE, fill=BLUE_DK, align=PP_ALIGN.CENTER)
style_cell(st.cell(1, 6), "セッション数", size=10, color=WHITE, fill=BLUE_DK, align=PP_ALIGN.CENTER)

srow = ["Webサーバ", "4", "305", "2コア/4GB", "4×1台＝4", "4コア/8GB", "4×2台＝8"]
for c, v in enumerate(srow):
    al = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
    style_cell(st.cell(2, c), v, size=11, color=NAVY, fill=BLUE_LT if c in (3, 4) else WHITE, align=al)

add_highlight(slide, s_x + sum(s_widths[:3]), s_y, sum(s_widths[3:5]), 1.04)

add_text(
    slide,
    0.5,
    7.06,
    12.45,
    0.3,
    "※　プラン1で同時セッション4を充足。同時アクセス305は接続保持・メモリ側の要件であり、プラン1で保持可能なことを負荷試験で確認します。",
    size=9.5,
    color=RED,
)

out = "/workspace/システム構成_サーバサイジング_20260915.pptx"
prs.save(out)
print("wrote", out)
for g, lab in zip(growths, labels):
    print(lab, [f"{v:,.1f}" if isinstance(v, float) else v for v in metrics(g)])


def render_preview(path):
    """PPTXと同じレイアウトのPNGプレビュー (LibreOfficeが無い環境用)。"""
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    from matplotlib.patches import Rectangle

    jp = font_manager.FontProperties(fname="/usr/share/fonts/truetype/wqy/wqy-microhei.ttc")
    plt.rcParams["font.family"] = jp.get_name()

    W, H = 13.333, 7.5

    def hexc(c):
        return "#%02X%02X%02X" % (c[0], c[1], c[2])

    fig, ax = plt.subplots(figsize=(W, H))
    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.invert_yaxis()
    ax.axis("off")
    fig.patch.set_facecolor("white")

    def t(x, y, s, size=11, color=NAVY, ha="left", va="center"):
        ax.text(x, y, s, fontproperties=jp, fontsize=size, color=hexc(color), ha=ha, va=va)

    t(0.45, 0.47, "4.対応内容説明 - システム構成（2/3）", size=24)
    ax.add_patch(Rectangle((0.45, 0.82), 12.45, 0.03, color=hexc(BLUE)))
    t(0.5, 1.13, "②　サーバサイジングの要素となる同時アクセス数・同時セッション数を、注文実績を根拠として想定しました。", size=12)

    hx = tbl_x
    xs = [hx]
    for w in widths:
        xs.append(xs[-1] + w)
    hy, hh, rh = tbl_y, 0.62, 0.35
    for c, h in enumerate(heads):
        fill = TEAL_DK if c == 6 else BLUE
        ax.add_patch(Rectangle((xs[c], hy), widths[c], hh, color=hexc(fill)))
        lines = h.split("\n")
        for i, ln in enumerate(lines):
            t((xs[c] + xs[c + 1]) / 2, hy + 0.22 + i * 0.22, ln, size=9.5, color=WHITE, ha="center")

    for r, (lab, g) in enumerate(zip(labels, growths)):
        month, avg_day, peak_day, peak_h, access, session = metrics(g)
        y = hy + hh + r * rh
        base = HILITE if g == 2.0 else (ROW_ALT if r % 2 == 1 else RGBColor(0xFF, 0xFF, 0xFF))
        ax.add_patch(Rectangle((xs[0], y), sum(widths), rh, color=hexc(base)))
        sess_fill = RGBColor(0xC8, 0xEB, 0xDF) if g == 2.0 else TEAL_LT
        ax.add_patch(Rectangle((xs[6], y), widths[6], rh, color=hexc(sess_fill)))
        cells = [
            (lab, NAVY, "left", 11),
            (f"{month:,.0f}", NAVY, "right", 11),
            (f"{avg_day:,.0f}", GRAY, "right", 11),
            (f"{peak_day:,.0f}", NAVY, "right", 11),
            (f"{peak_h:,.0f}", NAVY, "right", 11),
            (f"{access:,}", BLUE_DK, "right", 12),
            (f"{session}", TEAL_DK, "right", 13),
        ]
        for c, (v, col, al, sz) in enumerate(cells):
            px = xs[c] + 0.1 if al == "left" else xs[c + 1] - 0.12
            t(px, y + rh / 2, v, size=sz, color=col, ha=al)
        if g == 2.0:
            ax.add_patch(Rectangle((xs[0], y), sum(widths), rh, fill=False, edgecolor=hexc(RED), lw=1.8))
    for c in xs[1:-1]:
        ax.plot([c, c], [hy, hy + hh + 4 * rh], color="#D5DBE0", lw=0.6)

    ny = 3.6
    for ln in notes.split("\n"):
        t(0.5, ny, ln, size=8.2, color=GRAY)
        ny += 0.28

    t(0.5, 4.62, "③　上記①②の結果より、リリース当初は同時セッション数＝4をこなす能力のサーバを導入し、リリース後にレスポンス状況を監視して", size=12)
    t(0.5, 4.92, "　　スペックアップしていくことをご提案します。2倍成長（年58,464件）まで同時セッションは4で収まります。", size=12)
    t(0.5, 5.47, "④　以下にサイジング結果を提示します。", size=12)

    sxs = [s_x]
    for w in s_widths:
        sxs.append(sxs[-1] + w)
    r0, r1, r2 = s_y, s_y + 0.36, s_y + 0.68
    ax.add_patch(Rectangle((sxs[0], r0), sum(s_widths), 0.68, color=hexc(BLUE)))
    ax.add_patch(Rectangle((sxs[1], r1), sum(s_widths[1:]), 0.32, color=hexc(BLUE_DK)))
    t((sxs[0] + sxs[1]) / 2, r0 + 0.34, "サーバ", size=10.5, color=WHITE, ha="center")
    t((sxs[1] + sxs[3]) / 2, r0 + 0.18, "必要性能（ピーク時・2倍成長）", size=10, color=WHITE, ha="center")
    t((sxs[3] + sxs[5]) / 2, r0 + 0.18, "プラン1（標準）", size=10.5, color=WHITE, ha="center")
    t((sxs[5] + sxs[7]) / 2, r0 + 0.18, "プラン2（拡張）", size=10.5, color=WHITE, ha="center")
    sub = ["", "同時セッション", "同時アクセス", "SPEC", "セッション数", "SPEC", "セッション数"]
    for c in range(1, 7):
        t((sxs[c] + sxs[c + 1]) / 2, r1 + 0.16, sub[c], size=9.5, color=WHITE, ha="center")
    ax.add_patch(Rectangle((sxs[0], r2), sum(s_widths), 0.36, color="white"))
    ax.add_patch(Rectangle((sxs[3], r2), s_widths[3] + s_widths[4], 0.36, color=hexc(BLUE_LT)))
    for c, v in enumerate(srow):
        if c == 0:
            t(sxs[0] + 0.1, r2 + 0.18, v, size=10.5)
        else:
            t((sxs[c] + sxs[c + 1]) / 2, r2 + 0.18, v, size=10.5, ha="center")
    for c in sxs[1:-1]:
        ax.plot([c, c], [r0, r2 + 0.36], color="#D5DBE0", lw=0.6)
    ax.add_patch(Rectangle((sxs[3], r0), s_widths[3] + s_widths[4], 1.04, fill=False, edgecolor=hexc(RED), lw=1.8))

    t(0.5, 7.2, "※　プラン1で同時セッション4を充足。同時アクセス305は接続保持・メモリ側の要件であり、プラン1で保持可能なことを負荷試験で確認します。", size=8.6, color=RED)

    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("wrote", path)


render_preview("/workspace/システム構成_サーバサイジング_20260915.png")
