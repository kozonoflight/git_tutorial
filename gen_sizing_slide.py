#!/usr/bin/env python3
"""システム構成(2/3) サーバサイジングのスライドを生成する。

ピーク日・ピーク時間帯の実績が無いため、集中率は掛けずにピーク月を
月→日→時間と慣らして算定する。そのうえで「どこからオーバーするか」の
境目を別スライドで示し、偏りの実績はヒアリング事項とする。
"""

from math import ceil, exp, factorial

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

FONT = "メイリオ"

NAVY = RGBColor(0x1B, 0x3A, 0x4B)
BLUE = RGBColor(0x3D, 0x5A, 0x80)
BLUE_DK = RGBColor(0x2B, 0x40, 0x60)
BLUE_LT = RGBColor(0xEA, 0xF0, 0xF6)
TEAL_DK = RGBColor(0x1D, 0x7A, 0x70)
TEAL_LT = RGBColor(0xE0, 0xF1, 0xEE)
RED = RGBColor(0xC0, 0x39, 0x2B)
ORANGE = RGBColor(0xE0, 0x7A, 0x3D)
GRAY = RGBColor(0x5C, 0x67, 0x70)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ROW_ALT = RGBColor(0xF4, 0xF7, 0xFB)
HILITE = RGBColor(0xFF, 0xF4, 0xD6)
CREAM = RGBColor(0xFF, 0xF8, 0xEE)

PEAK_MONTH = 29232
DAYS = 30
HOURS = 8
PROC_SEC = 2.0
DWELL_MIN = 10.0
ADOPT_SESS = 4

LABELS = ["現状", "1.2倍", "1.5倍", "2倍"]
GROWTHS = [1.0, 1.2, 1.5, 2.0]


def flat(growth):
    month = PEAK_MONTH * growth
    day = month / DAYS
    hour = day / HOURS
    per_min = hour / 60
    per_sec = hour / 3600
    access = per_min * DWELL_MIN
    session = per_sec * PROC_SEC
    return month, day, hour, per_min, access, session


def burst_limit_per_sec(capacity=ADOPT_SESS, prob=0.001):
    """同時実行数が capacity を超える瞬間の割合が prob になる到着率を返す。"""

    def p_over(a):
        return 1 - sum(a**k * exp(-a) / factorial(k) for k in range(capacity + 1))

    lo, hi = 0.0, 50.0
    for _ in range(200):
        mid = (lo + hi) / 2
        if p_over(mid) < prob:
            lo = mid
        else:
            hi = mid
    return lo / PROC_SEC


AVG_LIMIT_SEC = ADOPT_SESS / PROC_SEC
BURST_LIMIT_SEC = burst_limit_per_sec()
BURST_LIMIT_HOUR = BURST_LIMIT_SEC * 3600
BURST_LIMIT_MIN = BURST_LIMIT_SEC * 60
BURST_LIMIT_ACCESS = BURST_LIMIT_MIN * DWELL_MIN


def style_cell(cell, text, size=11, color=NAVY, fill=None, align=PP_ALIGN.RIGHT):
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
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = FONT
        run.font.size = Pt(size if i == 0 else max(7.5, size - 2.5))
        run.font.color.rgb = color


def add_text(slide, x, y, w, h, text, size=11, color=NAVY, align=PP_ALIGN.LEFT):
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
    return box


def add_box(slide, x, y, w, h, fill=None, line=None, lw=1.6):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(lw)
    shp.shadow.inherit = False
    return shp


def title(slide, text):
    add_text(slide, 0.45, 0.22, 12.4, 0.5, text, size=26)
    add_box(slide, 0.45, 0.82, 12.45, 0.03, fill=BLUE)


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ============ スライド1: 慣らし基準の算定 ============
s1 = prs.slides.add_slide(prs.slide_layouts[6])
title(s1, "4.対応内容説明 - システム構成（2/3）")
add_text(
    s1,
    0.5,
    0.98,
    12.4,
    0.3,
    "②　サーバサイジングの要素となる同時アクセス数・同時セッション数を、注文実績を根拠として想定しました。",
    size=13,
)

heads = [
    "成長\n倍率",
    "ピーク月\n（件/月）",
    "日平均\n（件/日）",
    "時間平均\n（件/時）",
    "分平均\n（件/分）",
    "同時アクセス\n（閲覧人数）",
    "同時セッション\n（同時処理）",
]
widths = [1.05, 1.68, 1.68, 1.68, 1.55, 2.4, 2.41]
tbl_x, tbl_y = 0.5, 1.4
t1 = s1.shapes.add_table(5, 7, Inches(tbl_x), Inches(tbl_y), Inches(sum(widths)), Inches(2.02)).table
for i, w in enumerate(widths):
    t1.columns[i].width = Inches(w)
t1.rows[0].height = Inches(0.62)
for r in range(1, 5):
    t1.rows[r].height = Inches(0.35)
for c, h in enumerate(heads):
    style_cell(t1.cell(0, c), h, size=10.5, color=WHITE, fill=TEAL_DK if c == 6 else BLUE, align=PP_ALIGN.CENTER)

for r, (lab, g) in enumerate(zip(LABELS, GROWTHS), start=1):
    month, day, hour, per_min, access, session = flat(g)
    base = HILITE if g == 2.0 else (ROW_ALT if r % 2 == 0 else WHITE)
    sess_fill = RGBColor(0xC8, 0xEB, 0xDF) if g == 2.0 else TEAL_LT
    vals = [
        (lab, NAVY, PP_ALIGN.LEFT, base, 11),
        (f"{month:,.0f}", NAVY, PP_ALIGN.RIGHT, base, 11),
        (f"{day:,.0f}", NAVY, PP_ALIGN.RIGHT, base, 11),
        (f"{hour:,.0f}", NAVY, PP_ALIGN.RIGHT, base, 11),
        (f"{per_min:.1f}", GRAY, PP_ALIGN.RIGHT, base, 11),
        (f"{ceil(access):,}", BLUE_DK, PP_ALIGN.RIGHT, base, 12),
        (f"{session:.2f}", TEAL_DK, PP_ALIGN.RIGHT, sess_fill, 12),
    ]
    for c, (v, col, al, fl, sz) in enumerate(vals):
        style_cell(t1.cell(r, c), v, size=sz, color=col, fill=fl, align=al)
add_box(s1, tbl_x, tbl_y + 0.62 + 0.35 * 3, sum(widths), 0.35, line=RED, lw=2.0)

notes = (
    "※　営業日30日/月、営業8時間/日で慣らした値。ピーク日・ピーク時間帯の偏り（集中率）は実績が無いため考慮していません。\n"
    "※　同時アクセス＝分平均（件/分）×滞在10分。画面を開いている人数で、Web接続数・メモリの根拠。\n"
    "※　同時セッション＝秒平均（件/秒）×当該処理2秒。同一処理を同時に実行する件数で、在庫引当・注文確定の排他制御の根拠。"
)
add_text(s1, 0.5, 3.5, 12.45, 0.9, notes, size=9.5, color=GRAY)

add_text(
    s1,
    0.5,
    4.45,
    12.45,
    0.6,
    "③　上記①②の結果より、リリース当初は同時セッション数＝4をこなす能力のサーバを導入し、リリース後にレスポンス状況を監視して\n"
    "　　スペックアップしていくことをご提案します。慣らし値は1未満であり、4は時間帯の偏りを吸収するための余裕値です。",
    size=13,
)
add_text(s1, 0.5, 5.35, 12.45, 0.3, "④　以下にサイジング結果を提示します。", size=13)

s_widths = [1.62, 1.42, 1.42, 1.45, 1.45, 1.45, 1.45]
s_x, s_y = 0.5, 5.74
t2 = s1.shapes.add_table(3, 7, Inches(s_x), Inches(s_y), Inches(sum(s_widths)), Inches(1.04)).table
for i, w in enumerate(s_widths):
    t2.columns[i].width = Inches(w)
t2.rows[0].height = Inches(0.36)
t2.rows[1].height = Inches(0.32)
t2.rows[2].height = Inches(0.36)
t2.cell(0, 0).merge(t2.cell(1, 0))
t2.cell(0, 1).merge(t2.cell(0, 2))
t2.cell(0, 3).merge(t2.cell(0, 4))
t2.cell(0, 5).merge(t2.cell(0, 6))
style_cell(t2.cell(0, 0), "サーバ", size=11, color=WHITE, fill=BLUE, align=PP_ALIGN.CENTER)
style_cell(t2.cell(0, 1), "必要性能（採用値）", size=10.5, color=WHITE, fill=BLUE, align=PP_ALIGN.CENTER)
style_cell(t2.cell(0, 3), "プラン1（標準）", size=11, color=WHITE, fill=BLUE, align=PP_ALIGN.CENTER)
style_cell(t2.cell(0, 5), "プラン2（拡張）", size=11, color=WHITE, fill=BLUE, align=PP_ALIGN.CENTER)
for c, s in enumerate(["", "同時セッション", "同時アクセス", "SPEC", "セッション数", "SPEC", "セッション数"]):
    if c:
        style_cell(t2.cell(1, c), s, size=10, color=WHITE, fill=BLUE_DK, align=PP_ALIGN.CENTER)
srow = ["Webサーバ", "4", "220", "2コア/4GB", "4×1台＝4", "4コア/8GB", "4×2台＝8"]
for c, v in enumerate(srow):
    style_cell(
        t2.cell(2, c),
        v,
        size=11,
        color=NAVY,
        fill=BLUE_LT if c in (3, 4) else WHITE,
        align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER,
    )
add_box(s1, s_x + sum(s_widths[:3]), s_y, sum(s_widths[3:5]), 1.04, line=RED, lw=2.0)
add_text(
    s1,
    0.5,
    6.9,
    12.45,
    0.5,
    "※　採用値は慣らし値ではなく、次ページの境目（同時セッション4を超える水準）から設定しています。",
    size=9.5,
    color=RED,
)

# ============ スライド2: 境目 ============
s2 = prs.slides.add_slide(prs.slide_layouts[6])
title(s2, "4.対応内容説明 - システム構成（2/3 補足：見直しの境目）")
add_text(
    s2,
    0.5,
    0.98,
    12.4,
    0.3,
    "⑤　ピーク日・ピーク時間帯の偏りは実績が無いため、同時セッション4がどこまでの偏りを吸収できるかを整理しました。",
    size=13,
)

add_box(s2, 0.5, 1.4, 6.2, 2.35, fill=TEAL_LT, line=TEAL_DK, lw=1.5)
add_text(s2, 0.75, 1.58, 5.7, 0.35, "同時セッション4の限界", size=15, color=TEAL_DK)
add_text(
    s2,
    0.75,
    2.05,
    5.7,
    1.6,
    f"平均で4に達する到着率\n　4件 ÷ 処理2秒 ＝ {AVG_LIMIT_SEC:.1f}件/秒 ＝ {AVG_LIMIT_SEC * 3600:,.0f}件/時\n\n"
    f"瞬間的に4を超え始める到着率\n　約{BURST_LIMIT_MIN:.0f}件/分 ＝ 約{BURST_LIMIT_HOUR:,.0f}件/時\n"
    f"　（同時アクセス 約{BURST_LIMIT_ACCESS:.0f}人に相当）",
    size=11.5,
)

add_box(s2, 7.0, 1.4, 5.95, 2.35, fill=CREAM, line=ORANGE, lw=1.5)
add_text(s2, 7.25, 1.58, 5.45, 0.35, "つまり、この水準までは4で足りる", size=15, color=ORANGE)
add_text(
    s2,
    7.25,
    2.05,
    5.45,
    1.6,
    f"ピーク時間帯の注文が\n　約{BURST_LIMIT_HOUR:,.0f}件/時（約{BURST_LIMIT_MIN:.0f}件/分）\n"
    "を超えない限り、同時セッション4で処理できます。\n\n"
    "慣らし値（時間平均）に対する許容倍率は下表のとおりです。",
    size=11.5,
)

m_heads = ["成長倍率", "時間平均\n（件/時）", "同時セッション\n（慣らし値）", "境目\n（件/時）", "許容できる\n時間帯の偏り"]
m_widths = [1.75, 2.35, 2.75, 2.5, 3.1]
m_x, m_y = 0.5, 4.0
t3 = s2.shapes.add_table(5, 5, Inches(m_x), Inches(m_y), Inches(sum(m_widths)), Inches(2.02)).table
for i, w in enumerate(m_widths):
    t3.columns[i].width = Inches(w)
t3.rows[0].height = Inches(0.62)
for r in range(1, 5):
    t3.rows[r].height = Inches(0.35)
for c, h in enumerate(m_heads):
    style_cell(t3.cell(0, c), h, size=10.5, color=WHITE, fill=ORANGE if c == 4 else BLUE, align=PP_ALIGN.CENTER)
for r, (lab, g) in enumerate(zip(LABELS, GROWTHS), start=1):
    _, _, hour, _, _, session = flat(g)
    ratio = BURST_LIMIT_HOUR / hour
    base = HILITE if g == 2.0 else (ROW_ALT if r % 2 == 0 else WHITE)
    vals = [
        (lab, NAVY, PP_ALIGN.LEFT, base, 11),
        (f"{hour:,.0f}", NAVY, PP_ALIGN.RIGHT, base, 11),
        (f"{session:.2f}", TEAL_DK, PP_ALIGN.RIGHT, base, 11),
        (f"{BURST_LIMIT_HOUR:,.0f}", NAVY, PP_ALIGN.RIGHT, base, 11),
        (f"平均の約{ratio:.0f}倍まで", ORANGE, PP_ALIGN.RIGHT, base, 12),
    ]
    for c, (v, col, al, fl, sz) in enumerate(vals):
        style_cell(t3.cell(r, c), v, size=sz, color=col, fill=fl, align=al)
add_box(s2, m_x, m_y + 0.62 + 0.35 * 3, sum(m_widths), 0.35, line=RED, lw=2.0)

add_box(s2, 0.5, 6.25, 12.45, 0.95, fill=BLUE_LT, line=BLUE, lw=1.5)
add_text(
    s2,
    0.78,
    6.42,
    11.9,
    0.7,
    "【ご確認事項】　ピーク時間帯・繁忙日の偏りが平均の5倍を超える見込みがある場合はご提示ください。\n"
    "　　　　　　　　該当する実績（日次の注文件数、時間帯別の注文件数）をいただければ、上表の境目に対して再評価します。",
    size=12,
    color=BLUE_DK,
)

def render_preview(presentation, paths):
    """LibreOfficeが無い環境用に、生成したpptxをmatplotlibでPNG化する。"""
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    from matplotlib.patches import Rectangle

    jp = font_manager.FontProperties(fname="/usr/share/fonts/truetype/wqy/wqy-microhei.ttc")
    plt.rcParams["font.family"] = jp.get_name()
    sw = presentation.slide_width / 914400
    sh = presentation.slide_height / 914400

    def rgb(color):
        return "#%02X%02X%02X" % (color[0], color[1], color[2])

    def fill_of(obj):
        try:
            if obj.fill.type is not None and obj.fill.type == 1:
                return rgb(obj.fill.fore_color.rgb)
        except Exception:
            pass
        return None

    def line_of(shape):
        try:
            return rgb(shape.line.color.rgb), max(0.4, (shape.line.width or 0) / 12700)
        except Exception:
            return None, 0

    ALIGN = {PP_ALIGN.CENTER: "center", PP_ALIGN.RIGHT: "right"}

    def draw_tf(ax, tf, x, y, w, h, vcenter=False):
        paras = [p for p in tf.paragraphs]
        heights = []
        for p in paras:
            sz = max((r.font.size.pt for r in p.runs if r.font.size), default=12)
            heights.append(sz / 72 * 1.5)
        total = sum(heights)
        cy = y + (h - total) / 2 if vcenter else y + 0.02
        for p, lh in zip(paras, heights):
            txt = "".join(r.text for r in p.runs)
            if txt:
                sz = max((r.font.size.pt for r in p.runs if r.font.size), default=12)
                col = next((rgb(r.font.color.rgb) for r in p.runs if r.font.color and r.font.color.type is not None), "#000000")
                al = ALIGN.get(p.alignment, "left")
                px = x + 0.05 if al == "left" else (x + w / 2 if al == "center" else x + w - 0.06)
                ax.text(px, cy + lh / 2, txt, fontproperties=jp, fontsize=sz, color=col, ha=al, va="center")
            cy += lh

    for slide, path in zip(presentation.slides, paths):
        fig = plt.figure(figsize=(sw, sh))
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_xlim(0, sw)
        ax.set_ylim(sh, 0)
        ax.axis("off")
        fig.patch.set_facecolor("white")
        for shape in slide.shapes:
            x = shape.left / 914400
            y = shape.top / 914400
            w = shape.width / 914400
            h = shape.height / 914400
            if shape.has_table:
                tbl = shape.table
                colx = [x]
                for c in tbl.columns:
                    colx.append(colx[-1] + c.width / 914400)
                rowy = [y]
                for r in tbl.rows:
                    rowy.append(rowy[-1] + r.height / 914400)
                spans = {}
                for ri in range(len(tbl.rows)):
                    for ci in range(len(tbl.columns)):
                        cell = tbl.cell(ri, ci)
                        if cell.is_spanned:
                            continue
                        cw = colx[ci + cell.span_width] - colx[ci]
                        ch = rowy[ri + cell.span_height] - rowy[ri]
                        spans[(ri, ci)] = (colx[ci], rowy[ri], cw, ch, cell)
                for (ri, ci), (cx, cy, cw, ch, cell) in spans.items():
                    f = fill_of(cell)
                    ax.add_patch(Rectangle((cx, cy), cw, ch, facecolor=f or "white", edgecolor="#D5DBE0", lw=0.5))
                    draw_tf(ax, cell.text_frame, cx, cy, cw, ch, vcenter=True)
                continue
            f = fill_of(shape)
            lc, lw = line_of(shape)
            if f or lc:
                ax.add_patch(Rectangle((x, y), w, h, facecolor=f or "none", fill=f is not None, edgecolor=lc or "none", lw=lw))
            if shape.has_text_frame and shape.text_frame.text.strip():
                draw_tf(ax, shape.text_frame, x, y, w, h, vcenter=bool(f or lc))
        fig.savefig(path, dpi=110, facecolor="white")
        plt.close(fig)
        print("wrote", path)


out = "/workspace/システム構成_サーバサイジング_20260915.pptx"
prs.save(out)
render_preview(
    prs,
    [
        "/workspace/システム構成_サーバサイジング_20260915_p1.png",
        "/workspace/システム構成_サーバサイジング_20260915_p2.png",
    ],
)
print("wrote", out)
print("burst limit/h", round(BURST_LIMIT_HOUR, 1), "min", round(BURST_LIMIT_MIN, 1), "access", round(BURST_LIMIT_ACCESS, 1))
for lab, g in zip(LABELS, GROWTHS):
    month, day, hour, per_min, access, session = flat(g)
    print(f"{lab}: {month:,.0f} {day:,.1f} {hour:,.1f} {per_min:.2f} acc={access:.1f} sess={session:.3f} ratio={BURST_LIMIT_HOUR / hour:.1f}")
