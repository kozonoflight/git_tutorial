#!/usr/bin/env python3
"""同時アクセス / 同時セッション を分けたサイジング図を生成する。"""

from math import ceil

import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch, Rectangle

FONT = font_manager.FontProperties(fname="/usr/share/fonts/truetype/wqy/wqy-microhei.ttc")
plt.rcParams["font.family"] = FONT.get_name()
plt.rcParams["axes.unicode_minus"] = False

TEAL = "#2A9D8F"
TEAL_DK = "#1D7A70"
ORANGE = "#E07A3D"
RED = "#C0392B"
BLUE = "#3D5A80"
BLUE_DK = "#2B4060"
GRAY = "#5C6770"
BG = "#F7F8FA"
LINE = "#D5DBE0"
NAVY = "#1B3A4B"

PEAK_MONTH = 29232
DAYS = 30
DAY_F = 2.0
HOURS = 8
HOUR_F = 2.5
PROC_SEC = 2.0
BURST = 3.0
SAFE_SESS = 1.5
DWELL_MIN = 10.0
SAFE_ACC = 1.5
MIN_SESS = 3


def metrics(growth: float):
    month = PEAK_MONTH * growth
    avg_day = month / DAYS
    peak_day = avg_day * DAY_F
    avg_h = peak_day / HOURS
    peak_h = avg_h * HOUR_F
    per_min = peak_h / 60
    per_sec = peak_h / 3600
    sess_avg = per_sec * PROC_SEC
    sess_design = max(MIN_SESS, ceil(sess_avg * BURST * SAFE_SESS - 1e-9))
    acc_avg = per_min * DWELL_MIN
    acc_design = ceil(acc_avg * SAFE_ACC - 1e-9)
    return {
        "growth": growth,
        "month": month,
        "avg_day": avg_day,
        "peak_day": peak_day,
        "avg_h": avg_h,
        "peak_h": peak_h,
        "per_min": per_min,
        "per_sec": per_sec,
        "sess_avg": sess_avg,
        "sess_design": sess_design,
        "acc_avg": acc_avg,
        "acc_design": acc_design,
    }


def set_font(ax):
    for t in ax.get_xticklabels() + ax.get_yticklabels():
        t.set_fontproperties(FONT)


def save(fig, path):
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("wrote", path)


def text(ax, x, y, s, size=11, color=NAVY, ha="left", va="center", weight=None):
    ax.text(x, y, s, fontproperties=FONT, fontsize=size, color=color, ha=ha, va=va, weight=weight)


def title_bar(ax, label):
    ax.add_patch(Rectangle((0, 0.93), 0.012, 0.045, color=TEAL, transform=ax.transAxes, clip_on=False))
    text(ax, 0.03, 0.952, label, size=15, color=NAVY, va="center")


# --- 1. 用語の切り分け ---
fig, ax = plt.subplots(figsize=(13.2, 7.4))
ax.set_xlim(0, 13.2)
ax.set_ylim(0, 7.4)
ax.axis("off")
fig.patch.set_facecolor("white")
ax.add_patch(Rectangle((0, 7.05), 0.14, 0.28, color=TEAL))
text(ax, 0.28, 7.18, "同時アクセスと同時セッションの切り分け", size=16, color=NAVY)
text(ax, 0.28, 6.78, "弊社定義: 同時セッション = 同時処理。同時アクセス = 閲覧中の人数。混ぜない。", size=11, color=GRAY)

boxes = [
    {
        "x": 0.25,
        "title": "同時アクセス",
        "sub": "Concurrent access / 閲覧",
        "color": BLUE,
        "fill": "#EAF0F6",
        "lines": [
            "カタログ閲覧・検索・カート操作など、",
            "画面を開いている人数。",
            "",
            "同じボタンを押している必要はない。",
            "Web接続数・メモリの目安。",
            "",
            "算定: 到着率 x 滞在時間",
            "今回: 滞在10分仮定 (ログ未取得)",
        ],
    },
    {
        "x": 6.75,
        "title": "同時セッション",
        "sub": "同時処理 / 同一操作の同時実行",
        "color": TEAL_DK,
        "fill": "#E6F4F2",
        "lines": [
            "注文確定・在庫引当・採番など、",
            "同じ処理のボタンを同時に押す件数。",
            "",
            "ログイン中でも処理していなければ 0。",
            "在庫ロック・排他・採番の目安。",
            "",
            "算定: 到着率 x 当該処理の所要時間",
            "今回: 処理2秒 / 秒バースト3 / 安全率1.5",
        ],
    },
]
for b in boxes:
    ax.add_patch(
        FancyBboxPatch(
            (b["x"], 0.45),
            6.2,
            6.05,
            boxstyle="round,pad=0.02,rounding_size=0.08",
            facecolor=b["fill"],
            edgecolor=b["color"],
            linewidth=1.6,
        )
    )
    ax.add_patch(Rectangle((b["x"], 5.72), 6.2, 0.78, color=b["color"]))
    text(ax, b["x"] + 3.1, 6.22, b["title"], size=18, color="white", ha="center")
    text(ax, b["x"] + 3.1, 5.88, b["sub"], size=10, color="#F4F8F8", ha="center")
    y = 5.4
    for line in b["lines"]:
        text(ax, b["x"] + 0.28, y, line, size=11, color=NAVY if line else GRAY)
        y -= 0.48

save(fig, "/workspace/用語_同時アクセスと同時セッション.png")

# --- 2. 同時セッション積み上げ ---
cur = metrics(1.0)
fig, ax = plt.subplots(figsize=(13.4, 8.6))
ax.set_xlim(0, 13.4)
ax.set_ylim(0, 8.6)
ax.axis("off")
ax.add_patch(Rectangle((0, 8.22), 0.14, 0.28, color=TEAL))
text(ax, 0.28, 8.35, "同時セッションの算定 (同時処理 / 営業日30日)", size=15, color=NAVY)
text(ax, 0.28, 7.98, "同じ処理 (注文確定・在庫引当など) を同じ瞬間に押す件数。閲覧中の人数ではない。", size=10.5, color=GRAY)

rows = [
    ("1", "ピーク月 (実測)", f"{cur['month']:,.0f} 件/月", "2026年7月実績", TEAL),
    ("2", "÷ 営業日数 30日", f"{cur['avg_day']:,.0f} 件/日", "ピーク月の日平均", TEAL),
    ("3", "x 日集中率 2.0", f"{cur['peak_day']:,.0f} 件/日", "締切日への集中 (仮定)", ORANGE),
    ("4", "÷ 営業時間 8h", f"{cur['avg_h']:,.0f} 件/時間", "ピーク日の時間平均", TEAL),
    ("5", "x 時間集中率 2.5", f"{cur['peak_h']:,.0f} 件/時間", "時間帯の偏り (仮定)", ORANGE),
    ("6", "÷ 60分", f"{cur['per_min']:.1f} 件/分  =  {cur['per_sec']:.2f} 件/秒", "ピーク時の注文到着率 λ", TEAL),
    ("7", f"x 当該処理 {PROC_SEC:.0f}秒", f"{cur['sess_avg']:.2f} が同時処理中", "同時セッション (平均) = λ x 処理時間", TEAL),
    ("8", f"x 秒バースト {BURST:.0f} x 安全率 {SAFE_SESS}", f"同時セッション (設計)  {cur['sess_design']}", "同一ボタンの同時押しに耐える件数 (下限3)", RED),
]
y = 7.55
for num, left, mid, right, col in rows:
    ax.add_patch(FancyBboxPatch((0.22, y - 0.28), 12.96, 0.52, boxstyle="round,pad=0.01,rounding_size=0.05", facecolor="#F4FBFA" if col != RED else "#FDECEC", edgecolor=LINE, linewidth=0.8))
    ax.add_patch(plt.Circle((0.55, y - 0.02), 0.16, color=col))
    text(ax, 0.55, y - 0.02, num, size=10, color="white", ha="center")
    text(ax, 0.9, y - 0.02, left, size=11, color=NAVY)
    text(ax, 7.15, y - 0.02, mid, size=13, color=col, ha="center")
    text(ax, 12.95, y - 0.02, right, size=10, color=GRAY, ha="right")
    y -= 0.68

ax.add_patch(FancyBboxPatch((0.22, 0.18), 12.96, 1.55, boxstyle="round,pad=0.02,rounding_size=0.06", facecolor="#FFF8EE", edgecolor="#E8C9A0", linewidth=1.0))
text(ax, 0.45, 1.38, "読み方", size=11, color=RED)
text(ax, 0.45, 0.98, "平均は約6秒に1件の注文確定。2秒の処理中に重なるのは平均0.3件。秒単位の偏りを見ても同時セッションは3。", size=10.5, color=NAVY)
text(ax, 0.45, 0.55, "旧提案の「同時セッション8」は、この同時処理の余裕値として説明できる。同時アクセス (閲覧人数) には使わない。", size=10.5, color=NAVY)
save(fig, "/workspace/同時セッション算定_積み上げ.png")

# --- 3. 同時アクセス積み上げ ---
fig, ax = plt.subplots(figsize=(13.4, 8.2))
ax.set_xlim(0, 13.4)
ax.set_ylim(0, 8.2)
ax.axis("off")
ax.add_patch(Rectangle((0, 7.82), 0.14, 0.28, color=BLUE))
text(ax, 0.28, 7.95, "同時アクセスの算定 (閲覧人数 / 営業日30日)", size=15, color=NAVY)
text(ax, 0.28, 7.58, "画面を開いている人数。注文確定ボタンを押している人数ではない。滞在10分は仮定。", size=10.5, color=GRAY)

rows = [
    ("1-5", "ピーク時の到着率まで", f"{cur['peak_h']:,.0f} 件/時間", "同時セッションと同じ積み上げ", BLUE),
    ("6", "÷ 60分", f"{cur['per_min']:.1f} 件/分", "ピーク時の注文到着率 λ", BLUE),
    ("7", f"x 滞在時間 {DWELL_MIN:.0f}分", f"{cur['acc_avg']:.0f} 人が閲覧中", "同時アクセス (平均) = λ x 滞在", BLUE),
    ("8", f"x 安全率 {SAFE_ACC}", f"同時アクセス (設計)  {cur['acc_design']}", "未注文の閲覧は含まない下限見積", RED),
]
y = 7.12
for num, left, mid, right, col in rows:
    ax.add_patch(FancyBboxPatch((0.22, y - 0.38), 12.96, 0.68, boxstyle="round,pad=0.01,rounding_size=0.05", facecolor="#F3F6FA" if col != RED else "#FDECEC", edgecolor=LINE, linewidth=0.8))
    ax.add_patch(plt.Circle((0.62, y - 0.04), 0.2, color=col))
    text(ax, 0.62, y - 0.04, num, size=9, color="white", ha="center")
    text(ax, 1.05, y - 0.04, left, size=12, color=NAVY)
    text(ax, 7.15, y - 0.04, mid, size=14, color=col, ha="center")
    text(ax, 12.95, y - 0.04, right, size=10.5, color=GRAY, ha="right")
    y -= 0.92

ax.add_patch(FancyBboxPatch((0.22, 0.22), 12.96, 2.85, boxstyle="round,pad=0.02,rounding_size=0.06", facecolor="#F3F6FA", edgecolor="#B7C6D6", linewidth=1.0))
text(ax, 0.45, 2.72, "同時セッションとの差", size=12, color=BLUE_DK)
text(ax, 0.45, 2.22, "同じ609件/時でも、滞在10分なら同時アクセスは約100人。処理2秒なら同時セッションは1人未満。", size=11, color=NAVY)
text(ax, 0.45, 1.72, "旧提案の同時セッション8を同時アクセスに使うと、7月ピークの閲覧規模と桁が合わない。", size=11, color=NAVY)
text(ax, 0.45, 1.22, "この同時アクセスは「注文した人が滞在10分いた場合」の下限。見て帰った人はアクセスログが必要。", size=11, color=NAVY)
text(ax, 0.45, 0.68, "Webの接続・メモリは同時アクセス側。在庫ロック・採番の排他は同時セッション側。", size=11, color=NAVY)
save(fig, "/workspace/同時アクセス算定_積み上げ.png")

# --- 4. 依頼どおりの表: 月/日/時 + 同時アクセス + 同時セッション ---
growths = [1.0, 1.2, 1.5, 2.0]
labels = ["現状", "1.2倍", "1.5倍", "2倍"]
fig, ax = plt.subplots(figsize=(13.6, 8.15))
ax.set_xlim(0, 13.6)
ax.set_ylim(0, 8.15)
ax.axis("off")
ax.add_patch(Rectangle((0, 7.75), 0.14, 0.28, color=TEAL))
text(ax, 0.28, 7.88, "ピーク基準サイジング (営業日30日)", size=16, color=NAVY)
text(ax, 0.28, 7.48, "日集中率2.0 / 時間集中率2.5 / 同時アクセス=滞在10分x安全率1.5 / 同時セッション=処理2秒xバースト3x安全率1.5", size=9.5, color=GRAY)

x0, y0, tw, th = 0.22, 2.35, 13.16, 4.85
ax.add_patch(Rectangle((x0, y0), tw, th, facecolor="white", edgecolor=LINE, lw=0.8))

# 成長 / ピーク月 / ピーク日 / ピーク時 / 同時アクセス / 同時セッション
cols = [0.22, 1.85, 4.05, 6.25, 8.45, 10.75, 13.38]
header_h = 0.95
row_h = (th - header_h) / 4
ax.add_patch(Rectangle((x0, y0 + th - header_h), tw, header_h, color=BLUE, linewidth=0))
# 同時セッション列だけヘッダーをティールに
ax.add_patch(Rectangle((cols[5], y0 + th - header_h), cols[6] - cols[5], header_h, color=TEAL_DK, linewidth=0))
heads = ["成長倍率", "ピーク月\n(件/月)", "ピーク日\n(件/日)", "ピーク時\n(件/時)", "同時アクセス", "同時セッション"]
for i, h in enumerate(heads):
    cx = (cols[i] + cols[i + 1]) / 2
    text(ax, cx, y0 + th - header_h / 2, h, size=11, color="white", ha="center")

for r, (lab, g) in enumerate(zip(labels, growths)):
    m = metrics(g)
    y = y0 + th - header_h - (r + 1) * row_h
    if g == 2.0:
        ax.add_patch(Rectangle((x0, y), tw, row_h, facecolor="#FFF4D6", edgecolor="none"))
    elif r % 2 == 1:
        ax.add_patch(Rectangle((x0, y), tw, row_h, facecolor="#F4F7FB", edgecolor="none"))
    ax.add_patch(Rectangle((cols[5], y), cols[6] - cols[5], row_h, facecolor="#D8F0EC" if g != 2.0 else "#C8EBDF", edgecolor="none"))
    vals = [
        lab,
        f"{m['month']:,.0f}",
        f"{m['peak_day']:,.0f}",
        f"{m['peak_h']:,.0f}",
        f"{m['acc_design']:,}",
        str(m["sess_design"]),
    ]
    colors = [NAVY, NAVY, NAVY, NAVY, BLUE_DK, TEAL_DK]
    sizes = [13, 13, 13, 13, 14, 18]
    for i, (v, c, sz) in enumerate(zip(vals, colors, sizes)):
        cx = (cols[i] + cols[i + 1]) / 2
        text(ax, cx, y + row_h / 2, v, size=sz, color=c, ha="center")
    if g == 2.0:
        ax.add_patch(Rectangle((x0, y), tw, row_h, fill=False, edgecolor=RED, linewidth=1.8))

for c in cols[1:-1]:
    ax.plot([c, c], [y0, y0 + th], color=LINE, lw=0.7)
ax.plot([x0, x0 + tw], [y0 + th - header_h, y0 + th - header_h], color="white", lw=0.01)

ax.add_patch(
    FancyBboxPatch(
        (0.22, 0.22),
        13.16,
        1.95,
        boxstyle="round,pad=0.02,rounding_size=0.07",
        facecolor="#E6F4F2",
        edgecolor=TEAL_DK,
        linewidth=1.8,
    )
)
text(ax, 6.8, 1.72, "結論  同時セッションは 4 で足りる", size=18, color=TEAL_DK, ha="center")
text(ax, 6.8, 1.18, "現状から1.5倍までは3。2倍成長まで見ても最大4。旧提案の8は余裕値であり、採用は4。", size=12, color=NAVY, ha="center")
text(ax, 6.8, 0.68, "同時アクセスは閲覧人数 (現状153 / 2倍で305)。同時セッションは同時処理。混ぜない。ピーク時は件/時。", size=10.5, color=GRAY, ha="center")

fig.savefig("/workspace/サイジング根拠_ピーク基準.png", dpi=160, bbox_inches="tight", facecolor="white")
fig.savefig("/workspace/サイジング表_同時セッション4.png", dpi=160, bbox_inches="tight", facecolor="white")
plt.close(fig)
print("wrote /workspace/サイジング根拠_ピーク基準.png")
print("wrote /workspace/サイジング表_同時セッション4.png")

# --- 5. なぜ 153 対 3 か / いつ 4 になるか ---
fig, ax = plt.subplots(figsize=(13.4, 8.0))
ax.set_xlim(0, 13.4)
ax.set_ylim(0, 8.0)
ax.axis("off")
ax.add_patch(Rectangle((0, 7.58), 0.14, 0.28, color=TEAL))
text(ax, 0.28, 7.71, "なぜ同時アクセス153で、同時セッションは3なのか", size=15.5, color=NAVY)
text(ax, 0.28, 7.32, "同じ到着率でも、数える時間の長さが違う。セッションが4になるのは同時アクセス約300のとき。", size=10.5, color=GRAY)

# left access box
ax.add_patch(FancyBboxPatch((0.25, 4.55), 6.25, 2.55, boxstyle="round,pad=0.02,rounding_size=0.08", facecolor="#EAF0F6", edgecolor=BLUE, lw=1.5))
text(ax, 3.38, 6.78, "同時アクセス = 閲覧", size=14, color=BLUE, ha="center")
text(ax, 0.5, 6.28, "到着率 10.2件/分  x  滞在 10分", size=12, color=NAVY)
text(ax, 0.5, 5.78, "= 平均 102人が画面を開いている", size=12, color=NAVY)
text(ax, 0.5, 5.28, "x 安全率 1.5  =  設計 153", size=13, color=BLUE_DK)
text(ax, 0.5, 4.82, "「10分間、店に居る人数」", size=11, color=GRAY)

# right session box
ax.add_patch(FancyBboxPatch((6.9, 4.55), 6.25, 2.55, boxstyle="round,pad=0.02,rounding_size=0.08", facecolor="#E6F4F2", edgecolor=TEAL_DK, lw=1.5))
text(ax, 10.02, 6.78, "同時セッション = 同時処理", size=14, color=TEAL_DK, ha="center")
text(ax, 7.15, 6.28, "到着率 0.17件/秒  x  処理 2秒", size=12, color=NAVY)
text(ax, 7.15, 5.78, "= 平均 0.34件が処理中", size=12, color=NAVY)
text(ax, 7.15, 5.28, "x バースト3 x 安全率1.5  =  1.5", size=13, color=TEAL_DK)
text(ax, 7.15, 4.82, "切り上げし下限3を置くので 設計 3", size=11, color=GRAY)

# ratio bar
ax.add_patch(FancyBboxPatch((0.25, 3.35), 12.9, 0.95, boxstyle="round,pad=0.02,rounding_size=0.06", facecolor="#FFF8EE", edgecolor="#E8C9A0", lw=1.1))
text(ax, 6.7, 3.95, "滞在10分 = 600秒。処理は2秒。時間の長さが 300倍 違う。", size=13, color=NAVY, ha="center")
text(ax, 6.7, 3.55, "だから平均人数も約300倍 (102 対 0.34)。153と3は別物を数えている。", size=12, color=GRAY, ha="center")

# threshold
ax.add_patch(FancyBboxPatch((0.25, 0.28), 12.9, 2.85, boxstyle="round,pad=0.02,rounding_size=0.08", facecolor="#E6F4F2", edgecolor=TEAL_DK, lw=1.6))
text(ax, 6.7, 2.72, "同時セッションが 4 になるライン", size=15, color=TEAL_DK, ha="center")
text(ax, 6.7, 2.18, "設計式: 切り上げ(平均 x 4.5) が 4 を超える  =  平均が 0.67件 を超えたとき", size=11.5, color=NAVY, ha="center")
text(ax, 6.7, 1.62, "ピーク時 約1,201件/時   /   同時アクセス平均 約200   /   同時アクセス設計 約300", size=13, color=NAVY, ha="center")
text(ax, 6.7, 1.08, "現状153ではまだ半分。2倍 (同時アクセス305) で初めてセッション4。", size=12, color=NAVY, ha="center")
text(ax, 6.7, 0.55, "同時アクセス自体が4になることはない。4は同時処理側の話。", size=11, color=GRAY, ha="center")

save(fig, "/workspace/なぜ153対3か_セッション4のライン.png")

print("current", metrics(1.0))
print("x2", metrics(2.0))
