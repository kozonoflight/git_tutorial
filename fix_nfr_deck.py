#!/usr/bin/env python3
"""非機能要件PPTの矛盾を修正し、全ページのタイトルに通し番号を振る。

元ファイルはZIPのCRCが壊れていたため zip -FF で修復したものを入力とする。
デザイン・図・表はそのまま残し、テキストのみを置換する。
折り返さない設定のテキストボックスがあるため、1行あたりの文字数を増やさない。
"""

import copy
import sys

from pptx import Presentation

SRC = sys.argv[1] if len(sys.argv) > 1 else "/tmp/nfrfix.pptx"
DST = sys.argv[2] if len(sys.argv) > 2 else "/workspace/要件定義書_非機能要件_20260917_v2.pptx"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def iter_text_frames(slide):
    for shape in slide.shapes:
        if shape.has_text_frame:
            yield shape.text_frame
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    yield cell.text_frame
        if shape.shape_type == 6:
            for sub in shape.shapes:
                if sub.has_text_frame:
                    yield sub.text_frame


def _tf(obj):
    return obj.text_frame if hasattr(obj, "text_frame") else obj


def replace_in_frame(tf, old, new):
    hit = 0
    for para in tf.paragraphs:
        runs = para.runs
        if not runs:
            continue
        joined = "".join(r.text for r in runs)
        if old not in joined:
            continue
        runs[0].text = joined.replace(old, new)
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        hit += 1
    return hit


def replace_on_slide(slide, old, new, label=""):
    hit = sum(replace_in_frame(tf, old, new) for tf in iter_text_frames(slide))
    if hit == 0:
        print(f"  !! MISS {label}: {old[:40]!r}")
    return hit


def set_lines(obj, lines):
    """内容を行単位で入れ替える。2行目以降は1行目の段落書式を複製する。"""
    tf = _tf(obj)
    paras = tf.paragraphs
    base = None
    for p in paras:
        if p.runs:
            base = p
            break
    if base is None:
        return False
    base_xml = copy.deepcopy(base._p)
    body = tf._txBody
    for p in list(paras):
        body.remove(p._p)
    for line in lines:
        el = copy.deepcopy(base_xml)
        runs = el.findall(f"{A_NS}r")
        for extra in runs[1:]:
            el.remove(extra)
        for br in el.findall(f"{A_NS}br"):
            el.remove(br)
        runs[0].find(f"{A_NS}t").text = line
        body.append(el)
    return True


def append_line(obj, line):
    """末尾に1段落追加する。最後の段落の書式を複製する。"""
    tf = _tf(obj)
    base = None
    for p in tf.paragraphs:
        if p.runs:
            base = p
    if base is None:
        return False
    el = copy.deepcopy(base._p)
    runs = el.findall(f"{A_NS}r")
    for extra in runs[1:]:
        el.remove(extra)
    for br in el.findall(f"{A_NS}br"):
        el.remove(br)
    runs[0].find(f"{A_NS}t").text = line
    tf._txBody.append(el)
    return True


def find_title_shape(slide):
    for shape in slide.shapes:
        try:
            if shape.is_placeholder and shape.placeholder_format.type == 1:
                return shape
        except Exception:
            pass
    best = None
    for shape in slide.shapes:
        if not shape.has_text_frame or not shape.text_frame.text.strip():
            continue
        if shape.top is None or shape.top > 500000:
            continue
        if best is None or shape.width > best.width:
            best = shape
    return best


prs = Presentation(SRC)
slides = list(prs.slides)
if len(slides) != 15:
    sys.exit(f"想定外のページ数: {len(slides)}")

# ============================================================
# 1. 全ページのタイトルに通し番号を振る（表紙・裏表紙を除く）
#    タイトル枠は自動縮小設定のため、元の文字数を超えない範囲に収める
# ============================================================
TITLES = {
    2: "1.（参考）SLA（1/2）",
    3: "2.（参考）SLA（2/2）",
    4: "3.（参考）サービス対応",
    5: "4.全体俯瞰",
    6: "5.対応内容一覧（1/2）",
    7: "6.対応内容一覧（2/2）",
    8: "7.対応内容説明 - 監視（システム構成）",
    9: "8.対応内容説明 - バックアップ",
    10: "9.対応内容説明 - セキュリティ（1/2）",
    11: "10.対応内容説明 - セキュリティ（2/2）",
    12: "11.対応内容説明 - システム構成（1/3）",
    13: "12.対応内容説明 - システム構成（2/3）",
    14: "13.対応内容説明 - システム構成（3/3）",
}
print("== タイトル ==")
for no, text in TITLES.items():
    shape = find_title_shape(slides[no - 1])
    if shape is None:
        print(f"  !! p{no} タイトル未検出")
        continue
    before = shape.text_frame.text.replace("\n", " ")
    set_lines(shape, [text])
    print(f"  p{no}: {before}  ->  {text}")

# ============================================================
# 2. 1.（参考）SLA: 復旧時点の表現と本案件の想定値への参照
# ============================================================
print("== p2 SLA(1/2) ==")
s = slides[1]
replace_on_slide(s, "障害発生時の状態で別ハードウェアにて起動", "直近のバックアップ取得時点の状態で別ハードウェアにて起動", label="p2復旧時点")
tbl = next(sh.table for sh in s.shapes if sh.has_table)
for row in tbl.rows:
    if row.cells[1].text.strip() == "同時接続利用者数":
        append_line(row.cells[3], "本案件の想定は11.に記載")
        print("  同時接続利用者数の備考に本案件参照を追加")
    if row.cells[1].text.strip() == "重大障害時の代替手段":
        append_line(row.cells[3], "日次バックアップのため最大24時間分の戻りが発生")
        print("  重大障害時の備考に戻り時間を明記")

# ============================================================
# 3. 2.（参考）SLA: 表現の重複とバックアップ保存期間
# ============================================================
print("== p3 SLA(2/2) ==")
s = slides[2]
replace_on_slide(s, "保守はカスタマイズ費用の年額10%年額。", "保守はカスタマイズ費用の年額10%とする。", label="p3年額重複")
replace_on_slide(s, "実施する。なお。障害発生時は", "実施する。なお、障害発生時は", label="p3句点")
replace_on_slide(
    s,
    "バックアップした日から1週間保持する",
    "データ種別により7日間または30日間保持する（詳細は8.）",
    label="p3保存期間",
)

# ============================================================
# 4. 4.全体俯瞰: 誤字と参照先
# ============================================================
print("== p5 全体俯瞰 ==")
s = slides[4]
replace_on_slide(s, "赤字箇所を対象した対応内容", "赤字箇所を対象とした対応内容", label="p5誤字")
replace_on_slide(s, "1章に記載", "7.に記載", label="p5参照1")
replace_on_slide(s, "2章に記載", "8.に記載", label="p5参照2")
print(f"  3章参照 置換 {sum(replace_in_frame(tf, '3章に記載', '9.10.に記載') for tf in iter_text_frames(s))} 箇所")
replace_on_slide(s, "4章に記載", "11.12.13.に記載", label="p5参照4")

tbl = next(sh.table for sh in s.shapes if sh.has_table)
NOTE_FILL = {"信頼性": "1.に記載", "品質要件": "1.に記載", "性能要件": "11.12.に記載"}
for row in tbl.rows:
    head = row.cells[0].text.strip()
    if head in NOTE_FILL and not row.cells[3].text.strip():
        set_lines(row.cells[3], [NOTE_FILL[head]])
        print(f"  {head} の備考を補完 -> {NOTE_FILL[head]}")

# ============================================================
# 5. 5.対応内容一覧（1/2）: 誤字
# ============================================================
print("== p6 対応内容一覧(1/2) ==")
replace_on_slide(slides[5], "補完場所はクラウドサービス内", "保管場所はクラウドサービス内", label="p6保管")

# ============================================================
# 6. 6.対応内容一覧（2/2）: 動的リソースの実態を明記
# ============================================================
print("== p7 対応内容一覧(2/2) ==")
replace_on_slide(slides[6], "VulsによるCVEを対象に", "NESSUSによるCVEを対象に", label="p7ツール統一")
tbl = next(sh.table for sh in slides[6].shapes if sh.has_table)
for row in tbl.rows:
    req = row.cells[1].text.strip()
    if req in ("スケールアップの可能性", "動的なリソース拡張", "動的なリソース再編成") and not row.cells[3].text.strip():
        set_lines(row.cells[3], ["自動スケールではなく、計画停止を伴う手動の構成変更となります。"])
        print(f"  {req} の備考を補完")

# ============================================================
# 7. 9.セキュリティ（1/2）: 誤字とツール統一
# ============================================================
print("== p10 セキュリティ(1/2) ==")
s = slides[9]
replace_on_slide(s, "正なアクセスが同一IPアドレス", "不正なアクセスが同一IPアドレス", label="p10不正")
replace_on_slide(s, "XSS,SSQLインジェクション", "XSS,SQLインジェクション", label="p10SQL")
tbl = [sh.table for sh in s.shapes if sh.has_table][-1]
RELABEL = {"静的検査": ["脆弱性診断", "（開発完了時）"], "動的検査": ["脆弱性診断", "（システムテスト時）"]}
for row in tbl.rows:
    key = row.cells[0].text.strip()
    if key in RELABEL:
        set_lines(row.cells[0], RELABEL[key])
        print(f"  {key} -> {''.join(RELABEL[key])}")

# ============================================================
# 8. 11.システム構成（1/3）: 同時アクセスの定義と5年後の再計算
#    5年後は利用者2倍（12.の成長倍率2倍と対応）
#    フロント 3,400→6,800名、3%で204名
#    管理 25→50名。同時アクセスは★常時は全員、★随時は2割（最低2名）
#    合計 リリース時116名 / 5年後230名
# ============================================================
print("== p12 システム構成(1/3) ==")
s = slides[11]
replace_on_slide(
    s,
    "サーバサイジングの要素となるシステム利用者数の同時アクセス数の根拠を以下のとおり想定しました。",
    "サーバサイジングの要素となる同時アクセス数（ログインして画面を開いている人数）の根拠を以下のとおり想定しました。",
    label="p12冒頭",
)
replace_on_slide(
    s,
    "リリース5年後、アクティブユーザ数が20％増加する想定",
    "リリース5年後、アクティブユーザ数が2倍に増加する想定（12.の成長倍率2倍に対応）",
    label="p12成長前提",
)
replace_on_slide(
    s,
    "管理サイトの想定同時アクセス（ログイン）数は各担当2名と想定",
    "管理サイトの想定同時アクセス（ログイン）数は、★常時は全員、★随時は2割（最低2名）と想定",
    label="p12管理前提",
)
for shape in s.shapes:
    if shape.has_text_frame and "同時セッション数とは" in shape.text_frame.text:
        append_line(shape, "※SLA記載の同時接続利用者数100名（最善努力型）に対し、本案件は5年後230名を目標値として別途調整します。")
        print("  SLAとの差異に関する注記を追加")
        break

# key: (リリース時の同時アクセス, 5年後の利用者数, 5年後の同時アクセス)
ADMIN = {
    "営業担当者": ("2名", "20名", "4名"),
    "サイト管理者": ("2名", "10名", "2名"),
    "カスタマー": ("6名", "12名", "12名"),
    "入出庫担当者": ("4名", "8名", "8名"),
}
tbl = next(sh.table for sh in s.shapes if sh.has_table)
for ri, row in enumerate(tbl.rows):
    cells = row.cells
    head = cells[0].text.strip().replace("　", "")
    if "取引先担当者" in head or (head == "小計" and cells[3].text.strip() == "102名"):
        set_lines(cells[4], ["6,800名"])
        set_lines(cells[5], ["204名"])
        print(f"  r{ri} フロント5年後 -> 6,800名 / 204名")
    elif head == "小計" and cells[3].text.strip() == "8名":
        set_lines(cells[3], ["14名"])
        set_lines(cells[4], ["50名"])
        set_lines(cells[5], ["26名"])
        print(f"  r{ri} 管理小計 -> 14名 / 50名 / 26名")
    elif "【合計】" in head:
        set_lines(cells[3], ["116名"])
        set_lines(cells[4], ["6,850名"])
        set_lines(cells[5], ["230名"])
        print(f"  r{ri} 合計 -> 116名 / 6,850名 / 230名")
    else:
        for key, (rel_acc, y5_users, y5_acc) in ADMIN.items():
            if key in head:
                set_lines(cells[3], [rel_acc])
                set_lines(cells[4], [y5_users])
                set_lines(cells[5], [y5_acc])
                print(f"  r{ri} {key} -> {rel_acc} / {y5_users} / {y5_acc}")
                break

# ============================================================
# 9. 12.システム構成（2/3）: 同時セッションの位置づけと注記
# ============================================================
print("== p13 システム構成(2/3) ==")
s = slides[12]
replace_on_slide(
    s,
    "サーバサイジングの要素となるシステム利用者数の同時アクセス数、同時セッション数をピーク月（7月）の注文件数を根拠として想定しました。",
    "サーバサイジングの要素となる同時セッション数（同時に処理する件数）をピーク月（7月）の注文件数を根拠として想定しました。",
    label="p13冒頭",
)
replace_on_slide(s, "上記①②の結果より、", "11.および本ページの結果より、", label="p13結論")

tbl = next(sh.table for sh in s.shapes if sh.has_table)
hdr = tbl.rows[0].cells
set_lines(hdr[1], ["ピーク月", "(件/月)"])
set_lines(hdr[2], ["日平均", "(件/日)", "※30日/月換算"])
set_lines(hdr[3], ["時間平均", "(件/時)", "※8時間/日換算"])
print("  表頭の換算注記を該当列へ移動")

NOTE = [
    "※ 営業日30日/月、営業8時間/日で慣らした値。ピーク日・時間帯の偏りは実績がないため考慮しておりません。",
    "　 ピーク1時間1,200件（1分20件）でスペックアップを検討、2,400件（1分40件）が上限となります。",
    "※ 成長倍率2倍を5年後の想定とします（11.のアクティブユーザ2倍に対応）。",
    "※ 同時アクセス=分平均×滞在時間10分（注文基準の下限値。採用値は11.の116名／5年後230名）",
    "※ 同時セッション=秒平均×当該処理3秒（内訳はフロント2、管理2）",
]
for shape in s.shapes:
    if shape.has_text_frame and "営業日30日/月" in shape.text_frame.text:
        set_lines(shape, NOTE)
        print("  注記を差し替え")
        break

# ============================================================
# 10. 13.システム構成（3/3）: 初期構成の位置づけ
# ============================================================
print("== p14 システム構成(3/3) ==")
replace_on_slide(
    slides[13],
    "ご提案のシステム構成を記載します。",
    "ご提案のシステム構成を記載します。初期構成は同居1台構成、スケールアウトは有償オプションです。",
    label="p14初期構成",
)

prs.save(DST)
print("\nwrote", DST)
