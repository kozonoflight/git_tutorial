#!/usr/bin/env python3
"""シーズン管理のユーザー設計レビュー用シートを1枚で作る。

入力：
- シーズン管理の検討スライド（①商品コード ②商品データ運用 ③在庫運用）
- RJ_注文-在庫シミュレーション 展示会注文(シーズン対応)
- バッチ機能設計書 B05-EC-1-01 展示会受注締め / B05-EC-1-02 シーズン切替
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/ローリングスジャパン様_シーズン管理_運用とバッチ.xlsx"

NAVY = "1B3A4B"
TEAL = "1D7A70"
BLUE = "3D5A80"
ORANGE = "C46B2C"
BROWN = "6B4C2A"
WHITE = "FFFFFF"
INK = "1B3A4B"
GRAY = "5C6770"
PALE = "F4F7FB"
BLUE_LT = "E7EEF6"
TEAL_LT = "E5F3F1"
ORANGE_LT = "FDEFE4"
YELLOW = "FFF4D6"
CREAM = "FFF8EE"
RED_LT = "FBE6E3"
LINE = "D5DBE0"

LAST_COL = 7
WIDTHS = [7, 20, 22, 30, 42, 42, 38]

thin = Border(*(Side(style="thin", color=LINE),) * 4)
center = Alignment(wrap_text=True, vertical="center", horizontal="center")
left = Alignment(wrap_text=True, vertical="center", horizontal="left")
top_left = Alignment(wrap_text=True, vertical="top", horizontal="left")


def font(size=10, bold=False, color=INK):
    return Font(name="メイリオ", size=size, bold=bold, color=color)


def fill(color):
    return PatternFill("solid", fgColor=color)


wb = Workbook()
ws = wb.active
ws.title = "シーズン管理"
for i, w in enumerate(WIDTHS, 1):
    ws.column_dimensions[get_column_letter(i)].width = w


def cell(r, c, value, size=10, bold=False, color=INK, bg=None, align=left):
    x = ws.cell(r, c, value)
    x.font = font(size, bold, color)
    x.alignment = align
    x.border = thin
    if bg:
        x.fill = fill(bg)
    return x


def band(r, text, bg, size=12, height=24, color=WHITE):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=LAST_COL)
    x = ws.cell(r, 1, text)
    x.font = font(size, True, color)
    x.fill = fill(bg)
    x.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[r].height = height
    return r + 1


def note(r, text, bg=CREAM, color="7A4E00", height=34, size=9):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=LAST_COL)
    x = ws.cell(r, 1, text)
    x.font = font(size, False, color)
    x.fill = fill(bg)
    x.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
    ws.row_dimensions[r].height = height
    return r + 1


def table(r, spans, headers, rows, header_bg, row_bgs=None, height=48):
    """spans: 各列が結合する元の列数。合計は LAST_COL。"""
    def put(row_idx, values, bg_list, bold_first=False, hdr=False):
        col = 1
        for i, (span, v) in enumerate(zip(spans, values)):
            if span > 1:
                ws.merge_cells(start_row=row_idx, start_column=col, end_row=row_idx, end_column=col + span - 1)
            bg = bg_list[i] if isinstance(bg_list, list) else bg_list
            if hdr:
                cell(row_idx, col, v, 9, True, WHITE, bg, center)
            else:
                cell(row_idx, col, v, 9, bold_first and i == 0, INK, bg, center if (i == 0 and len(str(v)) <= 4) else left)
            for k in range(col + 1, col + span):
                ws.cell(row_idx, k).border = thin
                if bg:
                    ws.cell(row_idx, k).fill = fill(bg)
            col += span

    put(r, headers, header_bg, hdr=True)
    ws.row_dimensions[r].height = 28
    for i, values in enumerate(rows):
        rr = r + 1 + i
        bgs = row_bgs(i, values) if callable(row_bgs) else (row_bgs or [PALE if i % 2 else WHITE] * len(spans))
        put(rr, values, bgs, bold_first=True)
        ws.row_dimensions[rr].height = height if not isinstance(height, list) else height[i]
    return r + 1 + len(rows)


r = 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=LAST_COL)
x = ws.cell(r, 1, "シーズン管理　運用とバッチ（ユーザー設計レビュー）")
x.font = font(16, True, WHITE)
x.fill = fill(NAVY)
x.alignment = Alignment(vertical="center", indent=1)
ws.row_dimensions[r].height = 30
r += 1
r = note(
    r,
    "このシートでレビューいただきたいこと：①ECでシーズンをどう持つか、②年間のどの時点でバッチが動き、在庫と注文がどう変わるか、③ローリングスジャパン様に毎シーズン行っていただく作業。"
    "例は27FW（展示会1/1-2/15、シーズン開始7/1）。黄色は貴社にご判断いただきたい点です。",
    height=40,
)
r += 1

# ------------------------------------------------------------------ 1
r = band(r, "1. 管理の考え方　基幹にシーズンがないため、ECの商品コードにシーズンを持たせる", NAVY)
r = table(
    r,
    [1, 2, 2, 2],
    ["No", "決めたこと", "例", "これで何ができるか"],
    [
        ("1", "ECの商品はシーズンごとに別の商品として登録する。品番とJDE SKUは基幹と同じ、ECR商品コードだけ「品番_シーズン」にする。", "基幹：GH6HPA15W / J001011116 / 1,500円\nEC：GH6HPA15W_27SS / 1,000円\nEC：GH6HPA15W_27FW / 1,500円", "シーズンを指定した商品情報ダウンロードができる。定価をシーズンで変えられる。"),
        ("2", "シーズンを持たない継続品は、従来どおり1商品のまま。展示会の期間だけ展示会サプライヤ99999で扱う。", "シーズン無の商品（例：ボール、メンテナンス用品）", "継続品はシーズン切替の対象外。展示会受注締めだけが動く。"),
        ("3", "在庫は常に一つのシーズンに集約する。シーズン開始日に、前シーズンの在庫・入荷予定・引当・入出庫履歴を新シーズン商品へ移す。", "7/1に27SSの実在庫50、入荷予定100を27FWへ移し、27SSは0", "基幹は品番で在庫を1本で持つので、ECも1本に戻す。"),
        ("4", "フロントの商品情報ダウンロードは、入口でシーズンが決まる。", "商品詳細＝現シーズン／購入履歴から開く商品詳細＝注文時のシーズン／一括ダウンロード＝シーズン指定", "過去に注文したシーズンの情報を取り直せる。"),
    ],
    [NAVY, NAVY, NAVY, NAVY],
    height=[52, 44, 44, 44],
)
r += 1

# ------------------------------------------------------------------ 2
r = band(r, "2. 年間の流れ（27FWの例）　フロントの見え方・バッチ・貴社作業を時期ごとに", TEAL)
r = table(
    r,
    [1, 1, 1, 1, 1, 1, 1],
    ["No", "時期（例）", "通常注文の商品", "展示会注文の商品", "システム・バッチで起きること", "ローリングスジャパン様の作業", "確認ポイント"],
    [
        ("1", "展示会の前\n（12月中）", "27SS", "—", "なし", "27SSをコピーして27FWを作る（コード、シーズン、価格、画像、カテゴリ、販路、販売開始日）。\n27FWに展示会サプライヤ99999の取扱いと、未来在庫（入荷予定日2099/12/31）を設定する。\n展示会期間を設定する。", "27FWの販売開始日を、展示会開始日以前にする。シーズン切替は販売開始済みの商品だけを新シーズンとみなす。"),
        ("2", "展示会期間中\n1/1-2/15", "27SS（81999の実在庫）", "27FW（99999の架空在庫）", "展示会注文は30分ごとの受注連携で基幹へ。基幹へは81999として渡す。", "なし。展示会注文の変更は問い合わせで受ける。", "展示会の架空在庫は上限の大きい数で持ち、実在庫を減らさない。"),
        ("3", "展示会締め\n2/16 5:00", "27SS", "27FW（81999へ付け替え）", "【B05-EC-1-01 展示会受注締め】\n展示会注文の引当を99999から81999へ移し、展示会分の未来在庫（2099/12/31）を81999に作る。99999の在庫は次回用に初期値へ戻す。", "結果メールとログを確認する（失敗・見合わせ時は保守担当へメール）。", "実行中はフロントの注文を止める。展示会受注数が81999の引当として見える。"),
        ("4", "締め後からシーズン前\n2/16-6/30", "27SS", "27FW（81999）", "仕入先への発注は基幹で行い、入荷予定は毎日0:30（B02-IF-3-01）、在庫は毎日2:30（B01-IF-2-01）でECへ。", "受注分を基幹で仕入先へ発注する。", "同じJDE SKUの27SSと27FWが並ぶ期間。連携がどちらの商品に入るかを下の確認事項で決める。"),
        ("5", "シーズン切替\n7/1 5:00", "27FWへ切替", "27FW", "【B05-EC-1-02 シーズン切替】\n27SSの在庫・未来在庫・入出庫履歴・注文の引当を、JDE SKUが同じ27FWへ移す。27SSの販売終了日を6/30 23:59:59にする。", "結果を確認する。27FWが作られていない品番は切り替わらず、27SSが売れ続ける。", "実行中は注文を止める。7/1以降、通常注文の価格は27FWの定価になる。"),
        ("6", "切替のあと\n7/1以降", "27FW", "（次は28SS）", "毎日の在庫・入荷予定・出荷実績連携は27FWに入る。", "27FWをコピーして28SSを作り、次の展示会の設定をする。\n過去シーズン商品を残すか消すかを判断する。", "過去シーズンを消すと、購入履歴から開く商品詳細とダウンロードが使えなくなる。"),
    ],
    [TEAL] * 7,
    row_bgs=lambda i, v: [TEAL_LT, WHITE, BLUE_LT, BLUE_LT, ORANGE_LT, YELLOW, PALE],
    height=[92, 60, 92, 70, 92, 72],
)
r = note(
    r,
    "在庫の動き（例）：展示会締め前は27SS 実在庫50・入荷予定100、27FW 展示会受注200（99999）。締め後、27FWの200は81999の引当と2099/12/31の未来在庫になる。"
    "シーズン切替で27SSの実在庫50と入荷予定100が27FWへ移り、27SSは0。以降は基幹の入荷予定（例：1/10に50、2/10に50、3/10に100）で27FWの未来在庫が実日付に置き換わる。",
    bg=PALE,
    color=INK,
    height=46,
)
r += 1

# ------------------------------------------------------------------ 3
r = band(r, "3. シーズン管理のバッチ　起動条件と処理の中身", ORANGE)
r = table(
    r,
    [1, 2, 1, 2, 1],
    ["ID", "バッチ", "いつ実際に動くか", "処理内容", "失敗・待機のとき"],
    [
        ("B05-EC-1-01", "展示会受注締め\n（rawlings_exhibition_end）", "5分間隔で起動。設定した起動時刻（初期5:00）で、かつ展示会終了日の翌日だけ処理する。", "①メンテナンスモードにする ②DBバックアップ ③展示会注文（期間内・展示会区分）の引当を99999→81999 ④展示会分の未来在庫を81999に作成（2099/12/31） ⑤99999の在庫を次回用に戻す ⑥メンテナンス解除。③-⑤は一つのトランザクション。", "バックアップ失敗なら処理しない。メンテナンス中は5分おきに待ち、待機時間を超えたら当日は見合わせて保守担当へメール。"),
        ("B05-EC-1-02", "シーズン切替\n（rawlings_season_change）", "5分間隔で起動。設定した起動時刻で、かつシーズン開始日（SS 2/1、FW 7/1）だけ処理する。", "①メンテナンスモードにする ②DBバックアップ ③前シーズン商品ごとに、JDE SKUが同じ新シーズン商品へ未来在庫・在庫・入出庫履歴・注文の引当を移す ④前シーズン商品の販売終了日を前日23:59:59 ⑤メンテナンス解除。新シーズン商品がない品番は飛ばす。", "同上。"),
        ("参考", "毎日の連携", "入荷予定0:30、在庫2:30、出荷返品22:30。受注連携は8:00-20:00の30分間隔。", "シーズン管理のバッチは、これらと同じく実行中に注文を止める。受注連携は止めない。", "先行が終わらない場合の待機と見合わせは共通。"),
    ],
    [ORANGE] * 5,
    row_bgs=lambda i, v: [ORANGE_LT, WHITE, WHITE, WHITE, PALE],
    height=[82, 82, 46],
)
r += 1

# ------------------------------------------------------------------ 4
r = band(r, "4. ローリングスジャパン様に行っていただくこと（毎シーズン）", BROWN)
r = table(
    r,
    [1, 1, 2, 2, 1],
    ["No", "いつまでに", "作業", "設定する内容", "やらないと何が起きるか"],
    [
        ("1", "展示会開始の前", "新シーズン商品の作成（前シーズンをコピー）", "ECR商品コード「品番_シーズン」、シーズンコード、定価、画像、カテゴリ、販路、販売開始日。JDE SKUは前シーズンと同じ。", "展示会注文を受けられない。シーズン切替で在庫の移し先がない。"),
        ("2", "展示会開始の前", "展示会開始の設定", "展示会期間、展示会サプライヤ99999での取扱い、未来在庫（2099/12/31）。継続品も同じ。", "展示会ボタンが出ない、または架空在庫がなく注文できない。"),
        ("3", "毎年1回（変更時）", "シーズン開始日の設定", "SS・FWそれぞれの開始日時。初期はSS 2/1 0:00、FW 7/1 0:00。", "シーズン切替が想定外の日に動く、または動かない。"),
        ("4", "展示会締めのあと", "仕入先への発注", "展示会受注分を基幹で発注する。入荷予定は翌日0:30にECへ入る。", "展示会注文の納期が2099/12/31のまま進まない。"),
        ("5", "バッチの翌朝", "結果の確認", "締め・切替の結果、見合わせメールが来ていないか。", "見合わせに気づかず、前シーズンで販売が続く。"),
        ("6", "シーズン切替のあと", "過去シーズン商品の扱い", "残す（販売終了のまま）か、一定期間後に削除するか。", "削除すると購入履歴からの商品詳細・ダウンロードが使えない。"),
    ],
    [BROWN] * 5,
    row_bgs=lambda i, v: [YELLOW, WHITE, WHITE, WHITE, PALE],
    height=[52, 46, 40, 40, 34, 40],
)
r += 1

# ------------------------------------------------------------------ 5
r = band(r, "5. レビューで決めたいこと・設計書との食い違い", "8A6A12")
r = table(
    r,
    [1, 2, 3, 1],
    ["No", "論点", "内容", "決める人"],
    [
        ("1", "締め後からシーズン前の在庫連携", "2/16-6/30は同じJDE SKU・81999で27SSと27FWが並ぶ。基幹の在庫・入荷予定は品番単位なので、どちらのEC商品に入れるかを決める必要がある。シミュレーションでは入荷予定を27FWに入れているが、切替（7/1）より前の日付になっている。", "貴社・弊社"),
        ("2", "新シーズン商品がない品番", "シーズン切替は該当品番を飛ばし、前シーズン商品の販売終了日も入れない。前シーズンのまま売り続けてよいか、作成漏れとして検知するか。", "貴社"),
        ("3", "シーズンコードを入れる項目", "バッチは商品表示の「おすすめ」項目でシーズンを判定する。管理画面でどの項目に何を入れるかを貴社作業として明記する。", "弊社"),
        ("4", "過去シーズンの削除", "購入履歴から過去シーズンの商品詳細・ダウンロードを使う要望と両立しない。残す期間を決める。", "貴社"),
        ("5", "設定は誰が変えるか", "展示会期間、シーズン開始日、バッチ起動時刻は店舗設定。貴社が管理画面で変えるのか、弊社へ依頼するのか。", "貴社・弊社"),
        ("6", "展示会架空在庫の初期値", "検討資料とシミュレーションは10,000、バッチ設計の再設定値は100,000。どちらに揃えるか。", "弊社"),
        ("7", "シーズン切替の設定項目名", "B05-EC-1-02の起動時刻・待機時間（R020、R021）が「展示会受注締め」の名前で書かれている。設計書を「シーズン切替」に直す。", "弊社"),
    ],
    ["8A6A12"] * 4,
    row_bgs=lambda i, v: [YELLOW, WHITE, WHITE, PALE],
    height=[66, 46, 46, 40, 40, 34, 40],
)

ws.freeze_panes = "A4"
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 100
ws.sheet_properties.tabColor = TEAL
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins = PageMargins(0.4, 0.4, 0.5, 0.4, 0.25, 0.25)
ws.print_options.horizontalCentered = True
ws.oddFooter.left.text = "ローリングスジャパン様 シーズン管理 ユーザー設計レビュー"
ws.oddFooter.right.text = "&P / &N"

wb.properties.title = "シーズン管理 運用とバッチ"
wb.properties.creator = "フライトソリューションズ"
wb.save(OUT)
print("wrote", OUT, "rows", r)
