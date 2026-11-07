#!/usr/bin/env python3
"""ユーザー設計レビュー資料を、画面とバッチが分かれて話せる構成に整える。"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/ローリングスジャパン様_ユーザー設計レビュー資料_v2.xlsx"

NAVY = "1B3A4B"
TEAL = "1D7A70"
BLUE = "3D5A80"
ORANGE = "C46B2C"
WHITE = "FFFFFF"
INK = "1B3A4B"
GRAY = "5C6770"
PALE = "F4F7FB"
BLUE_LT = "E7EEF6"
TEAL_LT = "E5F3F1"
ORANGE_LT = "FDEFE4"
YELLOW = "FFF4D6"
CREAM = "FFF8EE"
LINE = "D5DBE0"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
center = Alignment(wrap_text=True, vertical="center", horizontal="center")
left = Alignment(wrap_text=True, vertical="center", horizontal="left")


def font(size=10, bold=False, color=INK):
    return Font(name="メイリオ", size=size, bold=bold, color=color)


def fill(color):
    return PatternFill("solid", fgColor=color)


def paint(cell, value, size=10, bold=False, color=INK, bg=None, align=None):
    cell.value = value
    cell.font = font(size, bold, color)
    cell.alignment = align or left
    cell.border = thin
    if bg:
        cell.fill = fill(bg)


def widths(ws, cols):
    for i, w in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def setup(ws, landscape=True):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.page_margins = PageMargins(0.4, 0.4, 0.55, 0.4, 0.25, 0.25)
    ws.oddFooter.left.text = "ローリングスジャパン様 ユーザー設計レビュー"
    ws.oddFooter.right.text = "画面とバッチを分けて説明"
    ws.sheet_properties.tabColor = TEAL
    ws.page_setup.scale = 70
    ws.print_options.horizontalCentered = True
    ws.sheet_format.defaultRowHeight = 18


def title(ws, last, text, note):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last)
    c = ws.cell(1, 1, text)
    c.font = font(16, True, WHITE)
    c.fill = fill(NAVY)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last)
    c = ws.cell(2, 1, note)
    c.font = font(9, False, "7A4E00")
    c.fill = fill(CREAM)
    c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
    ws.row_dimensions[2].height = 32


def header(ws, row, labels, colors):
    ws.row_dimensions[row].height = 32
    for i, (label, color) in enumerate(zip(labels, colors), 1):
        paint(ws.cell(row, i), label, 9, True, WHITE, color, center)


wb = Workbook()

# ----- 0 読み方 -----
ws = wb.active
ws.title = "0.読み方"
title(ws, 4, "設計レビューの話し方", "1-2はルール、1-3は時刻、1-4はシナリオです。同じ操作を二度説明しない。")
rows = [
    ("1", "1-1 システム概略", "誰と、どのシステムが境界かだけを見せる。操作の詳細には入らない。"),
    ("2", "1-2 制約事項", "誰が、いつまで、何をしてよいか。変更・キャンセル・返品のルールはここで言い切る。"),
    ("3", "1-4 シナリオ", "注文のあと、管理画面で人がすることと、バッチで戻ることを交互に話す。ルール全文は繰り返さない。"),
    ("4", "1-3 バッチ", "時刻、メンテナンスにする夜間処理、30分間隔の受注連携、締めとシーズン切替の順番だけを開く。"),
]
header(ws, 4, ["順", "開くシート", "そこで話すこと", "開かないこと"], [NAVY, NAVY, NAVY, NAVY])
body = [
    "操作手順、バッチID",
    "バッチの時刻、画面の項目",
    "1-2に書いた可否の読み上げ",
    "画面の項目説明",
]
for i, (a, b, c) in enumerate(rows):
    r = 5 + i
    ws.row_dimensions[r].height = 36
    paint(ws.cell(r, 1), a, 12, True, TEAL, TEAL_LT, center)
    paint(ws.cell(r, 2), b, 11, True, INK, WHITE)
    paint(ws.cell(r, 3), c, 10, False, INK, WHITE)
    paint(ws.cell(r, 4), body[i], 10, False, GRAY, PALE)
ws.merge_cells("A10:D10")
ws.cell(10, 1, "変更・キャンセル・返品は、起点が管理画面と基幹の二つある。ECへの反映は一つのバッチ。1-4の最後の表を画面に出したまま話す。").font = font(10, True, NAVY)
ws.cell(10, 1).alignment = left
ws.row_dimensions[10].height = 28
widths(ws, [8, 28, 78, 36])
setup(ws)
ws.freeze_panes = "A5"
ws.auto_filter.ref = None
ws.print_title_rows = "1:4"

# ----- 1-1 -----
ws = wb.create_sheet("1-1.システム概略")
title(ws, 3, "1-1. システム概略", "レビューの導入で、注文がフロント、管理画面、基幹、バッチのどこを通るかを見せる。")
header(ws, 4, ["境界", "役割", "この場で言うこと"], [NAVY, NAVY, NAVY])
actors = [
    ("フロント", "取引先（代理店含む）、代理店の取引先、営業", "注文、閲覧、連携前のキャンセル、連携後の依頼まで。変更は初回リリースでは不可。"),
    ("管理画面", "営業・CS", "代理注文、参照、納期回答までの変更、基幹結果の確認。与信・出荷・返品の確定はしない。"),
    ("基幹 JDE", "与信、納期、出荷、変更、キャンセル、返品の確定", "ECでできない確定はこちら。結果はバッチでECへ戻る。"),
    ("バッチ", "ECと基幹の往復", "受注の往復は30分間隔。在庫・入荷予定・出荷返品・展示会締めは夜間。"),
    ("EDI", "ALPEN、HIMARAYA、XEBIO、Shiratori、MegaSports", "ECでは購入履歴の閲覧のみ。与信以降は基幹側の手動対応。"),
    ("倉庫", "81999 日本トランスシティ、82999 社内、88999 海外", "展示会の引当は架空サプライヤー99999から81999へ夜間に切り替わる。"),
]
for i, rec in enumerate(actors):
    r = 5 + i
    ws.row_dimensions[r].height = 36
    bg = [BLUE_LT, TEAL_LT, ORANGE_LT, YELLOW, WHITE, PALE][i]
    for c, val in enumerate(rec, 1):
        paint(ws.cell(r, c), val, 10, c == 1, INK, bg)
widths(ws, [22, 62, 78])
setup(ws)
ws.freeze_panes = "A5"

# ----- 1-2 -----
ws = wb.create_sheet("1-2.制約事項")
title(ws, 9, "1-2. 業務・システムの制約事項", "ルールの表。バッチの時刻は1-3、画面遷移は1-4。左と同じ、は省略しない。")
labels = ["No", "項目", "フロント\n取引先（代理店含む）", "フロント\n代理店の取引先", "フロント\n営業", "管理画面で人がすること", "バッチで起きること", "基幹で人が始めること", "説明・確認"]
colors = [NAVY, NAVY, BLUE, BLUE, BLUE, TEAL, ORANGE, "6B4C2A", "8A6A12"]
header(ws, 4, labels, colors)
rules = [
    ("1", "注文モード", "TOPとグロナビで通常、カスタム、展示会を選ぶ。", "取引先（代理店含む）と同じ。", "取引先へのログインが前提。できることは同じ。", "モードなし。", "なし。", "なし。", "営業が代理注文以外でフロントを使うか確認する。"),
    ("2", "在庫表示", "手持在庫と入荷予定を表示。入荷予定日はサイト一律で加算。販売可能数は上限値。", "取引先（代理店含む）と同じ。", "取引先（代理店含む）と同じ。", "同じ内容を参照する。", "在庫はB01-IF-2-01、日次2:30。入荷予定はB02-IF-3-01、日次0:30。実行中は注文を止める。", "基幹が在庫と入荷予定の元。", "加算日数は未定。"),
    ("3", "展示会発注数", "未納シーズンごとの発注数。請求先の発注数。", "代理店が自分で抱えている分だけ。配下に紐づく分は含めない。", "取引先へのログインが前提。", "注文検索を展示会モードで絞って一覧する。", "なし。", "なし。", ""),
    ("4", "価格表示", "上代と下代。ユーザーマスタのフラグで判断。", "上代のみ。", "上代と下代。", "上代と下代。取引先別下代はCSVを随時アップロード。", "なし。下代の元データは基幹からダウンロード。", "下代の元を出す。", "機能は作る。初期リリースは全顧客を上代表示。"),
    ("5", "商品ダウンロード", "閲覧できるSKUを、商品ごとと一括で取得。下代あり。", "同じ範囲。下代なし。", "同じ範囲。下代あり。", "CSVダウンロード。", "なし。", "なし。", ""),
    ("6", "新規注文", "自社注文と、配下への代理注文。納品予定が複数なら順次か揃い出し。一つなら納品希望日。既存出荷先と一見出荷先。", "自社の注文だけ。", "代理注文。", "代理注文。", "B03-IF-4-01。30分間隔（0分、30分）。納品予定日でオーダーを分ける。", "受けるだけ。ECから届く。", "一見出荷先は出荷先コード＝請求先コードで連携される。基幹で付け替えたあと、ECの出荷先コードがどうなるか確認。"),
    ("7", "カスタム注文", "デザインID、管理番号、仕様書を5つまで。", "取引先（代理店含む）と同じ。", "取引先へのログインが前提。", "内容が戻ったとき、注文者へメール。納期と金額は管理画面で回答。", "変更はB03-IF-4-01とB03-IF-4-02の往復。どちらも30分間隔。", "デザインに合わせて基幹でも内容を変える。", "シミュレーションサイトはユニフォームとグラブで違う。"),
    ("8", "展示会注文", "期間中だけ注文できる。期間中の変更は問い合わせ。", "取引先（代理店含む）と同じ。", "取引先（代理店含む）と同じ。", "期間の制限なし。締めボタンはない。", "締めB05-EC-1-01、シーズン切替B05-EC-1-02。どちらも5:00。引当を99999から81999へ移す。", "なし。切り替えはECの夜間バッチ。", "バッチの中身は1-3で話す。"),
    ("9", "EDI注文", "購入履歴の閲覧だけ。", "取引先（代理店含む）と同じ。", "取引先（代理店含む）と同じ。", "閲覧だけ。", "出荷実績の日次B04-IF-5-01で取り込む。", "与信、納期、出荷、変更、キャンセル、返品は基幹側の手動。", "請求先コードしか分からない。請求先の注文として登録してよいか確認。"),
    ("10", "与信回答", "操作なし。", "操作なし。", "操作なし。", "判断はしない。戻ったJDEオーダーNo、またはキャンセルを見る。バッチと同じ結果を手入力できる。", "B03-IF-4-02。30分間隔（15分、45分）。", "OKならJDEオーダーNoを採番。NGなら基幹でキャンセル。", "手動は代替入力。原則はバッチ。"),
    ("11", "納期回答", "購入履歴で見る。", "取引先（代理店含む）と同じ。", "取引先へのログインが前提。", "人が回答を作るのではない。連携時に注文者へメール。手入力はバッチの代替。", "B03-IF-4-02。30分間隔（15分、45分）。", "納期を回答する。", ""),
    ("12", "出荷実績", "購入履歴で見る。出荷後に納品書をダウンロードできる。", "取引先（代理店含む）と同じ。", "取引先へのログインが前提。", "確定はしない。連携時にメールと納品書を作る。手入力はバッチの代替。EDI注文もここで見える。", "B04-IF-5-01。日次22:30。実行中は注文を止める。", "出荷を確定する。", ""),
    ("13", "注文変更", "機能なし。展示会だけの変更は初回リリースに含めない。", "機能なし。", "機能なし。", "納期回答まで変更できる。変更時にメール。", "画面で変えた内容はB03-IF-4-01。基幹で変えた結果はB03-IF-4-02。", "基幹でも変更できる。", "可能期間は納期回答まで。1-4の通常シナリオで経路を話す。"),
    ("14", "キャンセル", "連携前は受注待ちまで、その場でキャンセル。連携後は出荷待ち（納期回答）まで依頼。単位はJDEオーダー。履歴で見られる。どちらもメール。", "取引先（代理店含む）と同じ。", "取引先（代理店含む）と同じ。", "連携前はフロントが確定するので、管理画面では確定しない。連携後は結果が戻ってから成立させる。NGなら出荷待ちのまま。その後は基幹のキャンセルを再連携するか、返品へ回す。", "依頼はB03-IF-4-01。結果はB03-IF-4-02。", "基幹でもキャンセルできる。NGのときは基幹側もキャンセルしない。", "起点はフロント、管理結果処理、基幹の三つ。ECへの反映は戻りバッチ一つ。"),
    ("15", "返品", "出荷済み、または返品完了のあと、明細単位で依頼。メール。履歴で状況を見る。", "取引先（代理店含む）と同じ。", "取引先（代理店含む）と同じ。", "確定ボタンはない。実績が戻ったときメールする。", "B04-IF-5-01。日次22:30。出荷実績と同じ。実行中は注文を止める。", "返品を確定する。", "管理画面の数量変更では返品しない。"),
    ("16", "購入履歴", "閲覧、再注文、キャンセル依頼、返品依頼、納品書。カスタムのIDと仕様書。仕様書のアップロードは出荷待ち（納期回答）まで。", "取引先（代理店含む）と同じ。", "取引先（代理店含む）と同じ。", "同じ内容を参照する。", "なし。", "なし。", ""),
    ("17", "受注集計表", "なし。", "なし。", "なし。", "取引先別のExcelを随時作る。", "なし。", "なし。", ""),
]
for i, rec in enumerate(rules):
    r = 5 + i
    ws.row_dimensions[r].height = 72
    bgs = [PALE if i % 2 else WHITE] * 9
    bgs[5] = TEAL_LT
    bgs[6] = ORANGE_LT
    bgs[8] = YELLOW if rec[8] else bgs[0]
    for c, val in enumerate(rec, 1):
        paint(ws.cell(r, c), val, 9, c == 1, INK, bgs[c - 1], center if c == 1 else left)
widths(ws, [6, 16, 28, 24, 22, 32, 32, 28, 28])
setup(ws)
ws.freeze_panes = "C5"
ws.auto_filter.ref = "A4:I21"
ws.print_title_rows = "1:4"
ws.auto_filter.ref = "A4:I21"
ws.page_setup.fitToWidth = 1
ws.oddHeader.left.text = "1-2 制約事項"

# ----- 1-3 -----
ws = wb.create_sheet("1-3.バッチ概要")
title(ws, 6, "1-3. バッチ概要", "1-2の「30分間隔」「日次」は、この時刻を指す。受注連携だけは、実行中も注文を止めない。")
header(ws, 4, ["ID", "バッチ", "方向", "ECで起きること", "時刻", "実行中のフロント"], [NAVY, NAVY, NAVY, NAVY, NAVY, NAVY])
batches = [
    ("B03-IF-4-01", "受注情報連携", "EC→基幹", "新規、変更、キャンセル依頼を送る。", "30分間隔。0分、30分", "止めない"),
    ("B03-IF-4-02", "受注情報連携", "基幹→EC", "与信、納期、変更、キャンセルの結果を受ける。", "30分間隔。15分、45分", "止めない"),
    ("B02-IF-3-01", "入荷予定連携", "基幹→EC", "未来在庫を更新する。", "日次 0:30", "注文を止める"),
    ("B01-IF-2-01", "在庫連携", "基幹→EC", "在庫数と引当数を更新する。", "日次 2:30", "注文を止める"),
    ("B05-EC-1-01", "展示会受注締め", "EC内", "今季展示会を終え、引当を99999から81999へ移す。", "日次 5:00", "注文を止める"),
    ("B05-EC-1-02", "シーズン切替", "EC内", "前季の引当を、今季の一般販売商品へまとめる。", "日次 5:00。締めの完了後", "注文を止める"),
    ("B04-IF-5-01", "出荷・返品実績", "基幹→EC", "出荷と返品の実績を受ける。納品書とメールは取り込み後。", "日次 22:30", "注文を止める"),
]
for i, rec in enumerate(batches):
    r = 5 + i
    ws.row_dimensions[r].height = 32
    stop = rec[5] != "止めない"
    bg = ORANGE_LT if stop else TEAL_LT
    for c, val in enumerate(rec, 1):
        paint(ws.cell(r, c), val, 9, c == 1, INK, bg)
ws.merge_cells("A13:F13")
ws.cell(13, 1, "締めとシーズン切替はどちらも5:00。先行が終わっていなければ5分おきに待ち、店舗設定の待ち時間を超えたら当日はスキップして保守担当へメールする。").font = font(9, False, INK)
ws.cell(13, 1).alignment = left
ws.cell(13, 1).fill = fill(YELLOW)
ws.row_dimensions[13].height = 32
ws.merge_cells("A14:F14")
ws.cell(14, 1, "時刻の初期値は上表。受注連携の動かす時間帯は店舗設定「受注連携バッチ起動対象時間」、日次の時刻は「各バッチ起動時間」。").font = font(9, False, GRAY)
ws.cell(14, 1).alignment = left
ws.row_dimensions[14].height = 22
widths(ws, [20, 22, 14, 62, 32, 18])
setup(ws)
ws.freeze_panes = "A5"
ws.print_title_rows = "1:4"

# ----- 1-4 -----
ws = wb.create_sheet("1-4.シナリオ")
title(ws, 6, "1-4. シナリオ別の説明", "空欄は、上の共通画面と同じ。セルには説明ポイントだけを書き、ルールの全文は1-2のNoを指す。")
ws.merge_cells("A4:F4")
paint(ws.cell(4, 1), "A. 全シナリオ共通のフロント画面。ここは最初に一度だけ話す。", 11, True, WHITE, BLUE, left)
ws.row_dimensions[4].height = 22
header(ws, 5, ["画面", "説明するポイント", "画面", "説明するポイント", "画面", "説明するポイント"], [BLUE, BLUE, BLUE, BLUE, BLUE, BLUE])
common = [
    ("ログイン", "お知らせの登録と表示", "カート", "上代と下代", "注文完了", "表示内容"),
    ("TOP", "キーワード検索、展示会期間、購入履歴のある商品、おすすめ、バナー、営業日カレンダー、フッター", "お届け先", "既存出荷先、一見出荷先、JDEの桁、出荷方法による納品予定日", "購入履歴一覧", "検索条件、表示、上代と下代"),
    ("商品一覧", "注文モードで検索対象が変わる。サイズ、在庫あり、新商品。カート追加とNewの条件", "注文内容の確認", "表示内容、送料", "購入履歴詳細", "表示、上代と下代"),
    ("商品詳細", "上代と下代。展示会数はログイン者で変わる。在庫上限と入荷予定日の加算。カートイン、お気に入り、ダウンロード", "クイックオーダー", "SKU選択、注文モードによる絞り込み、カートイン", "", ""),
]
for i, rec in enumerate(common):
    r = 6 + i
    ws.row_dimensions[r].height = 48
    for c, val in enumerate(rec, 1):
        paint(ws.cell(r, c), val, 9, c % 2 == 1, INK, BLUE_LT if c % 2 else WHITE)

ws.merge_cells("A11:F11")
paint(ws.cell(11, 1), "B. シナリオ。フロント列は共通からの差分だけ。管理画面とバッチを交互に話す。", 11, True, WHITE, TEAL, left)
ws.row_dimensions[11].height = 22
header(ws, 12, ["No", "話す順番", "フロントの差分", "管理画面で人がすること", "バッチで起きること", "ここで言うポイント"], [TEAL, TEAL, BLUE, TEAL, ORANGE, "8A6A12"])
scenarios = [
    ("1", "通常。注文、与信、変更、納期、出荷。", "共通画面のまま。", "代理注文と参照。与信と納期と出荷は、人が判断しない。バッチと同じ結果の手入力だけが代替。変更は納期回答まで、注文編集で行う。", "注文はB03-IF-4-01。与信と納期の戻りはB03-IF-4-02。出荷はB04-IF-5-01。", "送料は別明細。納品書はJDEオーダー単位。変更できるのは納期回答まで。ルールは1-2の10、11、12、13。"),
    ("2", "通常。キャンセル依頼からキャンセル。", "連携前は購入履歴でその場のキャンセル。連携後は同じ画面で依頼。", "連携前は確定しない。結果が戻ったあと、成立させるか出荷待ちへ戻す。", "依頼はB03-IF-4-01。可否はB03-IF-4-02。", "基幹で先にキャンセルした場合も、ECへは同じ戻りバッチ。NGなら出荷待ちのままにし、再連携か返品。ルールは1-2の14。"),
    ("3", "通常。返品依頼から返品。", "出荷後、購入履歴から明細単位で依頼。", "確定しない。実績が戻るとメールする。数量変更では返品しない。", "B04-IF-5-01。日次22:30。実行中は注文を止める。", "確定の起点は基幹だけ。管理画面とバッチの両方が確定するわけではない。ルールは1-2の15。"),
    ("4", "カスタム。変更、納期、出荷。", "カートでデザインID、管理番号、仕様書5つ。確認画面と履歴に表示し、履歴から仕様書を出し入れする。", "納期と金額を回答。注文編集で数量、単価、送料、デザインID、仕様書を扱う。", "内容の往復はB03-IF-4-01と02。出荷はB04-IF-5-01。", "シミュレーションサイトはユニフォームとグラブで違う。仕様書のアップロードは納期回答まで。ルールは1-2の7。"),
    ("5", "展示会。締め、入荷、出荷、シーズン切替。", "期間中だけ注文できる。期間中の変更は問い合わせ。", "期間制限なし。締めボタンはない。", "入荷予定B02、在庫B01、締めB05-EC-1-01、切替B05-EC-1-02、出荷B04。夜間は注文を止める。", "事前設定を先に見せる。締めと切替は同じ5:00で、締めが終わってから切替。引当は99999から81999。ルールは1-2の8。"),
    ("6", "EDI。閲覧だけ。", "購入履歴に出る。変更もキャンセルも返品もできない。", "閲覧だけ。", "出荷実績の日次で載る。与信以降の往復バッチには乗らない。", "請求先コードしか来ない。請求先の注文として登録してよいか、1-2の9で確認する。"),
]
for i, rec in enumerate(scenarios):
    r = 13 + i
    ws.row_dimensions[r].height = 78
    bgs = [WHITE, PALE, TEAL_LT, ORANGE_LT, YELLOW]
    paint(ws.cell(r, 1), rec[0], 12, True, TEAL, TEAL_LT, center)
    for c, val in enumerate(rec[1:], 2):
        paint(ws.cell(r, c), val, 9, False, INK, bgs[c - 2])

ws.merge_cells("A20:F20")
paint(ws.cell(20, 1), "C. 変更・キャンセル・返品だけ、起点を並べる。Bのシナリオ2、3と1の変更箇所で、この表を出したまま話す。", 11, True, WHITE, "6B4C2A", left)
ws.row_dimensions[20].height = 22
header(ws, 21, ["操作", "フロント", "管理画面", "基幹", "ECへ戻すバッチ", "管理画面とバッチの違い"], ["6B4C2A"] * 6)
forks = [
    ("注文変更", "できない。", "納期回答まで編集し、基幹へ送る。", "基幹でも変更できる。", "画面起点はB03-IF-4-01。基幹起点はB03-IF-4-02。", "どちらが始めても、ECの確定データはバッチのあとでそろう。"),
    ("キャンセル\n連携前", "その場で確定する。", "確定しない。", "まだ渡っていない。", "送らない。", "バッチを待たない。"),
    ("キャンセル\n連携後", "依頼だけ出す。", "結果を見て、成立か出荷待ちへ戻す。", "可否を決める。基幹から始めることもできる。", "依頼はB03-IF-4-01。結果はB03-IF-4-02。", "管理画面は判断の結果を受ける。可否そのものは基幹。"),
    ("返品", "明細単位で依頼する。", "確定しない。メールだけ。", "確定する。", "B04-IF-5-01だけ。日次。", "管理画面に確定ボタンはない。"),
]
for i, rec in enumerate(forks):
    r = 22 + i
    ws.row_dimensions[r].height = 48
    for c, val in enumerate(rec, 1):
        bg = YELLOW if c == 6 else (TEAL_LT if c == 3 else ORANGE_LT if c == 5 else WHITE)
        paint(ws.cell(r, c), val, 9, c == 1, INK, bg, center if c == 1 else left)
widths(ws, [18, 36, 36, 32, 36, 36])
setup(ws)
ws.freeze_panes = "A6"
ws.print_title_rows = "1:2"
ws.page_setup.fitToWidth = 1
ws.oddHeader.left.text = "1-4 シナリオ"

wb.properties.title = "ローリングスジャパン様 ユーザー設計レビュー資料"
wb.properties.creator = "フライトソリューションズ"
wb.save(OUT)
print("wrote", OUT)
