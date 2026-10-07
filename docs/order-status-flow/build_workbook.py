#!/usr/bin/env python3
"""注文ステータス図を、遷移・業務・連携・メールに分けたブックにする。"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/docs/order-status-flow/注文ステータスと業務フロー.xlsx"

NAVY = "1B3A4B"
INK = "1F2933"
MUTED = "52606D"
WHITE = "FFFFFF"
PAPER = "F7F5F2"
YELLOW = "FFF4CC"  # ステータス
BLUE = "E7F0FA"  # 業務
PURPLE = "F3EAF8"  # 連携
GREEN = "E5F6EC"  # メール
RED = "FDECEC"  # 例外・赤線
AMBER = "FFF1DE"  # 差戻し
GRAY = "EEEEEE"  # 利用しない
LINE = "D5D1CB"
GOLD = "C4A35A"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
wrap = Alignment(wrap_text=True, vertical="center", horizontal="left")
wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")
font_name = "Yu Gothic"


def font(size=11, bold=False, color=INK, name=font_name):
    return Font(name=name, size=size, bold=bold, color=color)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def apply_header(ws, row, cols, fill_color=NAVY):
    for col in range(1, cols + 1):
        cell = ws.cell(row, col)
        cell.font = font(11, True, WHITE)
        cell.fill = fill(fill_color)
        cell.alignment = wrap_center
        cell.border = thin


def paint_body(ws, start, end, cols, zebra=PAPER):
    for r in range(start, end + 1):
        for c in range(1, cols + 1):
            cell = ws.cell(r, c)
            cell.font = font(10)
            cell.alignment = wrap
            cell.border = thin
            if cell.fill is None or cell.fill.fgColor is None or (
                getattr(cell.fill.fgColor, "rgb", None) in (None, "00000000")
                and cell.fill.patternType is None
            ):
                if r % 2 == 0:
                    cell.fill = fill(zebra)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def setup_page(ws, title, fit_width=1):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = fit_width
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.5, header=0.25, footer=0.25)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = "EC-Rider　注文ステータスと業務フロー"
    ws.oddFooter.left.text = title
    ws.oddFooter.right.text = "ページ &P / &N"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.page_setup.horizontalDpi = 300
    ws.page_setup.verticalDpi = 300
    ws.sheet_format.defaultRowHeight = 18


def write_title(ws, text, subtitle, cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    ws.cell(1, 1, text).font = font(18, True, NAVY)
    ws.cell(1, 1).alignment = Alignment(vertical="center", horizontal="left")
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
    ws.cell(2, 1, subtitle).font = font(10, False, MUTED)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 32
    for c in range(1, cols + 1):
        ws.cell(1, c).fill = fill(WHITE)
        ws.cell(2, c).fill = fill(WHITE)


def add_rows(ws, header_row, rows, fills_by_kind=None, kind_col=None):
    for i, row in enumerate(rows):
        r = header_row + 1 + i
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
        ws.row_dimensions[r].height = 48
    paint_body(ws, header_row + 1, header_row + len(rows), len(rows[0]) if rows else 1)
    if fills_by_kind and kind_col:
        for i, row in enumerate(rows):
            r = header_row + 1 + i
            kind = row[kind_col - 1]
            color = fills_by_kind.get(kind)
            if color:
                for c in range(1, len(row) + 1):
                    ws.cell(r, c).fill = fill(color)


def legend_row(ws, row, items, cols):
    """items: list of (label, fill hex)"""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    ws.row_dimensions[row].height = 22
    cell = ws.cell(row, 1, "　".join(label for label, _ in items))
    cell.font = font(9, False, MUTED)
    cell.alignment = Alignment(vertical="center")


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------

NORMAL = [
    [
        "0",
        "注文をつくる",
        "まだステータスはない。作った瞬間に承認待ちか受注待ちへ入る。",
        "顧客: フロントサイトで新規注文。\n運営: 管理サイトで代理注文。代理注文は、注文に使ったフロントユーザの承認権限に合わせて進む。\nEDI: 新規注文データを受け取って注文を作る。",
        "EDI → EC-Rider\n新規注文データ。",
        "この瞬間には送らない。入ったステータス側のメールを見る。",
        "承認権限がない → 承認待ち。\n承認権限がある → 受注待ち（図の赤線）。\nEDIの新規注文 → 受注待ち（図の赤線。承認待ちへの線はない）。",
        "—",
    ],
    [
        "1",
        "承認待ち",
        "承認権限を持たないフロントユーザが注文を登録した状態。",
        "顧客: 同一販売組織（shop）で承認権限を持つユーザが注文承認。\n運営: 代理承認。\n運営: 承認依頼（リクエスト注文のみ）。依頼しても承認待ちのまま。",
        "このステータスでは基幹へ渡さない。",
        "注文承認依頼メール\n宛先は店舗の承認者。",
        "注文承認・代理承認 → 受注待ち。",
        "顧客のキャンセル依頼 → キャンセル依頼中。\n運営のキャンセル → キャンセル完了。",
    ],
    [
        "2",
        "受注待ち",
        "運営の受付を待っている。外部サプライヤー商品で受注確認区分があるときは、確認を取る状態でもある。注文金額を変更できる唯一のステータス。",
        "運営: 注文変更。金額を変えても受注待ちのまま。\n基幹: 連携された注文を取り込み、与信チェックを行う。",
        "EC-Rider → 基幹\n新規注文・代理注文のあと、30分間隔で注文を連携し、受注扱いとする。\n基幹 → EC-Rider\n与信上限内なら基幹オーダー番号が返る。取り込むと出荷待ち。",
        "図では次の3通が、受注待ちから1本の線でまとまっている。\n・注文受付メール\n・注文承認完了メール\n・注文内容変更通知メール\n箱は顧客レーンにある。3通ごとの別線はない。",
        "与信回答を取り込む → 出荷待ち。\n与信が上限に達したときの遷移は、図にも説明にもない。",
        "顧客のキャンセル依頼 → キャンセル依頼中。\n運営のキャンセル → キャンセル完了。\n運営の差戻し先にはならない（差戻しで戻ってくる側）。",
    ],
    [
        "3",
        "出荷待ち",
        "基幹連携のあと、納期回答を待っている。",
        "顧客・運営の追加業務は図にない。納期回答の到着を待つ。",
        "基幹 → EC-Rider\n在庫があればピック番号が返る。取り込むと出荷指示済み。\n与信回答と納期回答は同じタイミングになることがある。",
        "なし。\n（入金確認メールは出荷待ちから線が出ているが、利用しない）",
        "納期回答を取り込む → 出荷指示済み。\n図では基幹の楕円から出荷待ちへ線があり、出荷待ちから出荷指示済みへ縦線が降りている。",
        "顧客のキャンセル依頼 → キャンセル依頼中。\n運営のキャンセル → キャンセル完了。\n運営の差戻し → 受注待ち。\nキャンセル依頼中からの差戻し、出荷指示済みからの差戻しは、ここへ戻る。",
    ],
    [
        "4",
        "出荷指示済み",
        "基幹から納期回答された状態。商品の出荷準備を行う。",
        "出荷準備。担当レーンは図に書かれていない。WMS列に箱はない。\n基幹: 納品予定日に合わせて出荷する。",
        "基幹 → EC-Rider\n出荷実績。取り込むと出荷済み。",
        "納期回答メール\n出荷指示済みから顧客レーンへ線が出ている。説明文は、納期回答の取込時に送る、とある。",
        "出荷実績を取り込む → 出荷済み。",
        "運営のキャンセル → キャンセル完了。\n顧客のキャンセル依頼は、出荷待ちまで。出荷指示済みには線がない。\n運営の差戻し → 出荷待ち。",
    ],
    [
        "5",
        "出荷済み",
        "出荷が完了した状態。",
        "顧客: 返品依頼（注文明細単位）。\n運営: 返品依頼（注文明細単位）。\n説明文は、出荷済み以降で依頼できる、とある。図の線は出荷済みから出ている。",
        "このステータスでは新たに渡すデータは図にない。出荷実績の取込でここに着く。",
        "出荷完了メール（顧客レーン。説明文は注文者へ）。\n出荷完了メール（店舗宛て）（運営レーン）。",
        "返品依頼 → 返品依頼中。",
        "返品依頼中からの差戻しは、ここへ戻る。",
    ],
    [
        "6",
        "返品依頼中",
        "返品を依頼したあとの状態。",
        "基幹: 返品対象が本社へ返送され、返品できるなら返品処理を行う。\n運営: 差戻し。",
        "基幹 → EC-Rider\n返品実績。取り込むと返品完了。",
        "返品依頼メール\n顧客レーンと運営レーンの両方に同じ箱があり、どちらも返品依頼中から線が出ている。説明文は注文者へ送付。",
        "返品実績を取り込む → 返品完了。",
        "運営の差戻し → 出荷済み。",
    ],
    [
        "7",
        "返品完了",
        "商品が返品された状態。",
        "図に次の業務はない。",
        "返品実績の取込でここに着く。",
        "返品完了メール（顧客レーン。説明文は注文者へ）。",
        "図に次のステータスはない。",
        "—",
    ],
    [
        "8",
        "キャンセル依頼中",
        "顧客がキャンセルを依頼したあとの状態。",
        "基幹連携前なら、EC側でキャンセル処理を行う。\n連携済みなら、基幹でキャンセルできるかを判定する。\n運営: 差戻し。\n説明文では、管理サイトからキャンセル依頼中の注文のキャンセルを受けて完了する、ともある。",
        "連携済みのとき\nEC-Rider → 基幹: キャンセル依頼（配送単位＝基幹オーダー単位）。\n基幹 → EC-Rider: キャンセルできるか、できないか。\nできないときの遷移は図にない。\n連携前は基幹へ渡さない。",
        "注文キャンセル依頼メール（顧客レーン。説明文は注文者へ）。",
        "基幹連携前 → キャンセル完了。\n基幹連携後 → 基幹のキャンセルを取り込んでキャンセル完了。",
        "運営の差戻し → 出荷待ち。",
    ],
    [
        "9",
        "キャンセル完了",
        "キャンセルが完了した状態。",
        "図に次の業務はない。",
        "基幹連携後のキャンセルは、基幹のキャンセルデータを取り込んでここに着く。\n運営が承認待ち〜出荷指示済みから直接キャンセルした線も、ここへ着く。",
        "注文キャンセル完了メール（顧客レーン。説明文は注文者へ）。",
        "図に次のステータスはない。",
        "—",
    ],
]

TRANSITIONS = [
    ["T01", "入口", "（なし）", "顧客", "新規注文", "承認権限がない", "承認待ち", "注文", "通常の線", "入ったあとに注文承認依頼メール"],
    ["T02", "入口", "（なし）", "顧客", "新規注文", "承認権限がある", "受注待ち", "注文", "赤線。承認待ちを飛ばす", "入ったあとに受注待ちのメール"],
    ["T03", "入口", "（なし）", "運営", "代理注文", "注文に使ったフロントユーザに承認権限がない", "承認待ち", "注文", "通常の線", "入ったあとに注文承認依頼メール"],
    ["T04", "入口", "（なし）", "運営", "代理注文", "注文に使ったフロントユーザに承認権限がある", "受注待ち", "注文", "赤線。図形への接続情報は欠けているが、顧客の赤線と同じ降り方", "入ったあとに受注待ちのメール"],
    ["T05", "入口", "（なし）", "EDI", "新規注文", "承認待ちへの線は描かれていない", "受注待ち", "注文", "赤線。終端は図形に未接続で、範囲は承認待ちの上から受注待ちの上まで", "入ったあとに受注待ちのメール"],
    ["T06", "通常", "承認待ち", "顧客", "注文承認", "同一販売組織で承認権限を持つ", "受注待ち", "注文", "通常の線", "受注待ち側のメール"],
    ["T07", "通常", "承認待ち", "運営", "代理承認", "—", "受注待ち", "注文", "通常の線", "受注待ち側のメール"],
    ["T08", "滞留", "承認待ち", "運営", "承認依頼", "リクエスト注文のみ", "承認待ち", "注文", "緑線の往復。ステータスは進まない。始点側は図形に未接続で、承認待ちと受注待ちの境目にある", "注文承認依頼メールは承認待ちに紐づく"],
    ["T09", "滞留", "受注待ち", "運営", "注文変更", "金額を変えられるのはここだけ", "受注待ち", "注文", "受注待ちの右で往復", "注文内容変更通知メールは、受注待ちにまとめて描かれている"],
    ["T10", "通常", "受注待ち", "基幹", "与信回答", "与信上限に達していない。基幹オーダー番号が返る", "出荷待ち", "注文", "受注待ちから基幹の楕円を経由", "—"],
    ["T11", "通常", "出荷待ち", "基幹", "納期回答", "在庫があり、ピック番号が返る。与信回答と同時のことがある", "出荷指示済み", "注文", "出荷待ちから出荷指示済みへの縦線。基幹の楕円自体の矢印は出荷待ちで止まっている", "納期回答メール"],
    ["T12", "通常", "出荷指示済み", "基幹", "出荷", "納品予定日に合わせて出荷し、出荷実績が返る", "出荷済み", "注文", "基幹の出荷から出荷済み", "出荷完了メール、出荷完了メール（店舗宛て）"],
    ["T13", "返品", "出荷済み", "顧客", "返品依頼", "出荷済み以降。図の線は出荷済みから", "返品依頼中", "注文明細", "通常の線", "返品依頼メール"],
    ["T14", "返品", "出荷済み", "運営", "返品依頼", "出荷済み以降。図の線は出荷済みから", "返品依頼中", "注文明細", "通常の線", "返品依頼メール（運営レーンにも同じ箱）"],
    ["T15", "返品", "返品依頼中", "基幹", "返品", "本社へ返送され、返品できる", "返品完了", "注文明細", "基幹の返品から返品完了", "返品完了メール"],
    ["T16", "キャンセル", "承認待ち", "顧客", "キャンセル依頼", "出荷待ちまで依頼できる、のうちの一つ", "キャンセル依頼中", "配送（基幹オーダー）", "通常の線", "注文キャンセル依頼メール"],
    ["T17", "キャンセル", "受注待ち", "顧客", "キャンセル依頼", "同上", "キャンセル依頼中", "配送（基幹オーダー）", "通常の線", "注文キャンセル依頼メール"],
    ["T18", "キャンセル", "出荷待ち", "顧客", "キャンセル依頼", "同上", "キャンセル依頼中", "配送（基幹オーダー）", "通常の線", "注文キャンセル依頼メール"],
    ["T19", "キャンセル", "キャンセル依頼中", "EC-Rider", "キャンセル処理", "基幹連携前（判断が YES）", "キャンセル完了", "配送（基幹オーダー）", "判断の下へ降りる線", "注文キャンセル完了メール"],
    ["T20", "キャンセル", "キャンセル依頼中", "基幹", "キャンセル", "基幹連携後（判断が NO）。基幹が結果を返す", "キャンセル完了", "配送（基幹オーダー）", "判断の右から基幹のキャンセルを経由", "注文キャンセル完了メール"],
    ["T21", "キャンセル", "承認待ち", "運営", "キャンセル", "出荷指示済みまで。依頼中は経由しない", "キャンセル完了", "配送（基幹オーダー）", "各ステータスから運営のキャンセルへ集まり、完了へ1本", "注文キャンセル完了メール"],
    ["T22", "キャンセル", "受注待ち", "運営", "キャンセル", "同上", "キャンセル完了", "配送（基幹オーダー）", "同上", "注文キャンセル完了メール"],
    ["T23", "キャンセル", "出荷待ち", "運営", "キャンセル", "同上", "キャンセル完了", "配送（基幹オーダー）", "同上", "注文キャンセル完了メール"],
    ["T24", "キャンセル", "出荷指示済み", "運営", "キャンセル", "顧客のキャンセル依頼はここには届かない。運営のキャンセルだけ", "キャンセル完了", "配送（基幹オーダー）", "同上", "注文キャンセル完了メール"],
    ["T25", "差戻し", "出荷指示済み", "運営", "差戻し", "説明文と同じ", "出荷待ち", "—", "緑線", "—"],
    ["T26", "差戻し", "出荷待ち", "運営", "差戻し", "説明文と同じ", "受注待ち", "—", "緑線", "—"],
    ["T27", "差戻し", "キャンセル依頼中", "運営", "差戻し", "説明文は出荷待ちへ戻す。元のステータスへ戻す線ではない", "出荷待ち", "—", "緑線", "—"],
    ["T28", "差戻し", "返品依頼中", "運営", "差戻し", "説明文と同じ", "出荷済み", "—", "緑線。差戻しの楕円はステータス列の右端", "—"],
    ["T29", "利用しない", "入金待ち", "顧客", "入金", "図に次ステータスへの線はない", "（線なし）", "—", "白抜きの図形", "—"],
    ["T30", "利用しない", "入金待ち", "運営", "入金確認", "—", "出荷待ち", "—", "白抜きの図形", "入金確認メール（利用しない。線は出荷待ちから）"],
    ["T31", "利用しない", "入金待ち", "顧客", "キャンセル依頼", "入金待ち自体を使わない", "キャンセル依頼中", "配送（基幹オーダー）", "ほかのキャンセル依頼と同じ線", "注文キャンセル依頼メール"],
    ["T32", "利用しない", "入金待ち", "運営", "キャンセル", "入金待ち自体を使わない", "キャンセル完了", "配送（基幹オーダー）", "ほかの運営キャンセルと同じ線", "注文キャンセル完了メール"],
]

WORK = [
    ["顧客", "新規注文", "注文前", "承認待ち、または受注待ち", "注文", "フロントサイト。承認権限がなければ承認待ち、あれば受注待ち。"],
    ["顧客", "注文承認", "承認待ち", "受注待ち", "注文", "注文者と同一販売組織で、承認権限を持つユーザが行う。"],
    ["顧客", "キャンセル依頼", "承認待ち、受注待ち、出荷待ち", "キャンセル依頼中", "配送（基幹オーダー）", "出荷指示済みでは依頼できない。"],
    ["顧客", "返品依頼", "出荷済み", "返品依頼中", "注文明細", "説明文は出荷済み以降。フロントサイトから明細単位。"],
    ["顧客", "入金", "入金待ち", "図に次の線はない", "—", "利用しない。"],
    ["運営", "代理注文", "注文前", "承認待ち、または受注待ち", "注文", "管理サイト。進み方は、注文に使ったフロントユーザの承認権限に従う。"],
    ["運営", "代理承認", "承認待ち", "受注待ち", "注文", "管理サイトから注文承認を代行する。"],
    ["運営", "承認依頼", "承認待ち", "承認待ちのまま", "注文", "リクエスト注文のみ。緑線。"],
    ["運営", "注文変更", "受注待ち", "受注待ちのまま", "注文", "金額変更ができる唯一のタイミング。"],
    ["運営", "キャンセル", "承認待ち、受注待ち、出荷待ち、出荷指示済み", "キャンセル完了", "配送（基幹オーダー）", "キャンセル依頼中を経由しない。説明文では、基幹からキャンセルデータが返って取り込み完了する、とも書かれている。図の線は基幹を経由していない。"],
    ["運営", "返品依頼", "出荷済み", "返品依頼中", "注文明細", "管理サイトから明細単位。"],
    ["運営", "差戻し", "出荷指示済み", "出荷待ち", "—", "戻せる組み合わせは説明文に4つだけ。"],
    ["運営", "差戻し", "出荷待ち", "受注待ち", "—", "同上。"],
    ["運営", "差戻し", "キャンセル依頼中", "出荷待ち", "—", "依頼前のステータスへ戻す線ではない。行き先は出荷待ち。"],
    ["運営", "差戻し", "返品依頼中", "出荷済み", "—", "同上。"],
    ["運営", "入金確認", "入金待ち", "出荷待ち", "—", "利用しない。"],
    ["運営", "出荷準備", "出荷指示済み", "出荷の実行は基幹側", "—", "説明文にある。担当レーンは図にない。WMS列に箱はない。"],
    ["EDI", "新規注文の取り込み", "注文前", "受注待ち", "注文", "業務としては注文を作ること。データの向きはシステム連携シート。"],
    ["基幹", "与信チェック", "受注待ち", "出荷待ち", "注文", "上限内なら基幹オーダー番号を返す。上限到達時は図にない。"],
    ["基幹", "納期回答", "出荷待ち", "出荷指示済み", "注文", "在庫があればピック番号を返す。与信回答と同時のことがある。"],
    ["基幹", "出荷", "出荷指示済み", "出荷済み", "注文", "納品予定日に合わせる。出荷実績を返す。"],
    ["基幹", "キャンセル判定と処理", "キャンセル依頼中（連携済み）", "キャンセル完了", "配送（基幹オーダー）", "可能か不能かを返す。不能のあとの遷移は図にない。"],
    ["基幹", "返品処理", "返品依頼中", "返品完了", "注文明細", "本社へ返送され、返品できる場合に行う。"],
]

LINKS = [
    ["L01", "EDI → EC-Rider", "新規注文データを受け取ったとき", "新規注文データ", "ECが注文を作る", "受注待ち", "承認待ちへの線は描かれていない。図の線は赤。"],
    ["L02", "EC-Rider → 基幹", "新規注文・代理注文のあと。受注待ち", "注文データ", "基幹が受注として取り込む", "受注待ちのまま（受注扱い）", "30分間隔。"],
    ["L03", "基幹 → EC-Rider", "与信チェックのあと", "基幹オーダー番号", "ECが取り込む", "出荷待ち", "与信上限に達していないとき。達したときの返却は図にない。"],
    ["L04", "基幹 → EC-Rider", "在庫があるとき", "ピック番号（納期回答）", "ECが取り込む", "出荷指示済み", "与信回答と同時になることがある。取込時に納期回答メール。"],
    ["L05", "基幹 → EC-Rider", "出荷したとき", "出荷実績", "ECが取り込む", "出荷済み", "注文者へ出荷完了メール。店舗宛てにも出荷完了メール。"],
    ["L06", "EC-Rider → 基幹", "キャンセル依頼が、すでに基幹連携済みのとき", "キャンセル依頼", "基幹がキャンセルできるか判定する", "キャンセル依頼中のまま結果待ち", "単位は配送（基幹オーダー）。基幹連携前はこの連携をしない。"],
    ["L07", "基幹 → EC-Rider", "キャンセル処理のあと", "キャンセルできるか、できないか。その後キャンセルデータ", "ECが取り込む", "キャンセル完了", "不能のときの遷移は図にない。説明文は、運営が行うキャンセルも基幹からデータが返る、と書いている。その経路は図では運営キャンセルから完了へ直結している。"],
    ["L08", "基幹 → EC-Rider", "返品処理のあと", "返品実績", "ECが取り込む", "返品完了", "返品依頼そのものはEC側の操作。基幹へ依頼データを渡す線は図にない。"],
    ["L09", "WMS", "—", "—", "—", "—", "列の見出しだけがある。箱も線もない。"],
]

MAILS = [
    ["M01", "注文承認依頼メール", "使う", "承認待ち", "承認待ちになったとき", "店舗の承認者", "顧客レーンの箱に「※店舗の承認者に送信」とある。", "リクエスト注文の承認依頼と同じタイミングとは限らない。メールの線は承認待ちから出ている。"],
    ["M02", "注文受付メール", "使う", "受注待ち", "受注待ちに紐づく。図は3通を1本の線にまとめている", "図は顧客レーン。宛先の文言は説明文にない", "受注待ちから封筒マークへ1本。", "注文を受け付けて受注待ちに入った連絡、と名前から読める。"],
    ["M03", "注文承認完了メール", "使う", "受注待ち", "同上。承認を経て受注待ちに入ったとき、と名前から読める", "図は顧客レーン。宛先の文言は説明文にない", "M02と同じ箱。", "単独の線はない。"],
    ["M04", "注文内容変更通知メール", "使う", "受注待ち", "同上。注文変更のとき、と名前から読める", "図は顧客レーン。宛先の文言は説明文にない", "M02と同じ箱。", "注文変更をしてもステータスは受注待ちのまま。"],
    ["M05", "納期回答メール", "使う", "出荷指示済み", "納期回答を取り込んで出荷指示済みになったとき", "図は顧客レーン。宛先の文言は説明文にない", "出荷指示済みから。枠が赤。", "説明文は取込時に送信。"],
    ["M06", "出荷完了メール", "使う", "出荷済み", "出荷実績を取り込んで出荷済みになったとき", "注文者（説明文）", "出荷済みから顧客レーン。", ""],
    ["M07", "出荷完了メール（店舗宛て）", "使う", "出荷済み", "同上", "店舗", "出荷済みから運営レーン。", "注文者向けとは別箱。"],
    ["M08", "返品依頼メール", "使う", "返品依頼中", "返品依頼をして返品依頼中になったとき", "注文者（説明文）", "顧客レーンと運営レーンの両方に同じ箱。どちらも返品依頼中から線が出ている。", "依頼したのが顧客でも運営でも、説明文の宛先は注文者。"],
    ["M09", "返品完了メール", "使う", "返品完了", "返品実績を取り込んで返品完了になったとき", "注文者（説明文）", "返品完了から顧客レーン。", ""],
    ["M10", "注文キャンセル依頼メール", "使う", "キャンセル依頼中", "キャンセル依頼をしてキャンセル依頼中になったとき", "注文者（説明文）", "キャンセル依頼中から顧客レーン。", ""],
    ["M11", "注文キャンセル完了メール", "使う", "キャンセル完了", "キャンセル完了になったとき", "注文者（説明文）", "キャンセル完了から顧客レーン。", "運営の直接キャンセルでも、依頼経由の完了でも、着地は同じステータス。"],
    ["M12", "入金確認メール", "利用しない", "出荷待ち", "図の線は出荷待ちから出ている", "図は顧客レーン", "白抜きの入金フローの近く。", "入金確認の楕円からメールへの線ではない。"],
]

GAPS = [
    ["図と説明の違い", "運営のキャンセル", "図は、承認待ち・受注待ち・出荷待ち・出荷指示済みから運営の「キャンセル」へ集まり、基幹を経由せずキャンセル完了へ進む。説明文は、基幹からキャンセルデータが連携され、ECが取り込むとキャンセル完了になる、と書いている。"],
    ["図と説明の違い", "基幹の楕円", "基幹レーンの楕円に「与信回答」「納期回答」「注文変更」が1つに入っている。矢印の先は出荷待ちだけ。納期回答で出荷指示済みになること、注文変更は受注待ちの運営操作であることは、別の線と説明文の側に書かれている。"],
    ["図にない", "与信が上限に達したとき", "上限内なら出荷待ち、とだけある。達したあとのステータスは書かれていない。"],
    ["図にない", "キャンセル不能", "基幹がキャンセル不能を返したあとのステータスは書かれていない。"],
    ["図にない", "返品不可", "返品できる場合に基幹が処理する、とある。できない場合のステータスは書かれていない。"],
    ["図にない", "WMS", "見出しだけがある。業務も連携も描かれていない。出荷準備の担当も図にはない。"],
    ["図にない", "出荷指示", "説明欄に「出荷指示（利用しない）」の見出しだけがある。本文も箱もない。使うステータスは出荷指示済み。"],
    ["線が欠けている", "代理注文の赤線", "承認権限があるときの受注待ち行きは、顧客の新規注文には接続がある。代理注文側の赤線は、同じ位置にあるが図形への接続情報が残っていない。説明文の承認権限ルールと合わせて読んだ。"],
    ["線が欠けている", "EDI新規注文の赤線", "始点はEDIの新規注文につながっている。終点は図形に未接続で、縦の範囲が承認待ちから受注待ちにかかっている。承認待ちへの別線はない。"],
    ["線が欠けている", "承認依頼の緑線", "承認依頼から承認待ちへは届いている。反対側の端は図形に未接続で、承認待ちと受注待ちの境目にある。ステータスが進む線としては描かれていない。"],
]


def build():
    wb = Workbook()

    # ----- 読み方 -----
    ws = wb.active
    ws.title = "読み方"
    setup_page(ws, "読み方", 1)
    ws.sheet_properties.tabColor = NAVY
    set_widths(ws, [22, 88])
    write_title(
        ws,
        "注文ステータスと業務フロー",
        "元の図は、顧客・ステータス・運営・EDI・基幹・WMS を横に並べ、そのあいだを線でつないでいる。ここではその線を追わなくて済むように、ステータス遷移・業務・システム連携・メールを分けて書いてある。",
        2,
    )
    sections = [
        ("先に見るシート", "「通常の流れ」。ステータスごとに、業務・連携・メール・次の行き先を1行にまとめてある。"),
        ("ステータス遷移", "入口、通常進行、ステータスが変わらない操作、返品、キャンセル、差戻し、利用しない経路。全32本。"),
        ("業務", "顧客、運営、EDI、基幹が、どのステータスで何をするか。"),
        ("システム連携", "EDI・基幹・EC-Rider のあいだで、何がどちらへ渡り、取り込んだあとのステータスは何か。"),
        ("メール", "11通と、利用しない1通。いつ送り、宛先はどこか。図が1本の線にまとめている3通は、その旨を書いてある。"),
        ("図の色", "黄い四角が使うステータス。楕円が操作。封筒がメール。赤線は承認待ちを飛ばして受注待ちへ入る経路。緑線は差戻しと、リクエスト注文の承認依頼。白抜きは利用しない。"),
        ("単位", "キャンセルは配送単位（基幹オーダー単位）。返品依頼は注文明細単位。"),
        ("使わないもの", "入金待ち、入金、入金確認、入金確認メール。説明欄の「出荷指示（利用しない）」は見出しだけで、箱はない。WMS列も見出しだけ。"),
    ]
    ws.cell(4, 1, "項目").font = font(11, True, WHITE)
    ws.cell(4, 2, "内容").font = font(11, True, WHITE)
    apply_header(ws, 4, 2)
    ws.row_dimensions[4].height = 22
    ws.auto_filter.ref = "A4:B12"
    ws.auto_filter.ref = "A4:B4"
    for i, (k, v) in enumerate(sections):
        r = 5 + i
        ws.cell(r, 1, k).font = font(11, True, NAVY)
        ws.cell(r, 1).fill = fill(YELLOW if i == 0 else PAPER)
        ws.cell(r, 1).alignment = wrap
        ws.cell(r, 1).border = thin
        ws.cell(r, 2, v).font = font(11)
        ws.cell(r, 2).alignment = wrap
        ws.cell(r, 2).border = thin
        ws.cell(r, 2).fill = fill(WHITE)
        ws.row_dimensions[r].height = 36
    ws.row_dimensions[5].height = 32
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = None
    ws.print_title_rows = "1:4"
    ws.page_setup.orientation = "portrait"
    ws.oddHeader.center.text = ""

    note_row = 14
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=2)
    ws.cell(note_row, 1, "元の図からは読み取れないこと").font = font(12, True, NAVY)
    notes = [
        "与信が上限に達したあとのステータス。",
        "基幹がキャンセル不能と返したあとのステータス。",
        "返品できないときのステータス。",
        "出荷準備をどのレーンが行うか。WMS列は空。",
    ]
    for i, text in enumerate(notes):
        r = note_row + 1 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(r, 1, text).font = font(10)
        ws.cell(r, 1).alignment = wrap
        ws.row_dimensions[r].height = 20

    # ----- 通常の流れ -----
    ws = wb.create_sheet("通常の流れ")
    setup_page(ws, "通常の流れ")
    ws.sheet_properties.tabColor = "C4A35A"
    headers = ["順", "ステータス", "意味", "業務", "システム連携", "メール", "次に進む", "ほかの行き先"]
    write_title(
        ws,
        "通常の流れ",
        "上から順に読む。キャンセル・返品・差戻しは右端の列に行き先だけ書き、詳細は「ステータス遷移」にある。入金待ちは使わないので、この表には入れてない。",
        len(headers),
    )
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, h)
    apply_header(ws, 4, len(headers))
    header_colors = [NAVY, "8A6A2F", NAVY, "2C5282", "6B3FA0", "276749", NAVY, "9B2C2C"]
    for c, color in enumerate(header_colors, 1):
        ws.cell(4, c).fill = fill(color)
    ws.row_dimensions[4].height = 22
    body_colors = [PAPER, YELLOW, WHITE, BLUE, PURPLE, GREEN, WHITE, RED]
    for i, row in enumerate(NORMAL):
        r = 5 + i
        for c, value in enumerate(row, 1):
            cell = ws.cell(r, c, value)
            cell.font = font(9, c == 2)
            cell.alignment = wrap
            cell.border = thin
            cell.fill = fill(body_colors[c - 1] if c != 1 else (GRAY if i % 2 else PAPER))
        ws.row_dimensions[r].height = 78
    set_widths(ws, [5, 16, 28, 36, 36, 32, 32, 32])
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = "A4:H14"
    ws.print_title_rows = "1:4"
    ws.auto_filter.ref = "A4:H14"
    ws.page_setup.fitToWidth = 1
    ws.sheet_view.zoomScale = 100
    ws.oddHeader.left.text = "EC-Rider　通常の流れ"

    # ----- ステータス遷移 -----
    ws = wb.create_sheet("ステータス遷移")
    setup_page(ws, "ステータス遷移")
    ws.sheet_properties.tabColor = "8A6A2F"
    headers = ["ID", "種類", "遷移前", "誰が", "操作", "条件", "遷移後", "単位", "図の線", "関係するメール"]
    write_title(
        ws,
        "ステータス遷移",
        "線を1本ずつ表にした。種類で絞り込める。赤線は承認待ちを飛ばす入口。緑線は差戻しと、リクエスト注文の承認依頼。",
        len(headers),
    )
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, h)
    apply_header(ws, 4, len(headers))
    ws.row_dimensions[4].height = 22
    kind_fill = {
        "入口": "FDE8E8",
        "通常": GREEN,
        "滞留": YELLOW,
        "返品": BLUE,
        "キャンセル": RED,
        "差戻し": AMBER,
        "利用しない": GRAY,
    }
    for i, row in enumerate(TRANSITIONS):
        r = 5 + i
        for c, value in enumerate(row, 1):
            cell = ws.cell(r, c, value)
            cell.font = font(9, c in (3, 7))
            cell.alignment = wrap
            cell.border = thin
            cell.fill = fill(kind_fill.get(row[1], WHITE))
        ws.row_dimensions[r].height = 36
    set_widths(ws, [6, 12, 16, 12, 16, 36, 16, 18, 36, 32])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:J{4 + len(TRANSITIONS)}"
    ws.print_title_rows = "1:4"
    ws.auto_filter.ref = f"A4:J{4 + len(TRANSITIONS)}"

    # ----- 業務 -----
    ws = wb.create_sheet("業務")
    setup_page(ws, "業務")
    ws.sheet_properties.tabColor = "2C5282"
    headers = ["誰が", "業務", "できるステータス", "したあとのステータス", "単位", "補足"]
    write_title(
        ws,
        "業務",
        "誰が、どのステータスで、何をするか。データの向きは「システム連携」に分けてある。",
        len(headers),
    )
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, h)
    apply_header(ws, 4, len(headers), "2C5282")
    ws.row_dimensions[4].height = 22
    who_fill = {"顧客": "E8F6EF", "運営": BLUE, "EDI": PURPLE, "基幹": "F8E8E8"}
    for i, row in enumerate(WORK):
        r = 5 + i
        for c, value in enumerate(row, 1):
            cell = ws.cell(r, c, value)
            cell.font = font(10, c == 1)
            cell.alignment = wrap
            cell.border = thin
            cell.fill = fill(who_fill.get(row[0], WHITE))
        ws.row_dimensions[r].height = 32
    set_widths(ws, [10, 24, 40, 28, 22, 55])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:F{4 + len(WORK)}"
    ws.print_title_rows = "1:4"

    # ----- システム連携 -----
    ws = wb.create_sheet("システム連携")
    setup_page(ws, "システム連携")
    ws.sheet_properties.tabColor = "6B3FA0"
    headers = ["ID", "向き", "いつ", "渡すもの", "受けた側の処理", "その結果のステータス", "補足"]
    write_title(
        ws,
        "システム連携",
        "EC-Rider、EDI、基幹のあいだの受け渡し。WMSは見出しだけで、連携は描かれていない。",
        len(headers),
    )
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, h)
    apply_header(ws, 4, len(headers), "6B3FA0")
    ws.row_dimensions[4].height = 22
    for i, row in enumerate(LINKS):
        r = 5 + i
        for c, value in enumerate(row, 1):
            cell = ws.cell(r, c, value)
            cell.font = font(10, c == 2)
            cell.alignment = wrap
            cell.border = thin
            cell.fill = fill(GRAY if row[1] == "WMS" else PURPLE)
        ws.row_dimensions[r].height = 40
    set_widths(ws, [6, 22, 36, 28, 28, 28, 48])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:G{4 + len(LINKS)}"
    ws.print_title_rows = "1:4"

    # ----- メール -----
    ws = wb.create_sheet("メール")
    setup_page(ws, "メール")
    ws.sheet_properties.tabColor = "276749"
    headers = ["ID", "メール", "使うか", "紐づくステータス", "いつ", "宛先", "図での結び方", "補足"]
    write_title(
        ws,
        "メール",
        "封筒マークへ向かう線を、1通ずつにした。受注待ちの3通は、元の図では1つの箱にまとまっている。",
        len(headers),
    )
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, h)
    apply_header(ws, 4, len(headers), "276749")
    ws.row_dimensions[4].height = 22
    for i, row in enumerate(MAILS):
        r = 5 + i
        for c, value in enumerate(row, 1):
            cell = ws.cell(r, c, value)
            cell.font = font(10, c == 2)
            cell.alignment = wrap
            cell.border = thin
            cell.fill = fill(GRAY if row[2] != "使う" else GREEN)
        ws.row_dimensions[r].height = 36
    set_widths(ws, [6, 28, 12, 18, 36, 32, 36, 40])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:H{4 + len(MAILS)}"
    ws.print_title_rows = "1:4"

    # ----- 図から読み取れないこと -----
    ws = wb.create_sheet("図の読み取り")
    setup_page(ws, "図の読み取り")
    ws.sheet_properties.tabColor = "9B2C2C"
    headers = ["区分", "対象", "内容"]
    write_title(
        ws,
        "図の読み取り",
        "線が途中で切れている箇所と、説明文と図で経路の書き方が違う箇所。表の本文は、図に接続がある線を優先し、ここだけ補足してある。",
        len(headers),
    )
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, h)
    apply_header(ws, 4, len(headers), "9B2C2C")
    ws.row_dimensions[4].height = 22
    gap_fill = {
        "図と説明の違い": AMBER,
        "図にない": RED,
        "線が欠けている": YELLOW,
    }
    for i, row in enumerate(GAPS):
        r = 5 + i
        for c, value in enumerate(row, 1):
            cell = ws.cell(r, c, value)
            cell.font = font(10, c == 1)
            cell.alignment = wrap
            cell.border = thin
            cell.fill = fill(gap_fill.get(row[0], WHITE))
        ws.row_dimensions[r].height = 48
    set_widths(ws, [18, 28, 100])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:C{4 + len(GAPS)}"
    ws.print_title_rows = "1:4"
    ws.page_setup.orientation = "landscape"

    # named range style cleanup: remove unused imports side effects
    wb.properties.title = "注文ステータスと業務フロー"
    wb.properties.subject = "ステータス遷移、業務、システム連携、メール"
    wb.active = wb["通常の流れ"]

    wb.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
