#!/usr/bin/env python3
"""1シートのフロー図。色は ステータス / ECアクション / メール / 連携。"""

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/docs/order-status-flow/注文ステータスと業務フロー.xlsx"

INK = "1F2933"
NAVY = "1B3A4B"
FONT = "Yu Gothic"

KIND = {
    "status": ("FFF4CC", "8A6A2F", "6B5420"),
    "action": ("E7F0FA", "2C5282", "234E78"),
    "mail": ("E5F6EC", "276749", "1F6B45"),
    "link": ("F3EAF8", "6B3FA0", "5B2D86"),
    "back": ("FFF1DE", "C05621", "9A5416"),
    "skip": ("FDECEC", "C53030", "9B2C2C"),
    "off": ("F3F4F6", "9CA3AF", "4B5563"),
    "head": ("F7F5F2", "E4E0D8", "52606D"),
}


def fill(h):
    return PatternFill("solid", fgColor=h)


def font(size=10, bold=False, color=INK):
    return Font(name=FONT, size=size, bold=bold, color=color)


def border(color, style="medium"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)


class Sheet:
    def __init__(self):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "フロー"
        self.used = set()

    def take(self, r1, c1, r2, c2):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) in self.used:
                    raise RuntimeError(f"重なり 行{r} 列{c}")
                self.used.add((r, c))

    def put(self, r1, c1, r2, c2, text, kind, size=10, bold=True):
        bg, line, fg = KIND[kind]
        self.take(r1, c1, r2, c2)
        if not (r1 == r2 and c1 == c2):
            self.ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        align = Alignment(wrap_text=True, vertical="center", horizontal="center")
        b = border(line)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.ws.cell(r, c)
                cell.fill = fill(bg)
                cell.border = b
                cell.alignment = align
        cell = self.ws.cell(r1, c1, text)
        cell.font = font(size, bold, fg)
        cell.alignment = align
        return cell

    def arrow(self, r, c1, c2, text, color, size=11):
        self.take(r, c1, r, c2)
        if c1 != c2:
            self.ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        align = Alignment(wrap_text=True, vertical="center", horizontal="center")
        for c in range(c1, c2 + 1):
            cell = self.ws.cell(r, c)
            cell.alignment = align
            cell.font = font(size, True, color)
        self.ws.cell(r, c1, text)


def build():
    s = Sheet()
    ws = s.ws

    # A余白 B-Dメール E矢印 F-H顧客 I矢印 J-Kステータス L矢印 M-Oキャンセル P矢印 Q-S運営 T矢印 U-W連携
    widths = {
        1: 2,
        2: 13, 3: 13, 4: 13,
        5: 3.2,
        6: 13, 7: 13, 8: 13,
        9: 3.2,
        10: 16, 11: 16,
        12: 4.5,
        13: 13, 14: 13, 15: 13,
        16: 3.6,
        17: 13, 18: 13, 19: 13,
        20: 3.2,
        21: 14, 22: 14, 23: 14,
    }
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    ws.merge_cells("A1:W1")
    t = ws.cell(1, 1, "注文ステータスと業務フロー")
    t.font = font(20, True, NAVY)
    t.alignment = Alignment(vertical="center", horizontal="left")
    ws.row_dimensions[1].height = 30

    # 色の意味は1行。レーンの見出しと重ならないようにする
    ws.merge_cells("A2:W2")
    def piece(text, color, size=12, bold=True):
        return TextBlock(InlineFont(rFont=FONT, sz=size, b=bold, color=color), text)

    ws.cell(2, 1).value = CellRichText(
        piece("色　　", "52606D", 11, False),
        piece("ステータス", "8A6A2F"),
        piece("　　", "52606D", 11, False),
        piece("EC上のアクション", "2C5282"),
        piece("　　", "52606D", 11, False),
        piece("メール", "276749"),
        piece("　　", "52606D", 11, False),
        piece("システム連携", "6B3FA0"),
        piece("　　", "52606D", 11, False),
        piece("差戻し", "C05621"),
        piece("　　", "52606D", 11, False),
        piece("承認待ちを飛ばす", "C53030"),
    )
    ws.cell(2, 1).alignment = Alignment(vertical="center", horizontal="left")
    for c in range(1, 24):
        ws.cell(2, c).fill = fill("F7F5F2")
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 8

    lanes = [
        (2, 4, "メール", "mail"),
        (6, 8, "ECアクション（顧客）", "action"),
        (10, 11, "ステータス", "status"),
        (13, 15, "キャンセル", "head"),
        (17, 19, "ECアクション（運営）", "action"),
        (21, 23, "システム連携", "link"),
    ]
    for c1, c2, text, kind in lanes:
        s.put(4, c1, 4, c2, text, kind, size=10)
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 8

    # 入口
    s.put(6, 6, 8, 8, "新規注文\nフロントサイト", "action")
    s.put(6, 17, 8, 19, "代理注文\n管理サイト\n権限は使ったユーザに従う", "action", size=9)
    s.put(6, 21, 8, 23, "EDI → EC\n新規注文", "link")
    s.arrow(9, 7, 7, "↓", "234E78")
    s.arrow(9, 18, 18, "↓", "234E78")
    s.arrow(9, 22, 22, "↓", "5B2D86")
    ws.row_dimensions[9].height = 16

    s.put(10, 10, 12, 11, "承認権限は？", "status", size=13)
    s.put(10, 21, 12, 23, "権限あり・EDI\n承認待ちを飛ばす", "skip", size=10)
    s.arrow(13, 10, 11, "権限なし  ↓", "6B5420", size=10)
    s.arrow(13, 21, 23, "権限あり・EDI  ↓", "C53030", size=10)
    ws.row_dimensions[13].height = 20

    # 承認待ち
    s.put(14, 10, 18, 11, "承認待ち", "status", size=16)
    s.put(14, 2, 18, 4, "注文承認依頼メール\n店舗の承認者へ", "mail")
    s.put(14, 6, 18, 8, "注文承認\n同じ販売組織の承認者", "action")
    s.arrow(16, 9, 9, "→", "234E78", size=14)
    s.put(14, 17, 15, 19, "代理承認", "action")
    s.arrow(14, 16, 16, "←", "234E78", size=14)
    s.put(16, 17, 18, 19, "↺ 承認依頼\nリクエスト注文のみ\nステータスはそのまま", "action", size=8)
    s.put(14, 21, 18, 23, "権限あり・EDIは\nここを飛ばして\n下の受注待ちへ", "skip", size=10)
    s.put(16, 13, 18, 15, "キャンセル依頼できる\n↓", "action", size=9)
    s.arrow(17, 12, 12, "→", "C53030", size=12)

    s.arrow(19, 10, 11, "↓   承認・代理承認", "6B5420", size=11)
    s.arrow(19, 21, 23, "↓   受注待ちへ直行", "C53030", size=10)
    s.arrow(19, 14, 14, "↓", "C53030", size=12)
    ws.row_dimensions[19].height = 22

    # 受注待ち
    s.put(20, 10, 28, 11, "受注待ち", "status", size=16)
    s.put(20, 2, 22, 4, "注文受付メール", "mail")
    s.put(23, 2, 25, 4, "注文承認完了メール", "mail")
    s.put(26, 2, 28, 4, "注文内容変更通知メール", "mail")
    s.put(20, 17, 24, 19, "↺ 注文変更\n金額を変えられるのはここだけ\nステータスはそのまま", "action", size=9)
    s.arrow(22, 16, 16, "←", "234E78", size=14)
    s.put(20, 21, 23, 23, "EC → 基幹\n30分ごとに注文連携\n受注扱い", "link", size=9)
    s.put(25, 21, 28, 23, "基幹 → EC\n与信回答\n上限内なら基幹オーダー番号", "link", size=9)
    s.put(24, 13, 28, 15, "承認待ち・受注待ち・\n出荷待ちから\n顧客がキャンセル依頼\n↓", "action", size=9)
    s.arrow(26, 12, 12, "→", "C53030", size=12)

    s.arrow(29, 10, 11, "↓   与信回答", "5B2D86", size=11)
    s.put(29, 12, 29, 12, "↑\n差戻し", "back", size=8)
    s.arrow(29, 14, 14, "↓", "C53030", size=12)
    ws.row_dimensions[29].height = 26

    # 出荷待ち
    s.put(30, 10, 34, 11, "出荷待ち", "status", size=16)
    s.put(30, 21, 34, 23, "基幹 → EC\n納期回答\n在庫があればピック番号\n与信と同時のこともある", "link", size=9)
    s.put(30, 13, 32, 15, "キャンセル依頼中", "status", size=14)
    s.put(33, 13, 34, 15, "注文キャンセル依頼メール\n注文者へ", "mail", size=9)
    s.arrow(31, 12, 12, "→", "C53030", size=14)
    s.put(33, 12, 34, 12, "←\n差戻し", "back", size=8)

    s.arrow(35, 10, 11, "↓   納期回答", "5B2D86", size=11)
    s.put(35, 12, 35, 12, "↑\n差戻し", "back", size=8)
    s.arrow(35, 13, 15, "連携前は、このまま完了へ  ↓", "C53030", size=9)
    ws.row_dimensions[35].height = 26

    # 出荷指示済み
    s.put(36, 10, 41, 11, "出荷指示済み", "status", size=16)
    s.put(36, 2, 38, 4, "納期回答メール", "mail")
    s.put(36, 17, 38, 19, "出荷準備", "action")
    s.put(39, 17, 41, 19, "キャンセル\n承認待ちからここまで\n依頼中は経由しない", "action", size=8)
    s.arrow(40, 16, 16, "←", "234E78", size=14)
    s.put(36, 21, 38, 23, "基幹 → EC\nキャンセル\n連携済みの依頼だけ", "link", size=9)
    s.put(39, 21, 41, 23, "基幹 → EC\n出荷実績", "link")
    s.put(36, 13, 38, 15, "連携後\n基幹がキャンセル", "link", size=10)
    s.put(39, 13, 41, 15, "キャンセル完了", "status", size=14)

    s.arrow(42, 10, 11, "↓   出荷", "5B2D86", size=11)
    s.put(42, 13, 42, 15, "注文キャンセル完了メール　注文者へ", "mail", size=8)
    ws.row_dimensions[42].height = 20

    # 出荷済み
    s.put(43, 10, 48, 11, "出荷済み", "status", size=16)
    s.put(43, 2, 45, 4, "出荷完了メール\n注文者へ", "mail")
    s.put(46, 2, 48, 4, "出荷完了メール\n店舗宛て", "mail")
    s.put(43, 6, 48, 8, "返品依頼\n注文明細ごと", "action")
    s.arrow(45, 9, 9, "→", "234E78", size=14)
    s.put(43, 17, 48, 19, "返品依頼\n注文明細ごと", "action")
    s.arrow(45, 16, 16, "←", "234E78", size=14)

    s.arrow(49, 10, 11, "↓   返品依頼", "234E78", size=11)
    s.put(49, 12, 49, 12, "↑\n差戻し", "back", size=8)
    ws.row_dimensions[49].height = 26

    # 返品
    s.put(50, 10, 54, 11, "返品依頼中", "status", size=16)
    s.put(50, 2, 54, 4, "返品依頼メール\n注文者へ", "mail")
    s.put(50, 21, 54, 23, "基幹 → EC\n返品実績\n本社へ返送され、返品できるとき", "link", size=9)
    s.arrow(55, 10, 11, "↓   返品実績", "5B2D86", size=11)
    ws.row_dimensions[55].height = 20

    s.put(56, 10, 59, 11, "返品完了", "status", size=16)
    s.put(56, 2, 59, 4, "返品完了メール\n注文者へ", "mail")

    ws.row_dimensions[60].height = 10
    s.put(
        61, 2, 62, 23,
        "利用しない　　入金待ち　　入金　　入金確認　　入金確認メール　　出荷指示　　WMSの列は空",
        "off",
        size=10,
        bold=False,
    )

    for r in range(6, 63):
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = 18

    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(0.3, 0.3, 0.4, 0.3, 0.15, 0.15)
    ws.print_area = "A1:W62"
    ws.sheet_properties.tabColor = "8A6A2F"
    ws.page_setup.horizontalCentered = True
    ws.oddFooter.left.text = "EC-Rider　注文の流れ"
    s.wb.properties.title = "注文ステータスと業務フロー"
    s.wb.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
