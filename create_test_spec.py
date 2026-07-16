#!/usr/bin/env python3
"""Primo dataLayer テスト仕様書 Excel生成"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- スタイル ---
HEADER_FILL = PatternFill("solid", fgColor="1A3A5C")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
SUBHEADER_FILL = PatternFill("solid", fgColor="2E86AB")
SUBHEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
SECTION_FILL = PatternFill("solid", fgColor="E8F4F8")
SECTION_FONT = Font(bold=True, size=10)
NORMAL = Font(size=10)
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

SITE = "primo-demo2-front.ec-rider2-demo.net"
GTM_ID = "GTM-TGF9D8WX"

# テストケース定義
TEST_CASES = [
    # (ID, カテゴリ, テスト項目, 対象, 優先度)
    ("TC-001", "GTM基本設定", "GTMスニペット（head）の設置", "GTM head", "高"),
    ("TC-002", "GTM基本設定", "GTMスニペット（body noscript）の設置", "GTM body", "高"),
    ("TC-003", "GTM基本設定", "User-ID初期化がGTMスニペットより前に実行される", "実装順序", "高"),
    ("TC-004", "GTM基本設定", "dataLayer初期化処理", "dataLayer init", "高"),
    ("TC-010", "共通ページ計測", "ログイン時：user_id/area_id/staff_id/member_r送信", "共通push", "高"),
    ("TC-011", "共通ページ計測", "未ログイン時：user_idキーを出力しない", "共通push", "高"),
    ("TC-012", "共通ページ計測", "member_rの数値マッピング（0〜3）", "member_r", "中"),
    ("TC-013", "共通ページ計測", "ページリロード時のdataLayer再出力", "共通push", "中"),
    ("TC-020", "認証", "ログイン成功時にloginイベント送信", "login", "高"),
    ("TC-021", "認証", "ログイン失敗時はloginイベントを送信しない", "login", "高"),
    ("TC-022", "認証", "会員登録成功時にsign_upイベント送信", "sign_up", "高"),
    ("TC-030", "リード", "見積書DL時にquote_downloadイベント送信", "quote_download", "高"),
    ("TC-031", "リード", "見積依頼成功時：generate_lead（lead_source=quote）+items", "generate_lead", "高"),
    ("TC-032", "リード", "問い合わせ成功時：generate_lead（lead_source=inquiry）", "generate_lead", "高"),
    ("TC-033", "リード", "再見積（re_quote）時はgenerate_leadを送信しない", "generate_lead", "高"),
    ("TC-040", "商品・カート", "商品詳細表示時：view_item + stock", "view_item", "高"),
    ("TC-041", "商品・カート", "カート追加成功時：add_to_cart（add_method=normal）", "add_to_cart", "高"),
    ("TC-042", "商品・カート", "add_to_cart：add_method=quick_order", "add_to_cart", "中"),
    ("TC-043", "商品・カート", "add_to_cart：add_method=order_history", "add_to_cart", "中"),
    ("TC-044", "商品・カート", "add_to_cart：add_method=quote_history", "add_to_cart", "中"),
    ("TC-045", "商品・カート", "カート削除時：remove_from_cart", "remove_from_cart", "高"),
    ("TC-046", "商品・カート", "お気に入り追加時：add_to_wishlist", "add_to_wishlist", "高"),
    ("TC-047", "商品・カート", "ecommerceイベント前にecommerce:nullをpush", "ecommerce clear", "高"),
    ("TC-050", "購入フロー", "注文プロセス開始：begin_checkout（cart_type=normal）", "begin_checkout", "高"),
    ("TC-051", "購入フロー", "begin_checkout（cart_type=quote）", "begin_checkout", "高"),
    ("TC-052", "購入フロー", "request_item_countの値（あり/なし=0）", "begin_checkout", "中"),
    ("TC-060", "購入フロー", "add_shipping_infoとadd_payment_infoの同時送信", "add_shipping_info/add_payment_info", "高"),
    ("TC-061", "購入フロー", "payment_typeの値（4種）", "add_payment_info", "中"),
    ("TC-062", "購入フロー", "クーポンコード（ecommerce.coupon）の送信", "add_shipping_info/add_payment_info", "中"),
    ("TC-070", "注文確定", "承認申請時：approval_request（order_id）", "approval_request", "高"),
    ("TC-071", "注文確定", "注文確定時：purchase（承認なし）", "purchase", "高"),
    ("TC-072", "注文確定", "承認あり：承認完了時にpurchase送信", "purchase", "高"),
    ("TC-073", "注文確定", "決済エラー時はpurchaseを送信しない", "purchase", "高"),
    ("TC-074", "注文確定", "customer_type（new/returning/guest）", "purchase", "高"),
    ("TC-075", "注文確定", "order_id = transaction_id = approval_request.order_id", "purchase/approval_request", "高"),
    ("TC-080", "その他", "mailmagazine_signup（all/text_only/none）", "mailmagazine_signup", "中"),
    ("TC-081", "その他", "一括注文フォーマットDL：bulk_order_download", "bulk_order_download", "高"),
    ("TC-082", "その他", "一括注文アップロード成功：bulk_order_upload + add_to_cart複数", "bulk_order_upload", "高"),
    ("TC-083", "その他", "一括注文アップロード失敗：bulk_order_uploadのみ（add_to_cartなし）", "bulk_order_upload", "高"),
    ("TC-084", "その他", "error_typeの複数値（カンマ区切り）", "bulk_order_upload", "中"),
    ("TC-090", "items配列", "items共通フィールドの存在・型", "items", "高"),
    ("TC-091", "items配列", "item_variantなし時はキーごと省略", "items", "中"),
    ("TC-092", "items配列", "サプライヤー異なる商品は別要素として配列に含む", "items", "中"),
]

# 詳細定義
DETAILS = {
    "TC-001": {
        "前提": f"対象サイト（{SITE}）の任意ページにアクセスできること",
        "手順": "1. ブラウザで任意ページを開く\n2. ページソースまたはDevTools Elementsで<head>内を確認",
        "コマンド": "document.querySelector('script[src*=\"googletagmanager.com/gtm.js\"]')",
        "期待": f"null以外の要素が返る。src属性に「{GTM_ID}」が含まれる。\n例: HTMLScriptElement {{ src: \"https://www.googletagmanager.com/gtm.js?id={GTM_ID}\" }}",
    },
    "TC-002": {
        "前提": "同上",
        "手順": "1. ページソースまたはElementsで<body>直後を確認",
        "コマンド": "document.querySelector('noscript iframe[src*=\"googletagmanager.com/ns.html\"]')",
        "期待": f"null以外のiframe要素が返る。srcに「{GTM_ID}」が含まれる。",
    },
    "TC-003": {
        "前提": "ログイン済み状態",
        "手順": "1. ページソースの<head>内でスクリプトの記述順序を確認\n2. user_id初期化のdataLayer.pushがGTMスニペットより前にあることを確認",
        "コマンド": "/* SourcesタブでHTMLを確認。または */\n[...document.querySelectorAll('head script')].map(s => s.textContent.slice(0,80))",
        "期待": "head内で「window.dataLayer = window.dataLayer || []」およびuser_idのpushが、GTMスニペット（googletagmanager.com/gtm.js）より前に記述されていること。",
    },
    "TC-004": {
        "前提": "任意ページを初回読み込み",
        "手順": "1. ページ読み込み直後にConsoleを開く",
        "コマンド": "window.dataLayer",
        "期待": "配列が存在する（undefinedでない）。\n例: [{gtm.start: ..., event: 'gtm.js'}, ...]",
    },
    "TC-010": {
        "前提": "取引先・担当者でログイン済み",
        "手順": "1. 任意ページを読み込み\n2. ConsoleでdataLayerを確認",
        "コマンド": "dataLayer.find(d => d.user_id)",
        "期待": "オブジェクトが返り、以下のキーが存在する:\n  user_id: '取引先ID'（文字列）\n  area_id: 'エリアID'（文字列）\n  staff_id: '担当者個人ID'（文字列）\n  member_r: 0|1|2|3（数値）\n\n例: {user_id: 'C00001', area_id: 'A01', staff_id: 'S00001', member_r: 0}",
    },
    "TC-011": {
        "前提": "未ログイン（ゲスト）状態",
        "手順": "1. ログアウトまたはシークレットウィンドウで任意ページを開く\n2. dataLayerを確認",
        "コマンド": "dataLayer.filter(d => 'user_id' in d)",
        "期待": "空配列 [] が返る（user_idキーを持つオブジェクトが存在しない）。",
    },
    "TC-012": {
        "前提": "会員ランクが異なる取引先のテストアカウントを用意",
        "手順": "1. 標準(0)/シルバー(1)/ゴールド(2)/その他(3)の各アカウントでログイン\n2. それぞれdataLayerを確認",
        "コマンド": "dataLayer.find(d => d.member_r !== undefined).member_r",
        "期待": "ランクに応じた数値が返る:\n  標準=0, シルバー=1, ゴールド=2, その他=3\n※文字列ではなく数値型であること（typeof === 'number'）",
    },
    "TC-013": {
        "前提": "ログイン済み",
        "手順": "1. ページ読み込み後にdataLayer件数を記録\n2. ページをリロード（F5）\n3. 再度件数を確認",
        "コマンド": "dataLayer.length  // リロード前後で実行",
        "期待": "リロード後はdataLayerがリセットされ、再度pushされる（リロードのたびに共通pushが出力される）。",
    },
    "TC-020": {
        "前提": "未ログイン。有効なテストアカウントあり",
        "手順": "1. ログイン画面で正しいID/PWを入力してログイン\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'login')",
        "期待": "1件以上の配列が返る。\n例: [{event: 'login'}]\n※eventキーのみ。他の余計なキーがないこと。",
    },
    "TC-021": {
        "前提": "未ログイン",
        "手順": "1. 意図的に誤ったID/PWでログイン試行\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'login')",
        "期待": "[] （空配列）。loginイベントはpushされない。",
    },
    "TC-022": {
        "前提": "会員登録フォームにアクセス可能",
        "手順": "1. 会員登録フォームに必要事項を入力して送信\n2. サーバー側で登録成功後、Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'sign_up')",
        "期待": "例: [{event: 'sign_up'}]",
    },
    "TC-030": {
        "前提": "見積書がDL可能な状態（見積依頼済み）",
        "手順": "1. 見積書DLリンクをクリック\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'quote_download')",
        "期待": "例: [{event: 'quote_download'}]",
    },
    "TC-031": {
        "前提": "ログイン済み。見積依頼フォームにアクセス可能",
        "手順": "1. 商品を選択して見積依頼フォームを送信（新規見積。再見積ではない）\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'generate_lead').map(d => ({event:d.event, lead_source:d.lead_source, items:d.ecommerce?.items}))",
        "期待": "例: [{event: 'generate_lead', lead_source: 'quote', items: [{item_id:'...', item_name:'...', item_brand:'...', item_category1:'...', ..., price:9000, quantity:1}]}]\n※lead_source='quote'、ecommerce.currency='JPY'、items配列あり",
    },
    "TC-032": {
        "前提": "問い合わせフォームにアクセス可能",
        "手順": "1. 問い合わせフォームを送信して成功\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'generate_lead')",
        "期待": "例: [{event: 'generate_lead', lead_source: 'inquiry'}]\n※lead_source='inquiry'。itemsは含まれない（またはecommerceなし）。",
    },
    "TC-033": {
        "前提": "再見積が可能な既存見積あり",
        "手順": "1. 既存見積に対して再見積を依頼\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'generate_lead')",
        "期待": "再見積操作後もgenerate_leadイベントは追加されない（件数が増えない）。",
    },
    "TC-040": {
        "前提": "商品詳細ページにアクセス可能",
        "手順": "1. 商品詳細ページを表示\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'view_item').map(d => d.ecommerce.items[0])",
        "期待": "items[0]に以下が含まれる:\n  item_id, item_name, item_brand, item_category1〜5, price(数値), quantity(数値), stock(数値)\n例: {item_id:'hinban00002', ..., stock: 216, price: 9000, quantity: 1}",
    },
    "TC-041": {
        "前提": "ログイン済み。商品詳細からカート追加可能",
        "手順": "1. 商品詳細から通常操作でカートに追加\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'add_to_cart')",
        "期待": "例: [{event:'add_to_cart', add_method:'normal', ecommerce:{value:9000, currency:'JPY', items:[...]}}]\n※直前に{ecommerce:null}がpushされていること",
    },
    "TC-042": {
        "前提": "クイックオーダー機能にアクセス可能",
        "手順": "1. クイックオーダー経由でカートに追加\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'add_to_cart').pop().add_method",
        "期待": "'quick_order'",
    },
    "TC-043": {
        "前提": "購入履歴からの再注文が可能",
        "手順": "1. 購入履歴から商品をカートに追加\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'add_to_cart').pop().add_method",
        "期待": "'order_history'",
    },
    "TC-044": {
        "前提": "見積履歴からのカート追加が可能",
        "手順": "1. 見積履歴から商品をカートに追加\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'add_to_cart').pop().add_method",
        "期待": "'quote_history'",
    },
    "TC-045": {
        "前提": "カートに商品が入っている",
        "手順": "1. カート画面で商品を削除\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'remove_from_cart')",
        "期待": "例: [{event:'remove_from_cart', ecommerce:{value:9000, currency:'JPY', items:[...]}}]",
    },
    "TC-046": {
        "前提": "商品詳細ページにアクセス可能",
        "手順": "1. お気に入りボタンをクリック\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'add_to_wishlist')",
        "期待": "例: [{event:'add_to_wishlist', ecommerce:{currency:'JPY', items:[...]}}]",
    },
    "TC-047": {
        "前提": "ecommerce系イベント（view_item等）を発火させる操作",
        "手順": "1. 商品詳細表示やカート追加を実行\n2. dataLayer全体を確認",
        "コマンド": "(() => { const i = dataLayer.findIndex(d => d.event === 'view_item'); return i >= 1 ? dataLayer[i-1] : null; })()",
        "期待": "ecommerceイベントの直前の要素が {ecommerce: null} であること。",
    },
    "TC-050": {
        "前提": "通常カートに商品あり",
        "手順": "1. カートから注文プロセスへ進む（決済・配送指定画面表示）\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'begin_checkout')",
        "期待": "例: [{event:'begin_checkout', cart_type:'normal', request_item_count:0, ecommerce:{value:..., currency:'JPY', items:[...]}}]",
    },
    "TC-051": {
        "前提": "見積カートに商品あり",
        "手順": "1. 見積カートから注文プロセスへ進む\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'begin_checkout').pop().cart_type",
        "期待": "'quote'",
    },
    "TC-052": {
        "前提": "リクエスト注文を含む/含まないケース両方",
        "手順": "1-A. 通常商品のみで注文プロセス開始\n1-B. リクエスト注文商品を含めて開始\n2. それぞれ確認",
        "コマンド": "dataLayer.filter(d => d.event === 'begin_checkout').pop().request_item_count",
        "期待": "通常のみ: 0（数値）\nリクエスト注文含む: 3等の正の整数",
    },
    "TC-060": {
        "前提": "注文プロセス中。配送・決済情報を入力可能",
        "手順": "1. 配送先・決済方法を入力して送信（サーバー受理成功）\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => ['add_shipping_info','add_payment_info'].includes(d.event)).map(d => d.event)",
        "期待": "['add_shipping_info', 'add_payment_info'] の両方が含まれる（同時送信）。\n※add_shipping_infoにshipping_tierは仕様に無いがcouponあり\n※add_payment_infoにpayment_typeあり",
    },
    "TC-061": {
        "前提": "各決済方法で注文可能なテスト環境",
        "手順": "1. 自社掛け/代引き/銀行振込/全額ポイントの各決済で送信\n2. それぞれpayment_typeを確認",
        "コマンド": "dataLayer.filter(d => d.event === 'add_payment_info').pop().payment_type",
        "期待": "on_credit / cash_on_delivery / bank_transfer / full_point のいずれか",
    },
    "TC-062": {
        "前提": "クーポン適用可能な注文",
        "手順": "1. クーポンコードを適用して配送・決済情報を送信\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'add_shipping_info').pop().ecommerce.coupon",
        "期待": "クーポンコード文字列が返る（例: 'COUPON001'）。未適用時はキーごと省略（推奨仕様①）。",
    },
    "TC-070": {
        "前提": "承認フロー対象の注文が可能",
        "手順": "1. 注文確認画面で承認申請ボタンを押下\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'approval_request')",
        "期待": "例: [{event:'approval_request', order_id:'T_00012345'}]",
    },
    "TC-071": {
        "前提": "承認不要の注文が可能",
        "手順": "1. 注文確定ボタンを押下して注文完了\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'purchase')",
        "期待": "例: [{event:'purchase', order_id:'T_00012345', customer_type:'new', point_amount:500, cart_type:'normal', request_item_count:3, ecommerce:{transaction_id:'T_00012345', value:9000, tax:900, shipping:500, currency:'JPY', coupon:'...', items:[...]}}]",
    },
    "TC-072": {
        "前提": "承認フロー対象の注文。承認者アカウントあり",
        "手順": "1. 承認申請 → 承認者が承認ボタン押下・処理完了\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'purchase')",
        "期待": "承認完了時にpurchaseが送信される（申請時ではない）。",
    },
    "TC-073": {
        "前提": "決済エラーを意図的に発生させられる環境",
        "手順": "1. 決済エラーになる条件で注文確定を試行\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'purchase')",
        "期待": "purchaseイベントは追加されない（件数が増えない）。",
    },
    "TC-074": {
        "前提": "初回購入/リピート/ゲストの各ケース",
        "手順": "1-A. 初回購入の取引先で注文\n1-B. 2回目以降の取引先で注文\n1-C. ゲスト購入\n2. customer_typeを確認",
        "コマンド": "dataLayer.filter(d => d.event === 'purchase').pop().customer_type",
        "期待": "new / returning / guest のいずれか",
    },
    "TC-075": {
        "前提": "承認フローありの注文完了",
        "手順": "1. 承認申請→承認→購入完了\n2. 各イベントのIDを比較",
        "コマンド": "(() => { const ar = dataLayer.find(d=>d.event==='approval_request'); const pu = dataLayer.find(d=>d.event==='purchase'); return {approval: ar?.order_id, transaction: pu?.ecommerce?.transaction_id, order: pu?.order_id}; })()",
        "期待": "3つの値がすべて同一（例: 'T_00012345'）。",
    },
    "TC-080": {
        "前提": "メルマガ設定変更可能（会員登録時 or マイページ）",
        "手順": "1. 受信設定を「全て受け取る」「テキストのみ」「受け取らない」にそれぞれ変更\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'mailmagazine_signup').map(d => d.subscription_status)",
        "期待": "['all'] / ['text_only'] / ['none'] がそれぞれ対応する",
    },
    "TC-081": {
        "前提": "一括注文画面にアクセス可能",
        "手順": "1. 一括注文フォーマット（CSV）をDL\n2. Consoleで確認",
        "コマンド": "dataLayer.filter(d => d.event === 'bulk_order_download')",
        "期待": "例: [{event:'bulk_order_download'}]",
    },
    "TC-082": {
        "前提": "正しいCSVファイルを用意",
        "手順": "1. 一括注文CSVをアップロードして成功\n2. カート画面遷移後、Consoleで確認",
        "コマンド": "({ upload: dataLayer.filter(d => d.event === 'bulk_order_upload'), carts: dataLayer.filter(d => d.event === 'add_to_cart' && d.add_method === 'bulk_order') })",
        "期待": "upload: [{event:'bulk_order_upload', upload_status:'success'}]\ncarts: アップロードした全商品分のadd_to_cartが存在し、add_method='bulk_order'",
    },
    "TC-083": {
        "前提": "エラーになるCSVファイルを用意（在庫不足等）",
        "手順": "1. 不正/エラーCSVをアップロード\n2. Consoleで確認",
        "コマンド": "({ upload: dataLayer.filter(d => d.event === 'bulk_order_upload'), carts: dataLayer.filter(d => d.event === 'add_to_cart') })",
        "期待": "upload: [{event:'bulk_order_upload', upload_status:'error', error_type:'...'}]\ncarts: add_to_cartは追加されない",
    },
    "TC-084": {
        "前提": "複数エラーが同時に発生するCSV",
        "手順": "1. 数量公倍数エラー+在庫不足等の複合エラーCSVをアップロード\n2. error_typeを確認",
        "コマンド": "dataLayer.filter(d => d.event === 'bulk_order_upload').pop().error_type",
        "期待": "カンマ区切りで複数エラー（例: 'quantity_invalid_multiple,insufficient_stock'）",
    },
    "TC-090": {
        "前提": "ecommerce.itemsを含む任意イベント発火済み",
        "手順": "1. view_itemまたはadd_to_cartを発火\n2. items[0]のフィールドを確認",
        "コマンド": "Object.keys(dataLayer.find(d => d.ecommerce?.items).ecommerce.items[0])",
        "期待": "['item_id','item_name','item_brand','item_category1','item_category2','item_category3','item_category4','item_category5','price','quantity'] が含まれる。\nprice/quantityは数値型。",
    },
    "TC-091": {
        "前提": "バリエーション（サイズ・色）がない商品",
        "手順": "1. バリエーションなし商品の詳細を表示\n2. itemsを確認",
        "コマンド": "'item_variant' in dataLayer.find(d => d.event === 'view_item').ecommerce.items[0]",
        "期待": "false（item_variantキー自体が存在しない）。",
    },
    "TC-092": {
        "前提": "異なるサプライヤーの商品を同時にカート追加可能",
        "手順": "1. サプライヤーAとBの商品をカートに追加\n2. items配列を確認",
        "コマンド": "dataLayer.filter(d => d.event === 'add_to_cart').pop().ecommerce.items.map(i => i.item_brand)",
        "期待": "配列に異なるサプライヤー名が別要素として含まれる（例: ['エビス倉庫', '別サプライヤー']）",
    },
}

# カテゴリ→シート名
CATEGORY_SHEETS = {
    "GTM基本設定": "TC-GTM",
    "共通ページ計測": "TC-COMMON",
    "認証": "TC-AUTH",
    "リード": "TC-LEAD",
    "商品・カート": "TC-ECOM",
    "購入フロー": "TC-CHECKOUT",
    "注文確定": "TC-PURCHASE",
    "その他": "TC-OTHER",
    "items配列": "TC-ITEMS",
}


def style_header_row(ws, row, cols, fill=HEADER_FILL, font=HEADER_FONT):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def create_matrix_sheet(wb):
    ws = wb.active
    ws.title = "テストマトリクス"
    headers = ["テストID", "カテゴリ", "テスト項目", "対象イベント/機能", "優先度", "実施結果", "実施日", "実施者", "備考"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header_row(ws, 1, len(headers))
    set_col_widths(ws, [10, 14, 42, 28, 8, 10, 12, 10, 20])

    for r, tc in enumerate(TEST_CASES, 2):
        for c, val in enumerate(tc, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = NORMAL
            cell.border = BORDER
            cell.alignment = WRAP
        for c in range(6, 10):
            ws.cell(row=r, column=c).border = BORDER

    # 凡例
    lr = len(TEST_CASES) + 3
    ws.cell(row=lr, column=1, value="【凡例】").font = Font(bold=True, size=10)
    ws.cell(row=lr + 1, column=1, value="実施結果: OK / NG / 未実施 / 保留")
    ws.cell(row=lr + 2, column=1, value=f"対象サイト: {SITE}")
    ws.cell(row=lr + 3, column=1, value=f"GTMコンテナID: {GTM_ID}")
    ws.cell(row=lr + 4, column=1, value="テスト環境: 手動テスト（ブラウザDevTools Console）")


def create_guide_sheet(wb):
    ws = wb.create_sheet("テスト実施手順（共通）", 1)
    rows = [
        ["Primo dataLayer 手動テスト実施手順"],
        [""],
        ["■ 事前準備"],
        ["1. Chrome（推奨）で対象サイトにアクセス"],
        ["2. F12 → Consoleタブを開く"],
        ["3. ページ遷移でdataLayerがリセットされるため、操作前にConsoleを開いた状態で実施"],
        ["4. Networkタブで「gtm」「collect」等のリクエストも併せて確認するとより確実"],
        [""],
        ["■ 基本確認コマンド"],
        ["コマンド", "用途", "期待される出力例"],
        ["dataLayer", "dataLayer全体を表示", "配列 [{...}, {...}, ...]"],
        ["JSON.stringify(dataLayer, null, 2)", "整形して全体表示", "読みやすいJSON形式"],
        ["dataLayer.filter(d => d.event === 'イベント名')", "特定イベントを抽出", "該当イベントの配列"],
        ["dataLayer.filter(d => d.event === 'イベント名').pop()", "直近の該当イベント", "最新1件のオブジェクト"],
        ["dataLayer.find(d => d.user_id)", "user_id付きpushを取得", "{user_id:'...', area_id:'...', ...}"],
        ["dataLayer.filter(d => 'user_id' in d)", "user_idキーの有無確認", "未ログイン時は []"],
        ["(() => { const i = dataLayer.findIndex(d => d.event === 'EVENT'); return i>=1 ? dataLayer[i-1] : null; })()", "ecommerce:null確認", "{ecommerce: null}"],
        [""],
        ["■ GTM設置確認コマンド"],
        ["document.querySelector('script[src*=\"googletagmanager.com/gtm.js\"]')", "head内GTMスニペット", "HTMLScriptElement（nullでない）"],
        ["document.querySelector('noscript iframe[src*=\"googletagmanager.com/ns.html\"]')", "body内noscript", "iframe要素（nullでない）"],
        [""],
        ["■ 判定基準"],
        ["・event名が指示書と完全一致（大文字小文字区別）"],
        ["・数値項目（price, quantity, stock, member_r等）は文字列ではなく数値型"],
        ["・未ログイン時はuser_idキー自体を出力しない（推奨仕様①）"],
        ["・値が存在しない任意項目はキーごと省略（nullではなくキー非出力）"],
        ["・ecommerce系イベントの直前に {ecommerce: null} がpushされている"],
        ["・JSONの最後の要素に余分なカンマがない（構文エラーなし）"],
    ]
    for r, row in enumerate(rows, 1):
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            if r == 1:
                cell.font = Font(bold=True, size=14, color="1A3A5C")
            elif r == 10:
                cell.fill = SUBHEADER_FILL
                cell.font = SUBHEADER_FONT
            elif row and row[0].startswith("■"):
                cell.font = Font(bold=True, size=11)
            else:
                cell.font = NORMAL
            cell.alignment = WRAP
    set_col_widths(ws, [70, 30, 40])
    ws.merge_cells("A1:C1")


def create_detail_sheet(wb, sheet_name, category):
    ws = wb.create_sheet(sheet_name)
    cases = [tc for tc in TEST_CASES if tc[1] == category]
    if not cases:
        return

    headers = ["テストID", "テスト項目", "前提条件", "操作手順", "確認コマンド（Console）", "期待結果（コマンド出力）", "実施結果", "備考"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header_row(ws, 1, len(headers))
    set_col_widths(ws, [10, 28, 22, 28, 38, 42, 10, 16])

    row = 2
    for tc_id, _, name, _, priority in cases:
        d = DETAILS[tc_id]
        values = [tc_id, f"【{priority}】{name}", d["前提"], d["手順"], d["コマンド"], d["期待"], "", ""]
        for c, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=c, value=val)
            cell.font = NORMAL
            cell.border = BORDER
            cell.alignment = WRAP
        ws.row_dimensions[row].height = 120
        row += 1

    ws.freeze_panes = "A2"


def main():
    wb = Workbook()
    create_matrix_sheet(wb)
    create_guide_sheet(wb)
    for category, sheet_name in CATEGORY_SHEETS.items():
        create_detail_sheet(wb, sheet_name, category)

    out = "/workspace/Primo_dataLayer_テスト仕様書.xlsx"
    wb.save(out)
    print(f"Saved: {out} ({len(TEST_CASES)} test cases)")


if __name__ == "__main__":
    main()
