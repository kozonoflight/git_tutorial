# さくらのクラウド環境複製 自動化手順書

元資料: `さくらのクラウド環境複製手順_20250311.xlsx`  
対象: EC-Rider Primo / 無料トライアル等の LB + WEB 環境複製

本書は「どこを自動化し、どこを手動にするか」と「実装・実行手順」をまとめたものです。  
**できるかわからない箇所は手動**として残しています。

---

## 1. 目的とゴール

### 目的
契約サイト／トライアル用に、既存アーカイブから LB・WEB サーバを複製し、ドメイン・SSL・アプリ・監視まで立ち上げる作業を、可能な範囲で自動実行する。

### ゴール（自動化後）
1. パラメータ（サイト名・ドメイン・プラン等）を渡す
2. インフラ作成〜OS設定〜アプリ投入の大半がスクリプトで完了する
3. 人手は「外部サービス登録」と「最終動作確認」に限定する

---

## 2. 全体方針

| 層 | ツール案 | 役割 |
|---|---|---|
| インフラ | Terraform または usacloud（さくらの公式 CLI） | アーカイブ／ディスク／サーバ作成 |
| OS・ミドルウェア | Ansible | netplan、SSH、ufw、Postfix、DKIM、Zabbix agent、nginx、certbot |
| アプリ | 既存 Capistrano + 補助シェル | デプロイ、env、初期 dump |
| 秘密情報 | 環境変数 / 社内シークレット管理 | APIキー、パス、証明書 |
| 手動 | ブラウザ・管理画面 | ドメイン契約、reCAPTCHA、Redmine 記載、最終確認 |

推奨リポジトリ構成（新規作成想定）:

```text
sakura-ecr-provision/
├── README.md
├── inventories/
│   └── trial/
│       └── hosts.yml          # 生成後の LB/WEB IP を入れる
├── group_vars/
│   └── all.yml                # ドメイン等の共通変数
├── terraform/                 # または usacloud スクリプト
│   └── ...
├── ansible/
│   ├── site.yml
│   ├── roles/
│   └── templates/
├── scripts/
│   ├── 00_prereq_check.sh
│   ├── 01_create_servers.sh   # usacloud/terraform ラップ
│   ├── 11_deploy_app.sh       # Capistrano 呼び出し
│   └── 18_import_dump.sh
└── docs/
    └── this-runbook.md
```

---

## 3. 事前準備（手動）

自動化の前に、人が決める／用意するものです。

| No | 項目 | 内容 |
|---|---|---|
| P1 | さくらの API キー | 会員 ID／パスワード直書きは禁止。APIキーを発行し CI/ローカルの秘密情報に格納 |
| P2 | リージョン | 石狩第1ゾーン固定（元手順どおり） |
| P3 | 複製元アーカイブ | LB用・WEB用マイアーカイブが存在すること |
| P4 | サイト識別子 | ホスト名、サイト区分コード、表示名 |
| P5 | ドメイン方針 | `ec-rider.net` / `ec-rider2-demo.net` 等、取得済みか新規か |
| P6 | サーバプラン | LB / WEB のプラン・ディスクサイズ |
| P7 | 初期 dump | 共有ドライブ上の init dump パス |
| P8 | env 雛形 | `ec_admin_env_primo` / `ec_front_env_primo` |
| P9 | 作業用 SSH 公開鍵 | 作業者／CI 用 |
| P10 | 実行マシン | usacloud/Terraform/Ansible が動く作業 PC または Jump ホスト |

### 入力パラメータ例（`group_vars/all.yml`）

```yaml
site_code: "exampleco"
region: "is1a"   # 石狩第1
lb_hostname: "exampleco-lb"
web_hostname: "exampleco-web"
domains:
  front: "lite-front.example.ec-rider2-demo.net"
  admin: "lite-admin.example.ec-rider2-demo.net"
  test: "lite-test.example.ec-rider2-demo.net"
  mail: "mail.example.ec-rider2-demo.net"
archive:
  lb: "ARCHIVE_NAME_OR_ID_LB"
  web: "ARCHIVE_NAME_OR_ID_WEB"
plan:
  lb: "..."
  web: "..."
ssh_user: "xec"
zabbix_server: "..."
dkim_selector: "anytime"
# dump_file / env templates は秘密情報扱いのパスで管理
```

---

## 4. 手順対照表（元手順 → 自動 / 手動）

| 元No | 作業項目 | 扱い | 実装手段 |
|---|---|---|---|
| 1 | アーカイブ作成（LB・WEB） | **自動** | usacloud / Terraform（既存ディスクから作成） |
| 2 | ディスク作成（LB） | **自動** | 同上 |
| 3 | ディスク作成（WEB） | **自動** | 同上 |
| 4 | サーバー追加（LB） | **自動** | 同上 |
| 5 | サーバー追加（WEB） | **自動** | 同上 |
| 6 | netplan 設定変更 | **自動** | cloud-init または Ansible（コンソール手作業を廃止） |
| 7 | 鍵作成〜SSH接続 | **半自動** | 鍵生成は自動、初回疎通確認は手動でも可 |
| 8 | ドメイン取得（LB） | **手動** | レジストラ契約・取得判断が必要 |
| 9 | SSL証明書適用 | **自動** | Ansible + certbot（DNS 準備後） |
| 10 | ECR環境設定 | **自動** | Ansible テンプレート（nginx / .env / DB更新） |
| 11 | Railsデプロイ | **自動** | 既存 Capistrano（IP を変数指定） |
| 12 | reCAPTCHA設定 | **手動** | Google 管理画面でのサイト登録 |
| 13 | ufw 設定変更 | **自動** | Ansible |
| 14 | Postfix 設定変更 | **自動** | Ansible |
| 15 | Zabbixエージェント | **半自動** | agent 導入は自動 / Zabbix UI 設定は当面手動 |
| 16 | DKIM | **半自動** | インストール〜設定は自動 / DNS TXT 登録は手動 |
| 17 | 動作確認 | **手動** | チェックリスト実施（将来 E2E 化検討） |
| 18 | 初期 dump 投入 | **自動** | スクリプト |
| 19 | サイト情報記載 | **手動** | Redmine / 共有ドライブ |

---

## 5. フェーズ別 自動化手順

作業はフェーズ順に実行する。失敗したらそのフェーズで止めて再実行可能にする。

### Phase 0. 事前チェック（自動）

**スクリプト:** `scripts/00_prereq_check.sh`

確認内容:
- API 認証できること
- 指定アーカイブが存在すること
- 変数ファイルに必須キーがあること
- dump / env 雛形ファイルが参照できること

合否: NG なら Phase 1 に進まない。

---

### Phase 1. インフラ作成（自動） … 元手順 1〜5

**担当ツール:** Terraform または usacloud

#### やること
1. （必要なら）複製元からマイアーカイブ作成  
2. LB 用ディスク作成（マイアーカイブ指定）  
3. WEB 用ディスク作成  
4. LB サーバ作成・起動  
5. WEB サーバ作成・起動  
6. 付与されたグローバル IP を `inventories/trial/hosts.yml` に書き出す

#### 出力成果物
```yaml
# inventories/trial/hosts.yml 例
all:
  children:
    lb:
      hosts:
        lb1:
          ansible_host: "x.x.x.x"
    web:
      hosts:
        web1:
          ansible_host: "y.y.y.y"
```

#### 手動に落とす条件
- さくらの API でアーカイブ作成が権限不足などでできない場合 → 管理画面で 1〜5 を実施し、IP だけ inventory に手投入

---

### Phase 2. OS 初期設定（自動） … 元手順 6, 7, 13, 14

**担当ツール:** Ansible（`ansible/site.yml` の一部）

#### 2-1. netplan（元 6）
- 複製元の固定 IP を捨て、DHCP（または新規採番方針）に変更
- `netplan apply`
- 疎通確認（SSH）

> 元手順は「コンソールから編集」だが、自動化では **cloud-init** で初回から正しい netplan を入れるか、コンソール相当の初期接続後に Ansible で置換する。

#### 2-2. SSH（元 7）
自動化する内容:
- 作業用公開鍵を `authorized_keys` に配置
- `PasswordAuthentication no` / `PubkeyAuthentication yes` / `PermitRootLogin no`
- `sshd` 再起動

当面手動でもよい内容:
- 基幹連携用の鍵方式（ed25519 可否）確認
- 秘密鍵のローカル保管運用

#### 2-3. ufw（元 13）
- LB: 22/25 等を WEB IP・許可オフィス IP 向けに再定義
- WEB: 80/22 を LB IP 向けに再定義
- `ufw reload`
- `ufw status numbered` で期待ルールと突合（自動アサート）

#### 2-4. Postfix（元 14）
- LB: `main.cf` のドメイン / `mynetworks`（WEB IP）更新
- WEB: `relayhost` を LB IP に更新
- `systemctl restart postfix`

---

### Phase 3. ドメイン（手動） … 元手順 8

**手動作業**

1. ドメイン未取得なら取得（`ec-rider.net` / `ec-rider2-demo.net` 等の方針に従う）
2. 少なくとも以下を登録  
   - front / admin / test の A レコード → **LB IP**  
   - mail 用レコード（既存流用可ならスキップ）  
3. 反映確認

```bash
dig +short lite-front.example... A
```

**自動化に進む条件:** front/admin/test が LB IP を向いていること。

---

### Phase 4. SSL（自動） … 元手順 9

**前提:** Phase 3 完了

Ansible で実施:
1. LB: ACME webroot 用ディレクトリ作成
2. nginx に `/.well-known/acme-challenge/` を設定
3. `certbot certonly --webroot`（対象ドメインは変数）
4. nginx に証明書パスを設定し reload
5. renew dry-run / post_hook（nginx restart）設定
6. WEB: `sites-enabled/ecr` の server_name 等を新ドメインへ変更し reload

失敗時:
- DNS 未反映が主因 → Phase 3 に戻る（手動）

---

### Phase 5. ECR 環境設定 + デプロイ（自動） … 元手順 10, 11, 18

#### 5-1. 設定ファイル更新（元 10）
Ansible テンプレートで置換:
- LB `/etc/nginx/nginx.conf`（WEB IP、ドメイン）
- WEB `/etc/nginx/sites-enabled/ecr`
- `/home/xec/capistrano/ec_admin/shared/.env.production`
- `/home/xec/capistrano/ec_front/shared/.env.production`
- `site_settings` の URL 類（DB 更新 SQL or rails runner。方式は既存運用に合わせる）

#### 5-2. デプロイ（元 11）
```bash
# 例: 既存手順の IP 指定デプロイをラップ
./scripts/11_deploy_app.sh --web-ip "$WEB_IP"
```

#### 5-3. 初期 dump（元 18）
```bash
./scripts/18_import_dump.sh --dump /path/to/dumpXXXX_primo_init_data.sql
```

**手動に落とす条件:**
- env 雛形が案件ごとに大きく違う
- site_settings の更新方法が環境依存で特定できない

---

### Phase 6. 外部サービス連携（手動中心） … 元手順 12, 15(UI), 16(DNS), 19

#### 6-1. reCAPTCHA（元 12）… **手動**
1. Google reCAPTCHA 管理画面でサイト追加
2. 発行キーを `.env` 等へ反映（反映作業自体は Ansible 化してよい）

#### 6-2. Zabbix（元 15）
- **自動:** エージェントインストール、`zabbix_agentd.conf`、AllowKey、restart
- **当面手動:** Zabbix 管理画面でのホスト追加、アイテム、Webシナリオ、トリガー  
  （後続で Zabbix API 化を検討）

#### 6-3. DKIM（元 16）
- **自動:** OpenDKIM 導入、鍵生成、設定ファイル、Postfix milter、サービス起動
- **手動:** DNS へ TXT（DKIM/DMARC）登録、送信テストで PASS 確認

#### 6-4. サイト情報記載（元 19）… **手動**
- Redmine（無料トライアル配下）へ顧客情報記載
- 「さくらのクラウドサービス状況」へサーバ追記

---

### Phase 7. 動作確認（手動） … 元手順 17

チェックリスト（元シート準拠）:

1. LB・WEB へ鍵 SSH できる
2. WEB DB へ SSH トンネル接続できる
3. フロント・管理サイトへブラウザアクセスできる
4. フロント一通り操作・メール送信（PC/SP 表示）
5. 管理サイト一通り操作・メール・申請オプション
6. reCAPTCHA 動作
7. 標準バッチが cron 起動
8. UFW が定義どおり
9. メンテナンスモード切替
10. ClamAV / cron
11. fail2ban
12. AIDE / cron
13. スワップファイル
14. LB↔WEB ログ・DBdump 転送
15. DB バックアップ cron
16. Zabbix 監視
17. 送信メール DKIM/DMARC PASS

合否を作業記録に残す。

---

## 6. 実行オーダー（オペレーション手順）

現場で回すときの順番です。

```text
[手動] 事前準備 P1〜P10
   ↓
[自動] Phase 0  事前チェック
   ↓
[自動] Phase 1  サーバ作成 → inventory 出力
   ↓
[自動] Phase 2  netplan / SSH / ufw / Postfix
   ↓
[手動] Phase 3  ドメイン取得・DNS
   ↓
[自動] Phase 4  SSL
   ↓
[自動] Phase 5  ECR設定・デプロイ・dump
   ↓
[手動] Phase 6  reCAPTCHA / Zabbix UI / DKIM DNS / ドキュメント
   ↓
[手動] Phase 7  動作確認
```

ワンライナーイメージ（自動部分のみ）:

```bash
./scripts/00_prereq_check.sh && \
./scripts/01_create_servers.sh && \
ansible-playbook -i inventories/trial/hosts.yml ansible/site.yml --tags os,app,ssl && \
./scripts/11_deploy_app.sh && \
./scripts/18_import_dump.sh
```

---

## 7. Ansible タグ設計（案）

| タグ | 内容 |
|---|---|
| `os` | netplan, sshd, ufw, postfix |
| `ssl` | certbot, nginx SSL |
| `app` | nginx upstream, env, site_settings |
| `monitor` | zabbix-agent |
| `mail` | opendkim |
| `all` | 上記一括（DNS 準備後） |

例:

```bash
# インフラ直後（SSL 前）
ansible-playbook -i inventories/trial/hosts.yml ansible/site.yml --tags os

# DNS 完了後
ansible-playbook -i inventories/trial/hosts.yml ansible/site.yml --tags ssl,app,monitor,mail
```

---

## 8. ロールバック・再実行方針

| 失敗箇所 | 対応 |
|---|---|
| Phase 1 途中失敗 | 作成済みリソースをタグで識別し削除、または管理画面で削除して再実行 |
| Phase 2 設定ミス | Ansible を冪等に書き、再実行で是正 |
| Phase 4 SSL 失敗 | DNS 確認後に `--tags ssl` のみ再実行 |
| Phase 5 デプロイ失敗 | Capistrano 再実行。DB は dump 投入前なら作り直し検討 |
| 本番相当データ破壊リスク | dump 投入は明示確認付き（`--yes`）にする |

---

## 9. セキュリティ注意

1. 元 Excel にある管理画面パスワード等は **自動化資材に転記しない**
2. APIキー・DBパス・`.env` は Git 管理外
3. 作業用に一時的に PasswordAuthentication を開ける手順は自動化しない（鍵前提で進める）
4. 証明書秘密鍵・DKIM 秘密鍵の権限（`opendkim` 所有者等）は role 内で固定

---

## 10. 導入ロードマップ（実装順）

手がける順番の提案です。

| Step | 内容 | 効果 |
|---|---|---|
| 1 | usacloud/Terraform でサーバ作成 + inventory 出力 | 画面クリック削減が大きい |
| 2 | Ansible `os`（netplan/SSH/ufw/postfix） | コンソール作業を撲滅 |
| 3 | Ansible `ssl` + `app` + dump/deploy ラップ | 構築時間の大半を短縮 |
| 4 | Zabbix agent / DKIM 自動 | ミドルウェア作業を削減 |
| 5 | （余力）Zabbix API・DNS API | 手動 Phase 6 を縮小 |

---

## 11. 手動チェックリスト（残作業用）

構築のたびに印刷／チケット化して使う。

- [ ] ドメイン取得・DNS（front/admin/test → LB）
- [ ] reCAPTCHA サイト登録 → キー受け渡し
- [ ] Zabbix ホスト / Web監視 / トリガー追加
- [ ] DKIM/DMARC の DNS TXT 登録
- [ ] 動作確認 17 項目
- [ ] Redmine 記載
- [ ] さくらのクラウドサービス状況へサーバ追記

---

## 12. 未確定事項（現時点は手動扱い）

実装時に調査が必要で、本書では自動化対象外としているもの。

1. さくらの「アーカイブ作成」を API で行う権限・操作の最終確認  
2. `site_settings` 更新の正式手段（SQL / rails runner / 管理画面）  
3. お名前.com（または利用レジストラ）の DNS API 利用可否  
4. Zabbix API の社内利用可否・認証方式  
5. 基幹連携向け SSH 鍵アルゴリズム制約（ed25519 可否）  
6. cloud-init をアーカイブ起動時に差し込めるか（できない場合は Ansible 初回接続方式）

これらが判明次第、該当 Phase を「自動」へ昇格する。

---

## 13. 完了条件

以下を満たしたら「自動化フローで構築完了」とする。

1. Phase 1〜2, 4〜5 がスクリプトで再現できる  
2. 手動は Phase 3, 6, 7 と事前準備のみ  
3. Phase 7 の必須確認（少なくとも SSH、サイト表示、メール、UFW、SSL）が Pass  
4. 秘密情報がリポジトリに含まれていない  

---

## 付録 A. 元手順シート対応

| シート名 | 本書 Phase |
|---|---|
| 1.アーカイブの作成（LB・WEB） | Phase 1 |
| 2.ディスクの作成（LB） | Phase 1 |
| 3.ディスクの作成（WEB) | Phase 1 |
| 4.サーバー追加（LB) | Phase 1 |
| 5.サーバー追加（WEB) | Phase 1 |
| 6.netplanの設定変更（LB・WEB） | Phase 2 |
| 7.鍵の作成～SSH接続（LB・WEB） | Phase 2 |
| 8.ドメイン取得（LB） | Phase 3（手動） |
| 9.SSL証明書の適用（LB・WEB） | Phase 4 |
| 10.ECR環境設定 | Phase 5 |
| （11. Railsデプロイ） | Phase 5 |
| 12.reCAPTCHA設定 | Phase 6（手動） |
| 13.ufwの設定変更 | Phase 2 |
| 14.Postfixの設定変更（LB・WEB） | Phase 2 |
| 15.Zabbixエージェントインストール | Phase 2/6 |
| 16.Dkimインストール | Phase 6 |
| 17.動作確認 | Phase 7（手動） |
| 18.初期dump投入（WEB） | Phase 5 |
| 19.サイト情報記載 | Phase 6（手動） |
