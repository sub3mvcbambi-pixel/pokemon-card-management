# 運用ガイド（日本語）

このドキュメントは、GitHub リポジトリと Google スプレッドシートを連携させて本システムを動かし始めるまでの手順を、初心者向けに丁寧に解説したものです。GitHub アカウントと対象スプレッドシート（ID: `1u8p0o9kCazC7_G45hmkxBSQFNnsEFvmArpuhfeADdg0`）はすでに用意されている前提です。

---

## 1. 必要なものの確認

- **GitHub アカウント**（既に所有している前提）
- **Google アカウント**（スプレッドシートと Apps Script を扱うため）
- **Google スプレッドシート**: ID `1u8p0o9kCazC7_G45hmkxBSQFNnsEFvmArpuhfeADdg0`
- **パソコン環境**: Node.js / npm がインストール済みであるとスムーズ（`npm` コマンドを使用）
- **Git**: コマンドラインで Git 操作ができる状態

> **ヒント:** もし Node.js や Git が未インストールの場合は、公式サイト（[Node.js](https://nodejs.org/ja) / [Git](https://git-scm.com/downloads)）からインストールしてください。

---

## 2. リポジトリの取得と初期設定

1. 作業用フォルダを決めて、ターミナル（またはコマンドプロンプト）を開きます。
2. リポジトリをクローンします。
   ```bash
   git clone https://github.com/sub3mvcbambi-pixel/pokemon-card-management.git
   cd pokemon-card-management
   ```
3. 依存パッケージをインストールします。
   ```bash
   npm install
   ```
4. `.clasp.json` を作成して、このプロジェクトを対象の Apps Script に紐付けます。
   ```bash
   npm run link -- --scriptId 1u8p0o9kCazC7_G45hmkxBSQFNnsEFvmArpuhfeADdg0
   ```
   - `scriptId` の値は、デプロイしたい Apps Script プロジェクトの ID に置き換え可能です。
   - コマンド実行後、プロジェクト直下に `.clasp.json` が作成されます。

---

## 3. Google 側の準備

1. スプレッドシート（ID: `1u8p0o9kCazC7_G45hmkxBSQFNnsEFvmArpuhfeADdg0`）を Google Drive で開き、名前を付けておきます。
2. スプレッドシートの「共有」設定で、GitHub Actions 用のサービスアカウント（後述）を **編集者** 権限で招待します。
3. Apps Script プロジェクト（スプレッドシートの「拡張機能 → Apps Script」から開く）の **スクリプト ID** を控えておくと安心です。

---

## 4. サービスアカウントと GitHub Secrets の設定

GitHub Actions が Google Apps Script へアクセスするために、サービスアカウントの認証情報を設定します。

1. **Google Cloud Console** で新しいプロジェクトを作成（例: `pokemon-card-management-ci`）。
2. 左上のメニューから「API とサービス → ライブラリ」を開き、`Google Apps Script API` を検索して **有効化**。
3. 同じく「API とサービス → 認証情報」から **サービスアカウント** を作成。
   - ロールは「編集者」を選択すると便利です。
4. 作成したサービスアカウントの詳細画面で「キー → 新しい鍵を追加 → JSON」を選ぶと、JSON ファイルがダウンロードされます。
5. GitHub リポジトリの **Settings → Secrets and variables → Actions** で以下 2 つのシークレットを登録します。
   - `GAS_SCRIPT_ID`: Apps Script のスクリプト ID。
   - `GAS_CREDENTIALS`: 先ほどダウンロードした JSON ファイルの中身をそのまま貼り付けます。

> **ポイント:** JSON は機密情報なので、社内で共有する際も取り扱いに十分注意してください。

---

## 5. GitHub Actions を利用したデプロイ

1. `main` ブランチに最新コードがあることを確認し、GitHub にプッシュします。
   ```bash
   git status      # 変更が無いか確認
   git push origin main
   ```
2. GitHub 上でリポジトリを開き、**Actions** タブをクリックします。
3. `Deploy to Google Apps Script` ワークフローが自動で実行され、完了すると緑色のチェックが表示されます。
4. 失敗（赤いバツ）になった場合は、ワークフローを開いてログを確認し、シークレット設定やサービスアカウント権限に問題がないか点検してください。

> **手動実行したい場合**: Actions タブで該当ワークフローを選び、「Run workflow」ボタンを押すと任意のタイミングでデプロイできます。

---

## 6. スプレッドシート側の初期化

1. デプロイが完了したら、対象スプレッドシートをブラウザで開きます。
2. メニューバーに **SalesOps** というメニューが追加されていることを確認します。
3. **SalesOps → システム初期化** をクリックします。
4. 初回実行時に表示される権限リクエストを承認します（Google アカウントでの許可が必要です）。
5. シートが自動生成され、「設定_定数」「マスター_顧客」「取引_注文ヘッダ」などのタブが作成されます。

---

## 7. 日々の運用フロー（例）

1. **顧客登録**: SalesOps メニューの「新規顧客作成」を実行し、顧客情報を入力します。
2. **注文作成**: 「新規注文作成」でヘッダを作成し、顧客 ID と注文日を設定します。
3. **明細取り込み**: 「明細貼り付け取り込み」で TSV/CSV データを貼り付け、注文に紐付けます。
4. **在庫引当**: 「在庫引当」を実行して SKU の在庫を引き当てます。必要に応じて「発送処理」まで進めます。
5. **決済照合**: PayPal / Wise CSV を取込シートに貼り付け、メニューから照合機能を使って注文に費用を反映させます。
6. **集計更新**: 必要に応じて「顧客集計更新」「Export_Looker 再生成」を実行し、レポートを最新化します。

---

## 8. よくあるトラブルと対処法

| 症状 | 原因候補 | 対処法 |
| ---- | -------- | ------ |
| Actions が失敗する | シークレット未設定 / JSON の貼り付けミス | `GAS_SCRIPT_ID` と `GAS_CREDENTIALS` を再確認し、JSON の改行含め正しいかチェックする |
| スプレッドシートにメニューが出ない | デプロイ未完了 / 権限エラー | Actions のログを確認し、デプロイが成功したか確認。必要であれば手動で `npm run deploy` を実行 |
| システム初期化でエラー | 権限未許可 / シートの保護設定 | 初回実行時の承認ダイアログで「許可」を選択。シートが既に存在する場合は重複に注意 |
| 決済 CSV が紐付かない | 取引 ID や金額が一致しない | 注文ヘッダの「決済参照ID」を手入力で補完し、再度照合を実行 |

---

## 9. 追加ドキュメント

- セットアップ全般: [docs/SETUP.md](SETUP.md)
- GitHub Actions と clasp によるデプロイ詳細: [docs/DEPLOYMENT.md](DEPLOYMENT.md)
- 英語での概要と補足情報: [README.md](../README.md)

困ったときは、Actions のログやブラウザのコンソールに表示されるエラー内容をメモし、チームで共有することで解決しやすくなります。