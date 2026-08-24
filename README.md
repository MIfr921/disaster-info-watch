# 避難所サイネージ更新監視 PoC（Google Apps Script版）

PoCです。

## できること

- 5つの自治体・防災ページを1時間ごとに監視
- 初回はbaseline保存（既存記事を更新扱いしない）
- ETag / Last-Modifiedがあれば条件付きGET
- 本文ハッシュとブロック差分を保存
- 給水、断水、避難所、入浴、物資、医療、交通、ライフライン、公的支援等の変更だけ候補化
- 新しく追加された関連リンクがあれば、その記事本文も取得して候補を作成
- 生成AIを使わず、原文の文を抽出してドラフト作成
- A4横・1ページにfit-to-pageしたPDFをGoogle Sheets経由で生成
- PDFをGoogle Driveへ保存
- メール通知（標準）
- Slack Incoming Webhook通知（任意）
- Web Appで原文・差分・PDFを確認
- Web Appでタイトル・本文を修正してPDF再生成
- 承認 / 対象外、確認者、確認時刻を記録
- 監視エラーと復旧を通知
- 検知→通知→初回閲覧→承認の時刻を記録

## 最短セットアップ

### 1. Apps Scriptプロジェクトを作る

1. https://script.google.com/ を開く
2. 「新しいプロジェクト」
3. `Code.gs` の内容を、このZIPの `Code.gs` で丸ごと置き換える
4. 左の「+」→ HTML を選び、ファイル名を `Index` にする
5. `Index.html` の内容を貼る
6. プロジェクト設定でタイムゾーンを `Asia/Tokyo` にする

`appsscript.json` は必須ではありません。マニフェストを表示して使う場合の参考として同梱しています。

### 2. `setupPoc()` を実行

Apps Scriptエディタ上部の関数選択で `setupPoc` を選んで実行します。

初回はGoogle Drive、Google Sheets、外部URL取得、メール、トリガー等への権限を求められます。自分のGoogleアカウントで許可してください。

この処理で以下が自動作成されます。

- Driveフォルダ `避難所サイネージPoC`
- DBスプレッドシート `避難所サイネージPoC_DB`
- PDFフォルダ
- Snapshotフォルダ
- 1時間間隔トリガー
- 5つの監視URL
- 初回baseline
- チーム共有URL用のランダムtoken

### 3. Webアプリとしてデプロイ

Apps Script右上の **デプロイ → 新しいデプロイ**。

種類: **ウェブアプリ**

- 次のユーザーとして実行: **自分**
- アクセスできるユーザー: **Googleアカウントを持つ全員**

組織内だけで使う場合は、Google Workspaceで許可されている組織内アクセスを選んでください。チームメンバーはGoogleへのログインが必要です。

### 4. チーム共有URLを取得

デプロイ後に、Apps Scriptエディタから `showPocInfo()` を実行します。

実行ログに次のようなURLが表示されます。

```
https://script.google.com/macros/s/.../exec?token=xxxxxxxx
```

**このURLだけをチームに共有**します。

tokenを知らない人がWeb App URLだけを開いてもレビュー画面は表示されません。

### 5. 動作テスト

Web画面の「テスト候補」を押します。

確認すること:

1. 更新候補が作られる
2. Google DriveにPDFが作られる
3. メール通知が届く
4. 通知のリンクから候補を開ける
5. 原文・差分・PDFが見える
6. テキストを修正してPDF再生成できる
7. 確認者名を入れて「承認」できる

## Slack通知を追加

SlackでIncoming Webhook URLを用意した後、Apps Scriptエディタから一度だけ次を実行します。

```javascript
setSlackWebhookForPoc('https://hooks.slack.com/services/....');
```

コードにWebhookを残したくない場合は、関数実行後にURL部分を消して保存して構いません。値はScript Propertiesに保存されます。

Slackを無効化する場合:

```javascript
setSlackWebhookForPoc('');
```

## 通知メール変更

標準では `setupPoc()` を実行したGoogleアカウントのメールを通知先に設定します。

変更する場合:

```javascript
setNotificationEmailForPoc('example@example.com');
```

## チーム共有URLを無効化したい

tokenを更新します。

```javascript
rotateReviewTokenForPoc();
```

その後 `showPocInfo()` で新しい共有URLを確認してください。古いURLは使えなくなります。

## 本PoCの安全設計

生成AIを使っていません。PDFの本文は取得した原文から文を抜き出すだけです。

ただし自動抽出なので、**配信前の人間確認は必須**です。

- 日時、住所、電話番号を推測しない
- 原文の文字列を優先
- 不要なナビゲーション変更はキーワード判定で除外
- 取得失敗と「更新なし」を区別
- 初回はbaselineのみ
- PDFはGoogle Sheetsの `scale=4` でA4横1ページにfit-to-page

## データ保存場所

`避難所サイネージPoC_DB` には3シートがあります。

### Sources
監視URL、最終確認、HTTP状態、ETag、hash、エラーなど。

### Candidates
更新候補、差分、PDF、ステータス、確認者、時刻など。

### Logs
監視・PDF・通知・承認などのログ。

本文snapshotはGoogle Sheetsのセル上限を避けるためDriveのテキストファイルに保存します。

## PoC KPI

Candidatesシートには以下を残します。

- `detected_at`: 更新検知
- `notified_at`: 通知
- `first_viewed_at`: 初回閲覧
- `reviewed_at`: 承認/対象外

この差から、例えば

- 更新検知→通知
- 通知→担当者閲覧
- 閲覧→承認

の時間を評価できます。

## 注意点

### 1. Apps Scriptは厳密なリアルタイム監視ではない

1時間トリガーは設定できますが、Google側の都合で実行が多少遅れることがあります。PoCでは「1時間前後での検知」を評価対象にしてください。

### 2. JavaScriptで後から描画されるサイト

Apps ScriptのURL Fetchは通常のHTTPレスポンスを取得します。ブラウザ上でJavaScript実行後に初めて本文が現れるサイトは取得できない場合があります。その場合は監視状態にエラー/本文抽出不足として現れます。

### 3. PDFのDrive共有

PoCではPDFを「リンクを知っている全員が閲覧」にしようとします。Google Workspaceの管理ポリシーで外部共有が禁止されている場合は失敗します。その場合はDriveフォルダをチームに共有してください。

### 4. Webアプリのtoken

PoC用の簡易アクセス制御です。本番システムの厳密な認証ではありません。実運用ではGoogle Workspace組織内アクセスや正式な認証を使ってください。

## 監視対象URL

初期値:

- https://www.city.yachiyo.lg.jp/life/1/9/index.html
- https://www.city.yachiyo.lg.jp/life/1/9/56/
- https://www.city.yachiyo.lg.jp/soshiki/92/
- https://www.bousai.pref.chiba.lg.jp/portal/X_PUB_VF_Top
- https://www.city.yachiyo.lg.jp/

監視対象を変更する場合はDBの `Sources` シートでURLを編集できます。

`active` を FALSE にすれば監視停止です。
