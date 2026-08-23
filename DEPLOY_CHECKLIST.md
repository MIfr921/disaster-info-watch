# デプロイチェックリスト

## 初回

- [ ] Google Apps Scriptで新規プロジェクトを作成
- [ ] `Code.gs` を貼り付け
- [ ] `Index.html` を作成して貼り付け
- [ ] タイムゾーンを Asia/Tokyo に設定
- [ ] `setupPoc()` を実行して権限承認
- [ ] 実行ログにエラーがない
- [ ] Google Driveに `避難所サイネージPoC` が作成された
- [ ] DBのSourcesに5URLが入っている
- [ ] Webアプリとしてデプロイ（実行: 自分）
- [ ] `showPocInfo()` でチーム共有URLを取得
- [ ] チーム共有URLをブラウザで開ける

## テスト

- [ ] 「テスト候補」を実行
- [ ] PDFが生成された
- [ ] PDFがA4横1ページで読める
- [ ] メール通知が届いた
- [ ] Slackを使う場合はSlack通知も届いた
- [ ] 候補の原文・差分・PDFを確認できる
- [ ] 文面修正→PDF再生成ができる
- [ ] 確認者名を入力して承認できる
- [ ] Candidatesに first_viewed_at / reviewed_at が残る

## 監視

- [ ] 5分トリガーが作成されている
- [ ] Sourcesの last_checked が更新される
- [ ] 取得失敗が「更新なし」扱いになっていない
- [ ] 初回baselineでは候補が大量発生していない

## PoC終了時

- [ ] Apps Scriptのトリガーを停止/削除
- [ ] 不要ならWebアプリのデプロイをアーカイブ
- [ ] 共有tokenを更新または無効化
- [ ] DriveのPoCフォルダを必要に応じて削除
