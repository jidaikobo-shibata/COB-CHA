## 2026-07-08 Manifest V2 通知の対象確認

### 何をしたか

- Google から届いた Chrome Web Store の Manifest V2 拡張削除通知について、このリポジトリ `COB-CHA` が対象かを確認した。
- リポジトリ内で Chrome 拡張の `manifest.json`、`manifest_version`、Chrome 拡張向け API らしき記述を検索した。
- `appsscript.json` と `.clasp.json` を確認し、このリポジトリが Google Apps Script / clasp 管理のプロジェクトであることを確認した。
- Google 公式の Manifest V2 support timeline を確認し、通知対象が Chrome Web Store の Manifest V2 extensions であることを確認した。

### なぜそうしたか

- メール本文が「Chrome Web Store に公開中の Manifest V2 Chrome extensions」を対象としていたため。
- `COB-CHA` は Google スプレッドシート用 Google Apps Script アドオンであり、Chrome 拡張の manifest とは異なる `appsscript.json` を使うため、直接対象かを切り分ける必要があった。

### 判断

- このリポジトリ内に Chrome 拡張の `manifest.json` は存在しない。
- `manifest_version` も見つからない。
- 存在する manifest は GAS 用の `appsscript.json` のみで、Chrome Manifest V2 / V3 とは別物。
- したがって、少なくともこの `COB-CHA` リポジトリ自体は Manifest V2 削除通知の直接対象ではないと判断した。
- メールは、同じ Google 開発者アカウントに紐づく別の Chrome Web Store 拡張に対して届いた可能性が高い。

### 未完了の事項

- Chrome Web Store Developer Dashboard 上で、どの拡張が Manifest V2 と判定されているかは未確認。

### 次にやるとよいこと

- Chrome Web Store Developer Dashboard を開き、公開中または非公開の拡張一覧を確認する。
- 各拡張のソースに `manifest.json` があり、`"manifest_version": 2` なら移行対象として扱う。
- 必要に応じて、2026-08-31 より前に Manifest V3 移行またはメトリクスのエクスポートを行う。

## 2026-08-06 保守作業の開始とクライアント依存の移行計画

### 何をしたか

- 作業開始前に `clasp pull` を実行し、Google Apps Script 側の25ファイルを取得した。
- pull後にGitとの差分がないことを確認した。
- `.claspignore` に `**/.codex/**` を追加した。
- 消化済みTODOに関連する識別コメントと、リンク生成用のコメントアウトコードを削除した。
- `getSelectedSheets()` の旧実装として残っていたコメントアウトコードを削除した。
- jQuery 1.9.1とFont Awesome 5.6.1の利用箇所を調査した。
- 整理後の25ファイルを `clasp push` でGoogle Apps Scriptへ反映した。

### なぜそうしたか

- `.codex/` 内に将来 `.js`、`.html`、`.json` を置いても、意図せずGASへpushされないようにするため。
- 完了済みのTODOやGit履歴で参照できる旧コードを除き、現在の実装を読みやすくするため。
- 古い外部依存を安全に除去する前に、影響範囲と移行順序を明確にするため。

### クライアント依存の調査結果

- jQueryを直接利用するHTMLは、`ui-common.html`を含む8ファイル。
- `$(` の出現は合計161箇所で、イベント登録、表示切替、フォーム値操作、DOM生成に使われている。
- Font Awesomeを読み込むHTMLは7ファイル。
- Font Awesomeは主に27個の進捗スピナーと、外部リンクアイコンに使われている。
- Google公式のApps Script用CSS `add-ons1.css` も各画面から読み込んでいるが、今回の移行対象には含めない。

### 移行方針

1. `ui-common.html` の進捗制御を標準DOM APIへ変更し、共通ヘルパーを整える。
2. 比較的小さい `ui-help.html`、`ui-credit.html`、`ui-issuelist.html` からjQueryを除去する。
3. `ui-additional-criteria.html` と `ui-lump-edit.html` を移行する。
4. 利用箇所が多い `ui-issue.html` と `ui-control-panel.html` を最後に移行する。
5. Font AwesomeのスピナーをCSSのみの進捗表示へ置換する。装飾目的のアイコンは支援技術から隠し、進捗状態は既存のライブリージョン内のテキストで伝える。
6. 外部リンクアイコンは、必要なら文字またはCSSによる装飾に置換し、リンク目的を可視テキストまたはアクセシブルな名前で維持する。
7. 全画面からjQueryとFont Awesomeの読み込みがなくなったことを検索で確認し、GAS上で各ダイアログとサイドバーを手動確認する。

### 未完了の事項

- jQueryとFont Awesomeの実際の置換。
- 各画面の操作、フォーカス、処理中表示、エラー時の復帰に関する回帰確認。
- クライアント側処理に対する自動検査の導入。

### 次にやるとよいこと

- まず `ui-common.html` と小規模画面を標準DOM APIへ移行し、共通パターンを確立する。
- 画面ごとの手動確認項目を作り、段階的にGASへ反映する。
