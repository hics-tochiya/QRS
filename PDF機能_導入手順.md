# 保存済みPDF機能

masterをgit pullして、これまでどおりアプリを再起動してください。依存ライブラリの追加、環境変数の手動設定は不要です。

トップ画面・照合結果画面の「保存済みPDF一覧」からdownloads直下のPDFを閲覧・ダウンロードできます。

## 管理者ログイン

IDはadminです。管理者パスワードと秘密鍵は初回起動時に自動生成され、instance/local_credentials.sqlite3へ保存されます。管理者パスワードは起動時のコンソール（PyCharmの実行画面）で確認できます。従来の固定パスワードから変更になります。

再起動やgit pullでも同じ値を使います。instanceフォルダはGit管理対象外です。削除すると認証情報が再生成されます。環境変数FLASK_SECRET_KEY、QRS_ADMIN_USERNAME、QRS_ADMIN_PASSWORDを設定済みなら、そちらを優先します。

## 検証

PDF機能：Flaskテストクライアントで一覧、空一覧、日本語名、inline/attachmentヘッダー、存在しないファイル、非PDF、外部シンボリックリンク、パストラバーサルを確認済み。

自動設定：初回生成、再読み込み時の値の維持、複数スレッドからの同時初期化、Python構文を確認済み。

WindowsのExcel連携、wkhtmltopdfによるPDF新規生成、実ブラウザの表示は未確認です。
