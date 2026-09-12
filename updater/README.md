# ClassPay Updater / Installer v1

ClassPay本体とは別のApps Scriptプロジェクトとして配置します。本体が自分自身を書き換えることを避け、先生本人のGoogle権限で対象プロジェクトを更新します。

## 導入

1. 新しいスタンドアロンApps Scriptプロジェクトを1つ作る。
2. `updater/` 内のファイルを配置する。
3. 標準Google Cloudプロジェクトまたは紐付けたCloudプロジェクトで Apps Script API を有効化する。
4. Webアプリとして「アクセスしているユーザーとして実行」「Googleアカウントを持つユーザー」でデプロイする。
5. 学校管理者にOAuthクライアントとscopeの許可を依頼する。
6. Updater URLをClassPay管理画面の「システム更新」に登録する。

本番配布前に、テスト用学級データで更新・rollbackを必ず検証してください。詳細は `docs/DISTRIBUTION_UPDATE_SYSTEM.md` を参照してください。
