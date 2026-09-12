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

## このリポジトリとGASのclasp同期

同期先Script IDは次のプロジェクトに設定済みです。

```text
1eNC3EuiClJWHoJTZ2-AuT0WTJc8gaiabhiLfm5VJJSgQT4SKEjCSJELh
```

Windowsでは、ClassPay3を取得したフォルダで次を実行します。

```bat
cd C:\Users\user\GAS\repos\ClassPay3
git pull origin main
cd updater
clasp login
clasp push
```

2回目以降は次だけで更新できます。

```bat
cd C:\Users\user\GAS\repos\ClassPay3
git pull origin main
updater\push_to_gas.cmd
```

`updater/.claspignore`により、Updater GASへ送られるのは`updater`直下の`.js`、`.html`、`appsscript.json`だけです。ClassPay本体、設計書、release manifestは送られません。

### GAS側の修正をGitHubへ戻す場合

`clasp pull`はローカルファイルを上書きします。通常運用ではGitHubを原本にして、`git pull`→`clasp push`を使ってください。GASエディタで緊急修正した場合だけ、先にGitの変更をcommitまたは退避してから次を実行します。

```bat
cd C:\Users\user\GAS\repos\ClassPay3\updater
clasp pull
git diff
```

内容を確認してからcommit・pushします。
