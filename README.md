# 消耗品管理・発注アプリ

Google Apps Script 製「消耗品管理・発注アプリ」の正式な開発リポジトリです。

現在の本番バージョンは **v6.1.0** です。

開発・運用上の詳細な引き継ぎ事項は [`HANDOFF.md`](HANDOFF.md) を参照してください。今後の開発では、作業前に必ず確認します。

## ディレクトリ構成

```text
src/
├── Code.gs          # サーバーサイド処理
├── Index.html       # Web アプリ画面
└── appsscript.json  # Apps Script マニフェスト
```

## ソース管理方針

次の3ファイルを唯一の正式ソースとして管理します。

- `src/Code.gs`
- `src/Index.html`
- `src/appsscript.json`

開発・修正は必ず `src/` のファイルに対して行います。ルートの `release-package-vX.X.X.json` は更新配信用の成果物であり、開発時の編集元には使用しません。

## リリース運用

更新を公開するときは、次の手順を守ります。

1. `src/` の正式ソースを更新する。
2. アプリのバージョンを上げる。
3. 正式ソースから、そのバージョンに対応する `release-package-vX.X.X.json` を生成する。
4. `release-manifest.json` のバージョン、公開日、更新内容、パッケージURLを最新版へ更新する。
5. `src/`、新しいリリースパッケージ、`release-manifest.json` を同じリリース変更として `main` に commit・push する。

過去の `release-package-vX.X.X.json` は、復旧・監査・バージョン確認に使用できるよう削除せず、リポジトリに保持します。

## 自動更新とデプロイ

既存アプリが `release-manifest.json` を確認し、指定された `release-package-vX.X.X.json` を取得して更新する現在の自動更新方式を維持します。

Web アプリの更新先には、設定済みの固定 Web アプリデプロイIDを継続して使用します。リリースのたびに別の Web アプリデプロイへ切り替えたり、固定デプロイIDを削除・上書きしたりしません。

詳細な安全要件、データ保護方針、リリース手順は `HANDOFF.md` を正とします。既存コードとの矛盾が見つかった場合は、推測で変更せず確認します。

## 復元元

`src/` の初期内容は、本番バージョン v6.0.6 の `release-package-v6.0.6.json` から復元しています。
