# Changelog

このプロジェクトは [Keep a Changelog](https://keepachangelog.com/ja/1.0.0/) の形式に、[Semantic Versioning](https://semver.org/lang/ja/) を沿わせています。

## [1.1.1] - 初回公開

### 追加

- 指定フォルダのツリー一覧生成（アウトライン折りたたみ対応）
- 走査結果を専用シートへ複製し、事前確認のうえ実行するファイル操作
  - ファイル名の変更
  - 走査ルート内でのファイル移動
  - 新規フォルダの作成
  - ファイル・空フォルダのゴミ箱への移動
  - ファイル名への連番付与（先頭・末尾）
- 実行前の安全検査（root 逸脱、reparse point、衝突、stale 状態、数式インジェクション対策）
- 削除開始前の失敗に対する逆順ロールバック
- append-only の実行ログ

### 変更（公開に伴うポータブル化）

- ビルド・テストスクリプトから開発環境固有の絶対パスを除去し、リポジトリルートを動的に解決するよう変更
- `#Const TEST_BUILD` によるビルドプロファイル切り替え（`Test` / `Release`）を導入し、release candidate からテスト専用 procedure を除外
- Recycle Bin 性能テストの許可パスを、開発環境固有の絶対パスから動的解決へ変更

### 既知の制限

- Shell `ExtendedProperty` の取得レイテンシは少数の実 Recycle Bin アイテムでのみ測定しており、10,000 件・50,000 件規模での実測値ではありません。
- 削除件数 1,001〜10,000 件の所要時間は「見積不能」と表示されます。
- 署名なし（unsigned）で配布します。インターネット経由でダウンロードしたファイルは Office の既定設定でマクロがブロックされます。README の手順に従って個別に Unblock してください。

[1.1.1]: https://github.com/Ryoichi-R/folder-tree-excel/releases/tag/v1.1.1
