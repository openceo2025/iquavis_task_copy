# VBA Port of iQUAVIS Tools

このフォルダには、既存の Python 製 iQUAVIS インポート/エクスポートツールを Excel VBA に移植するためのコードを格納しています。テンプレートとなる `xlsm` ブック内にモジュールをインポートし、ユーザーフォームなどから呼び出してください。

## 構成

- `IQuavisClient.cls` — iQUAVIS Web API へアクセスするクラス。認証、プロジェクト/タスクの取得、タスク更新を実装しています。
- `modJsonHelpers.bas` — JSON の変換、ディクショナリのフラット化/ネスト化などデータ整形のユーティリティです。
- `modExcelExport.bas` — タスク一覧をテンプレートブックに展開する処理をまとめています。
- `modExcelImport.bas` — Excel で編集したタスクを検出し、API へ更新を送信する処理をまとめています。
- `modIQuavisEntry.bas` — 上記モジュールを組み合わせたエントリーポイント例。フォーム側から呼び出す関数を記載しています。

## 依存関係

- **参照設定**: `Microsoft Scripting Runtime`, `Microsoft XML, v6.0`（または `WinHTTP`）にチェックしてください。
- **JSON ライブラリ**: VBA-JSON (`JsonConverter.bas`) を同じブックにインポートし、`JsonConverter` モジュールの `ParseJson` / `ConvertToJson` 関数を利用できるようにしてください。

## 利用方法の例

1. このフォルダ内の `.cls` / `.bas` ファイルを VBA エディタからインポートします。
2. `modIQuavisEntry` 内のサンプル関数をユーザーフォームやリボンボタンから呼び出します。
3. 認証情報、プロジェクト選択、テンプレートパスなどはフォーム側で受け取り、関数に引数として渡してください。

必要に応じてフォームや UI に合わせてエントリーポイントを拡張してください。
