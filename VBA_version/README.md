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

## ログインの基本フロー

1. `modIQuavisEntry.CreateClient` を呼び出して `IQuavisClient` インスタンスを生成します。必要に応じてベース URL やデバッグ出力の有効/無効を切り替えられます。
2. `modIQuavisEntry.AuthenticateClient` でユーザー ID とパスワードを渡すと、`IQuavisClient.Login` が `/token` エンドポイントに対してパスワード認証を実行します。
3. 成功するとアクセストークンがクラス内部のフィールド `mAccessToken` に保持され、`IQuavisClient.AccessToken` プロパティから取得できます。このトークンは後続の API 呼び出しで自動的に `Authorization: Bearer ...` ヘッダーに設定されます。

```vb
' 標準モジュールのサンプルコード
Public Sub LoginSample()
    Dim client As IQuavisClient
    Set client = CreateClient("https://example.com/iquavis-api", True)

    Call AuthenticateClient(client, "user@example.com", "password123")
    Debug.Print "AccessToken=" & client.AccessToken
End Sub
```

### ユーザーフォームとの連携例

VBA 側でログイン UI を提供する場合の一例です。

```vb
' UserFormLogin (ユーザーフォーム) のコード例
Option Explicit

Private mClient As IQuavisClient

Public Sub Initialize(ByVal client As IQuavisClient)
    Set mClient = client
End Sub

Private Sub btnLogin_Click()
    On Error GoTo Failed

    AuthenticateClient mClient, Me.txtUserId.Value, Me.txtPassword.Value
    MsgBox "ログインに成功しました", vbInformation
    Me.Hide
    Exit Sub

Failed:
    MsgBox "ログインに失敗しました: " & Err.Description, vbExclamation
End Sub
```

```vb
' 標準モジュールに配置するフォーム表示用プロシージャ
Public Sub ShowLoginForm()
    Dim client As IQuavisClient
    Set client = CreateClient()

    Dim frm As New UserFormLogin
    frm.Initialize client
    frm.Show

    If Len(client.AccessToken) > 0 Then
        MsgBox "アクセストークン取得済み: " & client.AccessToken
    End If
End Sub
```

## プロジェクト一覧の取得と履歴保存可否の判定

`IQuavisClient.ListProjects` は API からプロジェクトを取得し、プロジェクトごとのディクショナリまたはコレクションを返します。返却値を `modIQuavisEntry.FlattenProjects` でフラット化すれば、Excel への貼り付けや任意のプロパティ検査が容易になります。

```vb
Public Sub FetchProjectsSample()
    Dim client As IQuavisClient
    Set client = CreateClient()
    AuthenticateClient client, "user@example.com", "password123"

    Dim projects As Variant
    projects = FetchProjects(client)

    Dim flat As Collection
    Set flat = FlattenProjects(projects)

    Dim projectDict As Scripting.Dictionary
    For Each projectDict In flat
        Dim projectName As String
        projectName = CStr(projectDict("Name"))

        Dim historyEnabled As String
        If projectDict.Exists("HistoryEnabled") Then
            historyEnabled = CStr(projectDict("HistoryEnabled"))
        ElseIf projectDict.Exists("History") Then
            historyEnabled = CStr(projectDict("History"))
        Else
            historyEnabled = "不明"
        End If

        Debug.Print projectName & ": History=" & historyEnabled
    Next projectDict
End Sub
```

> **メモ**: プロジェクトの履歴保存可否は API レスポンスのフィールドに依存します。上記では `HistoryEnabled` または `History` といったキーが存在する想定で判定しています。実際の環境に合わせてキー名を調整してください。

## タスク一覧のエクスポート

`modIQuavisEntry.ExportTasksToWorkbook` は指定したプロジェクトのタスクを取得し、テンプレートブックへ展開します。

```vb
Public Sub ExportTasksSample()
    Dim client As IQuavisClient
    Set client = CreateClient()
    AuthenticateClient client, "user@example.com", "password123"

    Dim projects As Variant
    projects = FetchProjects(client)

    Dim firstProject As Variant
    firstProject = projects(1) ' 先頭のプロジェクトを使用する例

    Dim taskCount As Long
    taskCount = ExportTasksToWorkbook(client, firstProject, ThisWorkbook)

    MsgBox CStr(taskCount) & " 件のタスクをエクスポートしました", vbInformation
End Sub
```

エクスポート時には `tasks` シートに編集用データ、`_original` シートにバックアップが作成され、差分が黄色でハイライトされます。

## タスク更新の送信

`modIQuavisEntry.ApplyTaskUpdates` は `tasks` シートで変更されたセルを検知し、API へ `UpdateTask` を送信します。結果は `SummarizeResults` で集計できます。

```vb
Public Sub UpdateTasksSample()
    Dim client As IQuavisClient
    Set client = CreateClient()
    AuthenticateClient client, "user@example.com", "password123"

    Dim results As Collection
    Set results = ApplyTaskUpdates(client, ThisWorkbook.Worksheets("tasks"), _
                                   ThisWorkbook.Worksheets("_original"))

    Dim summary As String
    summary = SummarizeResults(results)
    MsgBox summary, vbInformation
End Sub
```

`ApplyTaskUpdates` 内部では、`CollectTaskRows` が変更行を抽出し、`BuildUpdatePayload` が JSON の更新ペイロードを生成します。API 応答の成否に応じてセルに青 (成功) / 赤 (失敗) のフィードバック色が設定されます。

---

必要に応じてフォームや UI に合わせてエントリーポイントを拡張してください。
