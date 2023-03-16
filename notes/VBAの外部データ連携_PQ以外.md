# VBAの外部データ連携_PQ以外

## 方法3選

1. Power Queryとの連携
2. 従来の方式
3. 外部ライブラリによる読み込み

### 方法1: Power Queryとの連携

- Excelデータタブの「データの取得と変換」欄にまとめてある
- ダイアログの「データの変換」ボタンを押す→Power Queryの画面が表示
- 作業フロー
  1. Power Queryで読み込み
  2. 結果をExcelの外部データ範囲として連携

### 方法2: 従来の方式

- Excelオプション設定→データ項目→レガシデータインポートウィザードの表示
- QueryTableオブジェクトを中心に、VBAオブジェクトのみで設定の作成から取り込みまで可能

### 方法3: 外部ライブラリによる読み込み

- テキストやファイルストリームなどを対象とすると良い
  - 理由: 動作が軽い、Power Queryや従来の方式は対象範囲が広く処理が大きい
- FileSystemObjectやStreamなどの外部ライブラリ

## テキストファイルからの読み込み

- 方法2: 従来の方式
  - QueryTableオブジェクトを使用し、区切り文字を指定して読み込む
  - QueryTableオブジェクト: QueryTablesコレクションに対して、Addメソッドで作成
  - 作業方針
    1. 読み込みたいテキストのパス情報を元にQueryTableオブジェクトを作成
    2. 各種プロパティ[^1]で読み込み設定を行い、Refreshメソッドで設定に従って読み込み
    3. 以降に必要なければDeleteメソッドで削除
  - 設定一部抜粋
    - 文字コード指定: TextFilePlatformプロパティ[^2]
    - フィールドデータ型を指定: TextFileColumnDataTypesプロパティ[^3]
      - よしなにやってくれる処理を省けるので、読み込み速度の向上が期待できる

  ```vba
  '基本
  読み込み先シート.QueryTables.Add _
    Connection:="TEXT;テキストファイルのパス文字列", _
    Destination:=読み込み先のセル

  With QueryTableオブジェクト
    .各種プロパティの設定
    .Refresh
  End With

  'テキストファイルの読み込み具体例
  Dim connectInfo As String
  '接続先の情報を作成
  connectInfo = "TEXT;" & ThisWorkbook.Path & "\外部データ.csv"
  'QueryTableを作成
  With ActiveSheet.QueryTables.Add(Connection:=connectInfo, Destination:=Range("B2"))
    '区切り文字の設定
    .TextFileParseType = xlDelimited
    .TextFileCommaDelimiter = True
    '読み込み
    .Refresh BackgroundQuery:=False
    '削除
    .Delete
  End With
  ```

- 方法3: 外部ライブラリによる読み込み
  - ADODB.Streamオブジェクト: Webサイトのログデータなど、一行ずつ大量に書き出されたデータなどに使うと良い
  - 各種プロパティ[^4]

  ```vba
  'テキストファイルをStreamで1行ずつ読み込み
  Dim textStream As Object, buf As String
  'Streamオブジェクトを生成
  Set textStream = CreateObject("ADODB.Stream")
  With textStream
    .Open
    .Type = 2 'adTypeText
    .Charset = "UTF-8"
    .LineSeparator = -1   'adCRLF
    '読込むファイルを指定
    .LoadFromFile ThisWorkbook.Path & "\外部データ(UTF-8).txt"
    'EOSプロパティがTrueになるまでループ処理
    Do While .EOS = False
      buf = .ReadText(-2) '1行読み込み
      If buf Like "*増田*" Then
        ActiveCell.Value = buf
        ActiveCell.Offset(1).Select
      End If
    Loop
    '閉じる
    .Close
  End With
  ```

## テキストファイルへの書き出し

- 方法2: 従来の方式
  - 書き出したいデータのみのブック作成
  - SaveAsメソッド、引数のFilename(パスとファイル名)とFileFormat(ファイル形式)を指定

  ```vba
  Dim saveRange As Range, saveBook As Workbook
  '書き出したいセル範囲をセット
  Set saveRange = Range("B2").CurrentRegion
  '新規ブックを作成し、コピー
  Set saveBook = Workbooks.Add
  saveRange.Copy saveBook.Worksheets(1).Range("A1")
  'CSV形式で保存
  saveBook.SaveAs _
    Filename:=ThisWorkbook.Path & "\ファイル名.拡張子", _
    'CSV形式.csv
    FileFormat:=xlCSV
    'タブ区切り形式.txt
    'FileFormat:=xlCurrentPlatformText
    'UTF8のCSV形式.csv
    'FileFormat:=xlCSVUTF8
  '保存が済んだ新規ブックを閉じる
  saveBook.Close
  ```

- 方法3: 外部ライブラリによる書き出し
  - 手間さえかければ、自分の好きな形式で書き出せる
  - ADODB.Streamオブジェクト

  ```vba
  'Streamオブジェクトでの基本の書き出し
  Dim textStream As Object, filePath As String
    '保存するテキストファイルのパス作成
  filePath = ThisWorkbook.Path & "\出力結果.txt"
    'Streamオブジェクトを生成して書き出し
  Set textStream = CreateObject("ADODB.Stream")
  With textStream
    .Open
    .Type = 2   'adTypeText
    .Charset = "UTF-8"
    '３回書き出し
    .WriteText "Hello", 1   'adWriteLine(改行アリ)
    .WriteText "Excel"
    .WriteText "VBA!"
    .SaveToFile filePath, 2 'adSaveCreateOverWrite（上書き）
    .Close
  End With

  'ArrayToText関数を使用したセル範囲からカンマ区切りの文字列を作成する関数
  Dim filePath As String, rng As Range
    '保存するテキストファイルのパス作成
  filePath = ThisWorkbook.Path & "\Streamで書出.txt"
    'Streamオブジェクトを生成して書き出し
  With CreateObject("ADODB.Stream")
    .Open
    .Type = 2   'adTypeText
    .Charset = "UTF-8"
    '見出しを書き出し
    .WriteText "ID,受注日,担当,商品名,価格,数量,小計", 1
    'セル範囲B3:H50の値を書き出し
    For Each rng In Range("B3:H50").Rows
      '1行分のデータをARRAYTOTEXT関数で文字列化して書き出し
      .WriteText WorksheetFunction.ArrayToText(rng.Value), 1
    Next
    .SaveToFile filePath, 2
    .Close
  End With
  ```

## Accessデータベースと連携

- 基本は、「方法1: Power Query連携」が良い、環境によって↓を使用
- 方法3: 外部ライブラリによる読み込み
  - DAO(Data Access Object)
  - 作業方針
    1. DBEngineオブジェクトを作成
    2. DBEngineオブジェクトのOpenDatabaseメソッドで任意のデータベースに接続、返り値はDatabaseオブジェクト
    3. Databaseオブジェクトの各種メソッドで、テーブルやクエリにアクセス
    4. 任意のテーブルを扱う→OpenRecordsetメソッドの引数にテーブル名orクエリ名を指定、返り値はRecordsetオブジェクト
    5. Recordsetオブジェクトから結果セットのデータを取り出す
    6. データ取り出し後は、RecordsetオブジェクトとDatabaseオブジェクトをClose
  - テーブルだろうがクエリだろうが、Recordsetオブジェクトに受け取って、CopyFromRecordsetメソッドでOK
  - パラメータクエリ: 上記作業方針の途中に以下の作業を追加
    - DatabaseオブジェクトにQueryDefsプロパティでパラメータを設定
    - Parametersプロパティを使用し、キーに対して値を渡す
    - `オブジェクト.Parameters("キー") = 値`
  - フィールド名も必要な場合
    - RecordsetオブジェクトのFieldsプロパティでFieldオブジェクトにアクセス
    - 各フィールドは「0」から始まるインデックス番号で管理
  - SQL文を使用したい場合
    - OpenRecordsetメソッドの引数にSQL文で渡す

  ```vba
  'テーブル/クエリに接続
  Dim DBE As Object, DB As Object, tmpRS As Object
    'DBEngineオブジェクト生成
  Set DBE = CreateObject("DAO.DBEngine.120")
    'データベースに接続
  Set DB = DBE.OpenDatabase(ThisWorkbook.Path & "\外部DB.accdb")
    'レコードセットにテーブルの内容を受け取る
  Set tmpRS = DB.OpenRecordset("T_社員")
    'データを転記
  Range("B2").CopyFromRecordset tmpRS
    '接続を切る
  tmpRS.Close
  DB.Close
    '列幅を自動調整
  Range("B2").CurrentRegion.Columns.AutoFit

  'パラメータクエリに接続
    '(省略)
  Set DB = DBE.OpenDatabase(ThisWorkbook.Path & "\外部DB.accdb")
    'パラメータクエリの定義を受け取り、パラメータを設定
  Set tmpQDef = DB.QueryDefs("PQ_社員明細")
  tmpQDef.Parameters("担当者入力") = "増田 宏樹"
    'レコードセットにパラメータクエリの内容を受け取る
  Set tmpRS = tmpQDef.OpenRecordset
    '(省略)

  'フィールド名も取り出す
    '(省略)
    'レコードセットにクエリの内容を受け取る
  Set tmpRS = DB.OpenRecordset("T_社員")
  'フィールド名書き出し
  For i = 0 To tmpRS.Fields.Count - 1
    Range("B2").Offset(0, i).Value = tmpRS.Fields(i).Name
  Next
    'データを転記
  Range("B3").CopyFromRecordset tmpRS
    '(省略)

  'SQL文で取り出す
    '(省略)
  Set DB = DBE.OpenDatabase(ThisWorkbook.Path & "\外部DB.accdb")
    'レコードセットにSQL文の結果を受け取る
  Set tmpRS = DB.OpenRecordset( _
    "SELECT *" & _
    " FROM" & _
      " Q_明細一覧" & _
    " WHERE" & _
      " 担当='増田 宏樹' AND" & _
      " 受注日 BETWEEN #2023/12/01# AND #2023/12/03#" _
  )
    'データを転記
  Range("B2").CopyFromRecordset tmpRS
    '(省略)
  ```

- 方法3: 外部ライブラリによる書き込み
  - DAO(Data Access Object)
  - 作業方針
    1. AddNewメソッドを実行
    2. `Recordsetオブジェクト!フィールド名 = 値`の形式
    3. Updateメソッドで確定
  - SQL文
    - DatabaseオブジェクトのExecuteメソッドの引数にSQL文を渡す
  - 既存レコードの修正
    1. ダイナセット形式で開く: OpenRecordsetの第2引数を「2(定数dbOpenDynasetの値)」
    2. 検索系メソッド(SeekメソッドやFindFirstメソッド)で対象レコードまで移動
    3. Editメソッドで編集、`Recordsetオブジェクト!フィールド名 = 値`の形式
    4. Updateメソッドで確定
  - トランザクション処理
    - Workspaceオブジェクトで以下を併用
      - BeginTransメソッド
      - CommitTransメソッド
      - Rollbackメソッド
      - エラー処理

```vba
'新規レコード追加
  '(省略)
  '「T_社員」テーブルに接続
Set tmpRS = DB.OpenRecordset("T_社員")
'新規レコードを追加
tmpRS.AddNew
tmpRS!ID = 11
tmpRS!担当 = "後藤 晋太郎"
tmpRS.Update
  '接続を切る
  '(省略)

'SQL文で新規レコード追加
  '(省略)
Set DB = DBE.OpenDatabase(ThisWorkbook.Path & "\外部DB.accdb")
  'SQL文で新規レコード追加
DB.Execute "INSERT INTO T_社員(ID, 担当) VALUES(11, '後藤 晋太郎') "
  '接続を切る
  '(省略)

'既存のレコードを修正
  '(省略)
  '「T_商品」テーブルに接続
  '第2引数にdbOpenDynasetの値「2」を指定しダイナセット形式で開く
Set tmpRS = DB.OpenRecordset("T_商品", 2)
  '「商品名」の値が「オリーブオイル」のレコードへ移動
tmpRS.FindFirst "商品名 = 'オリーブオイル'"
 '検索値が存在する場合には、その値を修正
If Not tmpRS.NoMatch Then
    tmpRS.Edit
    tmpRS!価格 = 1200
    tmpRS.Update
Else
    MsgBox "該当レコードは見つかりませんでした"
End If
  '接続を切る
  '(省略)

'トランザクション処理
  '(省略)
Set DB = DBE.OpenDatabase(ThisWorkbook.Path & "\外部DB.accdb")
  'WorkSpace取得
Set WS = DBE.WorkSpaces(0)
  'トランザクション処理＆エラートラップ開始
On Error GoTo ROLLBACK_DB
WS.BeginTrans

  '2つのテーブルに対する一連の処理実行
DB.Execute "UPDATE T_倉庫A" & _
           " SET 在庫数 = 在庫数 - 30" & _
           " WHERE ID = 1"

  '処理の途中でエラーを発生させる
Err.Raise 513

DB.Execute "UPDATE T_倉庫B" & _
           " SET 在庫数 = 在庫数 + 30" & _
           " WHERE ID = 1"
  '正常実行できた場合はコミットし終了処理へジャンプ
WS.CommitTrans
On Error GoTo 0
GoTo CLOSE_DB

  'エラートラップ用ラベル
ROLLBACK_DB:
  'ロールバックする
  WS.RollBack
  MsgBox "エラー発生。処理をロールバックします", vbExclamation
  '終了処理用ラベル
CLOSE_DB:
  '接続を切る
  '(省略)
```

---
[^1]: ExcelVBA[完全]入門, pp.512
[^2]: ExcelVBA[完全]入門, pp.514
[^3]: ExcelVBA[完全]入門, pp.516
[^4]: ExcelVBA[完全]入門, pp.518
