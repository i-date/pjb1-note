# VBAの外部データ連携_PQ以外

## 方法3選

1. Power Queryとの連携
2. 従来の方式
3. 外部ライブラリによる読み込み

### 方法1: Power Queryとの連携

- Excelデータタブの「データの取得と変換」欄にまとめてある
- ダイアログの「データの変換」ボタンを押す→Power Queryの画面が表示
- 元データの更新などはできない
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

## Power Query

- 大まかな流れ
  1. Power Queryで取り込みデータ指定
  2. Power Queryが自動的にナビゲーションテーブルを作成
  3. ナビゲーションテーブルをステップごとに削る
- 大概の最初のステップ: どの場所にある、どういう形式のデータを対象にするか
- 対応データ形式とデータ関数の一例[^1]、テキスト(csv形式)→Csv.Document関数
- 注意: クエリの削除順は参照している側からでないと削除できない

### Power Queryで外部データを取り込む手順

- 手順: クエリを登録→シート上に展開(外部データ範囲)
- ブックに登録されているクエリの確認: 「クエリと接続」ボタン
- VBAでは↑のクエリを、WorkbookQueryオブジェクトとして管理
- 手作業で作成したクエリのM言語確認: PowerQueryエディタ→ホーム→「詳細エディター」ボタン
- PowerQueryエディタ起動: Excel→データ→「データの取得」ボタン→ PowerQueryエディターの起動
- VBAでクエリの追加
  - Querysコレクション: WorkbookオブジェクトのQuerysプロパティで取得
  - QuerysコレクションのAddメソッド
  - `ブック.Querys.Add Name:=クエリ名. Formula:=コマンドテキスト`
- テーブルとして展開or値のみ転記
  - テーブルとして展開: クエリに関連付けたListObject作成
    - 構造化参照による各種データ範囲の取得も可能
    - 外部データ範囲と関連付けたテーブル(ListObject)→QueryTableオブジェクトのRefreshメソッドでデータの再読み込み
    - 例: `Activesheet.ListObjects("伝票データ").QueryTable.Refresh`
  - 値のみ転記: クエリに関連付けたQueryTableオブジェクトを作成、シート展開後に削除
- VBAでクエリをブックから削除
  - `ThisWorkbook.Queries("クエリ名").Delete`

```vba
'クエリの登録
Dim queryName As String, commandText As Variant
  'クエリ名設定
queryName = "伝票データ"
  'M言語のコマンドテキスト作成
  '※実際には「C:\excel\外部データ.csv」が無いと動作しません。あくまでもコードの例です。
commandText = Array( _
  "let", _
    "source = Csv.Document(" & _
      "File.Contents(""C:\excel\外部データ.csv"")), ", _
    "header = Table.PromoteHeaders(source)", _
  "in", _
    "header" _
)
commandText = Join(commandText, vbCrLf)
  '※サンプルブックと同フォルダー内の「外部データ.csv」をANSI文字コードセットで読み込むよう修正。
commandText = Replace( _
  commandText, _
  "File.Contents(""C:\excel\外部データ.csv"")", _
  "File.Contents(""" & ThisWorkbook.Path & "\外部データ.csv""),[Encoding=932]" _
)
  '新規WorbookQueryオブジェクトとしてブックに登録
ThisWorkbook.Queries.Add Name:=queryName, Formula:=commandText

'テーブルとして展開
Dim queryName As String, table As ListObject
queryName = "伝票データ"
  'クエリと関連付けたテーブル作成
Set table = ActiveSheet.ListObjects.Add( _
    SourceType:=xlSrcExternal, _
    Source:="OLEDB;" & _
             "Provider=Microsoft.Mashup.OleDb.1;" & _
             "Data Source=$Workbook$;" & _
             "Location=" & queryName & ";", _
    Destination:=Range("B2") _
)
  'テーブル名をクエリ名と同じものに変更
table.Name = queryName
  'テーブルに関連付けた外部接続先のデータをリフレッシュ（取り込み）
With table.QueryTable
    .CommandType = xlCmdSql
    .commandText = Array("SELECT * FROM [" & queryName & "]")
    .Refresh
End With

'クエリから結果の値のみを展開
Dim queryName As String, qt As QueryTable
  'クエリ名を指定
queryName = "伝票データ"
  'クエリをデータソースとするQueryTableを作成
Set qt = ActiveSheet.QueryTables.Add( _
    Connection:="OLEDB;" & _
                "Provider=Microsoft.Mashup.OleDb.1;" & _
                "Data Source=$Workbook$;" & _
                "Location=" & queryName & ";", _
    Destination:=Range("B2") _
)
  'クエリの結果全体を取り込み後にQueryTable自体を削除
With qt
    .CommandType = xlCmdSql
    .commandText = Array("SELECT * FROM [" & queryName & "]")
    '完全に読み込むまで待機してから次の処理を実行
    .Refresh BackgroundQuery:=False
    .Delete
End With
```

### M言語の基本的な記述方法

- 空クエリの作成: Excel→データ→「データの取得」ボタン→その他のデータソースから→空クエリ
- PowerQueryエディタ起動: Excel→データ→「データの取得」ボタン→ PowerQueryエディターの起動
- Power Queryのステップ: 一つ一つの操作の呼び方
- 以下が基本構文
  - letブロック: 各ステップの操作を記述
  - inブロック: 結果として何を出力するかを記述
  - M言語の関数はアロー形式で記述

  ```m
  let
    ステップ名1 = 処理,
    ステップ名2 = 処理, ...
  in
    出力となるステップ名

  関数名 = (引数) =>
    let
      // 関数内のステップ
    in
      // 出力となるステップ名
  ```
  
  ```m
  dataFromRecord = (rec as record) =>
    let
      // レコードのf、h、j列から値を取り出してリスト化
      nums = Record.ToList(Record.SelectFields(rec, {"f","h","j"})),
      // 数値を文字列化
      numText = List.transform(nums, each Nunber.ToText(_)),
      // 取り出したリストの値を使用して「令和○年○月○日」も文字列を作成
      dateText = Text.Combine(
        {"令和",numText{0},"年",numText{1},"月",numText{2},"日"}
      )
    in
      Date.FromText(dateText)
  ```

- M言語の基本ルール: テーブルやリストの内容をどうするかが中心、VBAのように多岐にわたる機能やオブジェクトはない
  - 組み込み関数
    - ほぼ関数によって処理
    - 関数を使用してテーブルを成形し、戻り値として成形後のテーブルを受け取る
    - ↑を繰り返して最終出力の形へと整える
  - リストとレコード
    - リスト: `{}`の中に`,`区切りで値を記述
      - `リスト = {値1, 値2, ...}`
      - 値の出力: `リスト{0始まりのインデックス番号}`
    - レコード: `[]`の中に`,`区切りで`名前=値`を記述
      - `レコード = [名前1=値1, 名前2=値2, ...]`
      - 値の出力: `レコード[名前]`
  - テーブルのレコードやフィールドを取り出す
    - テーブルの任意の1レコードの取得: `テーブル{インデックス番号}`
    - ユニークなフィールド値がある場合: `テーブル{[フィールド=フィールド値]}`
    - 任意の1レコードのフィールド値を取得: `テーブル{インデックス番号}[フィールド]`
    - レコードを指定せずにフィールド値リストを取得: `テーブル[フィールド]`
  - コメント
    - 一行: 先頭に`//`
    - ブロック: 開始`/*`、終了`*/`
  - 行継続文字は存在せずに、改行は継続されているとみなされる
  - 大文字小文字の区別がされる

### CSV形式のデータを扱う

- M言語
  - ナビゲーションテーブルを作成
  - フィールドデータ型を指定(データを読み込んだ時点ではAny型なので可能な限り指定[^2])
  - フィルター
    - 条件式はeachキーワードを使用してフィールド名・比較演算子・値をセット
    - and演算子とor演算子が使用可能
    - 日付シリアル値: `#date(年,月,日)`

```m
// ナビゲーションテーブルを作成するコード
// Csv.Document(File.Contents(パス文字列), [引数のリスト])

let
  filepath = "C:¥excel¥外部データ.csv",
  source = Csv.Document(File.Contents(filepath),[Encording=932]),
  // 一行目をヘッダーとして扱う
  header = Table.PromoteHeaders(source)
in
  header

// フィールドのデータを定義
// Table.TransformColumnTypes(テーブル, [フィールドとデータ型のリスト])

filetype = Table.TransformColumnTypes(
  header,
  {
    {"ID", Number.type},
    ...
  })

// フィルター設定
// Table.SelectRows(テーブル, 条件式)

filter1 = Table.SelectRows(fieldType, each([担当] = "担当者名"))
filter2 = Table.SelectRows(
  filter1,
  each [受注日] >= #date(2023,12,1) and [受注日] <= #date(2023,12,10)
)

// ソート
// Table.Sort(テーブル, {対象フィールドとソート順のリスト})

sort = Table.Sort(filter2, {"小計", Order.Descending})

// 上記をまとめて使用したクエリ例
let
  filePath = "##ファイルパス##",
  source = Csv.Document(File.Contents(filePath),[Encoding=932]),
  header = Table.PromoteHeaders(source),
  fieldType = Table.TransformColumnTypes(
      header,
      {
          {"ID", Number.Type},
          {"受注日", Date.Type},
          {"担当", Text.Type},
          {"商品名", Text.Type},
          {"価格", Currency.Type},
          {"数量", Number.Type},
          {"小計", Currency.Type}
      }),
  filter1 = Table.SelectRows(fieldType, each ([担当] = "増田 宏樹")),
  filter2 = Table.SelectRows(
    filter1,
    each [受注日] >= #date(2023, 12, 1) and [受注日] <= #date(2023, 12, 10)
  ),
  sort = Table.Sort(filter2,{"小計",Order.Descending})
in
  sort
```

- 上記のクエリをVBAで実行するマクロ

  ```vba
  Dim queryName As String, commandText As String, filePath As String

  '同じフォルダー内にあるテキストファイルのパスを指定
  filePath = ThisWorkbook.Path & "\外部データ.csv"

  'クエリ名設定
  queryName = "CSV取り込み"

  'M言語のコマンドテキスト作成
  commandText = Join( _
      WorksheetFunction.Transpose(Worksheets("テキストファイル").Range("A2:A24").Value), _
      vbCrLf _
  )
  commandText = Replace(commandText, "##ファイルパス##", filePath)

  '新規WorbookQueryオブジェクトとしてブックに登録
  ThisWorkbook.Queries.Add _
                  Name:=queryName, _
                  Formula:=commandText

  '追加したクエリを展開
  Dim table As ListObject

  'クエリと関連付けたテーブル作成
  Set table = ActiveSheet.ListObjects.Add( _
      SourceType:=xlSrcExternal, _
      Source:="OLEDB;" & _
               "Provider=Microsoft.Mashup.OleDb.1;" & _
               "Data Source=$Workbook$;" & _
               "Location=" & queryName & ";", _
      Destination:=Range("C1") _
  )

  'テーブル名をクエリ名と同じものに変更
  table.Name = queryName

  'テーブルに関連付けた外部接続先のデータをリフレッシュ（取り込み）
  With table.QueryTable
      .CommandType = xlCmdSql
      .commandText = Array("SELECT * FROM [" & queryName & "]")
      .Refresh
  End With
  ```

### Excelのデータを扱う

- 自ブックをデータソースとして使用
  - Excel.CurrentWorkbook関数を使用: `Excel.CurrentWorkbook()`
  - ↑現在のブック内の「テーブル」と「名付きセル範囲」からナビゲーションテーブルを作成
  - テーブルは、Excel側の名前を持つNameフィールドと対応するセルの値をテーブルで保持するContentフィールドを持つ
  - 各テーブルを取得: `ナビゲーションテーブル{Name="フィールド名"}[Content]`
  - 各テーブルのデータ範囲を取得: `Table.Range(テーブル, 開始位置, 取得レコード数)`

```m
let
  xlBook = Excel.CurrentWorkbook(),
  dataTable = xlBook{[Name="伝票"]}[Content],
  top5 = Table.Range(dataTable,0,5)
in
  top5
```

- 外部ブックをデータソースとして使用
  - Excel.Workbook関数: `Excel.Workbook(File.Contents(パス文字列))`
    - シート単位でテーブルとして扱えるようにピックアップ
    - 全シートのデータをひとまとめにしたい系の処理が手軽になる
  - 各シートに対応のテーブルを取得: `ナビゲーションテーブル{Name="シート名"}[Data]`
  - Table.Combine関数: `Table.Combine({テーブルのリスト})`
  - Table.RemoveRows関数: `Table.RemoveRows(テーブル, 削除行数0始まり)`
  - Table.PromptHeaders関数: `Table.PromptHeaders(テーブル)`
  - ループ処理
    - 方針
      1. 処理を行いたい対象を抽出/リスト化
      2. 抽出/リスト化した対象すべてにEachで任意の処理を適用
    - List.Transform関数: `List.Transform(リスト, each 処理)`
      - eachキーワードを指定すると、リストのメンバーを参照する演算子として`_`が使えるようになる
      - リストを一括処理できて便利

  ```m
  let
    xlBook = Excel.Workbook(File.Contents("##ファイルパス##")), 

    /* Combine使用例
    combineTable = Table.Combine({
      xlBook{[Name="支店A"]}[Data],
      Table.RemoveRows(xlBook{[Name="支店B"]}[Data],0),
      Table.RemoveRows(xlBook{[Name="支店C"]}[Data],0)
    })
    header = Table.PromptHeaders(combineTable)
    */

    //「注意事項」シートを除くシートのレコードのみ抽出
    // null値除外: [任意のキーのフィールド]<>null
    filter = Table.SelectRows(
        xlBook,each [Kind]="Sheet" and [Name]<>"注意事項"
    ),
    //「Data」列のみの値をリストとして取得
    dataList = filter[Data],
    //リスト内のメンバーに対してTable.PromoteHeaders適用
    transList = List.Transform(dataList,each Table.PromoteHeaders(_)),
    //リスト内のメンバー内のnull値を持つレコードを除外
    cleanList = List.Transform(transList,each Table.SelectRows(_,each [ID]<>null)),
    //リストのメンバーを全連結
    combine = Table.Combine(cleanList)
  in
    combine
  ```

- Excel上のテーブルを結合
  - Table.Join関数: `Table.Join(テーブル1, キー列1, テーブル2, キー列2[, 結合方法])`
  - Table.PrefixColumns関数: `Table.PrefixColumns(テーブル, プレフィックス文字列)`
    - フィールド名の衝突回避
    - 変換後フィールド名: `プリフィックス文字列.元のフィールド名`
  - Table.SelectColumn関数: `Table.SelectColumn(テーブル, [フィールド名のリスト])`
    - リスト内で指定した順にフィールドを並び替えてくれる

  ```m
  let
    xlBook = Excel.CurrentWorkbook(),
    //商品テーブルと明細テーブルをそれぞれ「left」「right」に格納
    left = xlBook{[Name="商品"]}[Content],
    right = xlBook{[Name="明細"]}[Content],
    //同じフィールド名があると衝突するため、明細テーブル側にプレフィックスを付加
    prefixRight = Table.PrefixColumns(right,"明細"),
    //結合
    join = Table.Join(left,"id",prefixRight,"明細.商品id"),
    //結合したテーブルから必要なフィールドのみをピックアップ
    result = Table.SelectColumns(join,{"明細.id","商品","価格","明細.数量"})
  in
    result
  ```

- 特定フォルダ内のブックをまとめて取り込む
  - Folder.Files関数: `Folder.Files(フォルダへのパス文字列)`
  - Table.ReplaceValue関数
    - `Table.ReplaceValue(テーブル, 検索値(古いデータ), 置換値(新しいデータ), リプレーサー, 対象列)`
  - 展開: 入れ子状になっているメンバーを同じ階層のデータとして昇格させる処理
  - Table.ColumnNames関数: `Table.ColumnNames(テーブル名)`
    - 返り値: 引数に指定したテーブルのフィールド名のリスト
  - Table.ExpandTableColumn関数
    - `Table.ExpandTableColumn(テーブル, フィールド名, 展開前フィールド名リスト, 展開後フィールド名リスト)`
  - Table.RenameColumns関数: `Table.RenameColumns(テーブル, {元の名前と変更後の名前のリスト})`

```m
let
  files = Folder.Files("##フォルダーパス##"),
  //Excelブックのみを抽出し、Name列とContent列のみ取り出し
  xlTable = Table.SelectColumns(
      Table.SelectRows(files,each [Extension]=".xlsx"),
      {"Name","Content"}
  ),
  //各ブック内の「売上データ」テーブルを取り出し
  pickup = Table.ReplaceValue(
          xlTable,
          each [Content],
          each Excel.Workbook([Content]){[Name="売上データ"]}[Data],
          Replacer.ReplaceValue,
          {"Content"}
  ),
  //Content列を「展開」、Power Query独特の処理
  fieldNames = Table.ColumnNames(pickup{0}[Content]),
  expand = Table.ExpandTableColumn(pickup,"Content",fieldNames,fieldNames),
  //Name列を「取り込み元」列にリネーム
  rename =Table.RenameColumns(expand,{{"Name", "取り込み元"}})
in
  rename
```

- Excelの方眼紙のデータを取り込む
  - Excel.Worksheet関数でシートごとにデータを取り込む
  - ↑入力されていない箇所はnull値
  - 力技で各値をピックアップ
  - M言語の関数はアロー形式で記述
  - 作業
    1. 処理を行いたい対象を抽出/リスト化: 扱いやすいようにテーブルを整える
       - 行番号をシート上と揃える→Table.InsertRows関数
       - 列名を変更→Table.TransformColumnNames関数
    2. 任意の処理を用意: レコードを作成するアロー関数を用意
    3. 抽出/リスト化した対象すべてにEachで任意の処理を適用

  ```m
  let
    xlBook = Excel.Workbook(File.Contents("##ファイルパス##")),
    //シートのデータのみからなるリストを作成
    sheets = Table.SelectRows(xlBook,each [Kind]="Sheet")[Data],  
    //行番号をシート上と揃える
    fixRow = List.Transform(
               sheets,each Table.InsertRows(_,0,{_{0},_{0}})
    ),
    //列文字列で指定できるようにレコード数と列名を調整
    colChars = {"dummy","b","c","d","e","f","g","h","i","j","k"},
    fixCol = List.Transform(
      fixRow,
      each Table.TransformColumnNames(
         _,
         each colChars{
           Number.FromText(Text.Replace( _,"Column",""))
         }
      )
    ),

    //---------------- 関数定義 ----------------
    //レコードから日付を算出する関数
    dateFromRecord = (rec as record)=>
      let
        nums = Record.ToList(
          Record.SelectFields(rec,{"f","h","j"})
        ),
        numText = List.Transform(nums,each Number.ToText(_)),
        dateText = Text.Combine(
          {"令和",numText{0},"年",numText{1},"月",numText{2},"日"}
        )
      in
        Date.FromText(dateText),
    //レコードのg,h,i,j,k列から金額を算出する関数
    amountFromRecord = (rec as record)=>
      let
        nums = Record.ToList(
          Record.SelectFields(rec,{"g","h","i","j","k"})
        ),
        numText = List.Transform(
          nums,
          each if _ = null then "0" else Number.ToText(_)
        )
      in
        Currency.From(Text.Combine(numText)),
    //シート情報からレコードを作成する関数
    getRecord = (sheet as table)=>
      let
        name = sheet{6}[f],
        date = dateFromRecord(sheet{2}),
        amount = amountFromRecord(sheet{16})
      in
        [#"氏名"=name, #"日付"=date, #"合計金額"=amount],
    //---------------- ここまで関数定義 ----------------

    //リスト内のデータを元にシートごとのデータのレコードのリストを作成
    records = List.Transform(fixCol,each getRecord(_)),
    //レコードのリストを元にテーブル作成
    table = Table.FromList(
      records,Record.FieldValues,{"氏名","日付","合計金額"}
    )
  in
    table
  ```

### いろいろな形式のデータを取り込む[^3]

- Accessデータベース
  - Access.Database関数: `Access.Database(File.Contents(パス文字列))`
  - 下記の例では、Nameフィールドがユニークな値としている
  - 重複がある場合は、いままでのようにTable.SelectRows関数など使用

  ```m
  let
    accessDB = Access.Database(
        File.Contents("##ファイルパス##")
    ),
    table = accessDB{[Name="T_社員"]}[Data]
  in
    table
  ```

- XMLやフィード情報
  - テーブル形式として扱えるXML形式のデータは、要素の値だけでなく、要素の属性(Attribute)の値も自動で個別フィールドに展開してくれる
  - 最上位の要素から、目的の階層までたどっていく(降りていく)
  - Xml.Tables関数
    - Web: `Xml.Tables(Web.Contents(URL文字列)`
    - file: `Xml.Tables(File.Contents(パス文字列))`

  ```m
  let
    rss = Xml.Tables(Web.Contents(
        "https://news.yahoo.co.jp/rss/topics/it.xml")
    ),
    channel = rss{0}[channel],
    item = channel{0}[item],
    table = Table.SelectColumns(item,{"title", "link"})
  in
    table

  // ファイルの場合
  let
    xml = Xml.Tables(File.Contents("##ファイルパス##")),
    itemsTable = xml{[Name="items"]}[Table],
    itemTable = itemsTable{[Name="item"]}[Table]
  in
    itemTable
  ```

- JSON形式
  - 最上位の要素(レコード)のリスト
  - シンプルな構成の場合、リストをTable.FromRecords関数でテーブル化すれば、要素の一覧テーブルとして取得可能
  - Json.Document関数: `Json.Document(File.Contents(パス文字列))`

  ```m
  let
    json = Json.Document(File.Contents("##ファイルパス##")),
    table = Table.FromRecords(json)
  in
    table
  ```

- PDF
  - ページ内に表がある場合、兵庫とテーブルで扱えるように取得してくれる
  - 画像形式PDFは、画像読み取り機能にかける(Excelデータタブ→「画像から」ボタン)
  - Pdf.Tables関数: `Pdf.Tables(File.Contents(パス文字列))`
  - Pdf.Tables関数の返り値は、Power Queryの解析結果の一覧がナビゲーションテーブルとして返る

  ```m
  let
    pdf = Pdf.Tables(File.Contents("##ファイルパス##")),
    filter = Table.SelectRows(pdf,each [Kind]="Table"),
    table = Table.PromoteHeaders(filter{0}[Data])
  in
    table
  ```

- クエリや接続を一括削除(VBA)
  - クエリ: Queriesコレクションで管理されている
  - 接続: Connectionsコレクションで管理されている
  - マウスポインタを「クエリと接続」ペインに乗せてクエリのプレビューを確認

  ```vba
  Dim book As Workbook
  Set book = ActiveWorkbook

  'Queriesコレクションについてループ
  Do While book.Queries.Count > 0
      book.Queries(1).Delete
      DoEvents
  Loop

  'Connectionsコレクションについてループ
  Do While book.Connections.Count > 0
      book.Connections(1).Delete
      DoEvents
  Loop
  ```

### かゆいところに手が届くPower Queryの仕組み

- 複数のクエリを組み合わせて使用できる、サブクエリのイメージ
- パラメータクエリのような仕組みもできる
  - Power Query上でクエリ「テーブルの選択」を選ぶと、パラメータの入力が促される
- Excel側からPower Query側に値を送る
  - Power Query: 特定のクエリをパラメータとして使用→クエリに文字列のみ入力
  - Excel側
    - WorkbookQueryオブジェクト: クエリを管理
    - Formulaプロパティ: 作成済みのクエリコマンドテキストを取得/変更が可能

  ```m
  let
    // このクエリとは別に「対象ブック」というクエリが存在している場合
    xlBook = 対象ブック
  in
    xlBook{[Name="テーブルA"]}[Content]

  // 自ブックから、パラメータ「tableName」で受け取ったテーブル名のテーブルを取得するクエリ
  let
    table = (tableName as text) =>
      Excel.CurrentWorkbook(){[Name=tableName]}[Content]
  in
    table
  ```

```vba
ThisWorkbook.Queries("クエリ名").Formula = 値
ActiveSheet.ListObjects("指定").QueryTable.Refresh
```

---
[^1]: ExcelVBA[完全]入門, pp.553
[^2]: ExcelVBA[完全]入門, pp.558
[^3]: ExcelVBA[完全]入門, 付録コードにVBAのコードあり
