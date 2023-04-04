# VBAのセルアクセス

- `[]`を使用したセル範囲のアクセス記法は使用しないほうが良い
- デバッグで使用するのはあり

## セル範囲

### 直接

| 概要 | コード | 備考 |
|:----|:--------|:-----|
| セル番地指定 | `Range("B2")` |  |
| 行列番号指定 | `Cells(2, 4)` |  |
| 2セルを囲む範囲1 | `Range(Range("B2"), Range("E4"))` |  |
| 2セルを囲む範囲2 | `Range("B2:E4")` | 先頭または終端のセルの調査が必要なときに便利 |
| 現在選択セル | `ActiveCell` |  |
| 現在選択範囲 | `Selection` |  |
| 行全体 | `Rows` | セル基準だと`Range("B2").EntireRow` |
| 行範囲 | ``Rows("2:4")`` |  |
| 列全体 | `Columns` | セル基準だと`Range("B2").EntireColumn` |
| 列範囲 | ``Columns("B:C")`` |  |

### 相対

| 概要 | コード | 備考 |
|:----|:--------|:-----|
| セル範囲の中のセル範囲 | `Range("B2:E6").Cells(2, 3)` |  |
| セル範囲の中の行 | `Range("B2:E6").Rows(3)` |  |
| セル範囲の中の列 | `Range("B2:E6").Columns(4)` |  |
| セル範囲の中のインデックス番号 | `Range("B2:E6").Cells(1)` | 終端セルはCellsの引数に`セル範囲.Cells.Count` |
| 基準セルからオフセット | `基準セル範囲.Offset(行オフセット数, 列オフセット数)` | マイナスの値もOK |
| サイズ拡張 | `基準セル範囲.Resize(行数, 列数)` |  |

### 表形式

| 概要 | コード | 備考 |
|:----|:--------|:-----|
| アクティブセル領域 | `Range("B2").CurrentRegion` | 注意: 表の周りに一行一列空欄がないと、余分な範囲も取得する |
| 最終行の次行を指定 | `Range("B2").End(引数).Offset(1)` |  |

- 次行/次列を選択（データ入力する時など）
  - Endプロパティ(引数: xlDown・xlUp・xlToLeft・xlToRight)
    - 下記コード(1): 見かけ上値がなくても、数式などがあれば取得
  - 表全体の行数を数えてオフセット、Withステートメントをあわせて使用するとさらに良い
    - 下記コード(2): 見かけ上値がなくても、数式などがあれば取得
  - Findで検索した位置をもとに取得
    - 下記コード(3): 見かけ上値で判断
  - 注意: フィルターがかかっているときはEndやFindは正しく動作しない
    - 対策1: CurrentRegionで行数を数える
    - 対策2: テーブル機能ベース

  ```vba
  コード(1)
  Range("B2").End(引数).Offset(1)

  コード(2)
  With Range("B2:F2")
    .Offset(.CurrentRegion.Rows.Count).Select
  End With

  コード(3)
  Dim lastRng As Range, targetField As Range
  Set targetField = Range(Range("F2"), Range("F2").End(xlDown))
  Set lastRng = Columns("F").Find( _
    "*", After:=targetField(1), LookIn:=xlValues, SearchDirection:=xlPrevious)
  lastRng.Offset(1).Select
  ```

### 異なるシートのセル範囲指定の注意

- Rangeプロパティに渡す2つの引数にセル`Range("B2")`を指定
- ↑の場合、アクティブシート上のセルになる

```vba
'2枚目のシートのセル範囲
'以下はダメ、2枚目のシート以外から実行できない
Worksheets(2).Range(Range("B2"), Range("E4")).Value = "VBA"

'以下はOK
With Worksheets(2)
  .Range(.Range("B2"), .Range("E4")).Value = "VBA"
End With

'↑のOKをWithを使用せずに一行で記述すると下記のようになる
Worksheets(2).Range(Worksheets(2).Range("B2"), Worksheets(2).Range("E4")).Value = "VBA"
```

## テーブル

- 表形式のセル範囲をテーブルとして登録
- 登録後、自動的に「名前付きセル範囲」として登録→構造化参照が可能
- ListObjectオブジェクトを使用、ListObjectsコレクションで管理されている
- 個別レコードはListRowオブジェクトを使用、ListRowsコレクションで管理されている
- テーブル作成の自動化

  ```vba
  'テーブル作成
  Sub createTable()
    Dim newTable As ListObject
    'アクティブシートに新規テーブルを作成
    Set newTable = ActiveSheet.ListObjects.Add(Source:=Range("B2:F7"))
    '後で扱いやすいように名前を設定
    newTable.Name = "売上テーブル"
    '書式をクリア
    newTable.TableStyle = ""
    'フィルター矢印をクリア
    newTable.ShowAutoFilterDropDown = False
  End Sub
  
  'テーブル解除
  Sub removeTable()
    'テーブル「売上テーブル」を解除
    With ActiveSheet.ListObjects("売上テーブル")
      .TableStyle = ""
      .Unlist
    End With
  End Sub
  ```

- 新規のテーブル範囲を作成
  - `ListObject.Add Source:=セル範囲`
- テーブル範囲にアクセス
  - シート経由: `シート.LinkObjects(テーブル名")`
  - セル経由: `テーブル内のセル.LinkObject`
  - 全体指定: `LinkObject.Range`
  - フィールド: `LinkObject.HeaderRowRange`
  - データ: `LinkObject.DataBodyRange`
  - 個別のデータ: `LinkObject.ListRows(インデックス番号)`

    ```vba
    '末尾のレコードにアクセス
    'テーブル取得
    Dim table As ListObject
    Set table = ActiveSheet.ListObjects("売上テーブル")
    'レコード取得
    Dim lastRec As ListRow
    Set lastRec = table.ListRows(table.ListRows.Count)  '末尾
    ```

- レコードの追加
  - 末尾: `ListRows.Add`、先頭: `ListRows.Add(1)`
  - 追加した位置のListRowオブジェクトを返す
  - ↑のオブジェクトのValueに値を代入
  - まとめて追加: Copy & PasteSpecial
- レコードの削除
  - `ListRows.Delete`

### 注意点

- テーブルデータをコピーする際の挙動
  - 全体を指定するとテーブルが別途作成される
  - 対策: HeaderRowRangeとDataBodyRangeなどのように分割してコピー
- EndとCurrentRegionプロパティの挙動
  - End:テーブル隣接セル範囲を含まない
  - CurrentRegion: テーブル隣接セル範囲を含む
- フィルター時の転記操作
  - アクティブなセルがテーブル内にある場合
    - 可視セルのみが転記
  - アクティブなセルがテーブル内にない場合
    - 全体が転記
    - 対策: SpecialCellsを用いて可視セルをコピー
- フィルター確認の違い
  - 通常: WorksheetオブジェクトのAutoFilterModeで確認可能
  - テーブル: ListObjectのAutoFilterのFilterModeプロパティ

## SpecialCells

- 条件を選択してジャンプ機能
- `セル範囲.SpecialCells(セルのタイプ[, オプション])`
- 入力漏れや数式であるはずのセルのチェックに使える
