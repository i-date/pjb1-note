# VBAのデータ処理

## ソート

- 方法
  1. RangeオブジェクトのSortメソッド: `データ入力セル範囲.Sort 各引数:=値`
  2. Sortオブジェクト: SortFieldオブジェクトを使用、SortFieldsコレクションで管理
     1. ソートしたいセル範囲を持つシートのSortオブジェクトを取得
     2. 対象セル範囲などの基本設定
     3. フィールドごとのソート方法をSortFieldsコレクションで作成`.SortFields.Add Key:=対象フィールドセル範囲 [, ほかの引数]`
     4. ソート実行`.Apply`
  3. テーブル機能と組み合わせたソート: ListObjectのSortプロパティから取得できるSortオブジェクトで管理
- 注意
  - 方法1・方法2において、ソート単位(行or列)が紛らわしいので注意

```vba
'セル範囲B2:B7に対して、「出荷数」フィールドを「降順」でソート
Range("B2:F7").Sort Header:=xlYes, Key1:="出荷数", Order1:=xlDescending

'Sortメソッドで複数条件でソート
Range("B2:F7").Sort Header:=xlYes, _
                    Key1:="担当", Order1:=xlAscending, _
                    Key2:="商品", Order2:=xlAscending, _
                    Key3:="日付", Order3:=xlDescending

'Sortオブジェクトで複数条件でSort
With ActiveSheet.Sort
  '既存のソート設定をクリア
  .SortFields.Clear
  '対象範囲とヘッダーの設定
  .SetRange Range("B2:F7")
  .Header = xlYes
  'ソート条件を追加 上から順に「担当」「商品」「日付」列
  .SortFields.Add Key:=.rng.Columns(2), Order:=xlAscending
  .SortFields.Add Key:=.rng.Columns(3), Order:=xlAscending
  .SortFields.Add Key:=.rng.Columns(4), Order:=xlDescending
  'ソート実行
  .Apply
End With

'テーブルを利用する際のソート
With ActiveSheet.ListObjects("出荷").Sort
  '既存のソート設定をクリア
  .SortFields.Clear
  'ソート条件を追加
  .SortFields.Add Key:=table.ListColumns("出荷数").Range
  'ソート実行
  .Apply
End With
```

## 抽出データとフィルター

- AutoFilterメソッド: `データ入力セル範囲.AutoFilter 各種引数:=値`
- 転記: フィルター適用セル範囲をそのままコピペ、可視部のコピーなどはしなくて良い
- 注意
  - 日付のフィルターはシリアル値ではなく、セルで表示されている表示形式での文字列を使う
  - 対策: 引数OperatorにxlAnd指定で、2つの日付の期間を計算させる

```vba
'フィルター作成とフィルター矢印をクリア
Dim fieldIndex As Long, rng As Range
Set rng = Range("B2:F50")
  'フィルター矢印を非表示に
For fieldIndex = 1 To rng.Columns.Count
  rng.AutoFilter fieldIndex, VisibleDropDown:=False
Next
  'フィルターをかける
rng.AutoFilter Field:=3, Criteria1:="=カレー"

'特定フィールドのフィルター(抽出条件)をクリア
Range("B2:F50").AutoFilter Field:=3

'フィルターがかかっている場合、フィルター設定そのものをクリア
Range("B2:F50").AutoFilter

'シート全体のフィルターをクリア
ActiveSheet.AutoFilterMode = False

'And
Range("B2:F50").AutoFilter Field:=4, _
  Criteria1:=">=1000", Operator:=xlAnd, Criteria2:="<2000"

'配列
Range("B2:F50").AutoFilter Field:=3, _
  Criteria1:=Array("ビール", "チャイ", "カレー", "コーヒー"), _
  Operator:=xlFilterValues

'上位何個
Range("B2:F50").AutoFilter Field:=5, _
  Criteria1:=3, Operator:=xlTop10Items

'抽出結果のコピー
  'フィルターのかかった状態で全体をコピー
Range("B2:F50").Copy
  '転記先に列幅を含めて貼り付け
With Worksheets("抽出結果").Range("B2")
  .PasteSpecial xlPasteColumnWidths
  .PasteSpecial xlPasteAll
End With
  '貼り付け待機状態を解除
Application.CutCopyMode = False

'抽出結果"以外"のコピー
With Range("B2:E10").Rows("2:10")
  '現在の可視セル(フィルター結果のセル)の行を非表示
  .SpecialCells(xlCellTypeVisible).EntireRow.Hidden = True
  'フィルターを解除
  .AutoFilter
  '可視セル(フィルター結果ではないセル)のみコピー
  .SpecialCells(xlCellTypeVisible).Copy
End With
  '転記先に列幅を含めて貼り付け
With Worksheets("抽出結果").Range("B2")
  .PasteSpecial xlPasteColumnWidths
  .PasteSpecial xlPasteAll
End With

'日付値のフィルター
Range("B2:D8").AutoFilter Field:=2, Criteria1:="=2023/4/1"

'日付値のフィルター_妥協策
Dim myDate As Date
myDate = #4/1/2023#　'変数myDateにはシリアル値が格納
  '期間を求める抽出条件として設定し、シリアル値ベースで抽出
Range("B2:D8").AutoFilter Field:=2, _
  Criteria1:=">=" & myDate, Operator:=xlAnd, Criteria2:="<=" & myDate
```

### フィルターの詳細設定

- AdvancedFilterメソッド: `データ入力セル範囲.AdvancedFilter 各種引数:=値`
  - 抽出と転記を一発で実行
  - 必要なフィールドのみも転記可能
    - 転記先のセル範囲に「抽出してほしいフィールド名」を事前に記述しておく
- 抽出条件をシートに記述
  - フィールド見出しを記述し、その下に抽出したい値を記述
  - 横方向: And条件
  - 縦方向: OR条件
  - 下記の例だと、「数量が200以上」または「合計が10000000以上」

|数量|合計|
|:---|:---|
| >=200 |  |
|  | >=10000000 |

```vba
'セル範囲B2:G50の表を、セル範囲K2:L3の条件で抽出
Range("B2:G50").AdvancedFilter _
  Action:=xlFilterInPlace, _
  CriteriaRange:=Range("K2:L3")


'セル範囲B2:G50の表を、セル範囲K2:L3の条件で抽出して転記
Worksheets("フィルターの詳細設定").Range("B2:G50").AdvancedFilter _
  Action:=xlFilterCopy, _
  CriteriaRange:=sht.Range("K2:L3"), _
  CopyToRange:=Worksheets("抽出先").Range("B2")

'セル範囲B2:G50の表を、セル範囲K2:L3の条件で抽出して転記
Worksheets("フィルターの詳細設定").Range("B2:G50").AdvancedFilter _
  Action:=xlFilterCopy, _
  CriteriaRange:=sht.Range("K2:L3"), _
  CopyToRange:=Worksheets("抽出先").Range("I2:J2")
```

## データの整理

### 重複削除

1. Dictionaryオブジェクトを使用
2. RemoveDuplicatesメソッドを使用

- リストから削除のセオリーは「後ろから削除」

```vba
'ユニークなリストを作成
Dim rng As Range, uniqueList() As Variant, dic As Object
  'Dictionnaryオブジェクト生成
Set dic = CreateObject("Scripting.Dictionary")
  'セル範囲B2:E5を走査してユニークな値をキー値としてピックアップ
For Each rng In Range("B2:E5")
  If Not dic.Exists(rng.Value) Then dic.Add rng.Value, "dummy"
Next
  'Dictionaryオブジェクトのキー値を配列で受け取る
uniqueList = dic.Keys

'1つ目のフィールドのみをチェックして重複削除
Range("B2:F9").RemoveDuplicates Columns:=1, Header:=xlYes

'2・3・4・5列目のフィールドをチェックして重複削除
Range("B2:F9").RemoveDuplicates _
  Columns:=Array(2, 3, 4, 5), Header:=xlYes

'重複削除_単一フィールド
  '処理対象行を管理する変数を宣言
Dim curRow As Long
  '「伝票ID」列の末尾のデータから2つ目のデータまで逆方向にループ処理
For curRow = 9 To 4 Step -1
  Cells(curRow, "B").Select '経過がわかりやすいよう対象セル選択
  'ひとつ上のセルと同じ値であれば、そのデータの範囲を削除
  If Selection.Value = Selection.Offset(-1).Value Then
      Selection.Resize(1, 5).Delete Shift:=xlShiftUp
  End If
Next

'重複削除_複数列
Dim curRow As Long, dic As Object, tmpKey As String, rng As Range
  'Dictionnaryオブジェクト生成
Set dic = CreateObject("Scripting.Dictionary")
  '末尾のデータから2つ目のデータまで逆方向にループ処理
For curRow = 9 To 4 Step -1
  'チェック対象レコードのセル範囲取得
  Set rng = Cells(curRow, "B").Resize(1, 5)
  'セル範囲の2～5番目の値（C～F列の値）を連結したキー文字列作成
  tmpKey = _
    rng.Cells(2).Value & _
    rng.Cells(3).Value & _
    rng.Cells(4).Value & _
    rng.Cells(5).Value
  'キーが重複した場合(登録済みだった場合)、セルを削除
  If dic.Exists(tmpKey) Then
    rng.Delete Shift:=xlShiftUp
  Else
  '重複していなければ行番号を登録
    dic.Add tmpKey, curRow
  End If
Next
```

## ワークシート関数を使用したデータ処理

- WorksheetFunction.Sortメソッド
  - `WorksheetFunction.Sort セル範囲, キー列, 並べ替え順, ソート方向`
  - 2次元配列を対象→戻り値も2次元配列、インデックスは「1」始まり
  - 出力確認
    - INDEXワークシート関数(後述)、TEXTJOINワークシート関数が便利
    - ARRAYTOTEXTワークシート関数だともっと便利

  ```vba
  'WorksheetFunction.Sortでソート
  Dim arr As Variant
    'ソートの結果配列を受け取る
  arr = WorksheetFunction.Sort(Range("B3:C7"), 2, -1)
    '結果を「1行」ずつ出力
  Dim idx As Long, wf As WorksheetFunction
  Set wf = WorksheetFunction
  For idx = LBound(arr) To UBound(arr)
    Debug.Print wf.TextJoin(",", False, wf.Index(arr, idx))
  Next
  ```

- WorksheetFunction.Indexメソッド
  - 配列から、任意の行・列のデータを取り出す
  - `WorksheetFunction.Index(2次元配列, 行インデックス)`
  - `WorksheetFunction.Index(2次元配列, 0, 列インデックス)`

- WorksheetFunction.Filterメソッド
  - `WorksheetFunction.Filter セル範囲, 条件, 一致しない場合の値`
  - 戻り値
    - ヒットなしの場合: 第3引数で指定した値→IsArray関数で判定可能
    - ヒット対象が1つの場合: 1次元配列
    - ヒット対象が複数の場合:2次元配列

  ```vba
  Dim filterArr As Variant
  'フィルターの結果配列を受け取る
  filterArr = WorksheetFunction.Filter(Range("B3:F50"), [D3:D50="カレー"], "")

  '[D3:D50="カレー"]を別シートから取得する場合: [シート名!D3:D50="カレー"]
  ```

- WorksheetFunction.Uniqueメソッド
  - `WorksheetFunction.Unique セル範囲, 方向, 回数`

  ```vba
  Dim uniqueList As Variant
  uniqueList = WorksheetFunction.Unique(Range("B3:B9"))

  '1列タテ方向のセル範囲なのでヨコ方向の一次元配列に変換
  uniqueList = WorksheetFunction.Transpose(uniqueList)
  'リスト数と始めの値を取り出す
  Debug.Print "リスト数：" & UBound(uniqueList)
  Debug.Print uniqueList(1)
  ```

## 表記

- 文字列変換: StrConv関数を使用
- フリガナ表記取得: PHONETICワークシート関数
- 形式変換: Format関数
- 正規表現: RegExpオブジェクト
- 置換: Replaceメソッド
  - `対象セル範囲.Replace 検索値, 置換値[, 各種引数:=値]`

```vba
Dim rng As Range
For Each rng In Range("B3:B6")
  '半角先頭大文字に統一
  rng.Value = StrConv(rng.Value, vbNarrow + vbProperCase)
  
  'カタカナ全角に統一
  rng.Value = StrConv(rng.Value, vbKatakana + vbWide)
  
  '英数字半角、カタカナ全角に統一
    'いったん全て半角・大文字に統一
  rng.Value = StrConv(rng.Value, vbNarrow + vbUpperCase)
    'PHONETICワークシート関数でフリガナ表記を取得して置き換え
  rng.Value = Application.WorksheetFunction.Phonetic(rng)
  
  '数値を元にして定形IDなどに統一
  If IsNumeric(rng.Value) Then
    rng.Value = Format(Val(StrConv(rng.Value, vbNarrow)), "VB-000")
  End If
  
  'シリアル値から復元
  rng.Value = Format(rng.Value, "'mm-dd") 'シングルクォーテーションで文字列を明示
Next

'正規表現で数値のみ取り出し
Dim rng As Range, regExpObj As Object, tmpMatches As Object
Set regExpObj = CreateObject("VBScript.RegExp")
regExpObj.Global = True
regExpObj.Pattern = "\d[\.\d]*"
For Each rng In Range("B3:B6")
  Set tmpMatches = regExpObj.Execute(StrConv(rng.Value, vbNarrow))
  If tmpMatches.Count > 0 Then
    rng.Value = tmpMatches(0).Value
  End If
Next

'シートの一覧表を元に置き換え
Dim sourceRng As Range, replaceTable() As Variant, i As Long
  '置換範囲セット
Set sourceRng = Range("B3:B6")
  '置換テーブルセット
replaceTable = Range("D3:E11").Value
  '置換
For i = 1 To UBound(replaceTable)
  sourceRng.Replace LookAt:=xlPart, _
                                  What:=replaceTable(i, 1), _
                                  Replacement:=replaceTable(i, 2), _
                                  MatchCase:=False, _
                                  MatchByte:=False
Next
```

## 検索

- Findメソッド: 最初に見つかったセルを選択
  - `検索対象セル範囲.Find 検索値[, 各種引数:=値]`
  - 各種引数を省略した場合: 「検索と置換」ダイアログの設定や前回の検索設定を引き継ぐ
- FindNextメソッド: 同じ検索条件で、引数に指定したセルの「次のセル」から検索を実施
  - `検索対象セル範囲.FindNext(After:=検索基準セル)`

```vba
'「Excel」と入力されているセルにジャンプ
Application.GoTo Cells.Find("Excel")

'検索の基本
Dim findCell As Range
  'Findメソッドの結果を変数で受ける
Set findCell = Columns("C").Find(What:="古川")
  'Nothingで無ければ対象セルが見つかっている
If Not findCell Is Nothing Then
  findCell.Select
  MsgBox "検索対象が見つかりました：" & findCell.Address
Else
  MsgBox "対象セルは見つかりませんでした"
End If

'すべて検索して処理
Dim findCell As Range, firstCell As Range, targetRng As Range
  '検索範囲をセット
Set targetRng = Columns("C")
  '初回検索はFindメソッド
Set findCell = targetRng.Find(What:="古川")
  '見つからなければ処理を終了
If findCell Is Nothing Then
  MsgBox "対象セルは見つかりませんでした"
  Exit Sub
End If
  '初回のセルを記録しておく
Set firstCell = findCell
  '2回目以降はFindNextメソッド
Do
  '見つかったセルに対する処理を記述
  '何らかの処理
  '「次のセル」を検索
  Set findCell = targetRng.FindNext(After:=findCell)
Loop Until findCell.Address = firstCell.Address
```
