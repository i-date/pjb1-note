# VBAの入力インターフェース

## 入力専用画面

- メリット
  - 利用者側: 入力しやすい
  - 利用者側: 問い合わせしやすい
  - 開発者側: チェックすれば良い「セル範囲」を限定できる
  - 開発者側: 転記処理の一環としてチェックや変換処理を通すことで、意図しないデータの蓄積を防ぐ
  - 開発者側: 目的が明確な短いマクロを作成することになり、管理しやすくなる
  - 開発者側: 入力シートにボタンを配置すると、アクティブシートを固定できる(しかし、アクティブシートに依存しない実装がベターではある)

- 転記の際に検討する項目
  1. 転記後の表見出しの整理
  2. 見出し項目に合わせた形式のデータを作成する仕組みの整理(入力用シートから、どのセル値をどのように拾ってくるか)
  3. 新規レコード入力位置を取得して入力する仕組みの整理
  4. 用意した仕組みを組み合わせる
  ---
  | 項目(フィールド) | タイプ(セル範囲の固定/可変) | 場所(参照値のセル範囲) |
  |:---|:-----|:---|
  | 伝票番号 | 固定 | セルF1 |
  | 発行日 | 固定 | セルF2 |
  | 取引先 | 固定 | セルB4 |
  | 商品 | 可変(セル範囲B14:F23) | 1列目(B列) |
  | 単価 | 可変(セル範囲B14:F23) | 2列目(C列) |
  | 数量 | 可変(セル範囲B14:F23) | 3列目(D列) |

## フォームコントロール

- Excel開発タブ→挿入→フォームコントロール
- 共通の仕組み: 右クリック→コントロールの書式設定→各種設定ができる
- VBAからアクセス: 固有機能を使用するなら方法2を使用
  1. WorksheetオブジェクトのShapesプロパティで図形として取得→ControlFormatプロパティでアクセス
  2. 各コントロールに対応したWorksheetオブジェクトの専用メソッド
- 代表的なコントロール
  - リストボックス
  - ドロップダウンリストボックス
  - チェックリストとオプションボタン: このオプションボタンは複数の候補から1つ選択(オプション選択)
  - グループボックスの中にオプションボタン: スコープがグループボックス内になるため、複数選択のオプションボタンが作成できる
  - スピンボタン

## シート自体をカスタムオブジェクトと捉える

- 入力用シートなど、用途が限定されたシートを作成する場合
- シート上の操作に関することは、シートのオブジェクトモジュールに記述するスタイル
- 特定のシートをカスタムオブジェクトと捉え、独自のプロパティやメソッドを追加し、利用する

- 作業方針
  1. VBEにてシートのオブジェクト名を用途に合わせて変更
     - 例: InputSheet、DataSheet
     - ↑で指定したオブジェクト名でアクセスできるようになる
  2. データをやり取りしやすいようにカスタムクラスを作成
     - クラスモジュール作成(プロパティとメソッドを記述)
  3. 1の入力用オブジェクトシートにデータをまとめる処理を追加
     - 入力シート(InputSheet)に、2のカスタムクラスを使用した処理を記述
  4. 1の蓄積用オブジェクトシートに転記処理を追加
     - 蓄積シート(DataSheet)に、2のカスタムクラスを使用した処理を記述
- 注意
  1. シートモジュールで優先されるのは、そのシートのプロパティやメソッドの「名前(識別子)」、オブジェクトのスコープで評価
     - 例: 普通だと`Range("A1")`はアクティブシートのセルA1→シートモジュールだと`Range("A1")`はそのシートのセルA1
  2. シート上に追加するプロパティやメソッド、そして引数のデータ型には、カスタムクラス型は指定できない

  ```vba
  '方針3のコード例
  '以下のプログラムでは、SalesDataとItemがカスタムクラス
  'ちなみに、この処理を記述した関数の返り値にはSalesData型は指定できない(上記の注意2)
  '新規SalesData作成
  Dim data As SalesData
  Set data = New SalesData

  'シート上の基本データをセットしていく
  data.ID = Range("F1").Value
  data.CreateDate = Range("F2").Value
  data.Customer = Range("B4").Value
  data.Amount = Range("C9").Value
  data.Tax = Range("F25").Value

  'シート上の明細データをセットしていく
  Dim itemRange As Range, record As Range, newItem As Item
  '明細のセル範囲をセット
  Set itemRange = Range("B14:F23")
  '明細の1行ずつについて走査
  For Each record In itemRange.Rows
    If record.Cells(1).Value = "" Then Exit For
    Set newItem = New Item
    newItem.Name = record.Cells(1)
    newItem.Price = record.Cells(3)
    newItem.Number = record.Cells(4)
    newItem.Subtotal = record.Cells(5)
    '明細としてコレクションに追加
    data.Items.Add newItem
  Next

  '作成したSalesDataを返す
  Set CreateSalesData = data
  ```

  ```vba
  '方針4のコード例
  '以下コードのSalesは引数として渡される
  '既に同じ伝票番号がある場合はエラーを出す
  If Not Columns("B").Find(Sales.ID) Is Nothing Then
   Err.Raise 513, Description:="既に同じ伝票番号のデータが存在しています"
  End If

  'カスタムクラスの型の変数に移し替える
  'コードヒントが表示されるようになる
  Dim data As SalesData
  Set data = Sales

  '明細の数だけループ
  Dim i As Long
  For i = 1 To data.Items.count
    With getNewRecordRange      '新規データ入力セル範囲を取得
      .Cells(1).Value = data.ID
      .Cells(2).Value = data.CreateDate
      .Cells(3).Value = data.Customer
      .Cells(4).Value = data.Items(i).Name
      .Cells(5).Value = data.Items(i).Price
      .Cells(6).Value = data.Items(i).Number
      .Cells(7).Value = data.Items(i).Subtotal
    End With
  Next
  ```
