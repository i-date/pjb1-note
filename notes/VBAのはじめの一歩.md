# VBAのはじめの一歩

## 基本

- ウィンドウ
  - モジュールウィンドウ
  - プロパティウィンドウ
  - コードウィンドウ

- モジュール...基本的には標準モジュール（Book全体で使用可能）、挿入→標準モジュール
- プロシージャ...マクロの最小単位(別言語でいう関数やルーチンのこと?)

  ```vba
  Sub 名前()
    処理
  End Sub
  ```

- 3要素
  - 関数
  - ステートメント...マクロ動作を制御する命令
  - オブジェクト式...エクセルを操作するときのコードの記述方法
    - プロパティ: `オブジェクト.プロパティ = 値`→`Range("A").Value = 100`
    - メソッド: `オブジェクト.メソッド オプション`→`Range("A").Delete Shift:=xlToLeft`

- 型注意
  - 配列
    - 静的配列
    - 動的配列
  - Variant型...Splitが使用可能、インデックスは1開始

## 基本構文

- 選択...特定範囲の場合`Range`、状況に応じて変化する場合`Cells`
  - `Range("A1")`
  - `Range("A1:C5")`or`Range("A1","C5")`
  - `Cells(2,3)`or`Cells(2,"C")`
  - `Rows(行数)`...行指定
  - `Columns(列名)`...列指定
    - 列を数値で指定したい場合: `Range(Columns(1), Columns(4+1))`
  - `Sheets("シート名")`or`Sheets(左からの順番)`
    - Sheetsは、挿入時に選択できる4種類すべてを対象
    - WorkSheetsは、ワークシートのみ
  - `Workbooks("ブック名+拡張子")`or`Workbooks(開いた順番)`

- プロパティ
  - Value...セルの値
  - Text...セルに表示されている値(カンマなどが含まれる)
  - End...control+shift+方向キーの処理
  - Offset...基準となるセルからの位置
  - ActiveCell...アクティブセルを表すRangeオブジェクトを返す
  - Selection...選択しているセル(範囲)を返す
  - Count...数を返す

- メソッド
  - Active...アクティブセルをどこにするか選択、複数選択でもアクティブセルは一つ
  - Select...指定したオブジェクトを選択
  - Copy...コピー&ペースト`コピー元セル.Copy Destination:=コピー先のセル`(`Destination:=`は省略可能)
  - PasteSpecial...形式を選択して貼り付け`ペースト先.PasteSpecial paste:=フォーマット形式`(`paste:=`は省略可能)
  - Clear...削除
  - ClearFormats...書式だけ削除
  - ClearContents...値だけ削除
  - Delete...オブジェクトごと削除、右クリックから削除を選択した場合の処理`セル.Delete shift:=フォーマット形式`(`shift:=`は省略可能)
  - Close...ブックを閉じる
  - Open...引数にpath¥ファイル名+拡張子

- 変数...宣言的
  - `Option Explicit`...最初の行で指定した場合、変数を宣言しないと使用できない
  - モジュールレベル変数...宣言セクションで指定した変数
  - `Dim 変数名 As データ型`
  - `Public 変数名 As データ型`
  - 配列
    - 0から始める場合:   `Dim 配列名(要素数-1) As データ型`
    - 0から始めない場合: `Dim 配列名(1 To 要素数) As データ型`

- ステートメント
  - For ~ Next

    ```vba
    Dim 変数名 As データ型
    For 変数 = 初期値 To 終了値
      処理
    Next 変数
    ```

  - If

    ```vba
    If 条件 Then 処理

    または

    If 条件1 Then
      条件1を満たした場合の処理
    ElseIf 条件2 Then
      条件2を満たした場合の処理
    Else
      何も満たしてない場合の処理
    End If
    ```

  - Select Case

    ```vba
    Select Case 値
      Case 条件1
        処理
      Case 条件2, 条件3
        処理
      Case Else
        その他の処理
    End Select

    条件で比較する場合: Case Is >= 10000
    ```

  - Do Loop...無限ループにならないように注意
    - 条件 Until...条件を満たすとループ終了
    - 条件 While...条件を満たす間ループ

    ```vba
    前判定
    Do 条件
      処理
    Loop

    後判定
    前判定
    Do
      処理
    Loop 条件
    ```
