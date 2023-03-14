# VBAのデータ単純入力

## セルの値と見た目の変更

- 各値の入力
  - 数式はValueでも大丈夫だが、厳密にはFormula
  - 日本独自の値は`FormulaLocal`や`FormulaR1C1Local`プロパティを使用
  - スピル形式は`Formula2`プロパティを使用

  ```vba
  'セルに入力
  Range("C2").Value = "VBA"
  Range("C3").Value = 1800
  Range("C4").Value = #6/5/2023#
  Range("C5").Formula = "=10*5"
  
  'セル範囲にまとめて入力
  Range("C7:E7").Value = Array(1, 2, 3)
  Range("C9:E10").Value = [{4,5,6;7,8,9}] '※2次元配列
  
  '相対参照で数式を入力
  Range("D3:D6").FormulaR1C1 = "=R[0]C[-2]*R[0]C[-1]"
  
  'スピル
  Range("E3").Formula2 = "=C3:C7*D3:D7"
  ```

- 各値の消去
  - 値のみ: `対象セル.ClearContents`
  - 結合セル対策
    - 対策1: 結合セル内の左上セルのValueプロパティに空白文字列を入力
    - 対策2: MergeAreaプロパティを使用、`Range("B4").MergeArea.ClearContents`

- フォント設定: `対象セル.Font.プロパティ = 設定値`
- 書式設定: `対象セル.NumberFormatLocal = 書式設定文字列`
  - 書式設定文字列は[プレースホルダー]を使用(`@`文字列、`#`数値など)
- 背景色設定: `対象セル.Interior.管理形式 = 色`、形式はRGB・パレット・テーマカラー
  - 背景色クリア: `対象セル.Interior.Pattern = xlNone`
- 罫線設定: `セル範囲.Borders.(場所を指定する定数).罫線の設定`、Borderオブジェクトを使用している
- セルサイズ設定
  - 幅: `対象セル.ColumnWidth = 幅`
  - 高さ: `対象セル.RowHeight = 高さ`
  - 自動調整: `対象セル.AutoFit`
- 位置揃え
  - 水平: `対象セル.HorizontalAlignment = 水平位置`
  - 垂直: `対象セル.VerticalAlignment = 垂直位置`
- 折り返しなど
  - 折り返し: `対象セル.WrapText = True`
  - 縮小表示: `対象セル.ShrinkToFit = True`

  ```vba
  '罫線を引く
  With Range("B2:D6")
    '上側の罫線を設定
    With .Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .Weight = xlMedium
      .ThemeColor = msoThemeColorAccent6
    End With
    '下側の罫線を設定
    With .Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .Weight = xlMedium
      .ThemeColor = msoThemeColorAccent6
    End With
    '中間のうち、横方向の罫線を設定
    With .Borders(xlInsideHorizontal)
      .LineStyle = xlContinuous
      .Weight = xlHairline
      .ThemeColor = msoThemeColorAccent6
    End With
  End With

  '罫線をクリア
  Range("B2:D6").Borders.LineStyle = xlNone
  
  'セルサイズ設定
  With Range("B2")
    .ColumnWidth = 10
    .RowHeight = 40
  End With
  
  'セル自動調整+α
  Dim rng As Range
  With Range("B2").CurrentRegion
    '列幅設定
    .EntireColumn.AutoFit
    For Each rng In .Columns
      rng.ColumnWidth = rng.ColumnWidth + 2
    Next
    '行の高さ設定
    .EntireRow.AutoFit
    For Each rng In .Rows
      rng.RowHeight = rng.RowHeight + 10
    Next
  End With
  
  '水平位置揃え
  Range("B2").HorizontalAlignment = xlLeft     '左
  Range("B4").HorizontalAlignment = xlCenter   '中央
  Range("B6").HorizontalAlignment = xlRight    '右
  '垂直位置揃え
  Range("D2").VerticalAlignment = xlTop      '上端
  Range("D4").VerticalAlignment = xlCenter   '中央
  Range("D6").VerticalAlignment = xlBottom   '下端
  ```

## コピー

1. 基本: `転記元セル範囲.Copy Destination:=転記先起点セル`
2. Valueを使用: `転記先セル範囲.Value = 転記元セル範囲.Value`
3. PasteSpecial
   - 転記元: `転記元セル範囲.Copy`
   - 転記先: `転記先セル範囲.PasteSpecial 形式`
   - コピー状態の解除: `Application.CutCopyMode = False`

```vba
'セル範囲のコピー
Range("D2:E5").Copy Destination:=Range("D7")

'数式の結果のみに変換
With Range("D3:D5")
  .Copy
  .PasteSpecial xlPasteValues
End With

'セル幅を含めてコピー
With Range("G2")
  .PasteSpecial xlPasteColumnWidths
  .PasteSpecial xlPasteAll
End With

'数式の内容をコピー
Range("G5").PasteSpecial xlPasteFormulas
```

---
[プレースホルダー]: https://learn.microsoft.com/ja-jp/office/vba/api/excel.range.numberformatlocal
