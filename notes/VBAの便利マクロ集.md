# VBAの便利マクロ集

- 2次元配列か判定

  ```vba
  Function isSingleDimension(arr As Variant) As Boolean

    Dim tmp As Variant
    '引数の2次元目にアクセスしてみてエラーならば1次元と判定
    On Error Resume Next
    tmp = LBound(arr, 2)
    isSingleDimension = IIf(Err.Number = 0, False, True)

    'エラートラップを元に戻す
    Err.Clear
    On Error GoTo 0

  End Function
  ```
