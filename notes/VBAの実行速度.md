# VBAの実行速度

## 簡易計測方法

- Timer関数
  - 午前0時からの経過時間を返す
  - 日を跨ぐタイミングでは計測できない
- コード内に記述した処理の実行時間を取得して表示

  ```vba
  Sub timeMacro()
    Dim startTime As Double
    startTime = Timer
    '何らかの処理
    MsgBox "処理速度: " & Format(Timer-startTime, "0.00秒")
  End Sub
  ```

## 高速化の基本

- セルへのアクセスを減らす、特に値の代入に時間がかかる
- 更新や再計算を一時止めて、処理を実行
  1. 画面更新をオフ/オン: マクロでOFFにしても実行後には自動的にONになる
  2. 数式の計算をストップ: 処理の実行後に元の設定に戻す
  3. イベント処理を止める: 処理の実行後にONに戻す
     - イベントの連鎖を止めるのにも使用、Changeイベントなど注意
  4. 警告・メッセージをスキップ: 処理の実行後にONに戻す

  ```vba
  'コード1
  '画面更新オフ
  Application.ScreenUpdating = False
  '※実行したい処理
  '画面更新オン
  Application.ScreenUpdating = True

  'コード2
  'マクロ実行時の再計算設定を保持
  Dim calcMode As XlCalculation
  calcMode = Application.Calculation
  '自動再計算オフ
  Application.Calculation = xlCalculationManual
  '※実行したい処理
  '元に戻す
  Application.Calculation = calcMode

  'コード3
  'イベント処理をオフ
  Application.EnableEvents = False
  '※実行したい処理
  'オンに戻す
  Application.EnableEvents = True

  'コード4
  '警告メッセージ表示をオフ
  Application.DisplayAlerts = False
  '※削除処理等の警告・確認が表示される処理
  'Worksheets.Add.Delete
  '警告メッセージ表示をオンに戻す
  Application.DisplayAlerts = True
  ```
