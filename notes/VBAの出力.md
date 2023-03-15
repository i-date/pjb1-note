# VBAの出力

## 印刷

- シートごとにPageSetupオブジェクトにアクセスして設定
- PageSetupオブジェクト: WorksheetのPageSetupプロパティ経由で取得
- プレビュー画面で確認: `印刷対象オブジェクト.PrintPreview`
- 印刷実行: `印刷対象オブジェクト.PrintOut`
- 印刷設定を行うマクロと印刷を行うマクロは分ける
  - 理由: 印刷設定を行う処理は、プリンタドライバとの通信に時間がかかるため

``` vba
'印刷設定
  'プリンターへの通信を一時的にオフ、PrintCommunicationはExcel2010以降
Application.PrintCommunication = False
  '各種印刷設定を行う
With ActiveSheet.PageSetup
  '印刷範囲をアドレス「文字列」で指定
  .PrintArea = ActiveSheet.Range("B4:I52").Address
  '印刷方向は「縦」
  .Orientation = xlPortrait
  'ズーム設定を自動に設定
  .Zoom = False
  '行・列全てが「1」ページに収まるように拡大率を自動調整
  .FitToPagesTall = 1
  .FitToPagesWide = 1
End With
  'プリンタへの通信を元に戻す
Application.PrintCommunication = True

'アクティブシートの印刷プレビュー
ActiveSheet.PrintPreview

'アクティブシートの印刷実行
ActiveSheet.PrintOut

'印刷関連の点線消去
ActiveSheet.DisplayPageBreaks = False
```

## PDF

- ExportAsFixedFormatメソッドを使用
  - 引数TypeにxlTypePDFを指定

  ```vba
  'PDF書き出し
  ActiveSheet.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=ThisWorkbook.Path & "\PDF出力.pdf"
  ```

## 結果Bookの送信準備

1. 非表示をチェック

   ```vba
   Sub checkHidden()

     Dim sht As Worksheet
     Dim isVisible As Boolean, hasHidden As Boolean, hasGroup As Boolean

     '非表示シートのチェック
     For Each sht In Worksheets
     
       'Visibleプロパティの値で非表示チェック
       isVisible = sht.Visible = xlSheetVisible
       'Areas.Countで非表示行・列のチェック
       hasHidden = sht.Cells.SpecialCells(xlCellTypeVisible).Areas.Count > 1
       'OutLineLevelでグループ化の有無をチェック
       hasGroup = IsNull(sht.Rows.OutlineLevel) Or IsNull(sht.Columns.OutlineLevel)
     
       '結果を出力
       Debug.Print "シート名：" & sht.Name
       Debug.Print "  表示      ：" & IIf(isVisible, "〇", "×")
       Debug.Print "  全表示    ：" & IIf(hasHidden, "非表示アリ", "〇")
       Debug.Print "  グループ化：" & IIf(hasGroup, "要確認", "〇") & vbCrLf
     
     Next
     
   End Sub
   ```

2. セル「A1」を選択

   ```vba
   Sub selectA1()

     Dim i As Long
     '最後に1枚目のシートが選択されるよう、逆順にループ
     For i = Worksheets.Count To 1 Step -1
       '非表示シートは処理から除く
        If Worksheets(i).Visible = xlSheetVisible Then
          Application.Goto Worksheets(i).Range("A1"), Scroll:=True
        End If
     Next

   End Sub
   ```

3. Bookの作成者や編集者をチェック
   - BuiltinDocumentPropertiesプロパティ経由でアクセス
   - 自社以外の人が作成してたりする場合に修正

   ```vba
   Sub getDocAuthor()

     With ActiveWorkbook.BuiltinDocumentProperties
       Debug.Print "作成者："; .Item("Author")
       Debug.Print "最終編集者：", .Item("Last author")
     End With

   End Sub

   'ドキュメントプロパティ設定
   ActiveWorkbook. _
     BuiltinDocumentProperties("Author").Value = "Authorの名前"
   ```
