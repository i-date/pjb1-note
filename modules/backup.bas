Attribute VB_Name = "backup"
Option Explicit

'=======================================================
'【概  要】バックアップ作成のメイン処理
'【引  数】引数名    説明
'          ---------------------------------------------
'          なし
'          ---------------------------------------------
'【戻り値】なし
'【備  考】なし
'=======================================================
Public Sub save()

  Dim book As Workbook, path As String
  Set book = ActiveWorkbook
  
  'バックアップフォルダの存在確認/作成
  path = book.path & "\" & Split(book.Name, ".")(0) & "_backup"
  makeFolder path
  
  'サブフォルダ(日付ごと)の存在確認/作成
  path = path & "\" & Format(Date, "yyyymmdd")
  makeFolder path
  
  'バックアップの作成
  path = path & "\" & getNameWithTimeStamp(book)
  book.SaveCopyAs path
  
End Sub

'=======================================================
'【概  要】フォルダ作成
'【引  数】引数名    説明
'          ---------------------------------------------
'          path      String型、フォルダパス
'          ---------------------------------------------
'【戻り値】なし
'【備  考】なし
'=======================================================
Private Sub makeFolder(path As String)
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  '対象フォルダの存在確認(ある: 処理なし、ない: 作成)
  If Not fso.FolderExists(path) Then
      fso.CreateFolder path
  End If
End Sub

'=======================================================
'【概  要】ファイル名に呼び出し時の時間を付与
'【引  数】引数名    説明
'          ---------------------------------------------
'          book      Workbook型
'          ---------------------------------------------
'【戻り値】呼び出し時の時間を付与したファイル名
'【備  考】マクロの有効/無効で拡張子を選択している
'=======================================================
Private Function getNameWithTimeStamp(book As Workbook) As String

  Dim baseName As String, suffix As String, extension As String
  
  '名前・接尾辞・拡張子の用意
  baseName = Split(book.Name, ".")(0)
  suffix = Format(Now, "_hhnn")
  extension = IIf(book.HasVBProject, ".xlsm", ".xlsx")
  
  getNameWithTimeStamp = baseName & suffix & extension
End Function
