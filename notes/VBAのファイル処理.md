# VBAのファイル処理

## ファイル処理をする方法の概要

1. Openステートメントを使用
2. FileSystemObjectを使用

## ブックの取得と保存

- 基本的な開く操作`Open`と閉じる操作`Close`
  - 注意: 開いたファイルがアクティブなブックになる
  - 対策: ThisWorkbookプロパティ、アクティブシートに依存しないで「そのマクロが記述されているBook」を返す
- 相対パス
  - Pathプロパティ: そのBookが保存されているフォルダまでのパスを返す
  - 例: `ActiveWorkbook.Path`や`ThisWorkbook.Path`

  ```vba
  '他のBookを開く
  Dim 変数 As Workbook
  Set 変数 = Workbooks.Open(パス)
  変数.Close
  
  '変更を保存せずに閉じる
  ActiveWorkbook.Close SaveChanges:=False
  '変更を保存して閉じる
  ActiveWorkbook.Close SaveChanges:=True
  
  '「このブック」に転記して閉じる
  Dim book As Workbook
  Set book = Workbooks.Open("C:\excel\支店データ.xlsx")
  book.Worksheets(1).Range("B2").CurrentRegion.Copy
  ThisWorkbook.Worksheets(1).Range("B2").PasteSpecial xlPasteAll
  book.Close
  ```

- 3種類の保存方法
  1. `Workbookオブジェクト.Save`: 上書き保存
  2. `Workbookオブジェクト.SaveAs ファイル名`: 名前をつけて保存
  3. `Workbookオブジェクト.SaveCopyAs ファイル名[, パスワード]`: ブックコピーを作成して保存
- マクロの有無判定
  - HasVBProjectプロパティ
  - 保存は、SaveAsプロパティの引数に`FileFormat:=xlOpenXMLWorkbookMacroEnabled`

  ```vba
  'パスワード付きで保存
  ActiveWorkbook.SaveAs _
    Filename:="C:\excel\バックアップ\売上データ.xlsx", _
    Password:="pass"

  'パスワードがかかったブックを開く
  Workbooks.Open _
    Filename:="C:\excel\バックアップ\売上データ.xlsx", _
    Password:="pass"
    
  '指定フォルダー内にxlsm形式で保存
  Dim book As Workbook: Set book = ActiveWorkbook
  book.SaveAs _
    Filename:=パス & ブック名 & ".xlsm", _
    FileFormat:=xlOpenXMLWorkbookMacroEnabled
  ```

## 複数ブックをまとめて扱う

- 基本的な考え方
  1. 処理対象のブックリストを作成
     - ブック名のリストを作成(Workbookオブジェクトのリストより手軽)
     - Array関数を使用したり、専用シートを作成
  2. ループ処理

## ファイルやフォルダーの操作

- FileSystemObject(FSO)
  - Microsoft Scripting Runtimeに用意されているファイル/フォルダ操作に特化したオブジェクト
  - FSOオブジェクト作成: `CreateObject("Scripting.FileSystemObject")`
  - FolderオブジェクトやFileオブジェクトを取得できる
- FSO使用方法
  1. CreateObjectでFSOオブジェクト作成
  2. FolderオブジェクトやFileオブジェクトを取得
  3. プロパティとメソッドでフォルダ/ファイルを操作

  ```vba
  'ファイルのコピー
  Dim fso As Object, Path As String
  Set fso = CreateObject("Scripting.FileSystemObject")
  Path = "C:\excel\バックアップ\売上データ.xlsx"
    'ファイルが存在する場合はコピー
  If fso.FileExists(Path) Then
    fso.GetFile(Path).Copy _
      Replace(Path, ".xlsx", "_バックアップ.xlsx")
  End If

  'フォルダーごとコピー
  Dim tmpFolder As Object
    'マクロを記述したブックと同じ場所にある「支店データ」フォルダーを取得
  Set tmpFolder = _
    CreateObject("Scripting.FileSystemObject") _
      .GetFolder(ThisWorkbook.Path & "\支店データ")
  tmpFolder.Copy ThisWorkbook.Path & "\支店データ_バックアップ"

  'ファイル名のリネーム
  CreateObject("Scripting.FileSystemObject") _
    .GetFile("C:\excel\バックアップ\売上データ.xlsx") _
    .Name = "変更後の名前.xlsx"
  ```

- フォルダ/ファイルの選択を行うダイアログ
  - FileDialogオブジェクト、4種類用意されている
  - `Application.FileDialog(ダイアログの種類)`
    1. ファイルを開く    : msoFileDialogOpen
    2. ファイルを保存する: msoFileDialogSaveAs
    3. ファイルを選択する: msoFileDialogFilePicker
    4. フォルダを選択する: msoFileDialogFolderPicker
- FileDialogの使用方法
  - Application.FileDialogでダイアログを取得
  - 下準備を各種プロパティで設定
  - Showメソッドで表示

  ```vba
  'フォルダ選択ダイアログ
  Dim fd As FileDialog
    'フォルダ選択ダイアログを取得
  Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    '表示タイトルと初期フォルダー設定
  With fd
    .Title = "フォルダーを選択してください"
    .InitialFileName = ThisWorkbook.Path
    .InitialView = msoFileDialogViewSmallIcons
  End With
    '表示して選択結果を取得
  If fd.Show = 0 Then
    Debug.Print "選択をキャンセルしました"
  Else
    Debug.Print "選択フォルダー名：", fd.SelectedItems(1)
  End If
  ```

## OneDrive上のファイルを扱う注意点

1. ローカルPC上のOneDriveに関連付けられたフォルダ内に保存
2. クラウド上のストレージとOneDrive用のフォルダを同期

- Pathプロパティ→クラウド側`https://～`が返る
  - クラウド判定: 「https」で始まるかどうか
  - 変換: VBAからはEnviron関数で取得できる
    - OneDrive for Business: `Environ("OneDriveCommercial")`
    - OneDrive: `Environ("OneDrive")`もしくは`Environ("OneDriveConsumer")`
    - 下記に、OneDriveを想定したURLの変換を行うコード

  ```vba
  Dim url As String, localPath As String
  Dim oneDrivePath As String, cloudPattern As String

  'アクティブなブックのPathを取得
  url = ActiveWorkbook.Path

  'クラウド上でなければ終了
  If Not url Like "https*" Then
    Debug.Print "ローカルのブックです"
    Exit Sub
  End If

  'OneDrive用フォルダーのパスを取得
  '個人用
  oneDrivePath = Environ("OneDrive")
  If oneDrivePath = "" Then oneDrivePath = Environ("OneDriveConsumer")

  '法人用
  'oneDrivePath = Environ("OneDriveCommercial")
  'If oneDrivePath = "" Then oneDrivePath = Environ("OneDrive")

  'パターン文字列指定
  cloudPattern = "https://d.docs.live.net/[^/$]+" '個人用
  'cloudPattern = "https://.+my.sharepoint.com/personal/[^/$]+"     '法人用

  'パターン文字列の範囲をOneDrive用フォルダーのパスに置換
  With CreateObject("VBScript.RegExp")
    .Pattern = cloudPattern
    .IgnoreCase = True
    .Global = False
     localPath = .Replace(url, oneDrivePath)
  End With

  '残る「/」をセパレーターに変換
  localPath = Replace(localPath, "/", Application.PathSeparator)

  '出力
  Debug.Print "変換後："; localPath
  Debug.Print "変換前："; url

  End Sub
  ```
