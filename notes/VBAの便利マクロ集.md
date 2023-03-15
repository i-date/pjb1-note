# VBAの便利マクロ集

- 2次元配列か判定する関数

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

- 引数に指定したフォルダー内のExcelブックのパス文字列のリストを返す関数

  ```vba
  Function getBookPathList(folderPath As String) As Variant

    Dim dic As Object, tmpfile As Object, tmpExtension As String
    Set dic = CreateObject("Scripting.Dictionary")

    '引数に指定したパスのフォルダーから、拡張子が「xlsx」のファイルのパスを辞書登録
    With CreateObject("Scripting.FileSystemObject")
      For Each tmpfile In .GetFolder(folderPath).Files
        tmpExtension = .GetExtensionName(tmpfile)
        If tmpExtension = "xlsx" Then dic.Add tmpfile.Path, "dummy"
      Next
    End With
    getBookPathList = dic.Keys

  End Function
  ```

- OneDriveを想定したURLの変換を行うコード

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
