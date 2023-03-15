# VBAとWebデータ

## 基本: コピーして整形

- Webブラウザ上で表形式で表示されているデータを貼付→「何かしらの規則性」を持って貼付
- 作業方針
  1. 取り出したいデータの表見出しを記述
  2. データの規則性を見つける
  3. 規則性に沿った大まかなループ処理の枠組みを作成(大抵行方向・列方向の2重ループ)
  4. 個々の値を取り出す際の変換処理を作成

## Power Query[^1]

- Power Queryによる取り込み・変換機能を使用するのが一番手軽
- WebページのURL指定→データの変換→必要部分の読み取り
- 方法
  1. Web.Page関数でWebページのデータを読み込む
     - `Web.Page(Web.Contents("WebページのURL"))`
     - 返り値: ナビゲーションテーブル
     - Data列にデータ候補がテーブル形式で保管`Web.Page(Web.Contents("WebページのURL"))[Data]{インデックス}`
  2. Table.UnpivotOtherColumns関数やTable.Unpivot関数でピボット解除
     - `Table.UnpivotOtherColumns(テーブル, 対象外の列名リスト, 見出し名の値の列名, クロス位置の値の列名)`
  3. Table.SplitColumn関数で区切り文字をキーとして列を分解
     - `Table.SplitColumn(テーブル, 列名, 分割ルール, 新しい列名のリスト)`
  4. Table.ReplaceValue関数で値を修正
     - `Table.ReplaceValue(テーブル, 検索値, 置換値, 対象列)`

## Webページのソース解析

- ソースを取り込み解析→各種ライブラリを利用

1. 任意のWebページのデータを取得する
   - Microsoft WinHTTP Services Library
   - WinHttpRequestオブジェクト
     - Openメソッド: URL指定
     - Sendメソッド: リクエスト送信
     - Statusプロパティ: ステータスコード確認
     - StatusTextプロパティ: ステータスコードのテキストを取得
     - ResponseTextプロパティ: レスポンス
   - 下記コード、Webページのソースを取得

   ```vba
   Dim httpRequest As WinHttpRequest

   '新規のWinHttpRequestリクエスト生成
   Set httpRequest = New WinHttpRequest

   'GETでリクエスト送信
   With httpRequest
     .Open "GET", "https://www.sbcr.jp/pc/"
     .send
   End With

   'ステータスコードを確認
   Do While httpRequest.Status <> 200
     Select Case httpRequest.Status
     Case Is > 399
       MsgBox "エラー発生：" & httpRequest.Status
       Exit Sub
     Case Else '状態を書き出し
       Debug.Print httpRequest.Status, httpRequest.StatusText
     End Select
     '読み込みフリーズしても実行停止できるようDoEvents
     DoEvents
   Loop

   'リクエストの結果テキストを表示
   MsgBox httpRequest.ResponseText
   Set httpRequest = Nothing
   ```

2. ソースを元にHTMLドキュメントとして解析
   - Microsoft HTML Object Library
   - HTMLDocumentオブジェクト
     - 変数: 頭に`I`が付与された、`IHTMLDocument`インターフェイス型
     - New演算子: 新規HTMLDocumentオブジェクト作成
     - writeメソッド: 読み込んだページのソースを書き込んでHTMLドキュメントとしてビルド
     - ビルド完了以降、各種プロパティ経由でHTMLの各要素にアクセスできる
     - getElementByIdメソッド: HTMLドキュメント内の任意の要素にアクセス
       - 返り値: HTMLElementオブジェクト
       - 個々のデータには、HTMLElementオブジェクトの各種プロパティ/メソッドでアクセス
     - getElementsByClassNameメソッド: HTMLドキュメント内の任意の要素のリストにアクセス
       - 返り値: HTMLElementCollectionオブジェクト
   - 注意
     1. JavaScriptなどを使用した動的なWebページは、ソースコードだけでは取得しきれない
     2. HTMLドキュメントとして正しく解析できないタイプのWebページ、正規表現など力技で対応

3. URLエンコードした値を取得
   - 検索エンジンなどでは、URLの末尾にパラメータ文字列を使用してアクセス
   - ↑のパラメータ文字列のURLエンコードが必須のケース
   - ENCODEURLワークシート関数を使用
     - 例: `Application.WorksheetFunction.EncodeURL("文字列")`

## XMLデータをHTMLDocumentで解析

- Webサイトのレスポンスに、XML形式のデータで返す場合もある
- 第一選択肢はPower Query
- VBA→DOMDocumentオブジェクト: 外部ライブラリMSXML2に用意されている
  - XML形式のドキュメントとしてパース
    - LoadXMLメソッド: 文字列からXML構築
    - Loadメソッド: ファイルからXML構築
  - ドキュメントツリーを移動して値にアクセス
    - FirstChildプロパティ
    - ChildNodesプロパティ
  - XPath式で指定
    - SelectSingleNodeメソッド
    - SelectNodesメソッド
- 名前空間が設定されている場合[^2]
  - SetPropertyメソッドでSelectionNamespaces属性を設定→名前空間のショートカットができる
- 下記コード、XML形式のデータを取り込み

  ```vba
  Dim httpRequest As WinHttpRequest, url As String
  Dim xml As Object, nodes As Object, node As Object

  '指定URLのソースを取得しHTMLドキュメントビルド
  url = "https://news.yahoo.co.jp/rss/topics/it.xml"
  Set httpRequest = New WinHttpRequest
  With httpRequest
    .Open "GET", url
    .send
    Do While .Status <> 200
        DoEvents
    Loop
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False
    xml.LoadXML .ResponseText
  End With

  '「item」要素をリストアップ
  Set nodes = xml.SelectNodes("//item")

  'リストアップしたentry要素から値をピックアップ
  Range("B3:C3").Select

  For Each node In nodes
    Selection.Value = Array( _
      node.SelectSingleNode("./title").Text(), _
      node.SelectSingleNode("./link").Text() _
    )
    Selection.Offset(1).Select
  Next
  ```

## JSON形式のデータ

- Power Queryにまかせる

---
[^1]: ExcelVBA[完全]入門, pp.605
[^2]: ExcelVBA[完全]入門, pp.620
