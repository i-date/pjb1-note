# 一時メモ

## Youtube

- エクセルVBAマクロとは？できることを完璧に学ぶ初心者入門講座【たった1動画で全てが分かるExcelVBAの教科書】[^1]
- マクロちゃんねる[^2]

- ボタンの作成
  - 開発タブ→挿入→ボタン
  - 動かしたいマクロを選択
  - ボタンを右クリックで、ボタンテキストやサイズを変更

- 最終行の取得

  ```vba
  Dim rowNo As long
  rowNo = Cells(Rows.Count, 1).End(xlUp).Row
  ```

## 命名規則

- https://excel-toshokan.com/vba-naming-rules/
- https://excel-ubara.com/excelvba4/EXCEL274.html

---
[^1]: https://www.youtube.com/watch?v=949u36vdN7U&t=114s
[^2]: https://www.youtube.com/@macro-chan/videos
