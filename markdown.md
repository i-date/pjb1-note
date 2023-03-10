# Markdown Note[^1]

## リスト

### 混合

番号なしリスト、番号ありリストは混ぜて使用できない  
ただし、インデントを設ければOK  

```markdown
- リスト
  1. 1
  2. 2

1. リスト
   - 1
   - 2
```

### インデント

番号の有無でインデントが異なる

```markdown
- 番号なしリスト
  インデント半角スペース2つ分
1. 番号ありリスト
   インデント半角スペース3つ分
```

## チェックボックス

```markdown
- [ ] 1
- [x] 2
```

## コードブロック

````text
```コードブロック内の言語名
内容
```
````

### コードブロック内のコードブロック

- 「開始と終了をバッククオート4個」で囲む
- 「行頭スペース4個」スタイルのコードブロックを使う

## テーブル

```markdown
| Left align | Right align | Center align |
|:-----------|------------:|:------------:|
| This       | This        | This         |
| column     | column      | column       |
| will       | will        | will         |
| be         | be          | be           |
| left       | right       | center       |
| aligned    | aligned     | aligned      |
```

以下表示例

| Left align | Right align | Center align |
|:-----------|------------:|:------------:|
| This       | This        | This         |
| column     | column      | column       |
| will       | will        | will         |
| be         | be          | be           |
| left       | right       | center       |
| aligned    | aligned     | aligned      |

## 画像

HTMLで入力すれば、画像サイズの調整も可能

```markdown
![代替テキスト](画像のURL "画像のタイトル")
```

## 複数リンク

```markdown
[ここ][link-1] と [この][link-1] リンクは同じ
[link-1] という書き方も可能

[link-1]: http://qiita.com/
```

## リンクカード

```markdown
(空行)
https://qiita.com/Qiita/items/c686397e4a0f4f11683d
(空行)
```

## 脚注

```markdown
本文中に [^1]
脚注として [^1]:...
```

### ダイアグラム

- PlantUML[^2]を使う
- Mermaid[^3]を使う

---
[^1]: https://qiita.com/Qiita/items/c686397e4a0f4f11683d
[^2]: https://plantuml.com/ja/
[^3]: https://mermaid.js.org/#/
