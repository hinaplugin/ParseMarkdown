# Office → Markdown 変換ツール

このツールは、`file` フォルダにある以下のファイルをすべて  
Markdown（`.md`）に変換し、`out` フォルダへ出力します。

対応形式:

- Word: `.docs`, `.docx`
- Excel: `.xlsx`
- PowerPoint: `.pptx`

すべて **完全オフライン** かつ **npm パッケージのみ** で動作します。  
社内文書など、外部へアップロードできない環境でも安全に利用できます。

---

## ディレクトリ構成

data/  
├ parse.js  
├ file/  
│ ├ sample.docs  
│ ├ sample.docx  
│ ├ sheet.xlsx  
│ └ slide.pptx  
└ out/

```yaml
- `file/`  
  変換したい Office ファイルを配置します。

- `out/`  
  変換後の `.md` ファイルが自動生成されます。

- `parse.js`  
  変換処理本体です。
```

---

## セットアップ

プロジェクトのルートで以下を実行します。

```bash
npm install mammoth turndown xlsx pptx2json
```
Node.js は v18 以降 を推奨します。

実行方法
```bash
node parse.js
```
実行すると、file フォルダ内の

- .docs

- .docx

- .xlsx

- .pptx

がすべて処理され、同名の .md ファイルとして out フォルダに出力されます。

※ 初回実行でfile/とout/が自動作成されます．

例:

```bash
file/word1.docs  -> out/word1.md
file/word2.docx  -> out/word2.md
file/sheet.xlsx  -> out/sheet.md
file/slide.pptx  -> out/slide.md
```
注意事項
- `.docs`は実体が`.docx`（OOXML）である前提です  
※ 古いバイナリ形式（.doc）の場合は変換できません

- 変換結果は元文書の構造を可能な限り保持しますが、  
Pandoc などの専用ツールほど完全ではありません

- 画像の抽出や高度なレイアウト保持が必要な場合は、  
追加の処理が必要になる場合があります

このツールは、

- 社内ネットワーク

- オフライン環境

- 機密文書

といった制約下でも安全に利用できる、ローカル完結型の変換スクリプトです。