# Bulk Rename From CSV

## 最初に使うファイル

- まずは [`start.cmd`](./start.cmd) を実行してください
- 画面で確認しながら使う通常の入口です
- `rename.ps1` はコマンドラインで使いたいとき用です

[`rename.ps1`](./rename.ps1) は、CSVをもとに同一フォルダ内のファイル名を安全に一括変更するための PowerShell スクリプトです。

特徴:

- まずプレビューだけ出せる
- HTML と CSV の両方で差分を確認できる
- 元ファイルがない、変更先が重複する、同名ファイルが既にある、などを事前に検出する
- 問題がある行が1つでもあると `Apply` は止まる

GUI で使いたい場合:

- [`start.cmd`](./start.cmd) をダブルクリック
- もしくは [`rename-gui.ps1`](./rename-gui.ps1) を `powershell -ExecutionPolicy Bypass -File ...` で起動

GUI では次のことができます:

- CSV と対象フォルダの選択
- CSV / フォルダのドラッグ&ドロップ
- 入力内容に応じた自動プレビュー
- 選択した実行可能な行だけ `適用`
- プレビュー用の HTML / CSV を作らずにそのまま確認

## 1. まずプレビュー

入力CSVの `Content` を新しい名前に使うなら、次のように実行できます。

```powershell
powershell -ExecutionPolicy Bypass -File .\rename.ps1 `
  -CsvPath .\inputs\rename-map.csv `
  -TargetDirectory ..\path\to\wav `
  -NameTemplate "{No:0000}_{Content}{Ext}" `
  -Mode Preview
```

出力:

- `outputs\bulk-rename-preview.html`
- `outputs\bulk-rename-preview.csv`

## 2. 問題なければ実行

```powershell
powershell -ExecutionPolicy Bypass -File .\rename.ps1 `
  -CsvPath .\inputs\rename-map.csv `
  -TargetDirectory ..\path\to\wav `
  -NameTemplate "{No:0000}_{Content}{Ext}" `
  -Mode Apply
```

## 3. 明示的な変更先列を使う

CSVに `TargetName` 列を追加した場合は、テンプレートではなくその列を使えます。

```powershell
powershell -ExecutionPolicy Bypass -File .\rename.ps1 `
  -CsvPath .\inputs\rename-map.csv `
  -TargetDirectory ..\path\to\wav `
  -SourceColumn "File" `
  -TargetColumn "TargetName" `
  -Mode Preview
```

## 使えるテンプレート例

- `{No:0000}_{Content}{Ext}`
- `{Category}_{No:0000}_{Content}{Ext}`
- `{Content}{Ext}`

使える主なトークン:

- `No`
- `Category`
- `File`
- `FileBase`
- `Content`
- `Ext`

補足:

- ファイル名に使えない文字は `_` に置き換えます
- 末尾のピリオドや前後の空白は自動で除去します
- 既に同名ファイルがある場合は上書きせず停止します
