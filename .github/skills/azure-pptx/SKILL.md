---
name: azure-pptx
description: "Microsoft Azure PowerPoint Templateを使ったPPTXの作成、編集、内容確認、画像化を支援する。WHEN: pptx, PowerPoint, スライド, プレゼンテーション, deck, Azureテンプレート"
---

# Azure PPTX Skill

このスキルは、Microsoft Azure PowerPoint Templateを保持したままプレゼンテーションを作成、編集するためのリポジトリ固有手順を定義する。実装には公開仕様と外部OSSを使用する。

## 適用条件

- `.pptx` を作成または編集するときに使用する。
- 新規作成では `.github/skills/azure-pptx/assets/Microsoft-Azure-PowerPoint-Template-Light.potx` を必ず基にする。
- ユーザーがテンプレート不要と明示した場合だけ、テンプレートを使わない方法を検討する。
- テンプレート固有のレイアウト、フォント、プレースホルダーは [references/azure-template.md](references/azure-template.md) に従う。

## 作成前の確認

1. 内容に合うテンプレートレイアウトを選ぶ。
2. 図を使う場合は、1スライドで文字を読める大きさになるか確認する。
3. 大型の構成図は無理に縮小せず、ファイル名を示して別資料を参照させる。
4. Copilot CLIやMCPツールを説明する場合は、機能名だけでなく提供元と用途を書く。

## 使用する仕様とツール

| 対象 | 用途 | 参照先 |
|---|---|---|
| ISO/IEC 29500、ECMA-376 | PPTXのパッケージとPresentationMLの仕様 | [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)、[PresentationML document structure](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document) |
| python-pptx 1.0.2（MIT） | テンプレートを保持したスライド操作 | [公式ドキュメント](https://python-pptx.readthedocs.io/en/latest/)、[GitHub](https://github.com/scanny/python-pptx) |
| LibreOffice（MPL 2.0） | PPTXからPDFへの変換 | [コマンドライン引数](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html) |
| Poppler `pdftoppm`（GPL-2.0-or-later） | PDFの各ページをJPEGへ変換 | [Poppler](https://poppler.freedesktop.org/) |
| MarkItDown（MIT） | スライド内テキストの抽出 | [GitHub](https://github.com/microsoft/markitdown) |

LibreOfficeとPopplerは外部プログラムとして呼び出す。このリポジトリには両者の実行ファイルを含めない。

## 新規プレゼンテーションの作成

`add_slide.py` は `.potx` のPresentation Partを `.pptx` のコンテンツタイプへ変更した一時パッケージを作り、python-pptxでスライドを追加する。入力ファイルは変更しない。

```bash
uv run .github/skills/azure-pptx/scripts/add_slide.py \
  .github/skills/azure-pptx/assets/Microsoft-Azure-PowerPoint-Template-Light.potx \
  output.pptx \
  --layout-index 12 \
  --title "タイトル" \
  --body "1つ目の要点" \
  --body "2つ目の要点"
```

`.potx` を入力した場合は、テンプレートのサンプルスライドとセクション情報を削除してから1枚追加する。`.pptx` を入力した場合は既存スライドを保持する。既定の動作を変更する場合は `--clear-slides` または `--keep-slides` を指定する。

本文プレースホルダーを明示する場合は `--body-placeholder-index` を使う。

```bash
uv run .github/skills/azure-pptx/scripts/add_slide.py \
  input.pptx output.pptx \
  --keep-slides \
  --layout-index 14 \
  --body-placeholder-index 12 \
  --title "比較" \
  --body "左列の要点"
```

`add_slide.py` は空のデッキの初期化と単純なスライド追加に使う。複数列、画像、表を含むプレゼンテーションでは、最初にこのスクリプトでテンプレート由来のPPTXを作り、そのPPTXを開く作業用スクリプトをPEP 723形式で作成する。汎用CLIへすべてのレイアウト操作を詰め込まない。

## 既存プレゼンテーションの編集

編集用スクリプトは作業内容ごとにPEP 723形式で作成し、`uv run` で実行する。次の条件を守る。

1. `Presentation("input.pptx")` で既存ファイルを開き、空の `Presentation()` から作り直さない。
2. テンプレートのプレースホルダーへ内容を設定し、位置、フォント、文字サイズを不要に上書きしない。
3. 箇条書きプレースホルダーへ手動の箇条書き記号を入力しない。
4. 出力先を入力とは別のファイルにし、保存後にpython-pptxで再度開けることを確認する。

python-pptxにはスライドを削除する公開APIがない。`add_slide.py` はテンプレートのサンプルを除去するため、python-pptx 1.0.2に限定して `Slides._sldIdLst` を使用する。python-pptxを更新するときは、テンプレート変換、全スライド削除、追加、保存、再オープンの一連の確認を行う。

## 内容と表示の確認

スライド内のテキストを抽出する。

```bash
uv run --with "markitdown[pptx]" python -m markitdown output.pptx
```

各スライドをJPEGへ変換する。

```bash
uv run .github/skills/azure-pptx/scripts/thumbnail.py output.pptx
```

既定の出力先は `output-slides/`、解像度は150 DPIである。特定範囲だけを出力する場合は `--first` と `--last` を指定する。

```bash
uv run .github/skills/azure-pptx/scripts/thumbnail.py \
  output.pptx \
  --outdir rendered \
  --first 2 \
  --last 4
```

抽出したテキストでは、欠落、誤字、スライド順を確認する。JPEGでは、文字切れ、要素の重なり、余白、コントラスト、プレースホルダーの残存を確認する。修正後は `--first`、`--last`、`--force` を指定し、影響したスライドだけを再度変換する。

## テンプレートを使わない例外

ユーザーがテンプレート不要と明示した場合は、[PptxGenJS公式ドキュメント](https://gitbrent.github.io/PptxGenJS/)など、採用するライブラリの公式資料を直接参照する。テンプレートを使わない成果物を通常の選択肢にはしない。
