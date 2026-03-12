# Splunk手順書生成用リポジトリ

このリポジトリは、既存の Word テンプレートを流用して
**Splunk Deployment Server whitelist設定手順書** を生成するためのリポジトリです。

## ディレクトリ構成

- `source/template/` : Word テンプレートを配置
- `source/input/` : 手順書作成の元資料（PDF など）を配置
- `output/` : 生成された Word ファイルの出力先
- `scripts/` : 文書生成・検証スクリプトを配置

## 入力ファイル配置場所

- テンプレート配置場所: `source/template/UniversalForwarderインストールマニュアル.docx`
- 入力PDF配置場所: `source/input/whitelist設定手順.pdf`

## 出力先

- `output/SplunkDS_whitelist設定手順書.docx`

## 生成コマンド

```bash
python scripts/generate_whitelist_doc.py
```

## 生成後の検証コマンド

```bash
python scripts/validate_generated_doc.py
```

## 生成方針（レビュー運用）

- `output/` 配下は生成物のため `.gitignore` で除外しています。
- PR ではテンプレート・入力資料・生成スクリプト・README の差分をレビュー対象とし、
  生成物は各環境で再生成して確認する運用です。

## 依存ライブラリ

- Python 3 標準ライブラリのみ（`zipfile`, `xml.etree.ElementTree`, `pathlib`）

## 補足

- 目次は Word のフィールド更新で再生成可能です（Word 上で「目次の更新」を実行）。
- テンプレート本体を直接上書きせず、必ず `output/` に新規出力してください。
