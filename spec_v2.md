# 要件定義書: pst-to-mbox-converter

## 1. 概要

Microsoft OutlookのPST (Personal Storage Table) ファイルを、多くのメールクライアントで利用可能な標準的なMBOX形式に変換する、Python製のコマンドラインツールを作成する。

## 2. 主な機能

-   入力PSTファイルと出力MBOXファイルのパスをコマンドライン引数で指定できる。
-   PSTファイル内のすべてのメールフォルダを再帰的に探索し、メールを抽出する。
-   抽出したメールをMBOX形式に変換し、指定されたファイルに出力する。
-   巨大なPSTファイルを効率的に処理するため、メモリ使用量を抑えた設計とする。
-   変換処理中の進捗状況がわかるように、コンソールにログを出力する。

## 3. 技術仕様

-   **開発言語**: Python
-   **パッケージ管理**: `pyproject.toml` を使用する
-   **主要ライブラリ**:
    -   `libpff` (Cライブラリ) を `ctypes` で直接バインディングして利用する
    -   `argparse`: コマンドライン引数の処理 (Python標準ライブラリ)
    -   `mailbox`: MBOXファイルの生成と書き込み (Python標準ライブラリ)

## 4. プロジェクト構成

パッケージとしてインストール可能で、メンテナンス性の高い構成とする。

```plaintext
pst-to-mbox-converter/
│
├── src/
│   └── pst_to_mbox_converter/
│       ├── __init__.py         # パッケージとして認識させるためのファイル
│       ├── main.py             # エントリーポイント、引数解析、処理フロー制御
│       ├── pst_reader.py       # ctypesを用いてlibpffを呼び出し、PSTファイルを読み込むモジュール
│       └── mbox_writer.py      # MBOXファイルへの書き込み責務を持つモジュール
│
├── pyproject.toml              # プロジェクト定義、依存関係、エントリーポイント設定
└── README.md                   # ツールの使い方や仕様を記述するドキュメント
```

## 5. エントリーポイントと実行方法

### 5.1. `pyproject.toml` の設定

プロジェクトをインストールした際に、カスタムコマンドでツールを実行できるようエントリーポイントを定義する。

```toml
[project.scripts]
pst-converter = "pst_to_mbox_converter.main:main"
```

### 5.2. 実行コマンド

プロジェクトルートで `pip install .` を実行してインストールした後、以下のコマンドでツールを実行できる。

```bash
pst-converter --input /path/to/your/file.pst --output /path/to/your/archive.mbox
```

## 6. モジュール設計

### 6.1. `pst_reader.py`

-   `PSTReader` クラスを定義する。
-   `ctypes` を使って `libpff` の関数を呼び出し、PSTファイルをオープンする。
-   フォルダとメッセージを抽出する責務を持つ。
-   `get_messages()` メソッドはジェネレータとして実装し、メモリ効率を向上させる。

### 6.2. `mbox_writer.py`

-   `MboxWriter` クラスを定義する。
-   指定されたパスにMBOXファイルを生成・オープンする責務を持つ。
-   `add_message()` メソッドを提供し、メールデータをMBOX形式で追記する。

### 6.3. `main.py`

-   `main()` 関数をエントリーポイントとして定義する。
-   コマンドライン引数を解析し、入力パスと出力パスを取得する。
-   `PSTReader` と `MboxWriter` のインスタンスを生成する。
-   `PSTReader` から受け取ったメッセージを `MboxWriter` に渡し、全体の変換処理を制御する。

