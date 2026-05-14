# 売上データ自動集計・月末売上管理ツール

CSV / Excel の売上データを読み込み、月末確認に必要な集計レポートを自動作成する Python 業務改善ツールです。  
Windows 環境での利用を想定し、CLI と Tkinter GUI の両方から実行できます。

## 概要

売上データを手作業でExcel集計している現場向けに、商品別・カテゴリ別・日別・月別の集計、月末サマリー、前月比較、未分類データ確認、検証エラー確認をまとめて出力します。  
出力はグラフ付きExcelレポートのため、そのまま月末確認や社内報告の下書きとして利用できます。

## 解決できる課題

- 毎月の売上集計に時間がかかる
- Excelの手作業集計でミスが起きやすい
- 商品別・カテゴリ別の集計を毎回作るのが大変
- 月末確認時にエラー行や未分類データを見落としやすい
- 前月比較を手作業で作っている
- CSVの文字化けや列名違いで集計作業が止まりやすい

## 主な機能

- CSV / Excel 入力対応: `.csv`, `.xlsx`, `.xls`
- CSV文字コード対応: UTF-8, UTF-8 BOM, CP932
- 列名マッピング
- 対象月、開始日、終了日指定
- 商品・カテゴリ絞り込み
- データプレビュー、集計プレビュー
- 商品別集計、カテゴリ別集計
- 商品別ランキング、カテゴリ別ランキング
- 日別売上推移、月別売上推移
- 月末サマリー
- 前月比較
- 件数、数量合計、平均単価
- 未分類データ一覧
- エラー行一覧
- Excelレポート出力、グラフ出力
- 集計CSV出力、検証エラーCSV出力
- Tkinter GUI
- CLI実行
- pytestによるテスト
- Windows向け実行スクリプト
- exe化・配布パッケージ作成用スクリプト

## 使用技術

- Python
- pandas
- openpyxl
- Tkinter
- pytest
- PowerShell

## セットアップ

Windows PowerShell で実行します。

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

GUIのドラッグ＆ドロップ機能を使う場合は、任意依存もインストールします。

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1 -InstallDragDrop
```

`.xls` を読み込む場合は `xlrd` が必要になることがあります。

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-optional.txt
```

## 使い方

### GUIで実行

```powershell
.\.venv\Scripts\python.exe gui.py
```

GUIでは、入力フォルダ、出力フォルダ、対象月、期間、商品、カテゴリ、ファイルパターンなどを指定して実行できます。  
CSV / Excelプレビュー、集計プレビュー、設定保存、履歴確認もGUIから操作できます。

### CLIで実行

環境チェック:

```powershell
.\.venv\Scripts\python.exe main.py --check-setup
```

プレビュー:

```powershell
.\.venv\Scripts\python.exe main.py --month 2026-04 --preview --preview-limit 20
```

検証のみ:

```powershell
.\.venv\Scripts\python.exe main.py --month 2026-04 --dry-run --all-summaries --monthly-trend --charts
```

Excelレポート出力:

```powershell
.\.venv\Scripts\python.exe main.py --month 2026-04 --all-summaries --monthly-trend --charts
```

設定ファイルとプリセットを使う場合:

```powershell
.\.venv\Scripts\python.exe main.py --config config.example.json --preset april_report
```

## 入力データ形式

`data/input` にCSVまたはExcelファイルを置きます。

対応形式:

- `.csv`
- `.xlsx`
- `.xls`

推奨形式は `.csv` または `.xlsx` です。

標準列:

| 列名 | 必須 | 内容 |
| --- | --- | --- |
| `date` | 必須 | 売上日。例: `2026-04-01` |
| `product` | 必須 | 商品名 |
| `category` | 任意 | カテゴリ |
| `quantity` | 必須 | 数量 |
| `unit_price` | 必須 | 単価 |

日本語列名にも対応しやすいよう、列名マッピングを用意しています。例:

- 日付
- 商品名
- カテゴリ
- 数量
- 単価

金額は `quantity * unit_price` で自動計算します。

## 出力レポート

Excelレポートは `data/output` に出力されます。

シート構成:

1. 月末サマリー
2. 前月比較
3. 詳細データ
4. 商品別集計
5. カテゴリ別集計
6. 日別推移
7. 月別推移
8. 未分類データ
9. エラー行一覧
10. 実行条件

月末サマリーでは、売上合計、明細件数、数量合計、平均単価、対象日数、商品数、カテゴリ数、売上トップ商品、売上トップカテゴリ、未分類データ件数、エラー行件数を確認できます。  
前月比較では、前月データがある場合に売上合計や平均単価などの差分・増減率を確認できます。前月データがない場合もエラーにはならず、比較不可として表示されます。

## テスト

```powershell
.\.venv\Scripts\python.exe -m pytest
```

## 補助スクリプト

月次実行タスク登録:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register_monthly_task.ps1 -DayOfMonth 1 -Time 09:00
```

exeビルド:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```

配布パッケージ作成:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\package_release.ps1
```

## ポートフォリオとしての見どころ

- Excel作業をPythonで自動化する業務改善ツール
- 月末処理に必要な集計、比較、確認用シートを一括生成
- CSV / Excel入力、文字化け対策、列名マッピングに対応
- 未分類データ・エラー行一覧により確認漏れを防止
- GUI / CLI の両方に対応
- pytestで主要処理をテスト
- GitHub公開を意識し、サンプルデータと出力データを分離

## 注意事項

- 実データ、実会社名、個人情報をGitHubに公開しないでください。
- `logs/` と `data/output/` は公開対象外です。
- `.env` やパスワード、APIキーをコミットしないでください。
- サンプルデータは架空データを使用してください。

## 関連ドキュメント

- 詳しい操作方法: [docs/operation_manual.md](docs/operation_manual.md)
- ポートフォリオ説明文: [docs/portfolio_overview.md](docs/portfolio_overview.md)
- 公開前チェックリスト: [docs/release_checklist.md](docs/release_checklist.md)
