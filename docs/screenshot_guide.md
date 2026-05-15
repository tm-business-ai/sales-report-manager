# スクリーンショット撮影ガイド

README用のスクリーンショットを更新するための撮影手順です。  
撮影には `data/input` のサンプルデータのみを使用し、実データや個人情報が写らないようにしてください。

## 1. GUIを起動する

通常版:

```powershell
.\.venv\Scripts\python.exe gui.py
```

EXE版:

```powershell
dist\SalesReportManager.exe
```

## 2. 初期画面を撮影

保存名:

```text
docs/images/gui_main.png
```

撮影内容:

- 実行タブ
- 入力フォルダ
- 出力フォルダ
- 対象月
- データをプレビュー
- Excelレポートを作成
- Excelレポートを開く
- 出力フォルダを開く

## 3. 列名設定画面を撮影

保存名:

```text
docs/images/gui_column_mapping.png
```

撮影内容:

- 列名設定画面
- 標準項目
- 入力ファイルの列名
- 入力ファイルの列名を読み取る
- 列名設定を保存

## 4. エラー確認タブを撮影

保存名:

```text
docs/images/gui_error_review.png
```

撮影内容:

- エラー確認タブ
- 入力データを検証
- エラーCSVを開く
- エラーCSVの保存先を開く
- エラーがない場合の案内、またはサンプルのエラー表示

## 5. Excelレポートを出力する

```powershell
.\.venv\Scripts\python.exe main.py --month 2026-04 --all-summaries --monthly-trend --charts
```

## 6. Excelシートを撮影

保存名:

- `docs/images/excel_month_summary.png`
- `docs/images/excel_previous_month_comparison.png`
- `docs/images/excel_product_summary.png`
- `docs/images/excel_category_summary.png`

撮影内容:

- 月末サマリー
- 前月比較
- 商品別集計
- カテゴリ別集計

必要に応じて、詳細データは `docs/images/excel_detail_data.png` として撮影します。

## 7. 配布パッケージフォルダを撮影

保存名:

```text
docs/images/release_package_folder.png
```

撮影内容:

- `release/SalesReportManager` のフォルダ構成
- `SalesReportManager.exe`
- `README_QUICK_START.txt`
- `config.example.json`
- `data/input`
- `data/output`
- `docs/operation_manual.md`

## 8. READMEに掲載する推奨画像

READMEでは、存在する画像のみリンクします。  
新しく撮影した画像を追加した後に、READMEのスクリーンショット欄へ掲載してください。

推奨ファイル名:

- `gui_main.png`
- `gui_column_mapping.png`
- `gui_error_review.png`
- `excel_month_summary.png`
- `excel_previous_month_comparison.png`
- `excel_product_summary.png`
- `excel_category_summary.png`
- `release_package_folder.png`

## 9. 注意事項

- 実データは使わない
- 個人情報が写らないようにする
- サンプルデータで撮影する
- 画像はPNG形式を推奨
- 画像名はREADMEと一致させる
- `data/output`, `logs`, `dist`, `build`, `release` はGit管理対象にしない
