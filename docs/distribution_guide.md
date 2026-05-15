# Windows向けEXE化・配布ガイド

## 目的

Python環境に詳しくない利用者でも起動しやすいように、GUIアプリをWindows向けEXEとして作成し、サンプルデータや簡易説明書を含む配布フォルダにまとめます。

## PyInstallerの導入

EXE化には任意依存の PyInstaller を使用します。

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-optional.txt
```

PyInstallerだけを入れる場合は次のコマンドでも構いません。

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
```

## EXE作成

次のスクリプトで `gui.py` から `dist/SalesReportManager.exe` を作成します。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```

内部では概ね次の処理を行います。

```powershell
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --noconsole --name SalesReportManager gui.py
```

PyInstallerが未導入の場合は、先に任意依存をインストールしてください。

## 配布パッケージ作成

EXE作成後、次のスクリプトで配布フォルダを作成します。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_release_package.ps1
```

作成される構成は次の通りです。

```text
release/SalesReportManager/
├─ SalesReportManager.exe
├─ README_QUICK_START.txt
├─ config.example.json
├─ data/
│  ├─ input/
│  │  ├─ sales_2026_03.csv
│  │  ├─ sales_2026_04.csv
│  │  ├─ sample_sales_2026_03.xlsx
│  │  └─ sample_sales_2026_04.xlsx
│  └─ output/
└─ docs/
   └─ operation_manual.md
```

## 配布時の注意点

- `release/`, `dist/`, `build/`, `*.spec`, `*.exe` はGitHubに含めません。
- `data/output/` や `logs/` に実行結果やログが残っていないか確認してください。
- 実データ、個人情報、認証情報を配布フォルダに含めないでください。
- 配布するサンプルデータは架空データのみ使用してください。
- Windows Defenderや社内セキュリティソフトで、個人作成EXEとして警告が表示される場合があります。

## README用画像の更新

README用の画像は `docs/images` にPNG形式で保存します。  
GUIの実行タブ、列名設定画面、エラー確認タブ、Excelレポート、配布フォルダ構成を撮影すると、利用イメージが伝わりやすくなります。

撮影手順は [docs/screenshot_guide.md](screenshot_guide.md) を参照してください。

## EXE化できない場合の通常起動

EXE化せずに利用する場合は、通常通りPythonからGUIを起動できます。

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
.\.venv\Scripts\python.exe gui.py
```

コマンドラインで動作確認する場合は次を実行します。

```powershell
.\.venv\Scripts\python.exe main.py --check-setup
.\.venv\Scripts\python.exe main.py --month 2026-04 --preview --preview-limit 5
```
