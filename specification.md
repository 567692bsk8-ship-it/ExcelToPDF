# 仕様書: 高速並列 Excel to PDF 一括変換システム (ExcelToPDF)

## 1. システム概要
本ソフトウェアは、任意のフォルダ内にある複数のExcelファイル（`.xlsx`, `.xls`, `.xlsm`, `.xlsb`）を検出し、マルチプロセス並列処理およびOffice最適化フラグを用いて極めて高速にPDFへ一括変換するデスクトップツールです。
CLI（コマンドライン）およびGUI（デスクトップ画面）の双方をサポートし、初心者から自動化バッチ処理まで幅広く対応します。

---

## 2. 背景と解決する課題
- **従来の課題**:
  - PowerShellやVBAの逐次実行では、1ファイルごとにExcelプロセスを起動・終了、シート再計算、外部リンク参照確認、プリンタ問い合わせ、画面描画などが発生し、1ファイルあたり15〜20秒以上要していた。
  - 大量ファイル（数十〜数百ファイル）の変換に数十分〜数時間がかかり、業務の大きなボトルネックになっていた。
- **本ツールの解決策**:
  - **並列ワーカープール**: CPUコア数に応じたマルチプロセスワーカーで並行処理。
  - **Excelプロセス常駐化**: ワーカーごとにExcelインスタンスを1度だけ起動してバッチ処理。
  - **オーバーヘッド抑止**: 画面描画停止、イベント停止、アラート無効化、自動計算停止、外部リンク更新スキップを徹底。
  - **クロスプラットフォーム対応**: WindowsではMicrosoft Excel COM、Mac/LinuxではLibreOfficeマルチインスタンスモードを自動切替。

---

## 3. 主要機能一覧
1. **一括変換エンジン**:
   - 指定フォルダ内のExcelファイルを再帰的/非再帰的に探索。
   - 隠しファイル（`~$...` 等のテンポラリファイル）の自動除外。
   - マルチプロセス（ProcessPoolExecutor）による高速並列処理。
   - 変換進捗率（パーセンテージ、完了数/総数、所要時間、平均処理速度）のリアルタイム表示。
2. **変換プロファイルのCRUD管理**:
   - 頻繁に利用する変換設定（対象拡張子、並列数、出力先フォルダ設定、除外パターン、上書き設定など）を名前付きプロファイルとして保存。
   - プロファイルの新規作成(Create)、一覧・詳細参照(Read)、更新・編集(Update)、削除(Delete)が可能。
3. **変換履歴のCRUD管理**:
   - 実行ごとのサマリー（変換日時、対象ファイル数、成功数、失敗数、所要時間）の自動記録。
   - 履歴の閲覧(Read)、メモ・タグの追加/更新(Update)、履歴の個別/一括削除(Delete)。
4. **二面インターフェース**:
   - **GUI**: 直感的なフォルダ選択ダイアログ、並列スライダー、進捗バー、プロファイル編集、履歴タブ、およびPCハードウェアスペック自動検出・診断タブ。
   - **CLI**: 引数指定（`--input`, `--output`, `--workers`, `--recursive` 等）によるバッチスクリプト組み込み対応。

---

## 4. アーキテクチャ構成
```
ExcelToPDF/
├── .agents/
│   └── AGENTS.md                  # プロジェクトローカルルール
├── excel_to_pdf/
│   ├── __init__.py
│   ├── core/
│   │   ├── __init__.py
│   │   ├── converter_base.py      # 抽象基底クラス (BaseConverter)
│   │   ├── windows_excel.py       # Windows COM最適化並列コンバータ
│   │   ├── libreoffice.py         # LibreOffice並列コンバータ (マルチプロファイル)
│   │   ├── engine_detector.py     # OS/利用可能エンジン判定
│   │   ├── system_info.py         # PCスペック・ハードウェア自動検出
│   │   └── batch_processor.py     # 並列実行制御・進捗コールバック
│   ├── config/
│   │   ├── __init__.py
│   │   ├── settings_manager.py    # 変換プロファイルCRUDマネージャー
│   │   └── history_manager.py     # 変換履歴CRUDマネージャー
│   ├── gui/
│   │   ├── __init__.py
│   │   └── app.py                 # Tkinterベースの洗練されたGUI
│   └── cli/
│       ├── __init__.py
│       └── main.py                # コマンドライン引数処理
├── main.py                        # アプリケーション起動エントリポイント
├── pyproject.toml                 # プロジェクトメタデータ & 依存定義
├── requirements.txt               # pip用依存リスト
├── README.md                      # 4大要素を含む公式ドキュメント
├── specification.md               # 本仕様書
└── tests/
    ├── __init__.py
    ├── test_config.py             # プロファイル & 履歴 CRUDテスト
    └── test_processor.py          # ファイル収集・バッチ制御テスト
```

---

## 5. データモデル (JSONストレージ)

### 5.1 プロファイルデータ (`profiles.json`)
```json
{
  "profiles": [
    {
      "id": "default",
      "name": "標準設定",
      "description": "CPUコア数に合わせた標準的な並列変換",
      "workers": 4,
      "recursive": true,
      "output_mode": "pdf_subfolder",
      "custom_output_dir": "",
      "overwrite": false,
      "supported_extensions": [".xlsx", ".xls", ".xlsm"],
      "created_at": "2026-10-13T18:00:00",
      "updated_at": "2026-10-13T18:00:00"
    }
  ]
}
```

### 5.2 変換履歴データ (`history.json`)
```json
{
  "records": [
    {
      "id": "hist_20261013_180000",
      "timestamp": "2026-10-13T18:00:00",
      "source_folder": "/path/to/excel",
      "output_folder": "/path/to/pdf",
      "total_files": 10,
      "success_count": 10,
      "failed_count": 0,
      "elapsed_seconds": 4.5,
      "workers_used": 4,
      "engine_used": "Windows Excel COM",
      "memo": "月次レポート変換",
      "failed_files": []
    }
  ]
}
```
