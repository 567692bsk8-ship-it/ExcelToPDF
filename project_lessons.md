# プロジェクト教訓・技術ナレッジ (ExcelToPDF)

本ドキュメントは、ExcelからPDFへの高速並列変換システム開発において得られた知見、ハマりポイント、および高速化技術の要点を記録します。

---

## 1. なぜ従来のPowerShell / VBAは1ファイル20秒もかかるのか？
1. **Excelプロセスの起動・終了コスト**:
   - `New-Object -ComObject Excel.Application` や外部スクリプトでファイルごとにプロセスを立ち上げると、それだけで数秒〜10秒以上の起動オーバーヘッドが発生する。
2. **デフォルトの同期待機と対話ダイアログ**:
   - 外部リンクの更新確認（`AskToUpdateLinks`）、マクロ警告、パスワード確認ダイアログなどで処理が停止する。
3. **全シートの自動再計算（Calculation = xlCalculationAutomatic）**:
   - ブックを開いた際、または印刷出力時にすべての数式が再計算されるため、数式が多いシートで著しい遅延が発生する。
4. **印刷スプーラ・ネットワークプリンタの通信ハング**:
   - ExcelのPDF出力（`ExportAsFixedFormat`）は、内部的にWindowsのデフォルトプリンタドライバと通信して用紙サイズやマージンを計算する。オフラインプリンタやネットワークプリンタがデフォルトの場合、タイムアウトまで10〜20秒フリーズする。
5. **COMのSTA（Single Threaded Apartment）制約**:
   - 単一のExcelプロセス内ではスレッドセーフではなく、直列処理しか行えない。

---

## 2. Pythonによる高速化・並列化の解決策

### 2.1 Windows COM (pywin32) の最適化
- **マルチプロセス化 (`ProcessPoolExecutor`)**:
  - COMはスレッド並列化できないため、CPUコア数に応じた**プロセス単位**で独立したExcelプロセスを起動する。
  - 各ワーカープロセス内で起動時に `pythoncom.CoInitialize()` を実行し、終了時に `pythoncom.CoUninitialize()` を実行する。
- **インスタンスの再利用 (Worker Batching)**:
  - 1ファイルごとにExcelを起動・破棄せず、1つのワーカープロセス内でExcelを常駐させ、複数ファイルを連続して `Workbooks.Open` → `ExportAsFixedFormat` → `Close(SaveChanges=False)` する。
- **徹底的なフラグ無効化**:
  ```python
  app.Visible = False
  app.ScreenUpdating = False
  app.DisplayAlerts = False
  app.EnableEvents = False
  app.AskToUpdateLinks = False
  app.Calculation = -4135  # xlCalculationManual (自動再計算を停止)
  ```
  ブックオープン時：
  ```python
  app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True, Notify=False)
  ```
  これにより、1ファイルあたり**0.5〜1.5秒程度**まで短縮可能。

### 2.2 LibreOffice (soffice) の並列化テクニック
- **同一プロファイル衝突の回避**:
  - 単純に `soffice --headless --convert-to pdf` を並列実行すると、共通のユーザープロファイルロックにより2つ目以降のプロセスが失敗・クラッシュする。
  - **解決策**: 各ワーカープロセスごとに一意の一時ディレクトリを作成し、`-env:UserInstallation=file:///tmp/soffice_profile_xxx` を渡して完全にプロファイルを分離する。
  - これにより、CPUコア数に応じた完全並列変換（マルチコア100%活用）が実現する。

### 2.3 出力先フォルダ（PDF/）の自動作成とサブディレクトリ階層維持
- **元データ混在の防止**:
  - 変換元Excelと同じフォルダに大量のPDFが混在すると視認性や管理性が損なわれるため、選択したフォルダ直下に自動的に `PDF/` フォルダを生成して保存する設計を標準化。
- **サブフォルダ階層の完全再現**:
  - `prepare_items` において `src.relative_to(input_base)` を用いて相対パスを算出し、`output_dir / rel_path` に配置することで、サブフォルダ（例: `営業部/週報.xlsx`）があっても `PDF/営業部/週報.pdf` のように階層を崩さず整然と格納される。
- **UIでの動的プレビュー**:
  - ユーザーがフォルダを選択した瞬間に、生成予定の `.../PDF` パスを視覚的にリアルタイム表示することで、意図しない場所への出力事故を防止。

### 2.4 並列処理におけるリアルタイム進捗ストリーミングとTkinter描画最適化
- **チャンク化と進捗通知のジレンマ**:
  - Excel COMやLibreOfficeの起動コストを削減するためにワーカーごと複数ファイルをバッチ（チャンク）で渡すと、通常はチャンク全体の完了まで `as_completed` が返らないため、GUI上ですべての処理が終わってから一括でログが表示されてしまう。
- **解決策 (Manager.Queueストリーミング)**:
  - 親プロセスで `multiprocessing.Manager().Queue()` を生成し、ワーカー関数へ渡す。
  - 各ワーカー内で1ファイル変換が完了するごとに即座に `progress_queue.put((src_path, success, err))` を送信。
  - 親プロセスは `ProcessPoolExecutor` の実行中、非同期にキューから結果を取り出して即座にGUIコールバックを呼ぶ。
- **Tkinterの即時再描画 (update_idletasks)**:
  - Tkinterのメッセージループポーリングを短縮（25ms）し、ログ追加およびプログレスバー更新時に `root.update_idletasks()` を呼ぶことで、UIスレッドが処理を待たずに即座に画面へ描画を反映。
  - 完了時のモーダルダイアログ（`messagebox.showinfo`）による描画ブロックを防ぐため、100%完了描画後にわずかな余韻（250ms）を設けてからダイアログを表示する。

---

## 3. CRUDアーキテクチャの適用
- 単なる変換スクリプトに留めず、実務で使い続けられるソフトウェアとして以下を実装：
  - **変換プロファイルCRUD**: 頻繁に使うフォルダや並列設定をプリセットとして保存・編集・削除可能。
  - **実行履歴CRUD**: 過去の変換ログ（成否、所要時間、失敗理由）の参照、個別・一括削除、メモ追加。

