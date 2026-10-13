"""
Tkinter-based Graphical User Interface (GUI) for Excel to PDF converter.
Features responsive background threading, real-time progress, profile CRUD, and history CRUD.
"""

from __future__ import annotations

import os
import queue
import sys
import threading
import time
from pathlib import Path
from typing import Any
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from excel_to_pdf.config.history_manager import HistoryManager, HistoryRecord
from excel_to_pdf.config.settings_manager import ConversionProfile, SettingsManager
from excel_to_pdf.core.batch_processor import BatchProcessor, BatchResult, ConversionItem
from excel_to_pdf.core.engine_detector import EngineDetector
from excel_to_pdf.core.system_info import SystemSpec, SystemSpecDetector


class ExcelToPdfApp:
    """Main desktop application window."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel to PDF 高速並列変換ツール")
        self.root.geometry("840x700")
        self.root.minsize(720, 600)

        self.settings_mgr = SettingsManager()
        self.history_mgr = HistoryManager()

        self.msg_queue: queue.Queue = queue.Queue()
        self.current_processor: BatchProcessor | None = None
        self.worker_thread: threading.Thread | None = None
        self.is_converting = False

        self._init_styles()
        self._build_ui()
        self._check_msg_queue()

    def _init_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # Configure custom modern colors and fonts
        style.configure(".", font=("Helvetica", 10))
        style.configure("TNotebook.Tab", font=("Helvetica", 10, "bold"), padding=[12, 6])
        style.configure("Accent.TButton", font=("Helvetica", 11, "bold"), foreground="#FFFFFF", background="#0066CC")
        style.map("Accent.TButton", background=[("active", "#0052A3")])
        style.configure("Danger.TButton", font=("Helvetica", 10), foreground="#CC0000")

    def _build_ui(self) -> None:
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab 1: Conversion execution
        self.tab_convert = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.tab_convert, text=" ⚡ 一括変換 ")
        self._build_convert_tab()

        # Tab 2: Profile Management (CRUD)
        self.tab_profiles = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.tab_profiles, text=" ⚙️ プロファイル設定 (CRUD) ")
        self._build_profiles_tab()

        # Tab 3: History Management (CRUD)
        self.tab_history = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.tab_history, text=" 📜 実行履歴 (CRUD) ")
        self._build_history_tab()

        # Tab 4: System Specification
        self.tab_spec = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.tab_spec, text=" 💻 PCスペック ")
        self._build_system_spec_tab()

    # ==========================================
    # TAB 1: CONVERSION
    # ==========================================
    def _build_convert_tab(self) -> None:
        # 1. Folder Selection Frame
        frame_dirs = ttk.LabelFrame(self.tab_convert, text="フォルダ選択", padding=10)
        frame_dirs.pack(fill="x", pady=(0, 10))

        # Source Dir
        lbl_src = ttk.Label(frame_dirs, text="変換元Excelフォルダ:")
        lbl_src.grid(row=0, column=0, sticky="w", pady=4)
        self.var_src_dir = tk.StringVar()
        self.var_src_dir.trace_add("write", lambda *args: self._update_output_preview())
        entry_src = ttk.Entry(frame_dirs, textvariable=self.var_src_dir, width=54)
        entry_src.grid(row=0, column=1, padx=6, pady=4, sticky="ew")
        btn_src = ttk.Button(frame_dirs, text="参照...", command=self._browse_src_dir)
        btn_src.grid(row=0, column=2, pady=4)

        # Dest Mode
        self.var_out_mode = tk.StringVar(value="pdf_folder")

        rb_pdf_folder = ttk.Radiobutton(
            frame_dirs,
            text="選択フォルダ内に「PDF」フォルダを自動作成して保存（推奨）",
            variable=self.var_out_mode,
            value="pdf_folder",
            command=self._on_out_mode_changed,
        )
        rb_pdf_folder.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 2))

        # Dynamic preview of target PDF directory
        self.lbl_pdf_target = ttk.Label(
            frame_dirs,
            text="📂 出力先: 選択フォルダ直下の「PDF」フォルダに保存されます",
            font=("Helvetica", 9),
            foreground="#1F497D",
        )
        self.lbl_pdf_target.grid(row=2, column=0, columnspan=3, sticky="w", padx=24, pady=(0, 4))

        rb_same = ttk.Radiobutton(
            frame_dirs,
            text="エクセルファイルと同じフォルダにPDFを出力",
            variable=self.var_out_mode,
            value="same",
            command=self._on_out_mode_changed,
        )
        rb_same.grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 2))

        rb_custom = ttk.Radiobutton(
            frame_dirs,
            text="別の出力先フォルダを指定する:",
            variable=self.var_out_mode,
            value="custom",
            command=self._on_out_mode_changed,
        )
        rb_custom.grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 2))

        self.var_out_dir = tk.StringVar()
        self.entry_out = ttk.Entry(frame_dirs, textvariable=self.var_out_dir, width=54, state="disabled")
        self.entry_out.grid(row=5, column=1, padx=6, pady=2, sticky="ew")
        self.btn_out = ttk.Button(frame_dirs, text="参照...", command=self._browse_out_dir, state="disabled")
        self.btn_out.grid(row=5, column=2, pady=2)

        frame_dirs.columnconfigure(1, weight=1)

        # 2. Options & Engine Frame
        frame_opt = ttk.LabelFrame(self.tab_convert, text="実行オプション & エンジン", padding=10)
        frame_opt.pack(fill="x", pady=(0, 10))

        # Profile selection
        lbl_prof = ttk.Label(frame_opt, text="プロファイル適用:")
        lbl_prof.grid(row=0, column=0, sticky="w", pady=4)
        self.combo_profile = ttk.Combobox(frame_opt, state="readonly", width=24)
        self.combo_profile.grid(row=0, column=1, sticky="w", padx=6, pady=4)
        self.combo_profile.bind("<<ComboboxSelected>>", self._on_profile_selected)

        # Workers slider
        lbl_workers = ttk.Label(frame_opt, text="並列プロセス数:")
        lbl_workers.grid(row=0, column=2, sticky="w", padx=(16, 4))
        self.var_workers = tk.IntVar(value=4)
        self.spin_workers = ttk.Spinbox(frame_opt, from_=1, to=16, textvariable=self.var_workers, width=5)
        self.spin_workers.grid(row=0, column=3, sticky="w")

        # Checkboxes
        self.var_recursive = tk.BooleanVar(value=False)
        chk_rec = ttk.Checkbutton(frame_opt, text="サブフォルダも再帰的に検索", variable=self.var_recursive)
        chk_rec.grid(row=1, column=0, columnspan=2, sticky="w", pady=4)

        self.var_overwrite = tk.BooleanVar(value=False)
        chk_over = ttk.Checkbutton(frame_opt, text="既存のPDFを上書き", variable=self.var_overwrite)
        chk_over.grid(row=1, column=2, columnspan=2, sticky="w", pady=4)

        # Detected Engine info
        best_eng = EngineDetector.get_best_engine()
        eng_text = (
            "Windows Excel COM (超高速)"
            if best_eng == "windows_excel"
            else ("LibreOffice ヘッドレス" if best_eng == "libreoffice" else "利用可能なエンジンがありません")
        )
        lbl_eng_info = ttk.Label(frame_opt, text=f"検出エンジン: {eng_text}", foreground="#006600" if best_eng else "#CC0000")
        lbl_eng_info.grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))

        # 3. Action Buttons & Progress
        frame_action = ttk.Frame(self.tab_convert)
        frame_action.pack(fill="x", pady=(0, 6))

        self.btn_start = ttk.Button(frame_action, text="🚀 一括変換を開始する", style="Accent.TButton", command=self._start_conversion)
        self.btn_start.pack(side="left", padx=(0, 10), ipady=4)

        self.btn_cancel = ttk.Button(frame_action, text="⏹️ 中断", command=self._cancel_conversion, state="disabled")
        self.btn_cancel.pack(side="left", ipady=4)

        self.lbl_status = ttk.Label(frame_action, text="待機中", font=("Helvetica", 10, "bold"))
        self.lbl_status.pack(side="right", padx=10)

        # Progress bar
        self.progress_bar = ttk.Progressbar(self.tab_convert, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(0, 6))

        # Log Text Box
        frame_log = ttk.LabelFrame(self.tab_convert, text="変換ログ & 詳細", padding=6)
        frame_log.pack(fill="both", expand=True)

        self.txt_log = tk.Text(frame_log, wrap="word", height=10, font=("Menlo" if sys.platform == "darwin" else "Consolas", 9))
        self.txt_log.pack(side="left", fill="both", expand=True)
        scroll_log = ttk.Scrollbar(frame_log, command=self.txt_log.yview)
        scroll_log.pack(side="right", fill="y")
        self.txt_log.config(yscrollcommand=scroll_log.set)

        self._refresh_profile_combo()

    def _browse_src_dir(self) -> None:
        d = filedialog.askdirectory(title="変換元Excelフォルダを選択")
        if d:
            self.var_src_dir.set(d)
            self._update_output_preview()

    def _browse_out_dir(self) -> None:
        d = filedialog.askdirectory(title="PDF出力先フォルダを選択")
        if d:
            self.var_out_dir.set(d)

    def _update_output_preview(self) -> None:
        mode = self.var_out_mode.get()
        src = self.var_src_dir.get().strip()

        if mode == "pdf_folder":
            if src:
                target = Path(src) / "PDF"
                self.lbl_pdf_target.config(
                    text=f"📂 出力先: {target} (自動作成されます)",
                    foreground="#1F497D",
                )
            else:
                self.lbl_pdf_target.config(
                    text="📂 出力先: 選択フォルダ直下の「PDF」フォルダに自動作成・保存されます",
                    foreground="#555555",
                )
        elif mode == "same":
            self.lbl_pdf_target.config(
                text="📂 出力先: 各Excelファイルと同じフォルダに出力されます",
                foreground="#555555",
            )
        elif mode == "custom":
            cust = self.var_out_dir.get().strip()
            self.lbl_pdf_target.config(
                text=f"📂 出力先: {cust if cust else '指定されたカスタムフォルダ'}",
                foreground="#555555",
            )

    def _on_out_mode_changed(self) -> None:
        if self.var_out_mode.get() == "custom":
            self.entry_out.config(state="normal")
            self.btn_out.config(state="normal")
        else:
            self.entry_out.config(state="disabled")
            self.btn_out.config(state="disabled")
        self._update_output_preview()

    def _refresh_profile_combo(self) -> None:
        profiles = self.settings_mgr.get_all_profiles()
        self.profile_map = {p.name: p for p in profiles}
        names = list(self.profile_map.keys())
        self.combo_profile["values"] = names
        if names:
            self.combo_profile.current(0)
            self._apply_profile(self.profile_map[names[0]])

    def _on_profile_selected(self, event: Any = None) -> None:
        selected_name = self.combo_profile.get()
        if selected_name in self.profile_map:
            self._apply_profile(self.profile_map[selected_name])

    def _apply_profile(self, p: ConversionProfile) -> None:
        self.var_workers.set(p.workers)
        self.var_recursive.set(p.recursive)
        self.var_overwrite.set(p.overwrite)
        if p.output_mode == "custom_directory" and p.custom_output_dir:
            self.var_out_mode.set("custom")
            self.var_out_dir.set(p.custom_output_dir)
        elif p.output_mode == "same_directory":
            self.var_out_mode.set("same")
        else:
            self.var_out_mode.set("pdf_folder")
        self._on_out_mode_changed()

    def _log(self, text: str) -> None:
        self.txt_log.insert("end", text + "\n")
        self.txt_log.see("end")

    def _start_conversion(self) -> None:
        src_path_str = self.var_src_dir.get().strip()
        if not src_path_str or not os.path.isdir(src_path_str):
            messagebox.showwarning("入力エラー", "有効な変換元フォルダを指定してください。")
            return

        src_path = Path(src_path_str).resolve()
        mode = self.var_out_mode.get()

        out_path: Path | None = None
        if mode == "pdf_folder":
            out_path = src_path / "PDF"
            try:
                out_path.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("フォルダ作成エラー", f"PDFフォルダの作成に失敗しました: {e}")
                return
        elif mode == "custom":
            custom_str = self.var_out_dir.get().strip()
            if not custom_str or not os.path.exists(custom_str):
                messagebox.showwarning("出力先エラー", "有効な出力先フォルダを指定してください。")
                return
            out_path = Path(custom_str).resolve()
        else:
            out_path = None

        engine = EngineDetector.get_best_engine()
        if not engine:
            messagebox.showerror(
                "エンジン未検出",
                "利用可能な変換エンジン（Microsoft Excel または LibreOffice）が見つかりません。",
            )
            return

        self.is_converting = True
        self.btn_start.config(state="disabled")
        self.btn_cancel.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.progress_bar["value"] = 0
        self.lbl_status.config(text="変換中...")

        self.current_processor = BatchProcessor(
            workers=self.var_workers.get(),
            engine_type=engine,
            overwrite=self.var_overwrite.get(),
        )

        # Run conversion in a separate thread so GUI remains responsive
        self.worker_thread = threading.Thread(
            target=self._run_conversion_worker,
            args=(Path(src_path), Path(out_path) if out_path else None, self.var_recursive.get()),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_conversion_worker(self, src_path: Path, out_path: Path | None, recursive: bool) -> None:
        processor = self.current_processor
        if processor is None:
            return

        if recursive:
            self.msg_queue.put(("log", f"🔍 ファイルを検索中（サブフォルダ含む）: {src_path}"))
        else:
            self.msg_queue.put(("log", f"🔍 指定フォルダ直下のファイルを検索中: {src_path}"))

        files = processor.scan_files(src_path, recursive=recursive)
        if not files:
            self.msg_queue.put(("log", "⚠️ 対象となるExcelファイルが見つかりませんでした。"))
            self.msg_queue.put(("done", None))
            return

        scope_desc = "サブフォルダ含む" if recursive else "指定フォルダ直下のみ"
        self.msg_queue.put(("log", f"📄 対象ファイル数: {len(files)} 件 ({scope_desc})"))
        self.msg_queue.put(("init_progress", len(files)))

        items = processor.prepare_items(files, src_path, out_path)

        def cb(completed: int, total: int, current_name: str, status_str: str) -> None:
            self.msg_queue.put(("progress", (completed, total, current_name, status_str)))

        result = processor.process_batch(items, progress_cb=cb)

        # Record History (Create)
        hist = HistoryRecord(
            source_folder=str(src_path),
            output_folder=str(out_path) if out_path else "same_directory",
            total_files=result.total_files,
            success_count=result.success_count,
            failed_count=result.failed_count,
            elapsed_seconds=result.elapsed_seconds,
            workers_used=processor.workers,
            engine_used=processor.engine_type,
            memo="GUI一括変換",
            failed_files=[it.source_path.name for it in result.items if it.status == "failed"],
        )
        self.history_mgr.add_record(hist)

        self.msg_queue.put(("done", result))

    def _cancel_conversion(self) -> None:
        if self.current_processor:
            self.current_processor.cancel()
            self._log("⚠️ ユーザーにより中断が要求されました...")
            self.lbl_status.config(text="中断処理中...")

    def _check_msg_queue(self) -> None:
        try:
            # Process up to 5 events per tick to yield control back to UI event loop quickly
            for _ in range(5):
                msg_type, payload = self.msg_queue.get_nowait()
                if msg_type == "log":
                    self._log(payload)
                    self.root.update_idletasks()
                elif msg_type == "init_progress":
                    self.progress_bar["maximum"] = payload
                    self.progress_bar["value"] = 0
                    self.root.update_idletasks()
                elif msg_type == "progress":
                    completed, total, cur_file, status_str = payload
                    self.progress_bar["value"] = completed
                    pct = int(completed / total * 100) if total > 0 else 0
                    self.lbl_status.config(text=f"{completed}/{total} 件完了 ({pct}%)")
                    self._log(f"[{completed}/{total}] {cur_file} -> {status_str}")
                    self.root.update_idletasks()
                elif msg_type == "done":
                    self._on_conversion_completed(payload)
                    break
        except queue.Empty:
            pass
        finally:
            self.root.after(25, self._check_msg_queue)

    def _on_conversion_completed(self, result: BatchResult | None) -> None:
        self.is_converting = False
        self.btn_start.config(state="normal")
        self.btn_cancel.config(state="disabled")

        if result is None:
            self.lbl_status.config(text="完了 (対象なし)")
            return

        self.lbl_status.config(text=f"完了 (成功: {result.success_count}, 失敗: {result.failed_count})")
        self._log("\n" + "=" * 45)
        self._log(f"🎉 処理完了! 総所要時間: {result.elapsed_seconds}秒")
        if result.total_files > 0:
            avg = result.elapsed_seconds / result.total_files
            self._log(f"⚡ 1ファイルあたり平均: {avg:.2f}秒")
        self._log(f"成功: {result.success_count} / スキップ: {result.skipped_count} / 失敗: {result.failed_count}")
        self._log("=" * 45)

        self._refresh_history_table()
        self.root.update_idletasks()

        # Slight delay so user can see 100% progress and final log before dialog appears
        def show_dialog() -> None:
            messagebox.showinfo(
                "変換完了",
                f"変換が完了しました！\n所要時間: {result.elapsed_seconds}秒\n成功: {result.success_count}件 / 失敗: {result.failed_count}件",
            )

        self.root.after(250, show_dialog)

    # ==========================================
    # TAB 2: PROFILES CRUD
    # ==========================================
    def _build_profiles_tab(self) -> None:
        frame_top = ttk.Frame(self.tab_profiles)
        frame_top.pack(fill="both", expand=True)

        cols = ("id", "name", "workers", "recursive", "overwrite", "desc")
        self.tree_profiles = ttk.Treeview(frame_top, columns=cols, show="headings", height=12)
        self.tree_profiles.heading("id", text="ID")
        self.tree_profiles.heading("name", text="プロファイル名")
        self.tree_profiles.heading("workers", text="並列数")
        self.tree_profiles.heading("recursive", text="再帰")
        self.tree_profiles.heading("overwrite", text="上書き")
        self.tree_profiles.heading("desc", text="説明")

        self.tree_profiles.column("id", width=80)
        self.tree_profiles.column("name", width=140)
        self.tree_profiles.column("workers", width=60, anchor="center")
        self.tree_profiles.column("recursive", width=60, anchor="center")
        self.tree_profiles.column("overwrite", width=60, anchor="center")
        self.tree_profiles.column("desc", width=260)

        self.tree_profiles.pack(side="left", fill="both", expand=True)
        scroll_prof = ttk.Scrollbar(frame_top, command=self.tree_profiles.yview)
        scroll_prof.pack(side="right", fill="y")
        self.tree_profiles.config(yscrollcommand=scroll_prof.set)

        # CRUD Action buttons
        frame_btn = ttk.Frame(self.tab_profiles, padding=(0, 10, 0, 0))
        frame_btn.pack(fill="x")

        btn_add = ttk.Button(frame_btn, text="➕ 新規作成 (Create)", command=self._create_profile_dialog)
        btn_add.pack(side="left", padx=4)

        btn_edit = ttk.Button(frame_btn, text="✏️ 編集 (Update)", command=self._edit_profile_dialog)
        btn_edit.pack(side="left", padx=4)

        btn_del = ttk.Button(frame_btn, text="🗑️ 削除 (Delete)", command=self._delete_profile_action)
        btn_del.pack(side="left", padx=4)

        btn_reload = ttk.Button(frame_btn, text="🔄 再読込 (Read)", command=self._refresh_profiles_table)
        btn_reload.pack(side="right", padx=4)

        self._refresh_profiles_table()

    def _refresh_profiles_table(self) -> None:
        for it in self.tree_profiles.get_children():
            self.tree_profiles.delete(it)

        profiles = self.settings_mgr.get_all_profiles()
        for p in profiles:
            self.tree_profiles.insert(
                "",
                "end",
                iid=p.id,
                values=(p.id, p.name, p.workers, "○" if p.recursive else "×", "○" if p.overwrite else "×", p.description),
            )

    def _create_profile_dialog(self) -> None:
        self._open_profile_editor(None)

    def _edit_profile_dialog(self) -> None:
        sel = self.tree_profiles.selection()
        if not sel:
            messagebox.showinfo("選択", "編集するプロファイルを一覧から選択してください。")
            return
        p_id = sel[0]
        profile = self.settings_mgr.get_profile_by_id(p_id)
        if profile:
            self._open_profile_editor(profile)

    def _delete_profile_action(self) -> None:
        sel = self.tree_profiles.selection()
        if not sel:
            messagebox.showinfo("選択", "削除するプロファイルを一覧から選択してください。")
            return
        p_id = sel[0]
        if messagebox.askyesno("削除確認", f"プロファイル (ID: {p_id}) を削除しますか？"):
            self.settings_mgr.delete_profile(p_id)
            self._refresh_profiles_table()
            self._refresh_profile_combo()

    def _open_profile_editor(self, profile: ConversionProfile | None) -> None:
        dlg = tk.Toplevel(self.root)
        dlg.title("プロファイル編集" if profile else "プロファイル新規作成")
        dlg.geometry("450x360")
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="プロファイル名:").grid(row=0, column=0, sticky="w", pady=6)
        var_name = tk.StringVar(value=profile.name if profile else "カスタム設定")
        ttk.Entry(frame, textvariable=var_name, width=32).grid(row=0, column=1, sticky="w", pady=6)

        ttk.Label(frame, text="並列プロセス数:").grid(row=1, column=0, sticky="w", pady=6)
        var_w = tk.IntVar(value=profile.workers if profile else 4)
        ttk.Spinbox(frame, from_=1, to=16, textvariable=var_w, width=6).grid(row=1, column=1, sticky="w", pady=6)

        var_rec = tk.BooleanVar(value=profile.recursive if profile else False)
        ttk.Checkbutton(frame, text="サブフォルダ再帰探索", variable=var_rec).grid(row=2, column=1, sticky="w", pady=6)

        var_ov = tk.BooleanVar(value=profile.overwrite if profile else False)
        ttk.Checkbutton(frame, text="既存PDFの上書き許可", variable=var_ov).grid(row=3, column=1, sticky="w", pady=6)

        ttk.Label(frame, text="出力先モード:").grid(row=4, column=0, sticky="w", pady=6)
        mode_options = [("選択フォルダ内のPDFフォルダ", "pdf_subfolder"), ("同階層", "same_directory"), ("カスタム指定", "custom_directory")]
        var_mode = tk.StringVar(value=profile.output_mode if profile else "pdf_subfolder")
        combo_mode = ttk.Combobox(
            frame,
            values=["選択フォルダ内のPDFフォルダ", "同階層", "カスタム指定"],
            state="readonly",
            width=26,
        )
        if var_mode.get() == "same_directory":
            combo_mode.current(1)
        elif var_mode.get() == "custom_directory":
            combo_mode.current(2)
        else:
            combo_mode.current(0)
        combo_mode.grid(row=4, column=1, sticky="w", pady=6)

        ttk.Label(frame, text="説明:").grid(row=5, column=0, sticky="nw", pady=6)
        txt_desc = tk.Text(frame, width=30, height=3, font=("Helvetica", 9))
        txt_desc.grid(row=5, column=1, sticky="w", pady=6)
        if profile:
            txt_desc.insert("1.0", profile.description)

        def save() -> None:
            name = var_name.get().strip()
            if not name:
                messagebox.showwarning("入力エラー", "プロファイル名を入力してください。", parent=dlg)
                return

            selected_idx = combo_mode.current()
            mode_val = "pdf_subfolder" if selected_idx == 0 else ("same_directory" if selected_idx == 1 else "custom_directory")

            if profile is None:
                # Create
                new_p = ConversionProfile(
                    name=name,
                    workers=var_w.get(),
                    recursive=var_rec.get(),
                    overwrite=var_ov.get(),
                    output_mode=mode_val,
                    description=txt_desc.get("1.0", "end-1c").strip(),
                )
                self.settings_mgr.create_profile(new_p)
            else:
                # Update
                profile.name = name
                profile.workers = var_w.get()
                profile.recursive = var_rec.get()
                profile.overwrite = var_ov.get()
                profile.output_mode = mode_val
                profile.description = txt_desc.get("1.0", "end-1c").strip()
                self.settings_mgr.update_profile(profile)

            self._refresh_profiles_table()
            self._refresh_profile_combo()
            dlg.destroy()

        btn_save = ttk.Button(frame, text="💾 保存", style="Accent.TButton", command=save)
        btn_save.grid(row=6, column=1, sticky="e", pady=16)

    # ==========================================
    # TAB 3: HISTORY CRUD
    # ==========================================
    def _build_history_tab(self) -> None:
        frame_top = ttk.Frame(self.tab_history)
        frame_top.pack(fill="both", expand=True)

        cols = ("id", "time", "success", "failed", "elapsed", "memo", "src")
        self.tree_history = ttk.Treeview(frame_top, columns=cols, show="headings", height=12)
        self.tree_history.heading("id", text="ID")
        self.tree_history.heading("time", text="日時")
        self.tree_history.heading("success", text="成功")
        self.tree_history.heading("failed", text="失敗")
        self.tree_history.heading("elapsed", text="時間(秒)")
        self.tree_history.heading("memo", text="メモ (編集可)")
        self.tree_history.heading("src", text="対象フォルダ")

        self.tree_history.column("id", width=80)
        self.tree_history.column("time", width=130)
        self.tree_history.column("success", width=50, anchor="center")
        self.tree_history.column("failed", width=50, anchor="center")
        self.tree_history.column("elapsed", width=60, anchor="center")
        self.tree_history.column("memo", width=140)
        self.tree_history.column("src", width=220)

        self.tree_history.pack(side="left", fill="both", expand=True)
        scroll_hist = ttk.Scrollbar(frame_top, command=self.tree_history.yview)
        scroll_hist.pack(side="right", fill="y")
        self.tree_history.config(yscrollcommand=scroll_hist.set)

        # History CRUD buttons
        frame_btn = ttk.Frame(self.tab_history, padding=(0, 10, 0, 0))
        frame_btn.pack(fill="x")

        btn_edit_memo = ttk.Button(frame_btn, text="✏️ メモ編集 (Update)", command=self._edit_history_memo)
        btn_edit_memo.pack(side="left", padx=4)

        btn_del_rec = ttk.Button(frame_btn, text="🗑️ 選択削除 (Delete)", command=self._delete_history_record)
        btn_del_rec.pack(side="left", padx=4)

        btn_clear_all = ttk.Button(frame_btn, text="⚠️ 全履歴クリア (Delete All)", command=self._clear_all_history)
        btn_clear_all.pack(side="left", padx=4)

        btn_refresh = ttk.Button(frame_btn, text="🔄 再読込 (Read)", command=self._refresh_history_table)
        btn_refresh.pack(side="right", padx=4)

        self._refresh_history_table()

    def _refresh_history_table(self) -> None:
        for it in self.tree_history.get_children():
            self.tree_history.delete(it)

        records = self.history_mgr.get_all_records()
        for r in records:
            self.tree_history.insert(
                "",
                "end",
                iid=r.id,
                values=(
                    r.id,
                    r.timestamp[:19].replace("T", " "),
                    r.success_count,
                    r.failed_count,
                    f"{r.elapsed_seconds:.1f}",
                    r.memo,
                    r.source_folder,
                ),
            )

    def _edit_history_memo(self) -> None:
        sel = self.tree_history.selection()
        if not sel:
            messagebox.showinfo("選択", "メモを編集する履歴行を選択してください。")
            return
        r_id = sel[0]
        rec = self.history_mgr.get_record_by_id(r_id)
        if not rec:
            return
        new_memo = simpledialog.askstring("メモ編集", "この変換履歴のメモを入力してください:", initialvalue=rec.memo)
        if new_memo is not None:
            self.history_mgr.update_record_memo(r_id, new_memo.strip())
            self._refresh_history_table()

    def _delete_history_record(self) -> None:
        sel = self.tree_history.selection()
        if not sel:
            messagebox.showinfo("選択", "削除する履歴行を選択してください。")
            return
        r_id = sel[0]
        if messagebox.askyesno("削除確認", "選択した履歴を削除しますか？"):
            self.history_mgr.delete_record(r_id)
            self._refresh_history_table()

    def _clear_all_history(self) -> None:
        if messagebox.askyesno("全削除確認", "本当にすべての変換履歴をクリアしますか？"):
            self.history_mgr.clear_all_records()
            self._refresh_history_table()

    # ==========================================
    # TAB 4: SYSTEM SPECIFICATIONS
    # ==========================================
    def _build_system_spec_tab(self) -> None:
        frame_toolbar = ttk.Frame(self.tab_spec)
        frame_toolbar.pack(fill="x", pady=(0, 10))

        lbl_desc = ttk.Label(
            frame_toolbar,
            text="お使いのPCハードウェア情報および変換エンジンの検出スペックです。",
            font=("Helvetica", 10),
        )
        lbl_desc.pack(side="left", pady=4)

        btn_reload = ttk.Button(
            frame_toolbar,
            text="🔄 スペック情報を再取得",
            command=self._load_and_display_specs,
        )
        btn_reload.pack(side="right")

        self.frame_spec_content = ttk.Frame(self.tab_spec)
        self.frame_spec_content.pack(fill="both", expand=True)

        self._load_and_display_specs()

    def _load_and_display_specs(self) -> None:
        for widget in self.frame_spec_content.winfo_children():
            widget.destroy()

        specs = SystemSpecDetector.get_specs()

        self.frame_spec_content.columnconfigure(0, weight=1)
        self.frame_spec_content.columnconfigure(1, weight=1)

        # 1. OS & Platform Card
        lf_os = ttk.LabelFrame(self.frame_spec_content, text="🖥️ OS & 実行環境", padding=12)
        lf_os.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        self._add_spec_row(lf_os, 0, "OS 種別:", f"{specs.os_name} (バージョン: {specs.os_version})")
        self._add_spec_row(lf_os, 1, "アーキテクチャ:", specs.architecture)
        self._add_spec_row(lf_os, 2, "Python バージョン:", f"Python {specs.python_version}")
        is_venv = hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix
        self._add_spec_row(lf_os, 3, "実行環境:", "仮想環境 (.venv)" if is_venv else "システム標準Python")

        # 2. CPU Card
        lf_cpu = ttk.LabelFrame(self.frame_spec_content, text="⚡ CPU (プロセッサ性能)", padding=12)
        lf_cpu.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)

        self._add_spec_row(lf_cpu, 0, "CPU 型番:", specs.cpu_model)
        self._add_spec_row(lf_cpu, 1, "論理コア数 (スレッド):", f"{specs.cpu_cores_logical} コア")
        self._add_spec_row(lf_cpu, 2, "物理コア数 (実コア):", f"{specs.cpu_cores_physical} コア")
        suitability = "極めて高い (高速マルチプロセス向き)" if specs.cpu_cores_logical >= 8 else "標準 (並列処理可能)"
        self._add_spec_row(lf_cpu, 3, "並列処理適性:", suitability)

        # 3. RAM & Storage Card
        lf_mem = ttk.LabelFrame(self.frame_spec_content, text="💾 メモリ & ストレージ", padding=12)
        lf_mem.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)

        ram_text = f"{specs.ram_total_gb} GB"
        if specs.ram_available_gb is not None:
            ram_text += f" (空き: 約 {specs.ram_available_gb} GB)"
        self._add_spec_row(lf_mem, 0, "実装メモリ (RAM):", ram_text)
        self._add_spec_row(lf_mem, 1, "作業ドライブ空き:", f"{specs.disk_free_gb} GB (総容量: {specs.disk_total_gb} GB)")
        self._add_spec_row(lf_mem, 2, "変換用一時領域:", "十分な空き容量があります" if specs.disk_free_gb > 5.0 else "空き容量にご注意ください")

        # 4. Engine & Recommendation Card
        lf_eng = ttk.LabelFrame(self.frame_spec_content, text="⚙️ 変換エンジン & パフォーマンス診断", padding=12)
        lf_eng.grid(row=1, column=1, sticky="nsew", padx=6, pady=6)

        has_engine = "未検出" not in specs.engine_name
        self._add_spec_row(lf_eng, 0, "検出エンジン:", specs.engine_name, highlight=has_engine)
        if specs.engine_path:
            self._add_spec_row(lf_eng, 1, "エンジン実行パス:", specs.engine_path)
        self._add_spec_row(lf_eng, 2, "推奨並列ワーカー数:", f"{specs.recommended_workers} プロセス (最適値)")

        # Recommendation action box
        frame_rec = ttk.LabelFrame(self.frame_spec_content, text="💡 パフォーマンス診断・推奨アドバイス", padding=12)
        frame_rec.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(10, 0))

        lbl_rec_comment = ttk.Label(
            frame_rec,
            text=f"判定: {specs.recommendation_comment}",
            font=("Helvetica", 9, "bold"),
            foreground="#1F497D",
        )
        lbl_rec_comment.pack(side="left", pady=4)

        btn_apply_rec = ttk.Button(
            frame_rec,
            text=f"⚡ 推奨並列数 ({specs.recommended_workers}) を一括変換に適用",
            style="Accent.TButton",
            command=lambda w=specs.recommended_workers: self._apply_recommended_workers(w),
        )
        btn_apply_rec.pack(side="right")

    def _add_spec_row(
        self, parent: ttk.LabelFrame, row: int, label: str, value: str, highlight: bool = False
    ) -> None:
        lbl_k = ttk.Label(parent, text=label, font=("Helvetica", 9, "bold"))
        lbl_k.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)

        fg = "#006600" if highlight else "#333333"
        lbl_v = ttk.Label(parent, text=value, font=("Helvetica", 9), foreground=fg)
        lbl_v.grid(row=row, column=1, sticky="w", pady=4)

    def _apply_recommended_workers(self, workers: int) -> None:
        self.var_workers.set(workers)
        messagebox.showinfo(
            "設定反映",
            f"推奨並列プロセス数 ({workers}) を一括変換タブの並列数に反映しました！",
        )


def run_gui() -> None:
    root = tk.Tk()
    app = ExcelToPdfApp(root)
    root.mainloop()
