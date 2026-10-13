"""
Command Line Interface (CLI) for Excel to PDF converter.
Supports batch conversion, profile CRUD, and history inspection.
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

from tqdm import tqdm

from excel_to_pdf.config.history_manager import HistoryManager, HistoryRecord
from excel_to_pdf.config.settings_manager import ConversionProfile, SettingsManager
from excel_to_pdf.core.batch_processor import BatchProcessor
from excel_to_pdf.core.engine_detector import EngineDetector


def setup_cli_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-to-pdf",
        description="高速並列 Excel → PDF 一括変換ツール (CLI & GUI)",
    )

    # Conversion arguments
    parser.add_argument(
        "-i", "--input",
        type=str,
        help="変換対象のエクセルファイルが格納されているフォルダパス",
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help="PDF出力先フォルダパス (省略時は入力フォルダ直下の『PDF』フォルダ)",
    )
    parser.add_argument(
        "-w", "--workers",
        type=int,
        default=None,
        help="並列ワーカー数 (デフォルト: プロファイル設定または4)",
    )
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        default=None,
        help="サブフォルダ内のエクセルファイルも再帰的に変換対象にする",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="サブフォルダを探索せず、直下のファイルのみ変換する",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        default=None,
        help="同名のPDFが既に存在する場合に上書きする",
    )
    parser.add_argument(
        "--engine",
        choices=["auto", "windows_excel", "libreoffice"],
        default="auto",
        help="使用する変換エンジン (デフォルト: 自動検出)",
    )
    parser.add_argument(
        "-p", "--profile",
        type=str,
        default=None,
        help="適用するプロファイル名またはID",
    )
    parser.add_argument(
        "-m", "--memo",
        type=str,
        default="",
        help="変換履歴に記録するメモ",
    )

    # Profile CRUD flags
    profile_group = parser.add_argument_group("プロファイル管理 (CRUD)")
    profile_group.add_argument(
        "--list-profiles",
        action="store_true",
        help="登録済みプロファイル一覧を表示 (Read)",
    )
    profile_group.add_argument(
        "--create-profile",
        nargs=3,
        metavar=("NAME", "WORKERS", "DESC"),
        help="新しいプロファイルを作成 (Create): 名前 並列数 説明",
    )
    profile_group.add_argument(
        "--delete-profile",
        type=str,
        metavar="PROFILE_ID",
        help="指定したIDのプロファイルを削除 (Delete)",
    )

    # History CRUD flags
    history_group = parser.add_argument_group("変換履歴管理 (CRUD)")
    history_group.add_argument(
        "--list-history",
        action="store_true",
        help="過去の変換実行履歴を表示 (Read)",
    )
    history_group.add_argument(
        "--clear-history",
        action="store_true",
        help="すべての変換履歴を削除 (Delete)",
    )

    return parser


def run_cli(args: argparse.Namespace) -> int:
    settings_mgr = SettingsManager()
    history_mgr = HistoryManager()

    # --- Profile CRUD commands ---
    if args.list_profiles:
        profiles = settings_mgr.get_all_profiles()
        print("\n📋 登録済み変換プロファイル一覧:")
        print("-" * 65)
        for p in profiles:
            print(f"ID: {p.id:<16} | 名前: {p.name:<18} | 並列数: {p.workers} | 再帰: {p.recursive}")
            print(f"   説明: {p.description}")
        print("-" * 65)
        return 0

    if args.create_profile:
        name, workers_str, desc = args.create_profile
        try:
            workers = int(workers_str)
        except ValueError:
            print(f"エラー: 並列数は整数で指定してください: {workers_str}", file=sys.stderr)
            return 1
        new_prof = ConversionProfile(
            name=name,
            workers=workers,
            description=desc,
        )
        created = settings_mgr.create_profile(new_prof)
        print(f"✅ プロファイルを作成しました: ID={created.id}, 名前={created.name}")
        return 0

    if args.delete_profile:
        success = settings_mgr.delete_profile(args.delete_profile)
        if success:
            print(f"✅ プロファイル '{args.delete_profile}' を削除しました。")
            return 0
        else:
            print(f"❌ プロファイル '{args.delete_profile}' が見つかりませんでした。", file=sys.stderr)
            return 1

    # --- History CRUD commands ---
    if args.list_history:
        records = history_mgr.get_all_records()
        print("\n📜 変換実行履歴一覧:")
        print("-" * 75)
        if not records:
            print("履歴はありません。")
        for r in records:
            print(f"ID: {r.id} | 日時: {r.timestamp[:19]} | 成功/総数: {r.success_count}/{r.total_files} | 所要時間: {r.elapsed_seconds}秒")
            print(f"   入力: {r.source_folder}")
            if r.memo:
                print(f"   メモ: {r.memo}")
        print("-" * 75)
        return 0

    if args.clear_history:
        history_mgr.clear_all_records()
        print("✅ すべての変換履歴をクリアしました。")
        return 0

    # --- Batch Conversion Execution ---
    if not args.input:
        print("エラー: 変換対象フォルダ (--input / -i) を指定してください。ヘルプは -h で確認できます。", file=sys.stderr)
        return 1

    input_path = Path(args.input).resolve()
    if not input_path.exists() or not input_path.is_dir():
        print(f"エラー: 指定された入力フォルダが存在しません: {input_path}", file=sys.stderr)
        return 1

    # Apply Profile if specified
    active_profile = settings_mgr.get_all_profiles()[0]
    if args.profile:
        found = None
        for p in settings_mgr.get_all_profiles():
            if p.id == args.profile or p.name == args.profile:
                found = p
                break
        if found:
            active_profile = found
            print(f"ℹ️ プロファイル '{active_profile.name}' を適用します。")
        else:
            print(f"⚠️ 指定プロファイル '{args.profile}' が見つかりません。デフォルトを使用します。")

    # Command line overrides profile defaults
    workers = args.workers if args.workers is not None else active_profile.workers
    recursive = active_profile.recursive
    if args.recursive:
        recursive = True
    elif args.no_recursive:
        recursive = False

    overwrite = args.overwrite if args.overwrite is not None else active_profile.overwrite
    output_dir: Path | None = None
    if args.output is not None:
        output_dir = Path(args.output).resolve()
    elif active_profile.output_mode == "custom_directory" and active_profile.custom_output_dir:
        output_dir = Path(active_profile.custom_output_dir).resolve()
    elif active_profile.output_mode == "same_directory":
        output_dir = None
    else:
        # デフォルト: 選択/入力フォルダ直下の "PDF" フォルダ
        output_dir = input_path / "PDF"

    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)

    # Engine detection
    engine_type = None
    if args.engine != "auto":
        engine_type = args.engine
    else:
        engine_type = EngineDetector.get_best_engine()

    if engine_type is None:
        print("❌ エラー: 利用可能な変換エンジンが見つかりません。", file=sys.stderr)
        print("  - Windows: Microsoft Excel がインストールされているか確認してください。", file=sys.stderr)
        print("  - macOS / Linux: LibreOffice をインストールし PATH に通してください。", file=sys.stderr)
        return 1

    processor = BatchProcessor(
        workers=workers,
        engine_type=engine_type,
        overwrite=overwrite,
        supported_extensions=active_profile.supported_extensions,
    )

    print(f"\n🚀 Excel → PDF 高速並列変換を開始します")
    print(f"  📂 入力フォルダ: {input_path}")
    print(f"  📁 出力先: {output_dir or '入力ファイルと同階層'}")
    print(f"  ⚡ 並列プロセス数: {workers}")
    print(f"  🔍 再帰探索: {'有効 (サブフォルダも探索)' if recursive else '無効 (指定フォルダ直下のみ)'}")
    print(f"  🔄 上書き設定: {'有効' if overwrite else 'スキップ'}")
    print(f"  ⚙️ 変換エンジン: {processor.engine_type}\n")

    files = processor.scan_files(input_path, recursive=recursive)
    if not files:
        print("⚠️ 対象となるExcelファイルが見つかりませんでした。")
        return 0

    scope_str = "サブフォルダ含む" if recursive else "指定フォルダ直下のみ"
    print(f"📄 対象ファイル数: {len(files)} 件 ({scope_str})")
    items = processor.prepare_items(files, input_path, output_dir)

    # Progress bar setup
    pbar = tqdm(total=len(items), unit="file", desc="変換進捗", ncols=80)

    def progress_callback(completed: int, total: int, current_file: str, status_text: str) -> None:
        pbar.set_postfix_str(f"{current_file[:20]} ({status_text})")
        pbar.update(1)

    result = processor.process_batch(items, progress_cb=progress_callback)
    pbar.close()

    # Record history
    history_record = HistoryRecord(
        source_folder=str(input_path),
        output_folder=str(Path(output_dir).resolve()) if output_dir else "same_directory",
        total_files=result.total_files,
        success_count=result.success_count,
        failed_count=result.failed_count,
        elapsed_seconds=result.elapsed_seconds,
        workers_used=workers,
        engine_used=result.engine_used,
        memo=args.memo,
        failed_files=[it.source_path.name for it in result.items if it.status == "failed"],
    )
    history_mgr.add_record(history_record)

    # Output Summary
    print("\n" + "=" * 50)
    print("🎉 変換処理が完了しました！")
    print(f"  ⏱️ 総所要時間: {result.elapsed_seconds:.2f} 秒")
    if result.total_files > 0:
        avg_time = result.elapsed_seconds / result.total_files
        print(f"  ⚡ 1ファイルあたり平均: {avg_time:.2f} 秒/ファイル")
    print(f"  ✅ 成功: {result.success_count} 件")
    print(f"  ⏭️ スキップ: {result.skipped_count} 件")
    print(f"  ❌ 失敗: {result.failed_count} 件")

    if result.failed_count > 0:
        print("\n⚠️ 失敗したファイル一覧:")
        for it in result.items:
            if it.status == "failed":
                print(f"  - {it.source_path.name}: {it.error_message}")

    print("=" * 50 + "\n")
    return 0 if result.failed_count == 0 else 2
