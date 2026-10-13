"""
Main entry point for Excel to PDF Converter.
Launches GUI by default, or CLI if arguments are supplied.
"""

from __future__ import annotations

import sys

from excel_to_pdf.cli.main import run_cli, setup_cli_parser
from excel_to_pdf.gui.app import run_gui


def main() -> None:
    # If explicit CLI arguments or flags are passed, run CLI
    cli_flags = {
        "-i", "--input", "-o", "--output", "-w", "--workers",
        "-r", "--recursive", "--no-recursive", "--overwrite",
        "--engine", "-p", "--profile", "--list-profiles",
        "--create-profile", "--delete-profile", "--list-history",
        "--clear-history", "-m", "--memo", "-h", "--help",
    }

    has_cli_args = any(arg in cli_flags for arg in sys.argv[1:])

    if "--cli" in sys.argv:
        # Explicit CLI request
        sys.argv.remove("--cli")
        parser = setup_cli_parser()
        args = parser.parse_args()
        sys.exit(run_cli(args))
    elif has_cli_args:
        # CLI flag detected
        parser = setup_cli_parser()
        args = parser.parse_args()
        sys.exit(run_cli(args))
    else:
        # Launch GUI by default
        try:
            run_gui()
        except tk.TclError as e:  # type: ignore
            print(f"GUIの起動に失敗しました（Display環境が見つからない可能性があります）: {e}")
            print("CLIモードをご利用ください: python main.py --help")
            sys.exit(1)


if __name__ == "__main__":
    main()
