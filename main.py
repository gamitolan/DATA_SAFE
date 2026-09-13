"""
main.py
Simplest-possible CLI for the OFS-based M580 tag backup/restore tool.
Backup discovers its tag list directly from OFS - no input file needed.

Usage:
    python main.py backup  --out backup_2026-09-08.xlsx
    python main.py restore --backup backup_2026-09-08.xlsx --report restore_report.xlsx

    python main.py init-config     (writes a starter config.json to edit)
"""

import argparse
import sys

from config import load_config, write_default_config
import backup
import restore


def main():
    parser = argparse.ArgumentParser(
        description="Backup/restore M580 tag values by name via OFS."
    )
    parser.add_argument("--config", default="config.json",
                         help="Path to config.json (default: ./config.json)")
    sub = parser.add_subparsers(dest="command", required=True)

    p_init = sub.add_parser("init-config", help="Write a starter config.json")

    p_backup = sub.add_parser("backup", help="Discover tags from OFS and back up their values")
    p_backup.add_argument("--out", required=True,
                           help="Output backup workbook path")
    p_backup.add_argument("--include-fb-members", action="store_true",
                           help="Include function block instance members "
                                "(e.g. 'InstanceName.Member') - excluded by default")

    p_restore = sub.add_parser("restore", help="Restore tag values to the CPU")
    p_restore.add_argument("--backup", required=True,
                            help="Backup workbook previously produced by 'backup'")
    p_restore.add_argument("--report", default=None,
                            help="Optional path for a restore report workbook")

    args = parser.parse_args()

    if args.command == "init-config":
        write_default_config(args.config)
        print(f"Wrote starter config to {args.config}.")
        return

    config = load_config(args.config)

    if args.command == "backup":
        backup.run_backup(config, args.out, include_fb_members=args.include_fb_members)
    elif args.command == "restore":
        restore.run_restore(config, args.backup, args.report)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001 - top-level CLI safety net
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
