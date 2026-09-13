"""
restore.py
Rebinds every tag from a backup workbook against OFS's CURRENT tag
dictionary (which should already have been refreshed with the
post-rebuild XVM before this runs), then attempts to write each one
back - classifying by the actual per-item Write() result.

Before attempting anything, tags are excluded from the write list if:
  - the tag name contains "#" (OFS's own internal/system items, e.g.
    "#OFSStatus", "#ClientAlive" - never meant to be restored)
  - the tag is explicitly listed in EXCLUDED_EXACT_TAGS (e.g.
    "M580!BMEP58_ECPU_EXT", the CPU data structure - parts of it are
    not writable, so it's excluded wholesale)
  - the tag's recorded access rights mark it read-only

Reports a summary via the `log` callback (everything - excluded,
read-only, bind failures, and write outcomes), but the persistent
report FILE only records tags that actually failed to write, so it
stays a short, actionable list rather than a full transcript.

Only writes anything after the `confirm` callback returns True.
"""

from ofs_client import OFSClient
import excel_io

# Tags excluded from restore wholesale, regardless of access rights.
# M580!BMEP58_ECPU_EXT is the CPU data structure - parts of its fields
# are not writable, so the whole tag is excluded rather than attempted
# and partially failing.
EXCLUDED_EXACT_TAGS = {"M580!BMEP58_ECPU_EXT"}


def _default_confirm(tag_value_pairs):
    """Fallback confirm() for plain script/console use."""
    answer = input(
        f"\nProceed with writing these {len(tag_value_pairs)} value(s) to the CPU? [y/N]: "
    ).strip().lower()
    return answer == "y"


def _is_writable(access_rights):
    """
    access_rights is whatever get_access_rights() recorded during
    backup: "readable", "writable", "readWritable", or None (lookup
    failed at backup time - unknown, so still attempt the write and
    let it fail cleanly if truly read-only, rather than skip it).
    """
    if access_rights is None:
        return True
    return access_rights != "readable"


def _is_excluded_by_name(tag):
    """
    Tags never eligible for restore regardless of access rights:
    OFS's own internal/system items (any tag containing "#"), and
    anything explicitly listed in EXCLUDED_EXACT_TAGS.
    """
    if "#" in tag:
        return True
    if tag in EXCLUDED_EXACT_TAGS:
        return True
    return False


def run_restore(config, backup_path, report_path=None, log=print, confirm=None, progress=None):
    """
    log: callable(str) for progress/status messages.
    confirm: callable(list[(tag, value)]) -> bool. Called ONCE, with
             every tag that's eligible and about to be attempted,
             before any write happens. Must return True to proceed.
             Defaults to a console y/N prompt if not supplied.
    progress: optional callable(phase: str, current: int, total: int),
             called periodically during binding and writing.
    """
    if confirm is None:
        confirm = _default_confirm
    if progress is None:
        progress = lambda phase, current, total: None

    backed_up = excel_io.read_backup(backup_path)
    if not backed_up:
        log(f"No usable rows found in {backup_path}. Nothing to restore.")
        return

    # The tag names recorded in the backup file are already the
    # fully-qualified item IDs (discovered via browsing at backup
    # time), so they are used as-is here.
    for r in backed_up:
        r["item_id"] = r["tag"]

    log(f"Connecting to OFS ({config['ofs_progid']} @ {config['ofs_node']})...")
    with OFSClient(node=config["ofs_node"], progid=config["ofs_progid"],
                   group_name=config["group_name"],
                   netapi_lib_dir=config["godsharp_netapi_lib_dir"],
                   base_lib_dir=config["godsharp_base_lib_dir"]) as client:

        item_ids = [r["item_id"] for r in backed_up]
        log(f"Binding {len(item_ids)} tag(s) against the current dictionary...")
        bind_results = client.bind_items(
            item_ids,
            progress=lambda cur, tot: progress("Binding", cur, tot))

        failed_bind = [r for r in backed_up if not bind_results[r["item_id"]]["ok"]]
        bound = [r for r in backed_up if bind_results[r["item_id"]]["ok"]]

        skipped_excluded = [r for r in bound if _is_excluded_by_name(r["tag"])]
        remaining = [r for r in bound if not _is_excluded_by_name(r["tag"])]

        skipped_readonly = [r for r in remaining if not _is_writable(r.get("access_rights"))]
        ready_to_write = [r for r in remaining if _is_writable(r.get("access_rights"))]

        log("\n=== Restore summary ===")
        log(f"  Ready to write   : {len(ready_to_write)}")
        log(f"  Skipped (excluded by name - system items / CPU structure): {len(skipped_excluded)}")
        log(f"  Skipped (read-only, recorded at backup time): {len(skipped_readonly)}")
        log(f"  Failed to bind (renamed/removed/retyped - CANNOT be restored): {len(failed_bind)}")

        if failed_bind:
            log("\n  Tags that failed to bind:")
            for r in failed_bind:
                log(f"    - {r['tag']}")

        # Excluded/read-only tags are deliberate, expected omissions -
        # the counts above are enough; no need to itemize each one.

        # The report FILE only records tags that failed to restore -
        # either because they no longer bind, or because the write
        # itself failed. Everything else (excluded, read-only,
        # successful writes) is still visible in the log for this run.
        report_rows = []
        for r in failed_bind:
            report_rows.append({"tag": r["tag"], "item_id": r["item_id"],
                                 "status": "failed_bind", "detail": ""})

        if not ready_to_write:
            log("\nNo eligible tags to write. Nothing to write.")
            if report_path:
                excel_io.write_restore_report(report_path, report_rows)
            return

        tag_value_pairs = [(r["tag"], r["value"]) for r in ready_to_write]
        log(f"\n{len(tag_value_pairs)} tag(s) are ready to be WRITTEN to the live M580 CPU.")

        if not confirm(tag_value_pairs):
            log("Restore cancelled. No values were written.")
            if report_path:
                excel_io.write_restore_report(report_path, report_rows)
            return

        item_ids_to_write = [r["item_id"] for r in ready_to_write]
        values_to_write = [r["value"] for r in ready_to_write]
        value_types_to_write = [r.get("value_type") for r in ready_to_write]
        errors = client.write(
            item_ids_to_write, values_to_write, value_types=value_types_to_write,
            progress=lambda cur, tot: progress("Writing", cur, tot), log=log)

        written = 0
        for r, err in zip(ready_to_write, errors):
            if err == 0:
                written += 1
            else:
                report_rows.append({"tag": r["tag"], "item_id": r["item_id"],
                                     "status": "failed_write", "detail": f"error {err}"})

        log(f"\nRestore complete: {written}/{len(ready_to_write)} value(s) written successfully.")
        if written < len(ready_to_write):
            log("Some writes failed - see the restore report for details.")

    if report_path:
        excel_io.write_restore_report(report_path, report_rows)
        log(f"Restore report written to {report_path}")
