"""
backup.py
Discovers every tag OFS currently knows about (via TitaniumAS's native
browse) and reads their values from the M580 CPU, writing the result
to a backup workbook. No externally supplied tag list is used.
Function block instance members are always excluded - see
_is_fb_instance_member() and run_backup() for why.
"""

from ofs_client import OFSClient
import excel_io


def _excel_safe(value):
    """
    openpyxl can only write plain scalars (str/int/float/bool/None)
    into a cell. Some items (e.g. OFS's own '#OFSStatus') return a
    .NET array rather than a scalar - convert those to a simple
    semicolon-separated string instead of crashing. Restoring such a
    value later will very likely just fail cleanly on a type mismatch
    (consistent with how any other unrestorable tag is handled), which
    is an acceptable tradeoff for what are OFS-internal diagnostic
    values rather than real M580 process data.
    """
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    try:
        items = list(value)  # works for .NET arrays via IEnumerable
        return ";".join(str(x) for x in items)
    except TypeError:
        return str(value)


def _is_fb_instance_member(item_id):
    """
    A "." in the leaf name (the part after the last "!" branch
    separator) indicates a member of a function block instance, e.g.
    "M580!TABLE_BIT_COMP_V1_1.SRC" - not a plain tag. These are always
    excluded from discovery, unconditionally - see run_backup()'s
    docstring for why this is no longer a user-facing option.
    """
    leaf = item_id.rsplit("!", 1)[-1]
    return "." in leaf


def run_backup(config, output_path, log=print, progress=None, device_alias=None):
    """
    log: callable(str) used for progress/status messages, so callers
    (CLI, GUI, etc.) can route output wherever they like. Defaults to
    print() for plain script use.
    progress: optional callable(phase: str, current: int, total: int),
    called periodically during binding, reading, and access-rights
    lookup, so a caller can show live chunk-by-chunk progress.
    device_alias: if given, scopes discovery to just that OFS device
    alias (see OFSClient.list_device_aliases()) instead of every
    device configured in OFS - useful when OFS has more than one
    device and only one is wanted for this backup.

    Function block instance members (item IDs with a "." in the leaf
    name, e.g. "M580!SomeInstance.SomeField") are always excluded from
    discovery. This used to be a user-facing "include FB instance
    data" option, removed after discovering it could never do what it
    was meant to: a function block instance's PRIVATE internal
    variables are never published to the XVM file or the CPU's Data
    Dictionary at all, regardless of Control Expert's "make internal
    variables accessible" project setting - that setting only affects
    whether the application's OWN code can reference them elsewhere in
    the project, not whether OFS (or any external tool) can see them.
    The dotted members that DID appear when this option was enabled
    were only ever the block's public I/O interface (inputs, outputs,
    in-outs) - never the private internals the option was meant to
    capture - so there was no real case left for keeping it.
    """
    if progress is None:
        progress = lambda phase, current, total: None

    log(f"Connecting to OFS ({config['ofs_progid']} @ {config['ofs_node']})...")
    with OFSClient(node=config["ofs_node"], progid=config["ofs_progid"],
                   group_name=config["group_name"],
                   netapi_lib_dir=config["godsharp_netapi_lib_dir"],
                   base_lib_dir=config["godsharp_base_lib_dir"],
                   titanium_lib_dir=config["titanium_lib_dir"]) as client:

        log("Discovering tags from OFS..." if not device_alias
            else f"Discovering tags from OFS (device: {device_alias})...")
        progress("Discovering", 0, 0)
        all_discovered = client.browse_all_items(
            progress=lambda current: progress("Discovering", current, 0),
            device_alias=device_alias)
        progress("Discovered", len(all_discovered), len(all_discovered))
        item_ids = [iid for iid in all_discovered if not _is_fb_instance_member(iid)]
        excluded_count = len(all_discovered) - len(item_ids)
        log(f"Discovered {len(all_discovered)} item(s).")
        if excluded_count:
            log(f"Excluding {excluded_count} function block instance member(s) "
                f"(e.g. 'InstanceName.Member') - only plain tags are backed up.")

        if not item_ids:
            log("No tags discovered. Nothing to back up.")
            return

        log(f"Binding {len(item_ids)} tag(s)...")
        bind_results = client.bind_items(
            item_ids,
            progress=lambda cur, tot: progress("Binding", cur, tot))

        bound_item_ids = [iid for iid, r in bind_results.items() if r["ok"]]
        unbound_item_ids = [iid for iid, r in bind_results.items() if not r["ok"]]

        if unbound_item_ids:
            log(f"WARNING: {len(unbound_item_ids)} tag(s) failed to bind and will be recorded with no value:")
            for iid in unbound_item_ids:
                log(f"  - {iid}  ({bind_results[iid]['error']})")

        values_by_item_id = {}
        errors_by_item_id = {}
        value_types_by_item_id = {}
        if bound_item_ids:
            log(f"Reading {len(bound_item_ids)} tag(s)...")
            values, errors, value_types = client.read(
                bound_item_ids,
                progress=lambda cur, tot: progress("Reading", cur, tot))
            for iid, val, err, vtype in zip(bound_item_ids, values, errors, value_types):
                values_by_item_id[iid] = val
                errors_by_item_id[iid] = err
                value_types_by_item_id[iid] = vtype

        rows = []
        for item_id in item_ids:
            bind = bind_results[item_id]
            ok = bind["ok"] and errors_by_item_id.get(item_id, 0) == 0
            rows.append({
                "tag": item_id,
                "item_id": item_id,
                "bind_ok": ok,
                "value": _excel_safe(values_by_item_id.get(item_id)),
                "value_type": value_types_by_item_id.get(item_id),
                "access_rights": None,  # no longer fetched - was the last unbatched, one-item-at-a-time
                                         # call in the pipeline; restore's read-only handling already
                                         # falls back gracefully to "attempt the write, let it fail
                                         # cleanly if truly read-only" when this is None
                "error_code": errors_by_item_id.get(item_id, bind["error"]),
            })

    excel_io.write_backup(output_path, rows)
    ok_count = sum(1 for r in rows if r["bind_ok"])
    failed_rows = [r for r in rows if not r["bind_ok"]]
    log(f"{ok_count}/{len(rows)} tags were saved properly. "
        f"{len(failed_rows)} tag(s) failed to bind.")
    if failed_rows:
        log("Failed tags:")
        for r in failed_rows:
            log(f"  - {r['tag']}")
    log(f"Written to {output_path}")
