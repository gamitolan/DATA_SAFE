"""
ofs_client.py
GodSharp.Opc.Da.OpcNetApi-based OFS client - uses the OPC Foundation's
official .NET API via native COM interop (through pythonnet), NOT
OFS's Automation Wrapper (Schneider's opcautosa2.dll / ofsauto.dll),
which was confirmed broken (AddItems fails universally, regardless of
content or call shape) after extensive troubleshooting.

Requires:
  - pythonnet (`pip install pythonnet`)
  - The extracted GodSharp.Opc.Da and GodSharp.Opc.Da.OpcNetApi NuGet
    packages present on disk - paths configured via config.json's
    "godsharp_netapi_lib_dir" / "godsharp_base_lib_dir".

All long-running methods (bind_items/read/write) accept an optional
`progress(current, total)` callback, called periodically so a caller
(e.g. the GUI) can show live chunk-by-chunk progress rather than a
long silent stretch with no feedback.
"""

import clr
import sys

import titanium_browser

_ASSEMBLIES_LOADED = False


def _ensure_assemblies_loaded(netapi_lib_dir, base_lib_dir):
    """Adds the GodSharp/.NET assembly search paths and loads them.
    Safe to call more than once - only does real work the first time.
    """
    global _ASSEMBLIES_LOADED
    if _ASSEMBLIES_LOADED:
        return

    for path in (netapi_lib_dir, base_lib_dir):
        if path not in sys.path:
            sys.path.append(path)

    clr.AddReference("OpcNetApi")
    clr.AddReference("OpcNetApi.Com")
    clr.AddReference("OpcComRcw")
    clr.AddReference("GodSharp.Opc.Da")
    clr.AddReference("GodSharp.Opc.Da.OpcNetApi")

    _ASSEMBLIES_LOADED = True


def _noop_progress(current, total):
    pass


def _clr_type_name(value):
    """
    Returns the value's real .NET type name (e.g. "Int16", "Single",
    "Boolean") if it can be determined, else falls back to the
    Python-side type name. Used to record each value's exact native
    type at read time, since once a value round-trips through Excel
    the distinction between e.g. Int16/Int32/Int64 or Single/Double
    is lost - Excel and openpyxl only know generic int/float/bool.
    """
    if value is None:
        return None
    try:
        return str(value.GetType().Name)
    except Exception:
        return type(value).__name__


_SCALAR_CTORS_BY_NAME = None  # populated lazily on first use, needs System imported


def _scalar_ctors():
    global _SCALAR_CTORS_BY_NAME
    if _SCALAR_CTORS_BY_NAME is None:
        import System
        _SCALAR_CTORS_BY_NAME = {
            "Boolean": System.Boolean,
            "Byte": System.Byte,
            "SByte": System.SByte,
            "Int16": System.Int16,
            "UInt16": System.UInt16,
            "Int32": System.Int32,
            "UInt32": System.UInt32,
            "Int64": System.Int64,
            "UInt64": System.UInt64,
            "Single": System.Single,
            "Double": System.Double,
            "String": System.String,
        }
    return _SCALAR_CTORS_BY_NAME


def _parse_scalar_text(text, element_type_name):
    """
    Parses one semicolon-split piece of an array's stored display
    string back into the right Python value for that element type,
    before handing it to the .NET constructor. Needs care for
    Boolean specifically - Python's bool("False") is True (any
    non-empty string is truthy), so a naive cast would silently
    corrupt every False element.
    """
    text = text.strip()
    if element_type_name == "Boolean":
        return text.lower() == "true"
    if element_type_name in ("Byte", "SByte", "Int16", "UInt16", "Int32", "UInt32", "Int64", "UInt64"):
        return int(text)
    if element_type_name in ("Single", "Double"):
        return float(text)
    return text  # String or unrecognized - leave as text


def _coerce_value(value, type_name):
    """
    Re-boxes a plain Python value back into the exact .NET type it was
    recorded as at backup time - either a scalar (e.g. Int16, Single)
    or an ARRAY (e.g. "Boolean[]" - as seen for an EBOOL array in
    Control Expert, which OFS exposes as a genuine array rather than
    individually-addressable elements). Without this:
      - a scalar value that started as e.g. an Int16 or Single gets
        written back as whatever pythonnet defaults a plain Python
        int/float to, which OFS can reject with OPC_E_BADTYPE even
        though the value itself is valid for that item.
      - an array value is unrecoverable: _excel_safe() (see backup.py)
        has to flatten any array into a semicolon-joined display
        string just to store it in an Excel cell at all, since neither
        Excel nor openpyxl can hold a raw array. Without rebuilding a
        genuine typed array here, that display string would just be
        written back as a plain string, which is certain to fail for
        an item that expects e.g. a real Boolean[].

    Falls back to the original value unchanged if coercion fails for
    any reason (e.g. an unrecognized type name, or the type wasn't
    recorded at all).
    """
    if value is None or not type_name:
        return value
    try:
        import System
        ctors = _scalar_ctors()

        if type_name.endswith("[]"):
            element_type_name = type_name[:-2]
            element_ctor = ctors.get(element_type_name)
            if element_ctor is None:
                return value  # unrecognized element type - leave unchanged

            # value is the semicolon-joined display string written by
            # _excel_safe() at backup time (or possibly already a
            # sequence, defensively handled too).
            if isinstance(value, str):
                parts = value.split(";") if value else []
            else:
                parts = list(value)

            arr = System.Array.CreateInstance(element_ctor, len(parts))
            for i, part in enumerate(parts):
                parsed = _parse_scalar_text(str(part), element_type_name)
                arr[i] = element_ctor(parsed)
            return arr

        ctor = ctors.get(type_name)
        if ctor is None:
            return value
        return ctor(value)
    except Exception:
        return value


class OFSError(Exception):
    """Raised for any OFS connection, binding, or read/write failure."""


class OFSClient:
    """
    Manages one connection to OFS via GodSharp.Opc.Da.OpcNetApi.

    Unlike the old Automation-Wrapper-based client, there is no
    separate "server handle" concept exposed here - tags are tracked
    internally (by GodSharp) keyed by item_id string, and read/write
    operate directly on item_id strings.
    """

    def __init__(self, node="localhost", progid="Schneider-Aut.OFS",
                 group_name="BackupRestoreGroup",
                 netapi_lib_dir=None, base_lib_dir=None, titanium_lib_dir=None):
        self.node = node
        self.progid = progid
        self.group_name = group_name
        self.netapi_lib_dir = netapi_lib_dir
        self.base_lib_dir = base_lib_dir
        self.titanium_lib_dir = titanium_lib_dir
        self._client = None
        self._subscription = None

    def connect(self):
        _ensure_assemblies_loaded(self.netapi_lib_dir, self.base_lib_dir)

        # Imported here (after assemblies are guaranteed loaded) rather
        # than at module level, since clr.AddReference must run first.
        from GodSharp.Opc.Da import OpcNetApiClient, DaClientOptions, Group
        from GodSharp.Opc.Da.Options import ServerData
        import System.Threading

        # A newly spawned thread (e.g. a GUI worker thread) defaults to
        # MTA on the .NET side, but OPC DA COM objects conventionally
        # require STA - set it explicitly before any COM object is
        # created on this thread. Must happen before the first COM
        # call, or it has no effect (silently ignored if already set).
        current_thread = System.Threading.Thread.CurrentThread
        if current_thread.GetApartmentState() != System.Threading.ApartmentState.STA:
            current_thread.TrySetApartmentState(System.Threading.ApartmentState.STA)

        data = ServerData()
        data.ProgId = self.progid
        data.Host = self.node
        data.Name = "OFSBackupTool"

        options = DaClientOptions()
        options.Data = data

        self._client = OpcNetApiClient(options)
        try:
            ok = self._client.Connect()
        except Exception as exc:
            self._client = None
            raise OFSError(f"Could not connect to OFS ({self.progid} @ {self.node}): {exc}") from exc

        if not ok:
            self._client = None
            raise OFSError(f"Could not connect to OFS ({self.progid} @ {self.node})")

        group = Group()
        group.Name = self.group_name
        group.UpdateRate = 1000
        self._subscription = self._client.Add(group)

    def disconnect(self):
        if self._client is not None:
            try:
                self._client.Disconnect()
            except Exception:
                pass
        self._client = None
        self._subscription = None

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()

    def browse_all_items(self, progress=None, device_alias=None):
        """
        Discovers every leaf item OFS currently exposes, via
        TitaniumAS.Opc.Client's native OpcDaBrowserAuto - NOT via
        GodSharp's BrowseNodes(), which turned out to secretly
        add/remove each item one at a time internally (explaining
        both its extreme slowness at scale and the per-item device
        reconnects seen in the OFS trace even when AddItems was
        never called). TitaniumAS browsed a 58,606-item project in
        54s (~1084 items/sec) with zero items created in OFS, vs.
        roughly 20 minutes via GodSharp for a similarly-sized project.

        device_alias: if given, scopes discovery to just that OFS
        device alias (see list_device_aliases()) instead of every
        device configured in OFS.

        progress(current), if given, is called periodically with the
        running count discovered so far (total is unknown until the
        browse finishes).

        Runs in its own dedicated MTA thread (see titanium_browser.py)
        since TitaniumAS requires MTA while this class's own
        connect() requires STA for GodSharp - the two can never share
        a thread, so this is fully self-contained and does not
        require self._client to be connected at all.
        """
        return titanium_browser.browse_all_items(self.titanium_lib_dir, self.progid, self.node,
                                                   progress=progress, device_alias=device_alias)

    def list_device_aliases(self):
        """
        Lists OFS's configured device aliases (e.g. ["M580"], or more
        than one if multiple devices are configured on this OFS
        instance) - a fast, single, non-recursive browse. Same
        dedicated-MTA-thread isolation as browse_all_items().
        """
        return titanium_browser.list_device_aliases(self.titanium_lib_dir, self.progid, self.node)

    def bind_items(self, item_ids, batch_size=50, progress=None):
        """
        Adds tags to the group, in batches. Returns a dict keyed by
        item_id: {"ok": bool, "error": str or None}.

        If a batch fails, falls back to adding that batch's items one
        at a time, so a single bad tag doesn't take the whole batch
        down with it.

        progress(current, total) is called after each batch, so a
        caller can show live chunk-by-chunk progress. batch_size
        defaults to 50 (smaller than the internal read/write chunking)
        specifically to give reasonably frequent progress updates
        without meaningfully hurting throughput at the tag counts this
        tool deals with.
        """
        if self._subscription is None:
            raise OFSError("Not connected. Call connect() first.")
        if progress is None:
            progress = _noop_progress

        from GodSharp.Opc.Da import Tag

        results = {}
        total = len(item_ids)
        for start in range(0, total, batch_size):
            chunk = item_ids[start:start + batch_size]
            try:
                tags = [Tag(item_id, i, 0) for i, item_id in enumerate(chunk, start=1)]
                self._subscription.Add(list(tags))
                for item_id in chunk:
                    results[item_id] = {"ok": True, "error": None}
            except Exception:
                for item_id in chunk:
                    try:
                        self._subscription.Add([Tag(item_id, 1, 0)])
                        results[item_id] = {"ok": True, "error": None}
                    except Exception as inner_exc:
                        results[item_id] = {"ok": False, "error": str(inner_exc)}
            progress(min(start + batch_size, total), total)
        return results

    def read(self, item_ids, progress=None, batch_size=500):
        """
        Reads a list of item_ids (must already be bound via
        bind_items). Returns (values, errors, value_types), all
        aligned with item_ids - errors[i] is 0 for success, or a
        nonzero/string marker on failure. value_types[i] is the
        value's real .NET type name (e.g. "Int16", "Single"), or None
        on failure - recorded so a later restore can re-box the value
        into its exact original type (see _coerce_value) rather than
        letting a generic Python int/float default to the wrong
        native width/precision, which OFS can reject on write.

        Uses GodSharp's batch Reads(String[]) per chunk rather than
        one Read() call per item - a single-item call pays the full
        COM/interop round-trip cost every time, which becomes the
        dominant cost at large tag counts (tens of thousands of
        items). If a batch call fails for any reason, falls back to
        one-at-a-time for that chunk so a single bad item (or an
        unexpected batch-call issue) doesn't take the whole chunk
        down with it.

        progress(current, total) is called after each chunk.
        """
        if self._subscription is None:
            raise OFSError("Not connected. Call connect() first.")
        if progress is None:
            progress = _noop_progress

        values, errors, value_types = [], [], []
        total = len(item_ids)
        for start in range(0, total, batch_size):
            chunk = item_ids[start:start + batch_size]
            try:
                results = self._subscription.Reads(list(chunk))
                if len(results) != len(chunk):
                    raise OFSError("Reads() returned a different count than requested - "
                                    "falling back to one-at-a-time for this chunk.")
                for result in results:
                    if result.Ok:
                        val = result.Result.Value
                        values.append(val)
                        errors.append(0)
                        value_types.append(_clr_type_name(val))
                    else:
                        values.append(None)
                        errors.append(result.Code)
                        value_types.append(None)
            except Exception:
                for item_id in chunk:
                    try:
                        result = self._subscription.Read(item_id)
                    except Exception as exc:
                        values.append(None)
                        errors.append(str(exc))
                        value_types.append(None)
                        continue
                    if result.Ok:
                        val = result.Result.Value
                        values.append(val)
                        errors.append(0)
                        value_types.append(_clr_type_name(val))
                    else:
                        values.append(None)
                        errors.append(result.Code)
                        value_types.append(None)
            progress(min(start + batch_size, total), total)
        return values, errors, value_types

    def get_access_rights(self, item_id):
        """
        Returns the item's access rights as GodSharp reports them
        ("readable", "writable", or "readWritable"), or None if the
        lookup itself failed. None should be treated as "unknown -
        attempt the write anyway and let it fail cleanly if truly
        read-only" rather than as a reason to skip it.
        """
        if self._client is None:
            raise OFSError("Not connected. Call connect() first.")
        try:
            props = self._client.GetItemProperties(item_id)
            return str(props[5].Value)  # 5 = standard OPC DA "Item Access Rights" property ID
        except Exception:
            return None

    def write(self, item_ids, values, value_types=None, progress=None, log=None, batch_size=25):
        """
        Writes values to a list of item_ids. Returns an errors list
        aligned with item_ids (0 = success).

        value_types: optional list aligned with item_ids/values,
        giving each value's recorded .NET type name (from read()'s
        value_types output at backup time). Each value is re-boxed
        into that exact type before writing (see _coerce_value) -
        without this, a value that started as e.g. an Int16 or Single
        gets written back as whatever pythonnet defaults a generic
        Python int/float to, which OFS can reject with OPC_E_BADTYPE
        even though the value itself is valid for that item.

        Uses GodSharp's batch Writes(KeyValuePair[]) with a fixed
        batch size of 25 - determined empirically to be the sweet
        spot on a large real project (13,488 tags): 10/batch was
        slower (too many round trips), 50/batch was slower (large
        batches occasionally land entirely inside a cluster of
        genuinely-bad tags - a bad item appears to disrupt the
        underlying Modbus request stream for its neighbors in the
        same batch, so a bigger batch means a bigger "blast radius"),
        and an adaptive version that shrank/grew the batch size in
        response to failure rate was also tried and came out slower
        still, since it spent time shrunk down to sizes already known
        to be worse. A plain fixed size of 25 is simpler and faster
        than either alternative tried so far.

        After a batch call, every failed item is retried individually
        regardless of how many failed - correctness matters more than
        speed for restore (capping retries by failure count was tried
        and measurably lowered the real restore success rate, since
        individual Write() calls are genuinely more tolerant of type
        mismatches than the batched call, not just a check for
        collateral damage from a bad neighbor).

        progress(current, total) is called after each chunk.
        log(str), if given, is called once per batch with a notable
        number of failures, purely for visibility into how much of
        the run is taking the slower individual-retry path.
        """
        if self._subscription is None:
            raise OFSError("Not connected. Call connect() first.")
        if progress is None:
            progress = _noop_progress
        if log is None:
            log = lambda msg: None
        if value_types is None:
            value_types = [None] * len(item_ids)

        LOG_FAILURE_THRESHOLD = 20  # note in the log when a batch has more failures than this - informational only, does not skip retrying

        from System.Collections.Generic import KeyValuePair
        from System import String, Object

        errors = []
        total = len(item_ids)
        for start in range(0, total, batch_size):
            chunk_ids = item_ids[start:start + batch_size]
            chunk_values = values[start:start + batch_size]
            chunk_types = value_types[start:start + batch_size]
            coerced_values = [_coerce_value(v, t) for v, t in zip(chunk_values, chunk_types)]

            chunk_errors = [None] * len(chunk_ids)
            try:
                pairs = [KeyValuePair[String, Object](item_id, value)
                         for item_id, value in zip(chunk_ids, coerced_values)]
                results = self._subscription.Writes(list(pairs))
                if len(results) != len(chunk_ids):
                    raise OFSError("Writes() returned a different count than requested - "
                                    "falling back to one-at-a-time for this chunk.")
                for i, result in enumerate(results):
                    chunk_errors[i] = 0 if result.Ok else result.Code
            except Exception:
                for i, (item_id, value) in enumerate(zip(chunk_ids, coerced_values)):
                    try:
                        result = self._subscription.Write(item_id, value)
                    except Exception as exc:
                        chunk_errors[i] = str(exc)
                        continue
                    chunk_errors[i] = 0 if result.Ok else result.Code

            # Retry every failed item individually, regardless of how
            # many failed in this batch - see docstring above for why
            # this isn't capped despite the speed cost.
            failed_indices = [i for i, err in enumerate(chunk_errors) if err != 0]
            if len(failed_indices) > LOG_FAILURE_THRESHOLD:
                log(f"Batch had {len(failed_indices)}/{len(chunk_ids)} write failures - "
                    f"retrying each individually (slower for this batch).")
            for i in failed_indices:
                try:
                    retry_result = self._subscription.Write(chunk_ids[i], coerced_values[i])
                    chunk_errors[i] = 0 if retry_result.Ok else retry_result.Code
                except Exception as exc:
                    chunk_errors[i] = str(exc)

            errors.extend(chunk_errors)
            progress(min(start + batch_size, total), total)
        return errors
