"""
titanium_browser.py
Discovery via TitaniumAS.Opc.Client's native OpcDaBrowserAuto, instead
of GodSharp's BrowseNodes(). GodSharp's browse turned out to secretly
add/remove each item one at a time internally - explaining both its
extreme slowness at scale and the per-item device reconnects seen in
the OFS trace log even in a run that never called AddItems at all.
TitaniumAS's browser discovered 58,606 items in 54 seconds
(~1084 items/sec) with ZERO items created in OFS during the browse,
vs. roughly 20 minutes via GodSharp for a similarly-sized project -
confirmed via the OFS trace log staying completely clean.

TitaniumAS requires MTA apartment state; GodSharp's OFS connection
(in ofs_client.py) requires STA. An OS thread's apartment state can
only be set once, before any COM activity on it, and never changed
afterward - so these two cannot share a thread. This module always
runs its browse in a dedicated, short-lived thread of its own,
isolated from whatever thread later does GodSharp-based work.

Requires:
  - The extracted TitaniumAS.Opc.Client NuGet package, plus its
    Common.Logging / Common.Logging.Core dependencies (same version,
    3.3.0), all in the same folder - path via config.json's
    "titanium_lib_dir".
"""

import threading

MAX_BROWSE_DEPTH = 20  # safety limit against runaway recursion
PROGRESS_INTERVAL = 25  # call progress() every N items discovered, not every single one


class TitaniumBrowseError(Exception):
    """Raised when the TitaniumAS-based browse fails."""


def _browse_thread_body(lib_dir, progid, node, result, progress=None, root_branch=""):
    import clr
    import sys

    if lib_dir not in sys.path:
        sys.path.append(lib_dir)

    import System.Threading
    current_thread = System.Threading.Thread.CurrentThread
    if current_thread.GetApartmentState() != System.Threading.ApartmentState.MTA:
        current_thread.TrySetApartmentState(System.Threading.ApartmentState.MTA)

    clr.AddReference("Common.Logging.Core")
    clr.AddReference("Common.Logging")
    clr.AddReference("TitaniumAS.Opc.Client")

    from TitaniumAS.Opc.Client import Bootstrap
    from TitaniumAS.Opc.Client.Da import OpcDaServer, OpcDaElementFilter, OpcDaPropertiesQuery
    from TitaniumAS.Opc.Client.Da.Browsing import OpcDaBrowserAuto

    try:
        Bootstrap.Initialize()
    except Exception:
        pass  # safe to ignore if already initialized from an earlier browse in this process

    server = None
    try:
        server = OpcDaServer(progid, node)
        server.Connect()

        browser = OpcDaBrowserAuto(server)
        filter_ = OpcDaElementFilter()
        props_query = OpcDaPropertiesQuery(False)  # don't fetch property values, just enumerate

        items = []

        def recurse(parent_id, depth=0):
            if depth > MAX_BROWSE_DEPTH:
                return
            for el in browser.GetElements(parent_id, filter_, props_query):
                if el.IsItem:
                    items.append(el.ItemId)
                    if progress is not None and len(items) % PROGRESS_INTERVAL == 0:
                        progress(len(items))
                if el.HasChildren:
                    recurse(el.ItemId, depth + 1)

        recurse(root_branch)
        result["items"] = items
    except Exception as exc:
        result["error"] = exc
    finally:
        if server is not None:
            try:
                server.Disconnect()
            except Exception:
                pass


def browse_all_items(lib_dir, progid, node, progress=None, device_alias=None):
    """
    Discovers every leaf item OFS currently exposes, via TitaniumAS,
    in a dedicated MTA thread. Blocks until that thread finishes.
    Returns a list of fully-qualified item ID strings.

    device_alias: if given (e.g. "M580"), scopes the browse to just
    that device's branch instead of walking the whole root - useful
    when OFS has more than one device alias configured and only one
    is wanted for this backup/restore.

    progress(current), if given, is called periodically (every
    PROGRESS_INTERVAL items) with the running count discovered so
    far. The total is unknown until the browse finishes, so there is
    no total in this callback's signature - callers that share a
    (phase, current, total) style progress protocol should wrap this
    with total=0 to mean "unknown, still counting".
    """
    result = {}
    root_branch = device_alias or ""
    t = threading.Thread(target=_browse_thread_body,
                          args=(lib_dir, progid, node, result, progress, root_branch))
    t.start()
    t.join()

    if "error" in result:
        raise TitaniumBrowseError(f"TitaniumAS browse failed: {result['error']}")
    return result.get("items", [])


def _list_device_aliases_thread_body(lib_dir, progid, node, result):
    import clr
    import sys

    if lib_dir not in sys.path:
        sys.path.append(lib_dir)

    import System.Threading
    current_thread = System.Threading.Thread.CurrentThread
    if current_thread.GetApartmentState() != System.Threading.ApartmentState.MTA:
        current_thread.TrySetApartmentState(System.Threading.ApartmentState.MTA)

    clr.AddReference("Common.Logging.Core")
    clr.AddReference("Common.Logging")
    clr.AddReference("TitaniumAS.Opc.Client")

    from TitaniumAS.Opc.Client import Bootstrap
    from TitaniumAS.Opc.Client.Da import OpcDaServer, OpcDaElementFilter, OpcDaPropertiesQuery
    from TitaniumAS.Opc.Client.Da.Browsing import OpcDaBrowserAuto

    try:
        Bootstrap.Initialize()
    except Exception:
        pass

    server = None
    try:
        server = OpcDaServer(progid, node)
        server.Connect()

        browser = OpcDaBrowserAuto(server)
        filter_ = OpcDaElementFilter()
        props_query = OpcDaPropertiesQuery(False)

        # Single, non-recursive call at the root - fast, since it
        # doesn't walk into any branch's contents, just lists what's
        # immediately there. "<<system>>" is OFS's own internal
        # housekeeping branch, not a real device, so it's excluded.
        aliases = [el.Name for el in browser.GetElements("", filter_, props_query)
                   if el.HasChildren and el.Name != "<<system>>"]
        result["aliases"] = aliases
    except Exception as exc:
        result["error"] = exc
    finally:
        if server is not None:
            try:
                server.Disconnect()
            except Exception:
                pass


def list_device_aliases(lib_dir, progid, node):
    """
    Lists OFS's configured device aliases (e.g. ["M580"], or
    ["M580", "PLC2"] if more than one device is configured) - a fast,
    single, non-recursive browse at the root, in its own dedicated MTA
    thread. Excludes OFS's own internal "<<system>>" branch, which is
    not a real device.
    """
    result = {}
    t = threading.Thread(target=_list_device_aliases_thread_body, args=(lib_dir, progid, node, result))
    t.start()
    t.join()

    if "error" in result:
        raise TitaniumBrowseError(f"TitaniumAS device list failed: {result['error']}")
    return result.get("aliases", [])
