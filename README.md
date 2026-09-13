# M580 Tag Backup/Restore via OFS

Backs up and restores M580 tag values **by symbolic name**, so a
Control Expert Rebuild All + full download doesn't silently reset
unlocated tags to their preset values.

## Architecture (revised)

This tool talks to OFS via **`GodSharp.Opc.Da.OpcNetApi`**, a .NET
library built on the OPC Foundation's own official .NET API, using
**native OPC DA COM interop** - NOT OFS's Automation Wrapper
(Schneider's `opcautosa2.dll` / `ofsauto.dll`).

That wrapper was extensively tested and found to be broken on this
installation: every call to its `AddItems` method failed identically
regardless of content, argument shape, or which of three different
Python COM approaches (`pywin32` dynamic dispatch, `pywin32` early
binding, `comtypes`) was used - confirmed via OFS's own connection
trace log that the calls never even reached the server process. A full
OFS repair install did not fix it. Renaming the wrapper DLLs and
confirming OFS's own bundled test client kept working proved OFS
itself doesn't use them either - they are a legacy, apparently
unmaintained convenience layer.

`GodSharp.Opc.Da.OpcNetApi`, reached from Python via **`pythonnet`**
(the `clr` module, a Python-to-.NET bridge), bypasses that layer
entirely and talks to OFS the same way OFS's own client and an
independent third-party OPC DA client both do - and this was confirmed
working end-to-end: connect, browse, add a real M580 tag, and read it
back with a valid quality and timestamp, all succeeded.

## Requirements

- Windows PC with OFS installed and running (Demo license is
  sufficient for this tool's usage pattern - see notes below).
- Python 3.9+ (Tkinter ships with the standard Windows installer).
- `pip install -r requirements.txt` (now includes `pythonnet`).
- The extracted **`GodSharp.Opc.Da`** and **`GodSharp.Opc.Da.OpcNetApi`**
  NuGet packages present on disk. A `.nupkg` is just a zip file -
  rename to `.zip` and extract. You need the `net46` subfolder from
  each package's `lib\` directory, which contains:
  - From `GodSharp.Opc.Da`: `GodSharp.Opc.Da.dll`
  - From `GodSharp.Opc.Da.OpcNetApi`: `GodSharp.Opc.Da.OpcNetApi.dll`,
    `OpcNetApi.dll`, `OpcNetApi.Com.dll`, `OpcComRcw.dll`

## Setup

Edit `config.json` (or use the defaults in `config.py`) to point at
wherever you extracted the two packages:

```json
{
  "ofs_progid": "Schneider-Aut.OFS",
  "ofs_node": "localhost",
  "group_name": "BackupRestoreGroup",
  "godsharp_netapi_lib_dir": "C:\\DNET\\godsharp.opc.da.opcnetapi.2022.308.10\\lib\\net46",
  "godsharp_base_lib_dir": "C:\\DNET\\godsharp.opc.da.2022.308.10\\lib\\net46"
}
```

`ofs_progid` here is **OFS's own ProgID** - not a wrapper ProgID, since
there is no wrapper in this architecture at all.

## Running the app

```
python app.py
```

Same GUI as before: choose **Backup** or **Restore**, pick a file
path, click **Run**. On Restore, after binding tags you get a popup
listing exactly what will be written - nothing touches the CPU until
you click **Yes**.

## Workflow

1. **Backup** (before making any changes): select Backup mode, pick an
   output path, Run. The tool connects to OFS, browses its full tag
   namespace (via GodSharp's native browse, which returns correct,
   fully-qualified item paths directly - e.g. `M580!R`,
   `<<system>>!#OFSStatus`), binds every discovered item, reads its
   value, and saves the result.
2. **Make your changes, Rebuild All, download to the M580.** Leave the
   CPU in STOP. Let the XVM auto-export as configured.
3. **Restart OFS** so it picks up the new XVM cleanly.
4. **Restore**: select Restore mode, pick the backup file (and
   optionally a report path), Run. Confirm the write-list popup. A
   genuinely read-only tag (e.g. a physical input) simply fails its
   write cleanly and shows up in the report - it doesn't need to be
   pre-filtered, since the native client's per-item Write() result is
   reliable in a way the old wrapper's batch calls never were.
5. **Put the CPU back in RUN.**

## On the OFS Demo license

Demo's only restriction is a 3-day runtime limit per execution session
- no item-count cap (unlike Small's 1000-item limit). Since this tool
connects, does its read or write pass, and disconnects, the 3-day
limit isn't a practical constraint, and a full-namespace discovery
easily exceeds 1000 items for a BMS this size anyway.

## Diagnostic scripts

The project folder also contains a number of `test_*.py` and
`explore_*.py` scripts from the troubleshooting process that led here.
They're not part of the application itself, but kept for reference -
particularly `explore_godsharp_methods.py` / `explore_godsharp_details.py`
(which reflect the real GodSharp API surface) and `test_godsharp_e2e.py`
(the first successful end-to-end proof).
