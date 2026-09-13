"""
config.py
Loads connection/config settings for the OFS backup/restore tool from a
simple JSON file, so nothing environment-specific is hardcoded in the
other modules.
"""

import json
import os

DEFAULT_CONFIG = {
    # ProgID of the OFS server itself (not a wrapper - GodSharp.Opc.Da.OpcNetApi
    # talks to it via native OPC DA COM interop, with no Automation
    # Wrapper in the middle at all).
    "ofs_progid": "Schneider-Aut.OFS",

    # "localhost" if this tool runs on the same PC as OFS.
    "ofs_node": "localhost",

    # Name for the OPC group this tool creates. Cosmetic only.
    "group_name": "BackupRestoreGroup",

    # Folders containing the extracted GodSharp.Opc.Da.OpcNetApi and
    # GodSharp.Opc.Da NuGet packages' net46 assemblies. Adjust these
    # to wherever you copied the extracted packages.
    "godsharp_netapi_lib_dir": r"C:\DNET\godsharp.opc.da.opcnetapi.2022.308.10\lib\net46",
    "godsharp_base_lib_dir": r"C:\DNET\godsharp.opc.da.2022.308.10\lib\net46",

    # Folder containing TitaniumAS.Opc.Client.dll (plus its
    # Common.Logging / Common.Logging.Core 3.3.0 dependencies, same
    # folder) - used for fast discovery/browsing only.
    "titanium_lib_dir": r"C:\DNET\titaniumas.opc.client.1.0.2\lib\net40",
}


def load_config(path="config.json"):
    """
    Load config.json if it exists, otherwise fall back to defaults.
    Any keys missing from the file are filled in from DEFAULT_CONFIG.
    """
    config = dict(DEFAULT_CONFIG)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            user_config = json.load(f)
        config.update(user_config)
    return config


def write_default_config(path="config.json"):
    """Write out a starter config.json for the user to edit."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, indent=2)
