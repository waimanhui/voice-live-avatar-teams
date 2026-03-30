"""
patch_acs.py — Re-apply the allowAccessRawMediaStream patch to acs-calling.js.

Run this script after updating acs-calling.js to a newer SDK version:
    python patch_acs.py

The patch enables raw MediaStream access in the ACS Calling SDK, which is
required for the Teams audio/video bridge to function.
"""

import os
import sys

# Reconfigure stdout to UTF-8 so Unicode characters print correctly on Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR  = os.path.join(SCRIPT_DIR, "..", "static")
TARGET_FILE = os.path.normpath(os.path.join(STATIC_DIR, "acs-calling.js"))

SEARCH  = "allowAccessRawMediaStream: false, deviceSelectionTimeoutInMs"
REPLACE = "allowAccessRawMediaStream: true, deviceSelectionTimeoutInMs"


def main():
    if not os.path.exists(TARGET_FILE):
        print(f"ERROR: {TARGET_FILE} not found.")
        sys.exit(1)

    with open(TARGET_FILE, encoding="utf-8") as f:
        content = f.read()

    count = content.count(SEARCH)
    if count == 0:
        # Already patched or search string changed in a newer SDK version
        if REPLACE in content:
            print("Already patched — no changes needed.")
        else:
            print("ERROR: Search string not found. The SDK may have changed.")
            print(f"  Looking for: {SEARCH!r}")
            sys.exit(1)
        return

    patched = content.replace(SEARCH, REPLACE, 1)  # only first occurrence
    with open(TARGET_FILE, "w", encoding="utf-8") as f:
        f.write(patched)

    print(f"Patched successfully ({count} occurrence found, 1 replaced).")
    print(f"  {SEARCH!r}")
    print(f"  -> {REPLACE!r}")


if __name__ == "__main__":
    main()
