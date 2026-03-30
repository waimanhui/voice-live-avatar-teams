"""
build_acs.py — Downloads and rebuilds acs-calling.js from npm packages.

This script produces the acs-calling.js browser bundle used by the
Voice Live Avatar app, then automatically applies the allowAccessRawMediaStream
patch required for the Teams audio/video bridge.

Usage:
    cd recuild-acs
    python build_acs.py                     # uses default (latest) SDK version
    python build_acs.py --version 1.34.1    # pin a specific version

Requirements:
    - Node.js + npm  (https://nodejs.org)
    - Python 3.8+

What it does:
    1. Creates a temporary npm project
    2. Installs @azure/communication-calling and @azure/communication-common
    3. Writes a minimal acs-entry.js entrypoint that re-exports both packages
    4. Bundles everything with esbuild into acs-calling.js (IIFE, global name AzureCommunicationCalling)
    5. Runs patch_acs.py to apply the allowAccessRawMediaStream patch
    6. Cleans up the temp directory
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile

# Reconfigure stdout to UTF-8 so Unicode characters print correctly on Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR   = os.path.join(SCRIPT_DIR, "..", "static")
OUTPUT_FILE  = os.path.normpath(os.path.join(STATIC_DIR, "acs-calling.js"))
PATCH_SCRIPT = os.path.normpath(os.path.join(SCRIPT_DIR, "patch_acs.py"))  # lives next to build_acs.py

ACS_ENTRY = """\
// Auto-generated entrypoint — do not edit
export * from '@azure/communication-calling';
export * from '@azure/communication-common';
"""

PACKAGE_JSON_TEMPLATE = """\
{{
  "name": "acs-bundle-build",
  "version": "1.0.0",
  "private": true,
  "dependencies": {{
    "@azure/communication-calling": "{calling_version}",
    "@azure/communication-common": "latest",
    "esbuild": "latest"
  }}
}}
"""

ESBUILD_CMD = (
    "npx esbuild acs-entry.js "
    "--bundle "
    "--format=iife "
    "--global-name=AzureCommunicationCalling "
    "--platform=browser "
    "--outfile={outfile}"
)


def run(cmd, cwd, description):
    print(f"\n>>> {description}")
    print(f"    {cmd}\n")
    result = subprocess.run(cmd, shell=True, cwd=cwd)
    if result.returncode != 0:
        print(f"ERROR: '{description}' failed with exit code {result.returncode}")
        sys.exit(result.returncode)


def main():
    parser = argparse.ArgumentParser(description="Rebuild acs-calling.js from npm")
    parser.add_argument(
        "--version", default="latest",
        help="@azure/communication-calling version to install (default: latest)"
    )
    args = parser.parse_args()

    calling_version = args.version
    print(f"Building acs-calling.js  (communication-calling@{calling_version})")

    tmpdir = tempfile.mkdtemp(prefix="acs-build-")
    try:
        # 1. Write package.json
        pkg_json = os.path.join(tmpdir, "package.json")
        with open(pkg_json, "w") as f:
            f.write(PACKAGE_JSON_TEMPLATE.format(calling_version=calling_version))

        # 2. Write entrypoint
        entry = os.path.join(tmpdir, "acs-entry.js")
        with open(entry, "w") as f:
            f.write(ACS_ENTRY)

        # 3. npm install
        run("npm install", cwd=tmpdir, description="npm install")

        # 4. Detect installed version
        installed_version = calling_version
        try:
            result = subprocess.run(
                "npm list @azure/communication-calling --depth=0 --json",
                shell=True, cwd=tmpdir, capture_output=True, text=True
            )
            import json
            data = json.loads(result.stdout)
            installed_version = (
                data.get("dependencies", {})
                    .get("@azure/communication-calling", {})
                    .get("version", calling_version)
            )
        except Exception:
            pass

        # 5. esbuild bundle
        outfile = OUTPUT_FILE.replace("\\", "/")
        run(
            ESBUILD_CMD.format(outfile=outfile),
            cwd=tmpdir,
            description=f"esbuild → {OUTPUT_FILE}"
        )

        print(f"\n✓ Bundle written to: {OUTPUT_FILE}")
        print(f"  @azure/communication-calling version: {installed_version}")

        # 6. Apply patch
        print("\n>>> Applying allowAccessRawMediaStream patch...")
        result = subprocess.run(
            [sys.executable, PATCH_SCRIPT],
            capture_output=True, text=True
        )
        print(result.stdout.strip())
        if result.returncode != 0:
            print(f"WARNING: patch_acs.py failed:\n{result.stderr}")

        # 7. Update acs-bundle-info.json if present
        bundle_info_path = os.path.normpath(os.path.join(STATIC_DIR, "acs-bundle-info.json"))
        if os.path.exists(bundle_info_path):
            import json
            with open(bundle_info_path, encoding="utf-8") as f:
                info = json.load(f)
            info["dependencies"]["@azure/communication-calling"] = installed_version
            with open(bundle_info_path, "w", encoding="utf-8") as f:
                json.dump(info, f, indent=2)
            print(f"  acs-bundle-info.json updated with version {installed_version}")

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    print("\nDone.")


if __name__ == "__main__":
    main()
