#!/usr/bin/env python3
from __future__ import annotations

import base64
import gzip
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PARTS_DIR = ROOT / "source-parts" / "v8.4.3.3"
OUTPUT = ROOT / "userscript" / "zujuan-print-workbench.user.js"
EXPECTED_SHA256 = "aa0e44004a7f92b285591f238aa55679648f128dab821a128bd70c9c8b8d2d7b"
EXPECTED_PARTS = 6


def main() -> None:
    parts = sorted(PARTS_DIR.glob("part-*.b64"))
    if len(parts) != EXPECTED_PARTS:
        raise SystemExit(f"Expected {EXPECTED_PARTS} source parts, found {len(parts)}")

    encoded = "".join(path.read_text(encoding="ascii").strip() for path in parts)
    try:
        compressed = base64.b64decode(encoded, validate=True)
        source = gzip.decompress(compressed)
    except Exception as exc:
        raise SystemExit(f"Failed to reconstruct userscript: {exc}") from exc

    digest = hashlib.sha256(source).hexdigest()
    if digest != EXPECTED_SHA256:
        raise SystemExit(
            "SHA-256 mismatch after reconstruction:\n"
            f"  expected: {EXPECTED_SHA256}\n"
            f"  actual:   {digest}"
        )

    if b"// @version      8.4.3.3" not in source[:4096]:
        raise SystemExit("Version marker 8.4.3.3 not found in Userscript header")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_bytes(source)
    print(f"Rebuilt: {OUTPUT.relative_to(ROOT)}")
    print(f"Bytes:   {len(source)}")
    print(f"SHA256:  {digest}")


if __name__ == "__main__":
    main()
