"""
clear_cache.py — wipe every __pycache__ folder under this directory.

Run from the tkinter folder:

    python clear_cache.py

or just double-click clear_cache.bat (which calls this script).

Useful after editing files in functions/ — Python caches compiled
.pyc files in __pycache__ and may keep using the stale version even
when the .py source has changed.
"""
from __future__ import annotations
import os
import shutil
import sys


_HERE = os.path.dirname(os.path.abspath(__file__))


def main() -> int:
    deleted = 0
    skipped = 0
    for root, dirs, _ in os.walk(_HERE):
        # Iterate over a copy because we're going to modify dirs in-place
        for name in list(dirs):
            if name == "__pycache__":
                target = os.path.join(root, name)
                try:
                    shutil.rmtree(target)
                    print(f"  removed: {target}")
                    deleted += 1
                except Exception as exc:
                    print(f"  SKIPPED: {target}  ({exc})")
                    skipped += 1
                # Don't recurse into a dir we just deleted
                dirs.remove(name)

    print()
    if deleted == 0 and skipped == 0:
        print("Nothing to clean. (No __pycache__ folders found.)")
    else:
        print(f"Done. Deleted {deleted} folder(s)"
              + (f", skipped {skipped}." if skipped else "."))
    return 0 if skipped == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
