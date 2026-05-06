"""
Full-content screenshot helper for the REMFluor-MD tkinter app.

Run from this folder:
    python screenshot_app.py

What it does:
  1. Imports REMFluorApp from main.py and boots it.
  2. Maximizes the window so the canvas viewport is as large as possible.
  3. Reads the inner-frame size from the canvas scrollregion (the *full*
     width/height of the laid-out content).
  4. Scrolls the canvas in tiles, grabs each visible region with
     PIL.ImageGrab, and pastes the tiles into a single PIL.Image of the
     full content size -- so the PNG contains everything, not just the
     visible viewport.
  5. Saves PNG to ./screenshot_<timestamp>.png and closes the app.

Requirements: Pillow (already in your env per requirements.txt).

Tips for a clean grab:
  * Don't overlap any other window on top of the app while it runs --
    ImageGrab captures actual screen pixels.
  * If the result has horizontal seams or duplicated regions, tweak
    SETTLE_MS or OVERLAP_PX below.
"""
from __future__ import annotations
import os
import sys
import time
import datetime as _dt

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)

import main as _main          # noqa: E402  -- REMFluorApp lives here
from PIL import Image, ImageGrab  # noqa: E402


# -- Tunables --------------------------------------------------------
INITIAL_SETTLE_MS = 1500   # wait after app starts for layout/images
SETTLE_MS         = 180    # wait between scroll moves before grabbing
OVERLAP_PX        = 20     # tile overlap to avoid 1-px seam artefacts


# -- Helpers ---------------------------------------------------------
def _grab_canvas_viewport(canvas):
    """Grab the canvas's currently-visible region as a PIL Image."""
    canvas.update_idletasks()
    canvas.update()
    x = canvas.winfo_rootx()
    y = canvas.winfo_rooty()
    w = canvas.winfo_width()
    h = canvas.winfo_height()
    return ImageGrab.grab(bbox=(x, y, x + w, y + h), all_screens=True)


def _scroll_targets(full_size, view_size):
    """Return left-edge pixel positions to scroll to so the union of
    viewport tiles covers [0, full_size)."""
    if full_size <= view_size:
        return [0]
    step = max(view_size - OVERLAP_PX, 1)
    targets = list(range(0, full_size - view_size + 1, step))
    if targets[-1] != full_size - view_size:
        targets.append(full_size - view_size)
    return targets


def capture_full_inner(app):
    """Scroll the inner canvas across its full scrollregion and stitch
    tiles into a single PIL.Image."""
    canvas = app.canvas
    inner  = app.inner

    # Force a fresh layout
    inner.update_idletasks()
    app.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))
    canvas.update_idletasks()
    app.update()

    sr = canvas.bbox("all")
    if not sr:
        return _grab_canvas_viewport(canvas)
    full_w, full_h = sr[2], sr[3]

    canvas.xview_moveto(0.0)
    canvas.yview_moveto(0.0)
    app.update_idletasks(); app.update()
    time.sleep(SETTLE_MS / 1000)

    vw = canvas.winfo_width()
    vh = canvas.winfo_height()

    full = Image.new("RGB", (full_w, full_h), color="white")

    x_targets = _scroll_targets(full_w, vw)
    y_targets = _scroll_targets(full_h, vh)

    visited = set()
    for y in y_targets:
        canvas.yview_moveto(y / full_h if full_h else 0)
        for x in x_targets:
            canvas.xview_moveto(x / full_w if full_w else 0)
            app.update_idletasks(); app.update()
            time.sleep(SETTLE_MS / 1000)

            actual_x = int(canvas.canvasx(0))
            actual_y = int(canvas.canvasy(0))
            key = (actual_x, actual_y)
            if key in visited:
                continue
            visited.add(key)

            tile = _grab_canvas_viewport(canvas)
            full.paste(tile, (actual_x, actual_y))

    return full


# -- Main ------------------------------------------------------------
def main():
    app = _main.REMFluorApp()
    try:
        app.state("zoomed")
    except Exception:
        try:
            app.attributes("-zoomed", True)
        except Exception:
            app.geometry("1920x1080")

    out_path = os.path.join(
        _HERE,
        "screenshot_{:%Y%m%d_%H%M%S}.png".format(_dt.datetime.now()),
    )

    def _shoot():
        try:
            img = capture_full_inner(app)
            img.save(out_path, "PNG")
            print("saved {}  ({}x{} px)".format(out_path, img.size[0], img.size[1]))
        except Exception as exc:
            print("ERROR while capturing: {!r}".format(exc))
        finally:
            app.after(50, app.destroy)

    app.after(INITIAL_SETTLE_MS, _shoot)
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
