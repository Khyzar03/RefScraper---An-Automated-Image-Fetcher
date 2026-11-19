"""
OMEGA watch image collector - Precision First-Result CDP Version (fixed)
"""

import asyncio
import re
import time
from collections import deque
from pathlib import Path
from io import BytesIO
from urllib.parse import quote

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage, _import_image
from openpyxl.utils import get_column_letter, coordinate_to_tuple
from PIL import Image as PILImage


# Patch openpyxl Image._data to avoid closing BytesIO stream during progressive saves
def _safe_image_data(self):
    img = _import_image(self.ref)
    if self.format in ["gif", "jpeg", "png"]:
        img.fp.seek(0)
        data = img.fp.read()
        img.fp.seek(0)
        return data
    else:
        fp = BytesIO()
        img.save(fp, format="png")
        data = fp.getvalue()
        return data


XLImage._data = _safe_image_data

# ---------- Settings ----------
INPUT_FILE = "Omega_Reference.xlsx"
OUTPUT_FILE = "Omega_Reference_with_images.xlsx"
SHEET_NAME = None
NAME_HEADER = "Reference"
PICTURE_HEADER = "Picture"
CANVAS_SIZE = (180, 180)      # Uniform square cell canvas size in pixels
MAX_WATCH_SIZE = (160, 160)   # Max width/height for the watch inside the cell
CACHE_DIR = Path("picture_cache")
DELAY_BETWEEN_WATCHES = 2.0   # Politeness delay to prevent Google from showing bot verification
PREVIEW_WAIT_SECONDS = 6      # Maximum wait for the enlarged preview to load
# ------------------------------

CACHE_DIR.mkdir(exist_ok=True)


def safe_name(text: str) -> str:
    return "".join(c if c.isalnum() else "_" for c in text)[:80]


def clean_reference(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def digits_only(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def remove_checkerboard_background(img: PILImage.Image) -> PILImage.Image:
    """
    Flood-fills from the corners/edges replacing any near-white or light-grey
    checkerboard pixels (R,G,B >= 215 with neutral grey saturation) with pure white (255, 255, 255).
    """
    img = img.convert("RGB")
    width, height = img.size
    pixels = img.load()

    def is_bg_pixel(rgb):
        r, g, b = rgb
        if r >= 215 and g >= 215 and b >= 215:
            # Must be neutral (grey or white, not a saturated watch element)
            if max(abs(r - g), abs(g - b), abs(r - b)) <= 12:
                return True
        return False

    corners = [(0, 0), (width - 1, 0), (0, height - 1), (width - 1, height - 1)]
    if not any(is_bg_pixel(pixels[x, y]) for x, y in corners):
        return img

    visited = bytearray(width * height)
    queue = deque()

    for x in range(width):
        for y in (0, height - 1):
            idx = y * width + x
            if not visited[idx] and is_bg_pixel(pixels[x, y]):
                visited[idx] = 1
                queue.append((x, y))

    for y in range(height):
        for x in (0, width - 1):
            idx = y * width + x
            if not visited[idx] and is_bg_pixel(pixels[x, y]):
                visited[idx] = 1
                queue.append((x, y))

    while queue:
        cx, cy = queue.popleft()
        pixels[cx, cy] = (255, 255, 255)

        for nx, ny in ((cx + 1, cy), (cx - 1, cy), (cx, cy + 1), (cx, cy - 1)):
            if 0 <= nx < width and 0 <= ny < height:
                nidx = ny * width + nx
                if not visited[nidx]:
                    visited[nidx] = 1
                    if is_bg_pixel(pixels[nx, ny]):
                        queue.append((nx, ny))

    return img


def process_high_res_image(pil_img: PILImage.Image) -> PILImage.Image:
    """Cleans background, ensures pure solid white behind the watch, and centers it in the canvas."""
    if pil_img.mode in ("RGBA", "LA", "P"):
        pil_img = pil_img.convert("RGBA")
        bg = PILImage.new("RGBA", pil_img.size, (255, 255, 255, 255))
        bg.alpha_composite(pil_img)
        pil_img = bg.convert("RGB")
    else:
        pil_img = pil_img.convert("RGB")

    pil_img.thumbnail(MAX_WATCH_SIZE, PILImage.Resampling.LANCZOS)
    pil_img = remove_checkerboard_background(pil_img)

    canvas = PILImage.new("RGB", CANVAS_SIZE, (255, 255, 255))

    paste_x = (CANVAS_SIZE[0] - pil_img.width) // 2
    paste_y = (CANVAS_SIZE[1] - pil_img.height) // 2

    canvas.paste(pil_img, (paste_x, paste_y))
    return canvas


async def get_watch_image_via_google_cdp(page, reference: str, save_path: Path) -> bool:
    """
    Clicks the absolute first organic search result and extracts the enlarged preview
    image from the top of the right-side display box, ensuring a pure white background.
    """
    try:
        raw_digits = digits_only(reference)
        query = f"OMEGA watch {raw_digits}" if raw_digits else f"OMEGA watch {reference}"
        search_url = f"https://www.google.com/search?q={quote(query)}&tbm=isch"

        await page.goto(search_url, wait_until="domcontentloaded", timeout=25000)

        # Detect and handle Google "unusual traffic" / CAPTCHA challenge
        if "/sorry/" in page.url or "google.com/sorry" in page.url:
            print("\n" + "=" * 70)
            print("[ACTION REQUIRED] Google verification / 'I am not a robot' challenge detected!")
            print("Please switch to your open Microsoft Edge window and complete the verification.")
            print("The script is paused and will resume automatically once you solve it...")
            print("=" * 70)
            while "/sorry/" in page.url or "google.com/sorry" in page.url:
                await page.wait_for_timeout(1500)
            print("\n[OK] Verification cleared! Resuming search...")
            await page.wait_for_timeout(2000)
            await page.goto(search_url, wait_until="domcontentloaded", timeout=25000)

        # Step 1: Find the first organic result thumbnail directly inside the browser V8 engine (fast & precise)
        js_find_thumbnail = """() => {
            const old = document.querySelector('[data-target-first-thumbnail="true"]');
            if (old) old.removeAttribute('data-target-first-thumbnail');

            const imgs = Array.from(document.querySelectorAll('img')).slice(0, 150);
            const boxes = [];
            for (let i = 0; i < imgs.length; i++) {
                const el = imgs[i];
                const r = el.getBoundingClientRect();
                // Skip icons, filter chips (< 60px) and large banners (> 500px)
                if (r.width < 60 || r.height < 60 || r.width > 500 || r.height > 500) continue;
                const style = window.getComputedStyle(el);
                if (style.visibility === 'hidden' || style.display === 'none' || style.opacity === '0') continue;
                const src = el.currentSrc || el.src || '';
                if (!src || src.includes('gstatic.com/images/branding')) continue;
                boxes.push({ index: i, box: { x: r.left, y: r.top, width: r.width, height: r.height } });
            }

            const clusters = {};
            for (const cand of boxes) {
                const key = `${Math.round(cand.box.width / 10)}_${Math.round(cand.box.height / 10)}`;
                if (!clusters[key]) clusters[key] = [];
                clusters[key].push(cand);
            }

            let best = [];
            for (const key in clusters) {
                if (clusters[key].length > best.length) {
                    best = clusters[key];
                }
            }

            if (best.length >= 5) {
                // Top row first, then left-to-right within that row
                best.sort((a, b) => {
                    const rowA = Math.round(a.box.y / 20);
                    const rowB = Math.round(b.box.y / 20);
                    if (rowA !== rowB) return rowA - rowB;
                    return a.box.x - b.box.x;
                });
                const targetIndex = best[0].index;
                imgs[targetIndex].setAttribute('data-target-first-thumbnail', 'true');
                return { index: targetIndex, src: imgs[targetIndex].currentSrc || imgs[targetIndex].src || '' };
            }
            return null;
        }"""

        target_info = None
        search_deadline = time.time() + 6
        while time.time() < search_deadline:
            target_info = await page.evaluate(js_find_thumbnail)
            if target_info:
                break
            await page.wait_for_timeout(100)

        if not target_info:
            raise RuntimeError(
                f"Could not confirm the first result thumbnail for {reference} "
                f"(no group of 5+ same-sized search thumbnails found in time)"
            )

        clicked_src = target_info.get("src", "")
        print(f"  clicked thumbnail: {clicked_src[:60]}")

        target = page.locator('img[data-target-first-thumbnail="true"]').first
        try:
            await target.click(timeout=5000)
        except Exception as e:
            raise RuntimeError(f"Could not click the first result thumbnail for {reference}: {e}")

        # Step 2: Extract the enlarged preview image from the right-side display box.
        # The main preview image is strictly at the top of the right panel (top < 500),
        # distinguishing it from the 'related images' / 'visually similar' grid below it.
        image_src = None
        best_fallback_src = None
        deadline = time.time() + PREVIEW_WAIT_SECONDS

        while time.time() < deadline:
            try:
                preview_info = await page.evaluate(
                    """() => {
                        // Priority 1: Check Google Images preview container selectors
                        const containerSelectors = [
                            'div[jsname="figiqf"] img.uT7N3d',
                            'a.YsLeY img.uT7N3d',
                            'div[jsname="figiqf"] img.sFlh5c',
                            'a.YsLeY img.sFlh5c',
                            'div[jsname="figiqf"] img',
                            'a.YsLeY img',
                            'img.uT7N3d',
                            'img.sFlh5c'
                        ];

                        for (const sel of containerSelectors) {
                            const imgs = Array.from(document.querySelectorAll(sel));
                            for (const img of imgs) {
                                const r = img.getBoundingClientRect();
                                const src = img.currentSrc || img.src || '';
                                if (!src || !src.startsWith('http') || src.includes('gstatic.com/images/branding') || src.includes('logo')) {
                                    continue;
                                }
                                // Main preview is in right panel (left > 300) and strictly above related items (top < 500)
                                if (r.top < 500 && r.left > 300 && r.width > 30 && r.height > 30) {
                                    const isHighRes = !src.includes('encrypted-tbn0.gstatic.com');
                                    return {
                                        src: src,
                                        isHighRes: isHighRes,
                                        w: img.naturalWidth || 0,
                                        h: img.naturalHeight || 0,
                                        top: r.top
                                    };
                                }
                            }
                        }

                        // Priority 2: Generic fallback across all images in the right-side preview area
                        const candidates = Array.from(document.querySelectorAll('img'))
                            .map(img => {
                                const r = img.getBoundingClientRect();
                                return {
                                    src: img.currentSrc || img.src || '',
                                    top: r.top,
                                    left: r.left,
                                    w: img.naturalWidth || 0,
                                    h: img.naturalHeight || 0,
                                    visible: r.width > 30 && r.height > 30
                                };
                            })
                            .filter(d => d.src && d.src.startsWith('http')
                                     && !d.src.includes('gstatic.com/images/branding')
                                     && !d.src.includes('logo')
                                     && d.visible
                                     && d.left > 350
                                     && d.top < 500
                                     && (d.w >= 100 || d.h >= 100));

                        if (candidates.length > 0) {
                            // Sort by topmost first (preview image is at the top), then by resolution
                            candidates.sort((a, b) => a.top - b.top || (b.w * b.h) - (a.w * a.h));
                            const topCand = candidates[0];
                            const isHighRes = !topCand.src.includes('encrypted-tbn0.gstatic.com');
                            return {
                                src: topCand.src,
                                isHighRes: isHighRes,
                                w: topCand.w,
                                h: topCand.h,
                                top: topCand.top
                            };
                        }

                        return null;
                    }"""
                )
            except Exception:
                await page.wait_for_timeout(100)
                continue

            if preview_info and preview_info.get("src"):
                best_fallback_src = preview_info["src"]
                # If high-res original loaded, or if deadline is approaching
                if preview_info.get("isHighRes") or (time.time() > deadline - 3.5):
                    image_src = preview_info["src"]
                    break

            await page.wait_for_timeout(100)

        if not image_src and best_fallback_src:
            image_src = best_fallback_src

        if not image_src:
            raise RuntimeError(
                f"The enlarged preview for {reference} did not load in time "
                f"(no valid preview image found in the right-side display box)."
            )

        print(f"  picked preview: {image_src[:80]}")

        # Neutralize Google's CSS checkerboard background in the preview panel to pure solid white
        try:
            await page.evaluate(
                """() => {
                    const targets = document.querySelectorAll(
                        'img.sFlh5c, img.uT7N3d, div[jsname="figiqf"] img, a.YsLeY, a.YsLeY img, div[jsname="figiqf"], div.p7sI2, div.v6bUne, div.lZvVKc'
                    );
                    for (const el of targets) {
                        el.style.setProperty('background-image', 'none', 'important');
                        el.style.setProperty('background', '#ffffff', 'important');
                        el.style.setProperty('background-color', '#ffffff', 'important');
                    }
                }"""
            )
        except Exception:
            pass

        # Download image bytes using browser context with Referer header
        headers = {
            "Referer": "https://www.google.com/",
        }
        data = None
        try:
            resp = await page.request.get(image_src, headers=headers, timeout=10000)
            if resp.ok:
                data = await resp.body()
        except Exception:
            data = None

        if not data:
            # Fallback: screenshot the preview element directly (now guaranteed on solid pure white)
            preview_loc = page.locator('div[jsname="figiqf"] img, a.YsLeY img, img.sFlh5c').first
            if await preview_loc.is_visible():
                data = await preview_loc.screenshot()
            else:
                raise RuntimeError(f"Could not download or capture preview image for {reference}")

        raw_img = PILImage.open(BytesIO(data))
        if raw_img.width < 50 or raw_img.height < 50:
            raise RuntimeError("Downloaded image dimensions too small.")

        final_img = process_high_res_image(raw_img)
        final_img.save(save_path, "PNG")
        return True

    except Exception as e:
        print(f" -> Error: {e}")
        return False


def get_image_row_col(img):
    """Returns (1-based row, 1-based col) for an openpyxl image anchor, or None."""
    anc = getattr(img, "anchor", None)
    if anc is None:
        return None
    if isinstance(anc, str):
        try:
            return coordinate_to_tuple(anc)
        except Exception:
            return None
    if hasattr(anc, "_from") and anc._from is not None:
        try:
            return (anc._from.row + 1, anc._from.col + 1)
        except Exception:
            pass
    if hasattr(anc, "row") and hasattr(anc, "column"):
        try:
            return (anc.row, anc.column)
        except Exception:
            pass
    return None


def get_existing_image_rows(ws, pic_col: int) -> set:
    """Returns a set of 1-based row numbers that already contain an anchored image in pic_col."""
    rows = set()
    for img in getattr(ws, "_images", []):
        rc = get_image_row_col(img)
        if rc and rc[1] == pic_col:
            rows.add(rc[0])
    return rows


def is_picture_tab_filled(ws, row: int, pic_col: int, existing_img_rows: set) -> bool:
    """Returns True if the cell already contains an anchored image or non-empty cell content."""
    if row in existing_img_rows:
        return True
    val = ws.cell(row=row, column=pic_col).value
    if val is not None and str(val).strip() != "":
        return True
    return False


def safe_save_workbook(wb, output_path: str) -> bool:
    """Safely saves the workbook, creating a fallback copy if the target file is open/locked."""
    path = Path(output_path)
    try:
        wb.save(str(path))
        return True
    except PermissionError:
        fallback = path.with_name(f"{path.stem}_autosave{path.suffix}")
        try:
            wb.save(str(fallback))
            print(f"\n[Warning] '{path.name}' is currently open/locked. Saved progress to '{fallback.name}' instead.")
            return True
        except Exception as e:
            print(f"\n[Error] Failed to save fallback workbook '{fallback.name}': {e}")
            return False
    except Exception as e:
        print(f"\n[Error] Failed to save workbook: {e}")
        return False


def ensure_edge_browser_running() -> bool:
    """Checks if Edge is running on port 9222; if not, automatically launches it."""
    import urllib.request
    try:
        with urllib.request.urlopen("http://localhost:9222/json/version", timeout=1):
            return True
    except Exception:
        pass

    print("Edge debug session on port 9222 not detected. Auto-launching Microsoft Edge...")
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    edge_exe = next((p for p in candidates if Path(p).exists()), "msedge")
    import subprocess
    try:
        subprocess.Popen([edge_exe, "--remote-debugging-port=9222", r"--user-data-dir=C:\EdgeDebug"])
    except Exception as e:
        print(f"Warning: could not auto-launch Edge ({e}). Please start it manually.")
        return False

    for _ in range(20):
        time.sleep(0.5)
        try:
            with urllib.request.urlopen("http://localhost:9222/json/version", timeout=1):
                print("Edge launched and CDP connection is ready.")
                return True
        except Exception:
            pass

    return False


async def main_async():
    t0 = time.time()
    input_file = INPUT_FILE if Path(INPUT_FILE).exists() else ("watches.xlsx" if Path("watches.xlsx").exists() else INPUT_FILE)
    if not Path(input_file).exists():
        raise SystemExit(f"Input file '{input_file}' not found.")

    output_file = OUTPUT_FILE if input_file == INPUT_FILE else "watches_with_images.xlsx"

    wb = load_workbook(input_file)
    ws = wb[SHEET_NAME] if SHEET_NAME else wb.active

    headers = {str(c.value).strip().lower(): c.column for c in ws[1] if c.value is not None}
    if NAME_HEADER.lower() not in headers:
        raise SystemExit(f'Header "{NAME_HEADER}" not found in row 1.')

    ref_col = headers[NAME_HEADER.lower()]
    pic_col = (
        headers.get(PICTURE_HEADER.lower())
        or headers.get("picture")
        or headers.get("pictures")
        or (ref_col - 1 if ref_col > 1 else ref_col + 1)
    )

    if ws.cell(row=1, column=pic_col).value is None:
        ws.cell(row=1, column=pic_col).value = PICTURE_HEADER

    col_letter = get_column_letter(pic_col)
    ws.column_dimensions[col_letter].width = 28

    # Preserve any images previously completed in output_file (if resuming from a previous run)
    if Path(output_file).exists() and Path(output_file).resolve() != Path(input_file).resolve():
        try:
            prev_wb = load_workbook(output_file)
            prev_ws = prev_wb[SHEET_NAME] if SHEET_NAME and SHEET_NAME in prev_wb.sheetnames else prev_wb.active
            preserved = 0
            for img in getattr(prev_ws, "_images", []):
                rc = get_image_row_col(img)
                if rc and rc[1] == pic_col:
                    r = rc[0]
                    ws.add_image(img, f"{col_letter}{r}")
                    ws.row_dimensions[r].height = 135
                    preserved += 1
            if preserved > 0:
                print(f"Preserved {preserved} previously completed images from '{output_file}'.")
        except Exception as e:
            print(f"Notice: could not load previous output file '{output_file}': {e}")

    existing_img_rows = get_existing_image_rows(ws, pic_col)

    rows, originals = [], {}
    skipped_count = 0
    for row in range(2, ws.max_row + 1):
        ref = clean_reference(ws.cell(row=row, column=ref_col).value)
        if not ref or ref.startswith("="):
            continue

        # If any picture tab already has an image inside it, skip this entry and do not change it
        if is_picture_tab_filled(ws, row, pic_col, existing_img_rows):
            skipped_count += 1
            continue

        key = digits_only(ref) or safe_name(ref)
        rows.append((row, ref, key))
        originals.setdefault(key, ref)

    if skipped_count > 0:
        print(f"Skipping {skipped_count} entries that already have an image in the Picture tab.")

    files = {k: CACHE_DIR / f"{safe_name(k)}_0.png" for k in originals}

    # Step A: Place any images that already exist in local picture_cache
    cached_placed = 0
    remaining_rows = []
    for row, ref, key in rows:
        img_path = files[key]
        if img_path.exists():
            try:
                xl_img = XLImage(str(img_path))
                cell_ref = f"{col_letter}{row}"
                ws.add_image(xl_img, cell_ref)
                ws.row_dimensions[row].height = 135
                existing_img_rows.add(row)
                cached_placed += 1
            except Exception as e:
                print(f"Warning: could not insert cached image for row {row} ({ref}): {e}")
                remaining_rows.append((row, ref, key))
        else:
            remaining_rows.append((row, ref, key))

    if cached_placed > 0:
        print(f"Inserted {cached_placed} images from picture cache into the spreadsheet.")
        safe_save_workbook(wb, output_file)

    # Step B: Identify references that actually need downloading via browser
    todo_keys = []
    for row, ref, key in remaining_rows:
        if not files[key].exists() and key not in todo_keys:
            todo_keys.append(key)

    placed_live = 0
    try:
        if todo_keys:
            ensure_edge_browser_running()
            print(f"Connecting to live browser session for {len(todo_keys)} references needing download...")
            from playwright.async_api import async_playwright

            async with async_playwright() as pw:
                try:
                    browser = await pw.chromium.connect_over_cdp("http://localhost:9222")
                    context = browser.contexts[0]
                    page = context.pages[0] if context.pages else await context.new_page()
                except Exception as e:
                    raise RuntimeError(
                        f"Could not connect to browser via CDP. Make sure Edge/Chrome is running with '--remote-debugging-port=9222'. Error: {e}"
                    )

                try:
                    for idx, key in enumerate(todo_keys, 1):
                        ref = originals[key]
                        print(f"[{idx}/{len(todo_keys)}] Processing {ref}...", end="")
                        success = await get_watch_image_via_google_cdp(page, ref, files[key])
                        if success:
                            print(" OK")
                            # Insert into worksheet for all matching rows
                            for row, r_ref, r_key in remaining_rows:
                                if r_key == key and row not in existing_img_rows:
                                    try:
                                        xl_img = XLImage(str(files[key]))
                                        cell_ref = f"{col_letter}{row}"
                                        ws.add_image(xl_img, cell_ref)
                                        ws.row_dimensions[row].height = 135
                                        existing_img_rows.add(row)
                                        placed_live += 1
                                    except Exception as e:
                                        print(f"\nWarning: could not insert image for {ref} (row {row}): {e}")

                            # Immediately save progress so failure at any point preserves work
                            safe_save_workbook(wb, output_file)
                        else:
                            print(" FAILED")

                        await asyncio.sleep(DELAY_BETWEEN_WATCHES)
                finally:
                    pass
    except KeyboardInterrupt:
        print("\n[Interrupted] Script stopped by user.")
    except Exception as e:
        print(f"\n[Error] Script encountered an issue: {e}")
    finally:
        # Guarantee that the modified excel file is saved no matter what caused termination
        safe_save_workbook(wb, output_file)
        total_images = len(getattr(ws, "_images", []))
        print(f"\nProgress saved to '{output_file}'. Total images now in sheet: {total_images} (finished in {time.time() - t0:.1f}s).")


def main():
    asyncio.run(main_async())


if __name__ == "__main__":
    main()
