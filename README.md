# RefScraper---An-Automated-Image-Fetcher
RefScraper - An automated tool that finds the exact watch image for every Ref ID and adds it to your sheet automatically.

Automates a task that would otherwise be done by hand: for every watch reference number in an Excel sheet, find a picture of that watch on Google Images and insert it into the matching row, so you end up with a spreadsheet of references with pictures sitting right next to them.

It connects to an already-open Chrome browser (via its remote-debugging port) rather than launching its own, searches Google Images for each reference, clicks the first result, grabs the enlarged preview that opens for it, downloads that image, and writes it straight into the Excel file using openpyxl — no copy-pasting, no manual downloading.

What it does, step by step

For every reference in the input spreadsheet:

Builds a Google Images search query from the reference (watch <reference digits>).
Waits for the results grid to load, then identifies the real thumbnail grid by grouping every picture on the page by size — the grid is whichever size group has the most members, which reliably separates it from unrelated UI elements like Google's filter bar.
Clicks the first thumbnail in that grid with a genuine (trusted) click, so the site's own JavaScript opens its inline preview panel normally instead of falling through to a default link navigation.
Waits for the enlarged preview image to appear, identifies it by position (next to the thumbnail that was clicked) and size (noticeably larger than the thumbnail), and downloads it directly — no clipboard, no screenshot.
Places the watch on a plain white background, centers it, and saves it as a small cached PNG.
Writes that PNG into the Picture column of the matching row.

Everything that's already been downloaded is cached, so re-running the script after a partial run or a failure only fetches what's still missing.

## Requirements
Python 3.9+
Google Chrome or Microsoft Edge
### Python packages:
pip install playwright openpyxl pillow
playwright install chromium
