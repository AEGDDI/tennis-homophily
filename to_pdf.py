"""
Convert report.html → report.pdf
Images are inlined as base64 so the PDF is fully self-contained.

Tries (in order):
  1. Microsoft Edge headless  – always installed on Windows 11
  2. Google Chrome headless   – if installed
  3. Playwright               – pip install playwright && playwright install chromium
  4. weasyprint               – pip install weasyprint  (needs GTK on Windows)

Run from the project root:  python to_pdf.py
"""
import base64
import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
HTML = ROOT / "report.html"
PDF  = ROOT / "report.pdf"

# ── Inline local images as base64 data URIs ────────────────────────────────────
def inline_images(html_text: str, base_dir: Path) -> str:
    """Replace <img src="relative/path"> with inline base64 data URIs."""
    def replace(m):
        src = m.group(1)
        if src.startswith("data:") or src.startswith("http"):
            return m.group(0)
        img_path = (base_dir / src).resolve()
        if not img_path.exists():
            return m.group(0)
        ext = img_path.suffix.lstrip(".").lower()
        mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "gif": "gif",
                "svg": "svg+xml", "webp": "webp"}.get(ext, ext)
        data = base64.b64encode(img_path.read_bytes()).decode()
        return f'src="data:image/{mime};base64,{data}"'

    return re.sub(r'src="([^"]+)"', replace, html_text)

def make_self_contained_html() -> Path:
    raw = HTML.read_text(encoding="utf-8")
    inlined = inline_images(raw, ROOT)
    tmp = tempfile.NamedTemporaryFile(suffix=".html", delete=False,
                                      mode="w", encoding="utf-8")
    tmp.write(inlined)
    tmp.close()
    return Path(tmp.name)

# ── 1. Edge / Chrome headless ──────────────────────────────────────────────────
BROWSER_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

def try_browser_headless(tmp_html: Path) -> bool:
    for exe in BROWSER_PATHS:
        if not Path(exe).exists():
            continue
        name = Path(exe).stem
        for flags in (
            ["--headless=new", "--disable-gpu", f"--print-to-pdf={PDF}",
             "--print-to-pdf-no-header", "--no-margins"],
            ["--headless", "--disable-gpu", f"--print-to-pdf={PDF}"],
        ):
            if PDF.exists():
                PDF.unlink()
            subprocess.run([exe] + flags + [str(tmp_html)],
                           capture_output=True, timeout=30)
            if PDF.exists():
                print(f"PDF saved to {PDF}  (via {name} headless)")
                return True
    return False

# ── 2. Playwright ──────────────────────────────────────────────────────────────
def try_playwright(tmp_html: Path) -> bool:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(tmp_html.as_uri(), wait_until="networkidle")
        page.pdf(
            path=str(PDF),
            format="A4",
            margin={"top": "15mm", "right": "15mm",
                    "bottom": "15mm", "left": "15mm"},
            print_background=True,
        )
        browser.close()
    print(f"PDF saved to {PDF}  (via playwright)")
    return True

# ── 3. weasyprint ──────────────────────────────────────────────────────────────
def try_weasyprint(tmp_html: Path) -> bool:
    from weasyprint import HTML as WH
    WH(filename=str(tmp_html)).write_pdf(str(PDF))
    print(f"PDF saved to {PDF}  (via weasyprint)")
    return True

# ── Main ───────────────────────────────────────────────────────────────────────
tmp = make_self_contained_html()
try:
    steps = [
        (try_browser_headless, "Edge/Chrome headless"),
        (try_playwright,       "playwright"),
        (try_weasyprint,       "weasyprint"),
    ]
    for fn, name in steps:
        try:
            if fn(tmp):
                sys.exit(0)
        except ImportError:
            print(f"{name} not installed — skipping")
        except Exception as e:
            print(f"{name} failed: {e}")
finally:
    tmp.unlink(missing_ok=True)

print()
print("Could not generate PDF automatically.")
print("Easiest option: open report.html in Edge or Chrome → Ctrl+P → Save as PDF")
sys.exit(1)
