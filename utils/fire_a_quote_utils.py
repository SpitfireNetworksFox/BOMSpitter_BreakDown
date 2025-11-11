from pathlib import Path
import argparse, base64, json, re
import pandas as pd

def write_pdf_via_playwright(out_html_path: Path, html: str | None = None) -> None:
    from playwright.sync_api import sync_playwright
    pdf_path = out_html_path.with_suffix(".pdf")
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        if html is not None:
            # Load from string; base_url makes relative paths work
            page.set_content(html, wait_until="load", base_url=str(out_html_path.parent.resolve()))
        else:
            # Fallback: load from disk
            page.goto(out_html_path.resolve().as_uri(), wait_until="load")
        page.emulate_media(media="print")
        page.pdf(
            path=str(pdf_path),
            format="Letter",
            margin={"top": "0.3in", "right": "0.3in", "bottom": "0.5in", "left": "0.3in"},
            print_background=True,
        )
        browser.close()

def _safe_filename(s: str) -> str:
      """
      Make a safe filename across Windows/macOS/Linux.
      """
      s = str(s).strip()
      s = re.sub(r'[\\/*?:"<>|]', "_", s)  # illegal chars -> underscore
      s = s.rstrip(" .")                   # no trailing space/dot on Windows
      reserved = {
          "CON","PRN","AUX","NUL",
          "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
          "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"
      }
      if s.upper() in reserved:
          s = f"_{s}"
      return s or "file"


def _b64_png(data: bytes) -> str: 
    return base64.b64encode(data).decode("ascii")

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df

def _pick(df: pd.DataFrame, names: list[str]) -> str|None:
    for n in names:
        if n in df.columns: 
            return n
    for n in names:
        for c in df.columns:
            if n in c: 
                return c
    return None

def _num(x):
    try: 
        return float(str(x).replace(",", "").replace("$",""))
    except Exception: 
        return None

def _prefer_sheet(xls: pd.ExcelFile, candidates: list[str]) -> str:
    lowers = [s.lower() for s in xls.sheet_names]
    for want in candidates:
        if want.lower() in lowers:
            return xls.sheet_names[lowers.index(want.lower())]
    return xls.sheet_names[0]


QUOTE_COUNTER_PATH = Path("utils/quote_number.txt")

def pop_and_increment_quote_number(path: Path = QUOTE_COUNTER_PATH,
                                   prefix: str = "Q-") -> str:
    """
    Reads the current integer from `path`, returns it as 'Q-<number>',
    and writes back (number + 1) atomically so the next run gets a new number.
    Starts at 1 if the file is missing/empty.
    """
    path.parent.mkdir(parents=True, exist_ok=True)

    # Read current number
    try:
        raw = path.read_text(encoding="utf-8").strip()
        # allow file to contain either just digits or a prefixed value like "Q-1234"
        m = re.search(r"(\d+)", raw) if raw else None
        current = int(m.group(1)) if m else 1
    except FileNotFoundError:
        current = 1
    except Exception:
        # Fallback safely if the file is corrupted
        current = 1

    # Write back next value (atomic-ish)
    next_val = current + 1
    tmp = path.with_suffix(".tmp")
    tmp.write_text(str(next_val), encoding="utf-8")
    tmp.replace(path)  # atomic on most OS/filesystems

    return f"{prefix}{current}"