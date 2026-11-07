"""Fleem Proxy Companion - turns MPCFill order XMLs (and proxy PDFs) into print-ready PDFs and prints them.

Everything is built in: it downloads each card image from MPC Autofill and lays the cards out on the
paper size you choose (A4, Letter, Legal, A3, Tabloid, A5; 63 x 88 mm cards, double-faced cards first then A-Z, a back page after every front page with
the double-faced cards' real backs, mirrored so they line up when the sheet is flipped). Printing goes
straight through Windows at exactly 100% size, on any printer, with that printer's own settings.

Right-click a PDF or XML > Print with Fleem (or Open with > Fleem Proxy Companion) to pick which fronts and
backs to print, on which printer - and to edit the cards (copies, images, backs, bleed, spacing).

Optionally (when settings.ini has a discord_token) it watches the order channels of the Fleem Discord
bot, makes each new order's PDF and asks you to load the paper and print it.

Usage:
    pythonw fleem_pdf_maker.py                    open Fleem's window (keep_running = yes: an icon by the clock)
    python  fleem_pdf_maker.py --once             check Discord once, print what happens, then exit
    pythonw fleem_pdf_maker.py --print FILE.pdf   open the print window for a PDF or XML
    python  fleem_pdf_maker.py --version          show which version this is
"""
import configparser
import gc
import json
import logging
import logging.handlers
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

try:  # blocks malicious XML; falls back to the standard library if not installed
    from defusedxml.ElementTree import fromstring as parse_xml
except ImportError:  # pragma: no cover
    from xml.etree.ElementTree import fromstring as parse_xml

HERE = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
SETTINGS_FILE = HERE / "settings.ini"
STATE_FILE = HERE / "processed_orders.json"
WORK_DIR = HERE / "work"
LOG_FILE = HERE / "fleem.log"


def trim_memory() -> None:
    """Hand idle memory back to Windows while Fleem waits. Windows pages back in only what's used
    when it wakes up, so this costs nothing but keeps the background footprint tiny."""
    gc.collect()
    if os.name != "nt":
        return
    try:
        import ctypes
        from ctypes import wintypes
        k32 = ctypes.windll.kernel32
        k32.GetCurrentProcess.restype = wintypes.HANDLE
        k32.SetProcessWorkingSetSize.argtypes = (wintypes.HANDLE, ctypes.c_size_t, ctypes.c_size_t)
        k32.SetProcessWorkingSetSize(k32.GetCurrentProcess(), ctypes.c_size_t(-1), ctypes.c_size_t(-1))
    except Exception as e:  # never let tidying up break anything
        log.debug("Couldn't trim memory: %s", e)

API = "https://discord.com/api/v10"
MPC_IMAGE_URL = "https://cdn.mpcautofill.com/images/google_drive/full/{id}.jpg?dpi=1500&jpgQuality=100"
USER_AGENT = "DiscordBot (https://discord.com, 1.0) FleemPdfMaker"
ORDER_CHANNEL = re.compile(r"^order-(\d+)-")
RETRY_FAILED_AFTER = 30 * 60      # seconds before retrying an order that failed

# The PDF layout: A4 portrait (or another paper, or landscape), standard cards (2.48 x 3.46 in), a bleed
# edge and card spacing of your choice, no guides, a header only if you ask for one, backs on every other page with the columns mirrored so they line up
# when the sheet is flipped (the same layout as Proxy PDF Maker).
PT_PER_MM = 72 / 25.4
PAPER_SIZES = {"A4": (210, 297), "Letter": (215.9, 279.4), "Legal": (215.9, 355.6), "A3": (297, 420),
               "Tabloid": (279.4, 431.8), "A5": (148, 210)}     # (width mm, height mm), portrait
DEFAULT_PAPER = "A4"
PAGE_W, PAGE_H = 210 * PT_PER_MM, 297 * PT_PER_MM      # A4 (the default; a layout can choose another paper)


# Card sizes (width, height in points, without bleed) - the same sizes as Proxy PDF Maker
# (github.com/Malacath-92/Proxy-PDF-Maker, MIT licence). Card images are stretched to fit a size.
CARD_SIZES = {
    "Japanese": (59 * PT_PER_MM, 86 * PT_PER_MM),       # e.g. Yu-Gi-Oh!
    "Miniature": (1.24 * 72, 1.73 * 72),                # half-size Magic cards
    "Oversized": (3.46 * 72, 4.96 * 72),                # oversized Magic cards
    "Poker": (2.5 * 72, 3.5 * 72),
    "Standard": (2.48 * 72, 3.46 * 72),                 # Magic, Pokemon and most other TCGs
}
CARD_SIZE_NOTES = {"Japanese": "59 x 86 mm, e.g. Yu-Gi-Oh!", "Miniature": "1.24 x 1.73 in, half-size Magic",
                   "Oversized": "3.46 x 4.96 in, oversized Magic", "Poker": "2.5 x 3.5 in",
                   "Standard": "2.48 x 3.46 in, Magic, Pokemon and most TCGs"}
DEFAULT_CARD = "Standard"
CARD_SIZE_ALIASES = {"Novelty": "Miniature"}       # old names (saved layouts from before a rename)


def custom_card_mm(value) -> tuple[float, float] | None:
    """(width mm, height mm) of a custom card size written as "63.5x88.9", or None."""
    m = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)\s*", value) if isinstance(value, str) else None
    if not m:
        return None
    w, h = float(m[1]), float(m[2])
    return (w, h) if 10 <= w <= 400 and 10 <= h <= 400 else None


def card_size_label(value) -> str:
    """How a card size shows in the Card box: "Standard", or "Custom (63.5 x 88.9 mm)"."""
    mm = custom_card_mm(value)
    value = CARD_SIZE_ALIASES.get(value, value)
    return f"Custom ({mm[0]:g} x {mm[1]:g} mm)" if mm else (value if value in CARD_SIZES else DEFAULT_CARD)


def card_size_named(width_pt: float, height_pt: float) -> str:
    """The card size value for a card this big: a preset's name, or a custom "WxH" (mm)."""
    for name, (w, h) in CARD_SIZES.items():
        if abs(width_pt - w) < 0.6 and abs(height_pt - h) < 0.6:
            return name
    return f"{round(width_pt / PT_PER_MM, 1):g}x{round(height_pt / PT_PER_MM, 1):g}"


# A layout is (bleed edge mm, card spacing mm[, paper[, orientation[, header[, numbering[, card]]]]]):
# orientation "portrait" or "landscape"; header None (no header), "" (the file's name) or the header's
# text; numbering "all" (every page has its page number) or "selected" (only the pages being printed,
# counted 1..n); card one of CARD_SIZES or a custom "WxH" in mm.
def layout_paper(layout) -> str:
    """The paper size name in a layout."""
    name = layout[2] if layout is not None and len(layout) > 2 else None
    return name if name in PAPER_SIZES else DEFAULT_PAPER


def layout_orientation(layout) -> str:
    """ "portrait" or "landscape"."""
    value = layout[3] if layout is not None and len(layout) > 3 else None
    return "landscape" if value == "landscape" else "portrait"


def layout_header(layout) -> str | None:
    """None: no header on the pages; "": a header with the file's name; else the header's text."""
    value = layout[4] if layout is not None and len(layout) > 4 else None
    return value if isinstance(value, str) else None


def layout_numbering(layout) -> str:
    """ "all" or "selected" (page numbers only on the pages being printed)."""
    value = layout[5] if layout is not None and len(layout) > 5 else None
    return "selected" if value == "selected" else "all"


def layout_card(layout) -> str:
    """The card size in a layout: a CARD_SIZES name or a custom "WxH" (mm)."""
    value = layout[6] if layout is not None and len(layout) > 6 else None
    value = CARD_SIZE_ALIASES.get(value, value)
    return value if value in CARD_SIZES or custom_card_mm(value) else DEFAULT_CARD


# Cut marks (Advanced settings > Cut marks), the layout's 8th part:
#  "none"       no marks
#  "crop"       crop marks in the page margins, lined up with every card edge (paper cutters/guillotines)
#  "silhouette" Silhouette Print & Cut: registration marks (type 1) on the front pages, the cards kept
#               clear of them, and cut lines (SVG + DXF) for Silhouette Studio
#  "cricut"     Cricut Print Then Cut: the cards laid out inside Design Space's 6.75 x 9.25 in area
#               (no bleed, room between cards) and exported as see-through PNG sheets; Design Space
#               adds its own sensor marks and bleed
CUT_MARKS = {"none": "None", "crop": "Crop marks (paper cutter)", "silhouette": "Silhouette Print & Cut",
             "cricut": "Cricut Print Then Cut"}
SILHOUETTE_MARKS = {"inset": 10.0, "length": 20.0, "thickness": 0.5}   # mm; match Silhouette Studio's
SILHOUETTE_SQUARE_MM = 5.0               # the filled square of a type 1 mark (top left)
CRICUT_AREA_IN = (6.75, 9.25)            # Design Space's biggest Print Then Cut image (portrait)
CRICUT_MIN_GAP_MM = 6.5                  # between cards on a Cricut sheet (room for Design Space's bleed)
CROP_MARK_MM = 5.0                       # crop mark length (shorter if the margin is small)
CARD_CORNER_MM = {"Standard": 2.5, "Japanese": 1.0, "Miniature": 2.5, "Oversized": 5.0, "Poker": 3.0}


def layout_marks(layout) -> str:
    """ "none", "crop", "silhouette" or "cricut"."""
    value = layout[7] if layout is not None and len(layout) > 7 else None
    return value if value in CUT_MARKS else "none"


def set_silhouette_marks(inset=None, length=None, thickness=None) -> None:
    """The registration marks' size in mm (from the settings), kept within sensible limits."""
    for key, value, lo, hi in (("inset", inset, 3, 30), ("length", length, 5, 40), ("thickness", thickness, 0.1, 2)):
        try:
            if value not in (None, ""):
                SILHOUETTE_MARKS[key] = min(hi, max(lo, float(value)))
        except (TypeError, ValueError):
            pass


def corner_radius(layout) -> float:
    """The cards' corner radius in points (for cut lines and Cricut sheets)."""
    return CARD_CORNER_MM.get(layout_card(layout), 2.5) * PT_PER_MM


def machine_layout(layout) -> tuple:
    """A layout fixed up for its cutting machine: Cricut sheets have no bleed (Design Space adds its own)
    and room between the cards."""
    layout = full_layout(layout)
    if layout_marks(layout) == "cricut":
        layout = (0.0, max(layout[1], CRICUT_MIN_GAP_MM)) + layout[2:]
    return layout


def print_area(layout) -> tuple[float, float, float, float]:
    """(left, bottom, right, top) in points: where the cards may go on the page."""
    PAGE_W, PAGE_H = page_size(layout)
    marks = layout_marks(layout)
    if marks == "silhouette":            # keep clear of the registration marks (and the strip they're in)
        mm = SILHOUETTE_MARKS
        near = (mm["inset"] + SILHOUETTE_SQUARE_MM + 2) * PT_PER_MM          # left/top: the square
        far = (mm["inset"] + mm["thickness"] + 2) * PT_PER_MM                # right/bottom: the L marks
        return near, far, PAGE_W - far, PAGE_H - near
    if marks == "cricut":                # Design Space's Print Then Cut area, in the middle of the page
        w, h = (v * 72 for v in CRICUT_AREA_IN)
        if layout_orientation(layout) == "landscape":
            w, h = h, w
        w, h = min(w, PAGE_W), min(h, PAGE_H)
        return (PAGE_W - w) / 2, (PAGE_H - h) / 2, (PAGE_W + w) / 2, (PAGE_H + h) / 2
    return 0.0, 0.0, PAGE_W, PAGE_H


def card_dims(layout=None) -> tuple[float, float]:
    """(width, height) in points of the layout's cards, without bleed."""
    value = layout_card(layout)
    mm = custom_card_mm(value)
    return (mm[0] * PT_PER_MM, mm[1] * PT_PER_MM) if mm else CARD_SIZES[value]


def card_geometry(layout=None) -> tuple[float, float, float, float]:
    """(card width, card height, image bleed x, image bleed y) in points. Card images are made for
    Standard cards with 0.12 in of bleed; for other sizes the whole image is stretched, bleed and all."""
    cw, ch = card_dims(layout)
    return cw, ch, IMAGE_BLEED * cw / CARD_W, IMAGE_BLEED * ch / CARD_H


def full_layout(layout) -> tuple:
    """Any layout as all eight parts (missing ones take their defaults)."""
    return (layout[0], layout[1], layout_paper(layout), layout_orientation(layout), layout_header(layout),
            layout_numbering(layout), layout_card(layout), layout_marks(layout))


def page_size(layout=None) -> tuple[float, float]:
    """(width, height) in points of the paper a layout uses, turned for landscape."""
    w, h = PAPER_SIZES[layout_paper(layout)]
    if layout_orientation(layout) == "landscape":
        w, h = h, w
    return w * PT_PER_MM, h * PT_PER_MM


def paper_named(width_pt: float, height_pt: float) -> str | None:
    """The paper size a page is (portrait or landscape), or None if it's none of PAPER_SIZES."""
    for name, (w, h) in PAPER_SIZES.items():
        W, H = w * PT_PER_MM, h * PT_PER_MM
        if (abs(width_pt - W) < 2 and abs(height_pt - H) < 2) or (abs(width_pt - H) < 2 and abs(height_pt - W) < 2):
            return name
    return None
CARD_W, CARD_H = 2.48 * 72, 3.46 * 72                    # Standard card, without bleed
IMAGE_BLEED = 0.12 * 72                                  # MPC images carry 0.12 in of bleed on each side
MAX_DPI = 1200                                           # images are capped at this (like Proxy PDF Maker)
JPEG_QUALITY = 95

APP_NAME = "Fleem Proxy Companion"
LEGACY_NAMES = ("Fleem PDF Maker",)      # what Fleem was called before (its .exe, startup entry, Open with)
RENDER_LOCK = threading.Lock()    # one PDF at a time (automatic orders and "Make PDF from XML…")
PDFIUM_LOCK = threading.RLock()   # pdfium isn't thread-safe: page previews (background) and printing take turns
log = logging.getLogger("fleem")


def _no_events(*_args) -> None:
    pass


# The tray icon replaces these to show progress; they do nothing when running without it.
on_status = _no_events      # on_status(text)
on_pdf_ready = _no_events   # on_pdf_ready(order_no, [pdf paths])
on_order_ready = None       # on_order_ready(order_no, [pdf paths]): the tray asks to print; else the PDF opens


def on_warning(text: str) -> None:   # the tray shows these as a notification too
    log.warning(text)


# ============================================================ settings
DEFAULT_LAYOUT = (1.0, 0.0, "A4", "portrait", None, "all", "Standard", "none")  # until you choose others (see layout_paper)
MAX_BLEED_MM = 3.0                # MPC images only have ~3 mm of bleed to show
MAX_SPACING_MM = 10.0


DEFAULT_SETTINGS_TEXT = """[helper]
# Where finished PDFs are saved.
output_folder = {output}

# Open each PDF made with "Make PDF from XML..." when it's done (yes/no).
open_pdfs = yes

# Which way to turn a portrait page when the printer's paper is set up sideways (landscape):
# counterclockwise or clockwise.
rotate_pages = counterclockwise

# Where Fleem keeps downloaded card images and temporary files (it makes a "Fleem cache" folder
# there, emptied whenever Fleem starts and quits). Empty: the work folder next to Fleem.
cache_folder =

# Card order: mdfc_first (MDFC then Alphabetical: double-faced cards first, so their backs fill as few
# back pages as possible) or alphabetical (Only Alphabetical). Advanced settings > Card order.
card_sort = mdfc_first

# Silhouette Print & Cut registration marks, in mm (Advanced settings > Cut marks > Silhouette). Set the
# same values in Silhouette Studio (Page Setup > Registration Marks: Type 1, Length, Thickness, Inset).
silhouette_mark_inset = 10
silhouette_mark_length = 20
silhouette_mark_thickness = 0.5

# Where Export as PDF saves (set with Advanced settings > Set PDF Export Location). Empty: Fleem asks
# where each time.
export_folder =

# The font for the writing on the RPG themes' scrolls (buttons, the banner). Empty: Pixeloid Sans, which
# comes with Fleem. Another font has to be installed, or its .ttf put in a "fonts" folder next to Fleem.
rpg_button_font =

# Keep Fleem running in the background after its window is closed (yes/no). Yes: its icon stays by the
# clock, Print with Fleem opens instantly and Discord orders are caught. No: Fleem only runs while one
# of its windows is open.
keep_running = no

# Check GitHub once a day for a new version of Fleem Proxy Companion and offer to install it (yes/no).
check_for_updates = yes

# Optional: watch the order channels of the Fleem Discord bot and turn new orders into PDFs.
# Fill in the bot's token and your server's ID (leave the token empty to turn this off).
discord_token =
guild_id =
check_every_minutes = 120
skip_orders_before = 0
"""


# ============================================================ add-on hooks
# Your own tools that aren't part of the public Fleem can live in an add-on: fleem_personal.py next to
# this file (it's in .gitignore, so it's never published). If it's there, Fleem loads it at the end of
# this file and calls its setup(core), which can replace any of these.
EDITION = "public"                 # "personal" with the add-on
KEEP_RUNNING_DEFAULT = "no"        # keep_running when settings.ini doesn't say
UNNAMED_LAYOUT_NOTE = None         # with layout_from_name: why last time's values are used ("The file
                                   # name doesn't say the material")
BLEED_TIP = "Extra picture around each card that's cut off, in mm (0 - 3)."
SPACING_TIP = "Gap between cards, in mm: 0 to cut whole sheets, ~1 to cut cards one by one."
LAYOUT_HINT = "Spacing is the gap between cards: 0 for cutting whole sheets, ~1 mm for cutting cards one by one."


def profile_of(name) -> str | None:
    """What kind of print a file is for, from its name. Printer settings can be kept per kind (see
    saved_settings_file). None: no kinds (the public build)."""
    return None


def profile_label(profile) -> str | None:
    """How the print window describes a kind's saved settings ("Classic PDFs"); None: per printer only."""
    return None


def layout_from_name(name, settings) -> tuple | None:
    """The (bleed mm, spacing mm, paper) a file's name calls for, or None for the usual rules."""
    return None


def order_layout_note(name, settings) -> str:
    """Added after an order's layout in the I'm ready window."""
    return ""


def loading_words(profile) -> str:
    """Added after "load N sheets" in the I'm ready window (" of card stock")."""
    return ""


class Settings:
    def __init__(self, path: Path, check: bool = True):
        if not path.exists():      # first run: start with the defaults
            output = Path.home() / "Documents" / "Fleem PDFs"
            try:
                path.write_text(DEFAULT_SETTINGS_TEXT.format(output=output), encoding="utf-8")
            except OSError:
                pass
        cp = configparser.ConfigParser(interpolation=None)
        cp.read(path, encoding="utf-8")
        s = cp["helper"] if cp.has_section("helper") else {}

        def get(key, default=""):
            return (s.get(key, default) or default).strip().strip('"')

        self.token = get("discord_token")
        self.guild_id = get("guild_id")
        self.output_folder = Path(get("output_folder", str(Path.home() / "Documents" / "Fleem PDFs")))
        self.check_every_minutes = max(1, int(get("check_every_minutes", "120") or 120))
        self.skip_orders_before = int(get("skip_orders_before", "0") or 0)
        self.open_pdfs = get("open_pdfs", "yes").lower() not in ("no", "false", "0", "off")
        self.check_for_updates = get("check_for_updates", "yes").lower() not in ("no", "false", "0", "off")
        self.keep_running = get("keep_running", KEEP_RUNNING_DEFAULT).lower() in ("yes", "true", "1", "on")
        self.cache_folder = get("cache_folder", "")
        self.export_folder = get("export_folder", "")
        self.card_sort = set_card_sort(get("card_sort", "mdfc_first"))
        set_silhouette_marks(get("silhouette_mark_inset", ""), get("silhouette_mark_length", ""),
                             get("silhouette_mark_thickness", ""))
        if get("rpg_button_font", ""):
            global RPG_BUTTON_FONT
            RPG_BUTTON_FONT = get("rpg_button_font", "")
        if self.cache_folder:                      # Advanced settings > Cache folder
            use_cache_folder(Path(os.path.expandvars(self.cache_folder)).expanduser() / CACHE_NAME)
        self.raw = {k: (v or "").strip().strip('"') for k, v in s.items()}    # every value, for an add-on
        # Which way to turn a portrait page when the printer's paper is set up sideways (landscape).
        self.rotate_pages = "clockwise" if get("rotate_pages").lower().startswith("clock") else "counterclockwise"

        if not check:
            return
        if self.token and not self.guild_id.isdigit():
            raise SystemExit(f"Fix {SETTINGS_FILE.name}: guild_id should be your server's ID (a number), "
                             "or empty discord_token to turn Discord watching off.")

    @property
    def watch_discord(self) -> bool:
        """Discord order watching is on only when a bot token is filled in."""
        return bool(self.token)


# ============================================================ Discord (read-only)
class Discord:
    def __init__(self, token: str):
        self.token = token

    def get(self, path: str):
        req = urllib.request.Request(API + path, headers={"Authorization": f"Bot {self.token}",
                                                         "User-Agent": USER_AGENT})
        for _ in range(5):
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except urllib.error.HTTPError as e:
                if e.code == 429:   # rate limited: wait as long as Discord asks
                    try:
                        wait = float(json.loads(e.read().decode("utf-8")).get("retry_after", 2))
                    except Exception:
                        wait = 2.0
                    time.sleep(wait + 0.25)
                    continue
                if e.code == 401:
                    raise SystemExit(f"Discord rejected the token in {SETTINGS_FILE.name} — copy DISCORD_TOKEN "
                                     "from Fleem's .env.")
                raise
        raise RuntimeError(f"Discord kept rate-limiting {path}")

    def order_channels(self, guild_id: str):
        """[(order_number, channel)] for Fleem's private order channels, oldest order first."""
        out = []
        for ch in self.get(f"/guilds/{guild_id}/channels"):
            m = ORDER_CHANNEL.match(ch.get("name", ""))
            if ch.get("type") == 0 and m:
                out.append((int(m.group(1)), ch))
        return sorted(out, key=lambda x: x[0])

    def channel_gone(self, channel_id: str) -> bool:
        """True only when Discord confirms the channel no longer exists (deleted order)."""
        try:
            self.get(f"/channels/{channel_id}")
            return False
        except urllib.error.HTTPError as e:
            return e.code == 404
        except Exception:
            return False   # network trouble: assume it still exists

    def order_post(self, channel_id: str, bot_id: str):
        """(xml attachment, status text) from the order bot's posts in this channel, or (None, None).
        The status comes from the order summary (the bot's first post with the XML); the XML is the
        newest one the bot posted - when staff change an order's material the bot posts the XML again
        under its new name ("<user> - <new material>.xml"), and Fleem makes the PDF again from it."""
        # the channel's first 50 messages (the summary) and its last 50 (a newer XML), in case it's busy
        seen = {m["id"]: m for m in self.get(f"/channels/{channel_id}/messages?after=0&limit=50")}
        if len(seen) >= 50:
            seen.update({m["id"]: m for m in self.get(f"/channels/{channel_id}/messages?limit=50")})
        posts = [m for m in seen.values()
                 if m.get("author", {}).get("id") == bot_id
                 and any(a.get("filename", "").lower().endswith(".xml") for a in m.get("attachments", []))]
        if not posts:
            return None, None
        first = min(posts, key=lambda m: int(m["id"]))
        newest = max(posts, key=lambda m: int(m["id"]))
        att = next(a for a in newest["attachments"] if a["filename"].lower().endswith(".xml"))
        status = ""
        for embed in first.get("embeds", []):
            for field in embed.get("fields", []):
                if field.get("name", "").strip().lower() == "status":
                    status = field.get("value", "")
        return att, status


def download(url: str, headers: dict | None = None, tries: int = 3) -> tuple[bytes, str]:
    """(bytes, content type)."""
    last = None
    for attempt in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "*/*", **(headers or {})})
            with urllib.request.urlopen(req, timeout=120) as resp:
                return resp.read(), resp.headers.get("Content-Type", "")
        except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
            last = e
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"Download failed after {tries} tries: {url} ({last})")


# ============================================================ order XML -> cards
def _local(tag) -> str:
    return str(tag).rsplit("}", 1)[-1].lower()


def _child_text(el, name: str) -> str:
    for c in el:
        if _local(c.tag) == name:
            return (c.text or "").strip()
    return ""


def safe_name(name: str) -> str:
    """Card name -> file-name-safe stem (no extension), like the MTG Card Downloader plugin."""
    name = re.sub(r'[/\\:*?"<>|]', "_", name).strip()
    name = re.sub(r"\.(png|jpe?g|webp|gif|bmp)$", "", name, flags=re.I).strip(" .")
    return name or "card"


def parse_order(xml_bytes: bytes):
    """Read an MPCFill order XML.

    Returns a list of card groups: {"front_id", "back_id" (None unless MDFC), "name", "count"}.
    The shared <cardback> is ignored on purpose — every non-MDFC back is printed blank.
    """
    root = parse_xml(xml_bytes)
    order = root if _local(root.tag) == "order" else next((e for e in root.iter() if _local(e.tag) == "order"), None)
    if order is None:
        raise ValueError("Not an MPCFill order XML (no <order>).")

    def section(name):
        return next((c for c in order if _local(c.tag) == name), None)

    def cards_in(sec):
        for card in (sec if sec is not None else []):
            if _local(card.tag) != "card":
                continue
            slots = [int(s) for s in _child_text(card, "slots").split(",") if s.strip().isdigit()]
            query = _child_text(card, "query")
            name = _child_text(card, "name") or query
            yield _child_text(card, "id"), slots, name, query

    fronts = section("fronts")
    if fronts is None:
        raise ValueError("The XML has no <fronts> section.")
    slot_front = {}
    for cid, slots, name, query in cards_in(fronts):
        for s in slots:
            slot_front[s] = (cid, name, query)
    slot_back, back_info = {}, {}
    for cid, slots, name, query in cards_in(section("backs")):
        back_info[cid] = (name, query)
        for s in slots:
            slot_back[s] = cid

    groups: dict[tuple, dict] = {}
    for slot in sorted(slot_front):
        front_id, name, query = slot_front[slot]
        key = (front_id, slot_back.get(slot))
        g = groups.setdefault(key, {"front_id": front_id, "back_id": key[1], "name": name, "query": query,
                                    "count": 0})
        if key[1] in back_info:
            g["back_name"], g["back_query"] = back_info[key[1]]
        g["count"] += 1
    return list(groups.values())


# ============================================================ project + render
EXT_FOR_TYPE = {"image/jpeg": "jpg", "image/jpg": "jpg", "image/png": "png", "image/webp": "webp"}


IMAGE_CACHE_DIR = WORK_DIR / "image_cache"     # card images downloaded since Fleem started


def placeholder_card(path: Path, name: str) -> None:
    """A grey stand-in card (MPC size, with bleed) for an image that couldn't be downloaded."""
    from PIL import Image, ImageDraw, ImageFont
    im = Image.new("RGB", (816, 1110), (200, 200, 200))
    d = ImageDraw.Draw(im)
    d.rectangle([36, 36, 780, 1074], outline=(120, 120, 120), width=6)
    try:
        font, small = ImageFont.load_default(size=44), ImageFont.load_default(size=30)
    except TypeError:                                  # older Pillow
        font = small = ImageFont.load_default()
    d.text((408, 420), "Image couldn't\nbe downloaded", fill=(60, 60, 60), font=font, anchor="mm", align="center")
    words, lines = Path(name).stem.split(), [""]
    for w in words:                                    # wrap the card name
        lines[-1] = (lines[-1] + " " + w).strip() if len(lines[-1]) + len(w) < 26 else lines[-1]
        if lines[-1] and not lines[-1].endswith(w):
            lines.append(w)
    d.text((408, 640), "\n".join(lines[:4]), fill=(40, 40, 40), font=small, anchor="mm", align="center")
    im.save(path, "JPEG", quality=90)


def download_images(groups, images_dir: Path) -> list[dict]:
    """Download fronts and MDFC backs into images_dir. Adds 'front_file'/'back_file' to each group.
    Images already downloaded since Fleem started come from its cache. An image that still won't
    download after retrying becomes a grey placeholder card (the group gets '_missing') instead of
    failing the whole order."""
    images_dir.mkdir(parents=True, exist_ok=True)
    IMAGE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    used = set()
    jobs = {}   # image id -> list of (group, field, stem)
    for g in groups:
        stem = safe_name(g["name"])
        n = 2
        while stem.lower() in used:            # same art name, different card/back: keep files apart
            stem = f"{safe_name(g['name'])} ({n})"
            n += 1
        used.add(stem.lower())
        g["_stem"] = stem
        jobs.setdefault(g["front_id"], []).append((g, "front_file", stem))
        if g["back_id"]:
            jobs.setdefault(g["back_id"], []).append((g, "back_file", "__" + stem))   # "__" = backside image
    failed = {}

    def fetch(image_id):
        cache_key = re.sub(r"[^\w-]", "_", image_id)
        cached = next(IMAGE_CACHE_DIR.glob(cache_key + ".*"), None)
        if cached is not None:
            data, ext = cached.read_bytes(), cached.suffix[1:]
        else:
            try:
                data, ctype = download(MPC_IMAGE_URL.format(id=image_id))
                ctype = ctype.split(";")[0].strip().lower()
                if not ctype.startswith("image/"):
                    raise RuntimeError(f"MPC Autofill didn't return an image (got {ctype or 'nothing'})")
            except Exception as e:
                failed[image_id] = str(e)
                return
            ext = EXT_FOR_TYPE.get(ctype, ctype[6:] or "jpg")
            tmp = IMAGE_CACHE_DIR / f"{cache_key}.{ext}.part"
            tmp.write_bytes(data)
            try:
                tmp.replace(IMAGE_CACHE_DIR / f"{cache_key}.{ext}")
            except OSError:
                pass
        failed.pop(image_id, None)
        # Write straight to disk so downloaded images never pile up in memory.
        users = jobs[image_id]
        g0, _, stem0 = users[0]
        raw = images_dir / f"{stem0}.{ext}"
        raw.write_bytes(data)
        use, suffix = raw, ext
        if needs_bleed(raw):                        # a card image with no (or too little) bleed: add some
            fixed = images_dir / f"{stem0}.sized.jpg"
            try:
                how = prepare_card_image(raw, fixed)
                log.info("%s is a card scan without bleed: %s", g0["name"], how)
                use, suffix = fixed, "sized.jpg"
            except Exception as e:
                log.warning("Couldn't resize %s: %s", g0["name"], e)
        for g, field, stem in users:
            target = images_dir / f"{stem}.{suffix}"
            if target != use:
                shutil.copy(use, target)
            g[field] = target.name

    with ThreadPoolExecutor(max_workers=8) as pool:
        list(pool.map(fetch, list(jobs)))
    if failed:                                  # one more go, one at a time, after a short pause
        time.sleep(5)
        for image_id in list(failed):
            fetch(image_id)
    if failed:                                  # still nothing: another art of the same card on MPCFill
        _use_other_art(failed, jobs, fetch)
    for image_id, why in failed.items():
        for g, field, stem in jobs[image_id]:
            fname = f"{stem}.missing.jpg"
            placeholder_card(images_dir / fname, g["name"] + (" (back)" if field == "back_file" else ""))
            g[field] = fname
            g["_missing"] = True
        log.warning("Card image %s (%s) couldn't be downloaded: %s - using a placeholder",
                    image_id, jobs[image_id][0][0]["name"], why)
    return groups


def missing_cards(groups) -> list[str]:
    """Names of cards whose image couldn't be downloaded (they have placeholders)."""
    return [Path(g["name"]).stem for g in groups if g.get("_missing")]


OTHER_ART_TRIES = 6        # other arts of a card to try when its chosen art won't download


def _use_other_art(failed: dict, jobs: dict, fetch) -> None:
    """For art that won't download: find the same card on MPCFill (fleem_mpcfill.py) and use another
    art of it - the first one that downloads. The groups get "_substituted" notes (see
    substituted_cards) and their new art's id."""
    module = mpcfill_module()
    if module is None:
        return
    api = globals().get("_mpcfill_api")
    if api is None:
        api = globals()["_mpcfill_api"] = module.MPCFillAPI(f"FleemProxyCompanion/{APP_VERSION}")
    known_bad = set(failed)
    for image_id in list(failed):
        users = jobs[image_id]
        g0, field0, _ = users[0]
        front_name = g0.get("query") or re.sub(r"\s*\(.*$", "", Path(g0.get("name") or "").stem)
        try:
            if field0 == "front_file":
                card, kind = front_name, "TOKEN" if str(g0.get("query") or "").lower().startswith("t:") else "CARD"
                card = card[2:] if kind == "TOKEN" else card
            else:                                   # a double-faced card's back: its back face's name
                card, kind = (g0.get("back_query") or api.dfc_back_name(front_name) or ""), "CARD"
            if not card:
                continue
            others = [i for i in api.search([card], kind).get(card, []) if i not in known_bad][:OTHER_ART_TRIES]
        except Exception as e:
            log.info("Couldn't look for another art of %s: %s", front_name, e)
            continue
        for other in others:
            before = jobs.get(other)               # (another card may use this art too)
            jobs[other] = users
            fetch(other)
            ok = other not in failed
            if before is not None:
                jobs[other] = before
            elif not ok:
                del jobs[other]
            if not ok:
                failed.pop(other, None)
                known_bad.add(other)
                continue
            failed.pop(image_id, None)
            for g, field, _ in users:
                g["back_id" if field == "back_file" else "front_id"] = other
                shown = card if field == "back_file" else \
                    (re.sub(r"\s*[\(\[].*$", "", Path(g.get("name") or "").stem) or card)
                g.setdefault("_substituted", []).append({"side": "back" if field == "back_file" else "front",
                                                         "card": shown, "old": image_id, "new": other})
            log.warning("%s: the chosen art (%s) wouldn't download, so another art of it was used (%s)",
                        card, image_id, other)
            break


def substituted_cards(groups) -> list[str]:
    """Cards that got another art of themselves because the chosen one wouldn't download."""
    out = []
    for g in groups:
        for sub in g.get("_substituted", []):
            label = sub["card"] + (" (back)" if sub["side"] == "back" else "")
            if label not in out:
                out.append(label)
    return out


def download_problems(groups) -> str | None:
    """What to tell you after downloading an order's images: other art used, placeholders. None if fine."""
    other, gone = substituted_cards(groups), missing_cards(groups)
    parts = []
    if other:
        parts.append(f"The chosen art for {len(other)} card{'s' * (len(other) != 1)} couldn't be downloaded, so "
                     "Fleem used a different art of the same card:\n" + "\n".join(other[:10])
                     + ("\n…" if len(other) > 10 else ""))
    if gone:
        parts.append(f"{len(gone)} card image{'s' if len(gone) > 1 else ''} couldn't be downloaded at all, so grey "
                     "placeholder cards were used:\n" + "\n".join(gone[:10]) + ("\n…" if len(gone) > 10 else "")
                     + "\n\nReplace them with Edit cards in the print window.")
    return "\n\n".join(parts) or None


# MPCFill art in the open projects: card image (see _src_id) -> {"id", "name", "query"}. Filled when an
# XML's cards are downloaded or a Fleem PDF that remembers them is opened; a card is exported as an
# MPCFill XML only if all its images are in here (replaced or edited images aren't).
MPC_ART = {}


def remember_mpc_art(src, art_id, name="", query="") -> None:
    if src and art_id:
        MPC_ART[_src_id(src)] = {"id": str(art_id), "name": name or "", "query": query or ""}


def mpc_art_of(src):
    return MPC_ART.get(_src_id(src)) if src else None


def can_export_xml(cards: list) -> bool:
    """Is every card (front, and back if it has one) MPCFill art? Then the project can be an XML."""
    return bool(cards) and all(mpc_art_of(c["front"]) and (not c["back"] or mpc_art_of(c["back"])) for c in cards)


def mpcfill_xml(cards: list, stock: str = "(S30) Standard Smooth", foil: bool = False) -> str:
    """The project as an MPCFill order XML (the same form MPCFill saves): every card copy is a slot, in
    print order; cards without a back get MPCFill's project card back."""
    from xml.sax.saxutils import escape
    fronts, backs = {}, {}
    for slot, c in enumerate(cards):
        f = mpc_art_of(c["front"])
        fronts.setdefault(f["id"], [f, []])[1].append(slot)
        if c["back"]:
            b = mpc_art_of(c["back"])
            backs.setdefault(b["id"], [b, []])[1].append(slot)

    def card_xml(art, slots):
        name = art["name"] or "card.png"
        if not re.search(r"\.(png|jpe?g|webp)$", name, re.I):
            name += ".png"
        return ("    <card>\n"
                f"      <id>{escape(art['id'])}</id>\n"
                "      <sourceType>Google Drive</sourceType>\n"
                f"      <slots>{','.join(map(str, slots))}</slots>\n"
                f"      <name>{escape(name)}</name>\n"
                f"      <query>{escape((art['query'] or Path(name).stem).lower())}</query>\n"
                "    </card>\n")
    xml = ("<order>\n  <details>\n"
           f"    <quantity>{len(cards)}</quantity>\n    <stock>{escape(stock)}</stock>\n"
           f"    <foil>{'true' if foil else 'false'}</foil>\n  </details>\n"
           "  <fronts>\n" + "".join(card_xml(a, sl) for a, sl in fronts.values()) + "  </fronts>\n")
    if backs:
        xml += "  <backs>\n" + "".join(card_xml(a, sl) for a, sl in backs.values()) + "  </backs>\n"
    return xml + "  <cardback></cardback>\n</order>\n"


CACHE_NAME = "Fleem cache"       # the folder Fleem makes inside a cache folder you choose


def cache_root(settings=None) -> Path:
    """The folder with Fleem's temporary files and downloaded card images."""
    return WORK_DIR


def use_cache_folder(folder: Path) -> None:
    global WORK_DIR, IMAGE_CACHE_DIR
    WORK_DIR = folder
    IMAGE_CACHE_DIR = WORK_DIR / "image_cache"


def set_cache_folder(settings, chosen: Path | None) -> None:
    """Keep the cache in a "Fleem cache" folder inside `chosen` (None: the work folder next to Fleem).
    Only that folder is ever emptied, never the one you picked."""
    new = Path(chosen) / CACHE_NAME if chosen else HERE / "work"
    if new == WORK_DIR:
        return
    new.mkdir(parents=True, exist_ok=True)
    probe = new / ".fleem-write-test"
    probe.write_text("ok", encoding="utf-8")       # can Fleem write there?
    probe.unlink()
    clean_work_folder()                            # the old cache goes
    if WORK_DIR.name == CACHE_NAME:
        try:
            WORK_DIR.rmdir()
        except OSError:
            pass
    use_cache_folder(new)
    value = str(chosen) if chosen else ""
    try:
        save_setting("cache_folder", value)
    except OSError as e:
        log.warning("Couldn't save cache_folder: %s", e)
    if settings is not None:
        settings.cache_folder = value
    log.info("Cache folder: %s", new)


def clean_work_folder() -> None:
    """Remove everything left in the work folder (temporary files and the image cache). Done when
    Fleem starts and when it quits, so nothing piles up."""
    if not WORK_DIR.exists():
        return
    for item in WORK_DIR.iterdir():
        try:
            shutil.rmtree(item) if item.is_dir() else item.unlink()
        except OSError as e:
            log.info("Couldn't remove %s: %s", item, e)


def sort_key(group) -> tuple:
    """Alphabetical by card name: the XML's <query> (the real card name), else the art's file name."""
    text = (group.get("query") or group.get("name") or "").casefold()
    text = re.sub(r"\.(png|jpe?g|webp)$", "", text)
    text = re.sub(r"^[^0-9a-z]+", "", text)        # ignore leading quotes, brackets, etc.
    return text, (group.get("name") or "").casefold()


# How cards are put in order (Advanced settings > Card order; card_sort in settings.ini):
# MDFCs first so their backs share as few back pages as possible, or simply everything A-Z.
CARD_SORTS = {"mdfc_first": "MDFC then Alphabetical", "alphabetical": "Only Alphabetical"}
CARD_SORT = "mdfc_first"


def set_card_sort(value) -> str:
    """Use this card order from now on (an unknown value means MDFC then Alphabetical)."""
    global CARD_SORT
    CARD_SORT = value if value in CARD_SORTS else "mdfc_first"
    return CARD_SORT


def mdfcs_first() -> bool:
    return CARD_SORT != "alphabetical"


def print_order(groups) -> list:
    """MDFCs first (alphabetical), then every other card alphabetically, filling the next open slot -
    or, with Only Alphabetical, every card A-Z."""
    if not mdfcs_first():
        return sorted(groups, key=sort_key)
    mdfcs = sorted((g for g in groups if g["back_id"]), key=sort_key)
    others = sorted((g for g in groups if not g["back_id"]), key=sort_key)
    return mdfcs + others


def sorted_cards(cards: list) -> list:
    """A deck's cards (one entry per copy) in the card order, ignoring where they were moved to. Cards
    without a name keep their order, after the named ones of their kind."""
    order = sorted(range(len(cards)), key=lambda i: ((0 if cards[i]["back"] else 1) if mdfcs_first() else 0,
                                                     cards[i].get("sort") or "~", i))
    return [cards[i] for i in order]


def card_slots(layout=(1.0, 0.0)) -> list[tuple[float, float]]:
    """Bottom-left corner (in points) of each card's area (card + bleed) on a page, left to right and
    top to bottom; the grid is centred on the page (or in the area a cutting machine can use: see
    print_area). `layout` = (bleed edge mm, card spacing mm[, paper...])."""
    x0, y0, x1, y1 = print_area(layout)
    area_w, area_h = x1 - x0, y1 - y0
    bleed, spacing = layout[0] * PT_PER_MM, layout[1] * PT_PER_MM
    card_w, card_h = card_dims(layout)
    cell_w, cell_h = card_w + 2 * bleed, card_h + 2 * bleed
    cols = max(1, int((area_w + spacing) // (cell_w + spacing)))
    rows = max(1, int((area_h + spacing) // (cell_h + spacing)))
    left = x0 + (area_w - (cols * cell_w + (cols - 1) * spacing)) / 2
    top = y1 - (area_h - (rows * cell_h + (rows - 1) * spacing)) / 2       # the grid's top edge
    return [(left + c * (cell_w + spacing), top - r * (cell_h + spacing) - cell_h)
            for r in range(rows) for c in range(cols)]


def trim_boxes(layout, count: int | None = None) -> list[tuple[float, float, float, float]]:
    """(x, y, w, h) in points of each card itself (without bleed) on a front page: where it's cut."""
    bleed = layout[0] * PT_PER_MM
    card_w, card_h = card_dims(layout)
    slots = card_slots(layout)
    return [(x + bleed, y + bleed, card_w, card_h) for x, y in slots[:len(slots) if count is None else count]]


def mark_shapes(layout, count: int) -> list[tuple[float, float, float, float]]:
    """The cut marks on a front page with `count` cards, as filled black rectangles (x, y, w, h) in
    points (PDF coordinates: y up). Crop marks: thin lines in the margins in line with every card
    edge. Silhouette: type 1 registration marks (a square top left, L shapes top right and bottom
    left). Cricut and none: nothing (Design Space prints its own marks)."""
    marks = layout_marks(layout)
    PAGE_W, PAGE_H = page_size(layout)
    out = []
    if marks == "crop" and count:
        boxes = trim_boxes(layout, count)
        line = 0.3                                            # pt
        gap = 1.5 * PT_PER_MM                                 # between the cards' area and a mark
        left = min(b[0] for b in boxes) - layout[0] * PT_PER_MM
        right = max(b[0] + b[2] for b in boxes) + layout[0] * PT_PER_MM
        bottom = min(b[1] for b in boxes) - layout[0] * PT_PER_MM
        top = max(b[1] + b[3] for b in boxes) + layout[0] * PT_PER_MM
        edge = 1.0 * PT_PER_MM                                # stay off the very edge of the paper
        xs = sorted({round(v, 3) for b in boxes for v in (b[0], b[0] + b[2])})
        ys = sorted({round(v, 3) for b in boxes for v in (b[1], b[1] + b[3])})
        most = CROP_MARK_MM * PT_PER_MM
        for x in xs:
            up = min(most, PAGE_H - edge - (top + gap))                  # above the cards
            if up >= 1.5:
                out.append((x - line / 2, top + gap, line, up))
            down = min(most, bottom - gap - edge)                        # below
            if down >= 1.5:
                out.append((x - line / 2, bottom - gap - down, line, down))
        for y in ys:
            r_len = min(most, PAGE_W - edge - (right + gap))             # right of the cards
            if r_len >= 1.5:
                out.append((right + gap, y - line / 2, r_len, line))
            l_len = min(most, left - gap - edge)                         # left
            if l_len >= 1.5:
                out.append((left - gap - l_len, y - line / 2, l_len, line))
    elif marks == "silhouette":
        inset, length, th = (SILHOUETTE_MARKS[k] * PT_PER_MM for k in ("inset", "length", "thickness"))
        sq = SILHOUETTE_SQUARE_MM * PT_PER_MM
        out.append((inset, PAGE_H - inset - sq, sq, sq))                       # top left: the square
        out.append((PAGE_W - inset - length, PAGE_H - inset - th, length, th))  # top right: an L
        out.append((PAGE_W - inset - th, PAGE_H - inset - length, th, length))
        out.append((inset, inset, length, th))                                  # bottom left: an L
        out.append((inset, inset, th, length))
    return out


def shows_header(layout) -> bool:
    """Headers stay off pages a cutting machine reads (they could be taken for a registration mark)."""
    return layout_marks(layout) not in ("silhouette", "cricut")


def _rounded_rect_points(x, y, w, h, r, steps: int = 6) -> list[tuple[float, float]]:
    """The outline of a rounded rectangle as points (for cut files), counter-clockwise from bottom left."""
    import math
    r = max(0.0, min(r, w / 2, h / 2))
    pts = []
    for cx, cy, a0 in ((x + w - r, y + r, -90), (x + w - r, y + h - r, 0), (x + r, y + h - r, 90), (x + r, y + r, 180)):
        for i in range(steps + 1):
            a = math.radians(a0 + 90 * i / steps)
            pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts


def write_cut_files(layout, count: int, folder: Path, stem: str) -> list[Path]:
    """Cut lines for a front page with `count` cards, for Silhouette Studio and other cutter software:
    an SVG (Silhouette Studio Designer edition and up, Brother CanvasWorkspace, Inkscape...) and a DXF
    (Silhouette Studio's free Basic edition). Both are exactly the page's size, in mm, with each card's
    outline (rounded corners) where it's printed. Returns the files written."""
    PAGE_W, PAGE_H = page_size(layout)
    r = corner_radius(layout)
    boxes = trim_boxes(layout, count)
    mm = lambda pt: pt / PT_PER_MM
    folder.mkdir(parents=True, exist_ok=True)
    svg = folder / f"{stem}.svg"
    rects = "\n".join(
        f'  <rect x="{mm(x):.3f}" y="{mm(PAGE_H - y - h):.3f}" width="{mm(w):.3f}" height="{mm(h):.3f}" '
        f'rx="{mm(r):.3f}" ry="{mm(r):.3f}" fill="none" stroke="#ff0000" stroke-width="0.25"/>'
        for x, y, w, h in boxes)
    svg.write_text(f'<?xml version="1.0" encoding="UTF-8"?>\n'
                   f'<svg xmlns="http://www.w3.org/2000/svg" width="{mm(PAGE_W):.3f}mm" height="{mm(PAGE_H):.3f}mm" '
                   f'viewBox="0 0 {mm(PAGE_W):.3f} {mm(PAGE_H):.3f}">\n'
                   f'  <title>Fleem cut lines: {count} card(s) on {layout_paper(layout)} '
                   f'{layout_orientation(layout)}</title>\n{rects}\n</svg>\n', encoding="utf-8")
    dxf = folder / f"{stem}.dxf"
    lines = ["0", "SECTION", "2", "HEADER", "9", "$ACADVER", "1", "AC1009", "9", "$INSUNITS", "70", "4",
             "9", "$EXTMIN", "10", "0.0", "20", "0.0", "9", "$EXTMAX", "10", f"{mm(PAGE_W):.3f}", "20",
             f"{mm(PAGE_H):.3f}", "0", "ENDSEC", "0", "SECTION", "2", "ENTITIES"]
    for x, y, w, h in boxes:                 # R12 polylines: read by nearly everything
        lines += ["0", "POLYLINE", "8", "CUT", "66", "1", "70", "1"]
        for px, py in _rounded_rect_points(x, y, w, h, r):
            lines += ["0", "VERTEX", "8", "CUT", "10", f"{mm(px):.4f}", "20", f"{mm(py):.4f}"]
        lines += ["0", "SEQEND", "8", "CUT"]
    lines += ["0", "ENDSEC", "0", "EOF"]
    dxf.write_text("\r\n".join(lines) + "\r\n", encoding="ascii")
    return [svg, dxf]


def cricut_sheet(cards: list, layout, renderer, dpi: int = 300):
    """One Cricut Print Then Cut sheet: Design Space's area (6.75 x 9.25 in, or turned) as a PIL RGBA
    image, every card at its real size with rounded corners and see-through everywhere else - Design
    Space cuts around each card and adds its own bleed and sensor marks."""
    from PIL import Image, ImageDraw
    layout = machine_layout(layout)
    x0, y0, x1, y1 = print_area(layout)
    s = dpi / 72
    sheet = Image.new("RGBA", (round((x1 - x0) * s), round((y1 - y0) * s)), (0, 0, 0, 0))
    card_w, card_h, bleed_x, bleed_y = card_geometry(layout)
    full_w, full_h = round((card_w + 2 * bleed_x) * s), round((card_h + 2 * bleed_y) * s)
    cw, ch = round(card_w * s), round(card_h * s)
    mask = Image.new("L", (cw, ch), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, cw - 1, ch - 1), radius=round(corner_radius(layout) * s), fill=255)
    for (bx, by, bw, bh), card in zip(trim_boxes(layout, len(cards)), cards):
        im = renderer.card_image(card["front"], full_w, full_h)
        ox, oy = round(bleed_x * s), round(bleed_y * s)
        face = im.crop((ox, oy, ox + cw, oy + ch)).convert("RGBA")
        sheet.paste(face, (round((bx - x0) * s), round((y1 - by - bh) * s)), mask)
    return sheet


def cricut_size_text(layout) -> str:
    """The Cricut sheet's size, e.g. "6.75 x 9.25 in" (to type into Design Space)."""
    x0, y0, x1, y1 = print_area(machine_layout(layout))
    return f"{(x1 - x0) / 72:.2f} x {(y1 - y0) / 72:.2f} in"


def pdf_image(path: Path) -> tuple[bytes, int, int, str]:
    """A card image as JPEG data for the PDF: the original file if it's already a plain JPEG within
    MAX_DPI, otherwise converted (and scaled down) once. Returns (data, width, height, colour space)."""
    from PIL import Image
    data = path.read_bytes()
    max_w = round((CARD_W + 2 * IMAGE_BLEED) / 72 * MAX_DPI)
    with Image.open(path) as im:
        if im.format == "JPEG" and im.mode in ("RGB", "L") and im.width <= max_w:
            return data, im.width, im.height, "/DeviceRGB" if im.mode == "RGB" else "/DeviceGray"
        im = im.convert("RGB")
        if im.width > max_w:
            im = im.resize((max_w, round(im.height * max_w / im.width)), Image.LANCZOS)
        import io
        out = io.BytesIO()
        im.save(out, "JPEG", quality=JPEG_QUALITY, subsampling=0)
        return out.getvalue(), im.width, im.height, "/DeviceRGB"


IMAGE_TYPES = (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tif", ".tiff", ".jfif", ".avif")


def _edge_strips(im, k: int):
    """The outer k pixels of each side (skipping the corners), as images: {"left", "right", "top", "bottom"}."""
    w, h = im.size
    mx, my = max(1, w // 8), max(1, h // 8)
    return {"left": im.crop((1, my, 1 + k, h - my)), "right": im.crop((w - 1 - k, my, w - 1, h - my)),
            "top": im.crop((mx, 1, w - mx, 1 + k)), "bottom": im.crop((mx, h - 1 - k, w - mx, h - 1))}


def _edge_kinds(im):
    """For each side: ("solid", colour) when the card's very edge is one plain colour (a border or a thin
    outline - it's continued in that colour), else ("art", None) when the picture runs to the edge
    (it's continued as a mirror image)."""
    im = _small_copy(im)
    k = max(2, round(im.width * 0.004))       # just the outermost few pixels
    out = {}
    for side, strip in _edge_strips(im.convert("RGB"), k).items():
        data = list(strip.getdata())
        mid = len(data) // 2
        colour = tuple(sorted(c[i] for c in data)[mid] for i in range(3))
        close = sum(1 for c in data if sum((a - b) ** 2 for a, b in zip(c, colour)) <= 45 * 45)
        out[side] = ("solid", colour) if close >= 0.88 * len(data) else ("art", None)
    return out


def _small_copy(im, width: int = 816):
    """A smaller copy for looking at (the edge checks don't need full resolution)."""
    from PIL import Image
    if im.width <= width * 1.2:
        return im
    return im.resize((width, max(1, round(im.height * width / im.width))), Image.BILINEAR)


def _border_colour(im):
    """The colour of the card's border: the median of all four edges."""
    from PIL import ImageStat, Image
    im = _small_copy(im)
    k = max(2, round(im.width * 0.004))
    strips = list(_edge_strips(im.convert("RGB"), k).values())
    joined = Image.new("RGB", (sum(s.width * s.height for s in strips), 1))
    x = 0
    for s in strips:
        data = s.resize((s.width * s.height, 1))
        joined.paste(data, (x, 0))
        x += data.width
    return tuple(int(v) for v in ImageStat.Stat(joined).median)


def _fill_corners(im):
    """Rounded card corners (white, black or see-through outside the curve) are filled with the colour of
    the card's edge next to each corner, so they don't turn into wedges or streaks in the bleed.
    Returns an RGB image."""
    from PIL import Image, ImageDraw
    w, h = im.size
    if im.mode == "RGBA":                     # see-through corners: put the border colour behind them
        bg = Image.new("RGBA", im.size, _border_colour(im.convert("RGB")) + (255,))
        bg.alpha_composite(im)
        im = bg
    im = im.convert("RGB")
    px = im.load()
    dist = lambda a, b: sum((x - y) ** 2 for x, y in zip(a, b)) ** 0.5
    radius_max = min(w, h) * 0.08             # card corners are ~4.8 % of the width; allow a bit more
    t_max = int(radius_max * 0.293) + 1       # the curve meets the diagonal 0.293 r from the corner
    k = max(2, round(w * 0.004))
    for flip_x, flip_y in ((False, False), (True, False), (False, True), (True, True)):
        X = lambda x: w - 1 - x if flip_x else x
        Y = lambda y: h - 1 - y if flip_y else y
        corner = px[X(0), Y(0)]
        # the card's edge colour just past this corner's curve, along both sides
        a, b = int(radius_max), int(radius_max * 2)
        samples = [px[X(x), Y(d)] for x in range(a, b) for d in range(1, k + 1)] + \
                  [px[X(d), Y(y)] for y in range(a, b) for d in range(1, k + 1)]
        local = tuple(sorted(s[i] for s in samples)[len(samples) // 2] for i in range(3))
        if dist(corner, local) < 50:
            continue                          # square corner, or already the edge's colour
        t = next((t for t in range(1, t_max) if dist(px[X(t), Y(t)], corner) > 60), None)
        if t is None or t < 2:
            continue                          # no rounded corner here (just a detail in the picture)
        r = t / 0.293 + 1.5
        inset = max(1.5, w * 0.003)           # also cover the soft (anti-aliased) rim just inside the curve
        size = int(r) + 2
        scale = 4                             # draw the curve 4x bigger, then shrink: smooth edge
        mask = Image.new("L", (size * scale, size * scale), 255)
        ImageDraw.Draw(mask).ellipse((inset * scale, inset * scale, (2 * r - inset) * scale, (2 * r - inset) * scale),
                                     fill=0)
        mask = mask.resize((size, size), Image.LANCZOS)
        if flip_x:
            mask = mask.transpose(Image.FLIP_LEFT_RIGHT)
        if flip_y:
            mask = mask.transpose(Image.FLIP_TOP_BOTTOM)
        im.paste(local, (w - size if flip_x else 0, h - size if flip_y else 0), mask)
    return im


def _pad_smart(im, left: int, top: int, right: int, bottom: int, kinds=None):
    """Grow an image outwards for bleed: a side whose edge is one plain colour (a border) continues in
    that colour; a side where the picture runs to the edge is continued as its mirror image."""
    from PIL import Image
    im = im.convert("RGB")
    if not (left or top or right or bottom):
        return im
    w, h = im.size
    kinds = kinds or _edge_kinds(im)
    out = Image.new("RGB", (w + left + right, h + top + bottom))
    W, H = out.size

    # mirrored continuation everywhere first (sides, then corners), then plain colours on top
    if left:
        out.paste(_tile_mirror(im, "left", left), (0, top))
    if right:
        out.paste(_tile_mirror(im, "right", right), (left + w, top))
    out.paste(im, (left, top))
    if top:
        out.paste(_tile_mirror(out.crop((0, top, W, top + h)), "top", top), (0, 0))
    if bottom:
        out.paste(_tile_mirror(out.crop((0, top, W, top + h)), "bottom", bottom), (0, top + h))
    boxes = {"left": (0, 0, left, H), "right": (left + w, 0, W, H), "top": (0, 0, W, top), "bottom": (0, top + h, W, H)}
    sizes = {"left": left, "right": right, "top": top, "bottom": bottom}
    for side, (kind, colour) in kinds.items():
        if kind == "solid" and sizes[side]:
            out.paste(colour, boxes[side])
    return out


def _tile_mirror(im, side: str, n: int):
    """A strip n pixels deep that continues `im` past one side as its mirror image (repeated if the
    strip is deeper than the image)."""
    from PIL import Image
    w, h = im.size
    flip = Image.FLIP_LEFT_RIGHT if side in ("left", "right") else Image.FLIP_TOP_BOTTOM
    depth = w if side in ("left", "right") else h
    pieces, made, src_flip = [], 0, True
    while made < n:
        take = min(depth, n - made)
        if side == "left":
            piece = im.crop((0, 0, take, h))
        elif side == "right":
            piece = im.crop((w - take, 0, w, h))
        elif side == "top":
            piece = im.crop((0, 0, w, take))
        else:
            piece = im.crop((0, h - take, w, h))
        pieces.append(piece.transpose(flip) if src_flip else piece)
        made += take
        src_flip = not src_flip
    if side in ("left", "right"):
        strip = Image.new("RGB", (n, h))
        x = n if side == "left" else 0
        for p in pieces:
            if side == "left":
                x -= p.width
                strip.paste(p, (x, 0))
            else:
                strip.paste(p, (x, 0))
                x += p.width
    else:
        strip = Image.new("RGB", (w, n))
        y = n if side == "top" else 0
        for p in pieces:
            if side == "top":
                y -= p.height
                strip.paste(p, (0, y))
            else:
                strip.paste(p, (0, y))
                y += p.height
    return strip


def _bleed_in_image(aspect: float):
    """How much bleed (inches, each side) an image of this shape has, if it's clearly a card with some
    bleed - more than none (a plain card is 2.48 x 3.46 in) and not MPC's 0.12 in. None otherwise."""
    full = (CARD_W + 2 * IMAGE_BLEED) / (CARD_H + 2 * IMAGE_BLEED)
    trim = CARD_W / CARD_H
    if abs(aspect - full) <= 0.006 or abs(aspect - trim) <= 0.006 or aspect >= 1:
        return None
    w, h = CARD_W / 72, CARD_H / 72
    b = (h * aspect - w) / (2 - 2 * aspect)           # (w + 2b) / (h + 2b) = aspect
    return b if 0.02 <= b <= 0.35 else None


def needs_bleed(path: Path) -> bool:
    """Should Fleem make bleed for this downloaded image? Yes for a card scan without bleed (see
    is_bare_card_scan), and for any image that is clearly card-shaped with no or too little bleed
    (less than 60 % of MPC's 0.12 in) - those would otherwise be stretched to fill the bleed too."""
    from PIL import Image
    try:
        with Image.open(path) as im:
            aspect = im.width / im.height
    except Exception:
        return False
    if abs(aspect - CARD_W / CARD_H) <= 0.004:              # exactly a card's shape: no bleed at all
        return True
    b = _bleed_in_image(aspect)
    if b is not None and b * 72 < 0.6 * IMAGE_BLEED:
        return True
    return is_bare_card_scan(path)


def is_bare_card_scan(path: Path) -> bool:
    """A downloaded image that's clearly a card with no bleed at all: card-shaped, with rounded (or
    see-through) corners. Those get bleed added. Everything else from MPC Autofill is used the way MPC
    prints it (the image fills the card plus its bleed) - including art whose uploader gave it a
    different border colour or slightly different proportions."""
    from PIL import Image
    try:
        with Image.open(path) as im:
            trim = CARD_W / CARD_H
            if abs(im.width / im.height - trim) > 0.012:
                return False
            if im.mode in ("RGBA", "LA", "PA") or "transparency" in im.info:
                a = im.convert("RGBA").getchannel("A")
                if sum(a.getpixel(p) < 128 for p in ((0, 0), (im.width - 1, 0), (0, im.height - 1),
                                                     (im.width - 1, im.height - 1))) >= 3:
                    return True
            small = _small_copy(im.convert("RGB"), 400)
    except Exception:
        return False
    edge = _border_colour(small)
    w, h = small.size
    px = small.load()
    dist = lambda a, b: sum((x - y) ** 2 for x, y in zip(a, b)) ** 0.5
    rounded = 0
    for cx, cy, dx, dy in ((0, 0, 1, 1), (w - 1, 0, -1, 1), (0, h - 1, 1, -1), (w - 1, h - 1, -1, -1)):
        corner = px[cx, cy]
        inside = px[cx + dx * int(w * 0.03), cy + dy * int(w * 0.03)]   # just inside the corner's curve
        if dist(corner, edge) > 60 and dist(inside, edge) < 60:
            rounded += 1
    return rounded >= 3


def prepare_card_image(src: Path, out: Path, as_card: bool = False) -> str:
    """Turn any card image into one like MPC Autofill's (2.72 x 3.70 in: the card plus 0.12 in of bleed
    on each side), so it lines up with the other cards. Landscape images are turned upright. An image
    that is already that shape is used as is; otherwise it's taken as the card itself: sized to the
    card's shape (cropped only if it's clearly another shape), its rounded corners filled with the
    border colour, and the bleed made by continuing the border colour (or mirroring the picture where
    it runs to the edge). Saved as a JPEG, at most MAX_DPI. Returns a short description of what was done.
    `as_card`: the picture is the card itself whatever its shape (Make bleed: for images that are MPC's
    shape but have no bleed in them)."""
    from PIL import Image, ImageOps
    with Image.open(src) as im:
        im = ImageOps.exif_transpose(im)
        if im.mode in ("RGBA", "LA", "P", "PA") or "transparency" in im.info:
            im = im.convert("RGBA")
        else:
            im = im.convert("RGB")
    done = []
    if im.width > im.height:
        im = im.rotate(90, expand=True)
        done.append("turned upright")
    full = (CARD_W + 2 * IMAGE_BLEED) / (CARD_H + 2 * IMAGE_BLEED)      # 2.72 / 3.70
    trim = CARD_W / CARD_H                                               # 2.48 / 3.46
    aspect = im.width / im.height
    if abs(aspect - full) <= 0.006 and not as_card:  # already has MPC-style bleed
        if im.mode == "RGBA":
            bg = Image.new("RGBA", im.size, "white")
            bg.alpha_composite(im)
            im = bg
        im = im.convert("RGB")
        max_w = round((CARD_W + 2 * IMAGE_BLEED) / 72 * MAX_DPI)
        if im.width > max_w:
            im = im.resize((max_w, round(im.height * max_w / im.width)), Image.LANCZOS)
        done.append("already had bleed")
    elif _bleed_in_image(aspect) is not None and not as_card:   # some bleed already (less or more than MPC's)
        b_in = _bleed_in_image(aspect)
        card_px = im.width * CARD_W / (CARD_W + 2 * b_in * 72)        # the card's own width in the image
        dpi = min(MAX_DPI, max(300, card_px / (CARD_W / 72)))
        scale = (CARD_W / 72 * dpi) / card_px
        im = im.resize((max(1, round(im.width * scale)), max(1, round(im.height * scale))), Image.LANCZOS)
        if im.mode == "RGBA":
            bg = Image.new("RGBA", im.size, _border_colour(im.convert("RGB")) + (255,))
            bg.alpha_composite(im)
            im = bg
        im = im.convert("RGB")
        tw = round((CARD_W + 2 * IMAGE_BLEED) / 72 * dpi)
        th = round((CARD_H + 2 * IMAGE_BLEED) / 72 * dpi)
        if b_in * 72 >= IMAGE_BLEED:                 # more bleed than needed: trim it evenly
            x0, y0 = (im.width - tw) // 2, (im.height - th) // 2
            im = im.crop((x0, y0, x0 + tw, y0 + th))
            done.append(f"extra bleed trimmed ({b_in * 25.4:.1f} mm to {IMAGE_BLEED / 72 * 25.4:.1f} mm)")
        else:                                       # less: keep it and add the rest
            px, py = max(0, (tw - im.width) // 2), max(0, (th - im.height) // 2)
            im = _pad_smart(im, px, py, max(0, tw - im.width - px), max(0, th - im.height - py))
            done.append(f"bleed topped up ({b_in * 25.4:.1f} mm to {IMAGE_BLEED / 72 * 25.4:.1f} mm)")
    else:
        if abs(aspect - trim) / trim > 0.04:        # clearly another shape: crop the middle to the card's shape
            if aspect > trim:
                w = round(im.height * trim)
                im = im.crop(((im.width - w) // 2, 0, (im.width - w) // 2 + w, im.height))
            else:
                h = round(im.width / trim)
                im = im.crop((0, (im.height - h) // 2, im.width, (im.height - h) // 2 + h))
            done.append("cropped to card shape")
        dpi = min(MAX_DPI, max(300, im.width / (CARD_W / 72)))
        cw, ch = round(CARD_W / 72 * dpi), round(CARD_H / 72 * dpi)
        b = round(IMAGE_BLEED / 72 * dpi)
        im = im.resize((cw, ch), Image.LANCZOS)      # (a few % off the card's shape: sized, not cropped)
        im = _fill_corners(im)
        kinds = _edge_kinds(im)
        im = _pad_smart(im, b, b, b, b, kinds)
        solid = [s for s, (k, _) in kinds.items() if k == "solid"]
        done.append("bleed added (" + ("border colour" if len(solid) == 4 else
                                       "mirrored" if not solid else "border colour + mirrored") + ")")
    im.save(out, "JPEG", quality=JPEG_QUALITY, subsampling=0)
    return ", ".join(done)


def write_cards_pdf(groups, images_dir: Path, out_path: Path, layout=(1.0, 0.0), keep_order: bool = False,
                    header_name: str | None = None, numbered=None) -> None:
    """Lay the cards out like Proxy PDF Maker does and write the PDF: MDFCs first, then A-Z (or in the
    given order with keep_order); every front page is followed by its back page, which only shows the
    MDFCs' real backs (the rest stay white). `layout` = (bleed edge mm, card spacing mm). Each image is
    stored once however many copies use it."""
    import zlib
    if not keep_order:
        groups = print_order(groups)
    copies = [g for g in groups for _ in range(g["count"])]
    slots = card_slots(layout)
    PAGE_W, PAGE_H = page_size(layout)
    BLEED = layout[0] * PT_PER_MM
    pages = [copies[i:i + len(slots)] for i in range(0, len(copies), len(slots))] or [[]]

    files = []                       # each image file once, in first-use order
    names = {}                       # front image -> (card name, alphabetical sort text), kept in the PDF
    arts = {}                        # image -> its MPCFill art (id, name, query), kept in the PDF (for XML export)
    for g in groups:
        for f in (g.get("front_file"), g.get("back_file")):
            if f and f not in files:
                files.append(f)
        if g.get("front_file") and g["front_file"] not in names and (g.get("name") or g.get("query")):
            names[g["front_file"]] = (g.get("name") or g.get("query"), g.get("_sort") or sort_key(g)[0])
        if "_art" in g:                          # from a deck in memory: what's known about its images
            for side in ("front", "back"):
                if g["_art"].get(side) and g.get(f"{side}_file"):
                    arts.setdefault(g[f"{side}_file"], g["_art"][side])
        elif not g.get("_missing"):              # straight from an XML
            if g.get("front_file"):
                arts.setdefault(g["front_file"], {"id": g["front_id"], "name": g.get("name") or "",
                                                  "query": g.get("query") or ""})
            if g.get("back_file") and g.get("back_id"):
                arts.setdefault(g["back_file"], {"id": g["back_id"], "name": g.get("back_name") or "",
                                                 "query": g.get("back_query") or ""})

    def pdf_text(text: str) -> str:
        return "<FEFF" + text.encode("utf-16-be").hex().upper() + ">"
    header = layout_header(layout) if shows_header(layout) else None
    if header == "":
        header = header_name or Path(out_path).stem
    total_pages = 2 * len(pages)

    def marks_ops(count: int) -> str:            # cut marks (front pages only)
        shapes = mark_shapes(layout, count)
        return ("q\n0 g\n" + "".join(f"{x:.3f} {y:.3f} {w:.3f} {h:.3f} re f\n" for x, y, w, h in shapes) + "Q\n"
                if shapes else "")
    selected_only = layout_numbering(layout) == "selected"
    numbered = sorted(numbered or [])          # with "selected": the pages being printed (1 = first page)

    def header_ops(page_no: int) -> str:
        if not header:
            return ""
        x, y, size = header_place(layout)
        if selected_only:                      # only the pages being printed, counted 1..n
            number = numbered.index(page_no) + 1 if page_no in numbered else None
            line = header_line(header, number, len(numbered))
        else:
            line = header_line(header, page_no, total_pages)
        raw = line.encode("cp1252", "replace")
        text = raw.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)").decode("latin1")
        return f"q\n0 g\nBT\n/FH {size:.2f} Tf\n{x:.3f} {y:.3f} Td\n({text}) Tj\nET\nQ\n"
    first_image = 3                  # 1 = catalog, 2 = page tree
    image_obj = {f: first_image + i for i, f in enumerate(files)}
    first_page = first_image + len(files)
    page_objs = [first_page + 2 * i for i in range(2 * len(pages))]   # each page + its content stream

    card_w, card_h, bleed_x, bleed_y = card_geometry(layout)

    def draw(card_file, x, y):
        img_x, img_y = x + BLEED - bleed_x, y + BLEED - bleed_y
        return (f"q\n{x:.6f} {y:.6f} {card_w + 2 * BLEED:.6f} {card_h + 2 * BLEED:.6f} re\nW\nn\nq\n"
                f"{card_w + 2 * bleed_x:.6f} 0 0 {card_h + 2 * bleed_y:.6f} {img_x:.6f} {img_y:.6f} cm\n"
                f"/Im{image_obj[card_file]} Do\nQ\nQ\n")

    tmp = out_path.with_suffix(".part")
    offsets = {}
    with open(tmp, "wb") as f:
        def obj(num, body: bytes, stream: bytes | None = None):
            offsets[num] = f.tell()
            f.write(b"%d 0 obj\n" % num + body)
            if stream is not None:
                f.write(b"\nstream\n" + stream + b"\nendstream")
            f.write(b"\nendobj\n")

        f.write(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
        obj(1, b"<< /Type /Catalog /Pages 2 0 R >>")
        kids = " ".join(f"{n} 0 R" for n in page_objs)
        obj(2, f"<< /Type /Pages /Kids [{kids}] /Count {len(page_objs)} >>".encode())
        for name in files:
            data, w, h, cs = pdf_image(images_dir / name)
            label = (f"/FleemName {pdf_text(names[name][0])} /FleemSort {pdf_text(names[name][1])} "
                     if name in names else "")
            if name in arts and arts[name].get("id"):
                a = arts[name]
                label += (f"/FleemArt {pdf_text(a['id'])} /FleemArtName {pdf_text(a.get('name') or '')} "
                          f"/FleemArtQuery {pdf_text(a.get('query') or '')} ")
            obj(image_obj[name], f"<< /Type /XObject /Subtype /Image /Width {w} /Height {h} /ColorSpace {cs} "
                                 f"/BitsPerComponent 8 /Filter /DCTDecode {label}/Length {len(data)} >>".encode(), data)
        n = 0
        for cards in pages:
            front = "".join(draw(c["front_file"], *slots[i]) for i, c in enumerate(cards))
            back = "".join(draw(c["back_file"], PAGE_W - slots[i][0] - (card_w + 2 * BLEED), slots[i][1])
                           for i, c in enumerate(cards) if c.get("back_file"))
            front += marks_ops(len(cards))
            for content in (front, back):
                used = sorted({int(x) for x in re.findall(r"/Im(\d+) Do", content)})
                res = " ".join(f"/Im{u} {u} 0 R" for u in used)
                content += header_ops(n + 1)
                fonts = (" /Font << /FH << /Type /Font /Subtype /Type1 /BaseFont /Helvetica "
                         "/Encoding /WinAnsiEncoding >> >>") if header else ""
                page_num, content_num = page_objs[n], page_objs[n] + 1
                obj(page_num, (f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PAGE_W:.6f} {PAGE_H:.6f}] "
                               f"/Resources << /XObject << {res} >>{fonts} >> /Contents {content_num} 0 R >>").encode())
                packed = zlib.compress(content.encode("latin1"))     # (the header's text is cp1252)
                obj(content_num, b"<< /Filter /FlateDecode /Length %d >>" % len(packed), packed)
                n += 1
        xref = f.tell()
        total = page_objs[-1] + 2
        f.write(b"xref\n0 %d\n0000000000 65535 f \n" % total)
        for num in range(1, total):
            f.write(b"%010d 00000 n \n" % offsets[num])
        remembered = (f" /FleemHeader {pdf_text(header)}" + (" /FleemNumbering /selected" if selected_only else "")).encode() \
            if header else b""
        if layout_marks(layout) != "none":
            remembered += f" /FleemMarks /{layout_marks(layout)}".encode()
        f.write(b"trailer\n<< /Size %d /Root 1 0 R /Info << /Producer (Fleem Proxy Companion)%s >> >>\n"
                b"startxref\n%d\n%%%%EOF\n" % (total, remembered, xref))
    os.replace(tmp, out_path)


# ============================================================ main loop
def load_state() -> dict:
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_state(state: dict) -> None:
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, indent=2), encoding="utf-8")
    os.replace(tmp, STATE_FILE)


def pdf_name(xml_filename: str) -> str:
    """Discord turns spaces in attachment names into underscores; put them back.

    'kim_-_My_Order.xml' -> 'kim - My Order' (the username part keeps any underscores it really has).
    """
    stem = Path(xml_filename).stem
    if "_-_" in stem:
        user, rest = stem.split("_-_", 1)
        return f"{user} - {rest.replace('_', ' ')}"
    return stem


def open_pdfs(settings: Settings, paths) -> None:
    """Open finished PDFs in your normal PDF app."""
    if not settings.open_pdfs:
        return
    for path in paths:
        try:
            os.startfile(str(path))
        except Exception as e:
            log.warning("Couldn't open %s: %s", Path(path).name, e)


# ============================================================ printing
LAST_LAYOUT_FILE = HERE / "last_layout.json"


def last_layout():
    """The bleed edge and card spacing used last time (DEFAULT_LAYOUT the first time)."""
    try:
        d = json.loads(LAST_LAYOUT_FILE.read_text(encoding="utf-8"))
        paper = d.get("paper") if d.get("paper") in PAPER_SIZES else DEFAULT_PAPER
        orientation = "landscape" if d.get("orientation") == "landscape" else "portrait"
        card = d.get("card")
        card = CARD_SIZE_ALIASES.get(card, card)
        card = card if card in CARD_SIZES or custom_card_mm(card) else DEFAULT_CARD
        return (float(d["bleed_mm"]), float(d["spacing_mm"]), paper, orientation, "" if d.get("header") else None,
                "selected" if d.get("numbering") == "selected" else "all", card,
                d.get("marks") if d.get("marks") in CUT_MARKS else "none")
    except Exception:
        return DEFAULT_LAYOUT


def remember_layout(layout) -> None:
    """Remember the layout for next time (the header only as on/off: its text belongs to the file)."""
    try:
        LAST_LAYOUT_FILE.write_text(json.dumps({"bleed_mm": layout[0], "spacing_mm": layout[1],
                                                "paper": layout_paper(layout), "orientation": layout_orientation(layout),
                                                "header": layout_header(layout) is not None,
                                                "numbering": layout_numbering(layout),
                                                "card": layout_card(layout),
                                                "marks": layout_marks(layout)}), encoding="utf-8")
    except OSError:
        pass


def describe_layout(layout) -> str:
    text = f"{layout[0]:.2f} mm bleed edge, {layout[1]:.2f} mm card spacing, {layout_paper(layout)} paper"
    if layout_card(layout) != DEFAULT_CARD:
        text += f", {card_size_label(layout_card(layout))} cards"
    if layout_orientation(layout) == "landscape":
        text += " (landscape)"
    if layout_header(layout) is not None and shows_header(layout):
        text += ", with a header" + (" (numbering only the pages printed)" if layout_numbering(layout) == "selected" else "")
    marks = {"crop": "crop marks", "silhouette": "Silhouette registration marks",
             "cricut": "laid out for Cricut Print Then Cut"}.get(layout_marks(layout))
    if marks:
        text += f", {marks}"
    return text


def header_line(text: str, page_no: int | None, total: int | None) -> str:
    """The header's text with its page number (just the text when the page gets no number)."""
    return f"{text}   ·   page {page_no} of {total}" if page_no else text


def header_place(layout) -> tuple[float, float, float]:
    """Where the header goes on a page: (x, baseline y, font size), all in points - in the gap above the
    cards, lined up with their left edge."""
    page_w, page_h = page_size(layout)
    slots = card_slots(layout)
    cell_h = card_dims(layout)[1] + 2 * layout[0] * PT_PER_MM
    gap = page_h - (max(y for _, y in slots) + cell_h)          # paper above the top row
    size = max(4.0, min(8.0, gap * 0.55))
    return min(x for x, _ in slots), page_h - gap / 2 - size * 0.35, size


def parse_layout(bleed_text: str, spacing_text: str):
    """Text from the two boxes -> (bleed mm, spacing mm), or an error message."""
    try:
        bleed = float(bleed_text.replace("mm", "").replace(",", ".").strip())
        spacing = float(spacing_text.replace("mm", "").replace(",", ".").strip())
    except ValueError:
        return "Bleed edge and card spacing must be numbers (in mm)."
    if not 0 <= bleed <= MAX_BLEED_MM:
        return f"Bleed edge must be between 0 and {MAX_BLEED_MM:.2f} mm."
    if not 0 <= spacing <= MAX_SPACING_MM:
        return f"Card spacing must be between 0 and {MAX_SPACING_MM:.0f} mm."
    return (round(bleed, 3), round(spacing, 3))


_DRAW = re.compile(r"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re\s+W\s+n\s+q\s+([-\d.]+) 0 0 ([-\d.]+) "
                   r"([-\d.]+) ([-\d.]+) cm\s+/(\S+) Do")


def read_layout(path: Path):
    """Read the layout of a PDF made by Fleem (or Proxy PDF Maker with the same settings): its bleed
    edge and card spacing in mm, and every card copy in order with its front/back image. None if the
    PDF isn't laid out that way (then its spacing can't be changed)."""
    from pypdf import PdfReader
    try:
        reader = PdfReader(str(path))
        pages = reader.pages
        if not len(pages) or len(pages) % 2:
            return None
        copies, bleeds, xs, ys, cell, cards = [], set(), set(), set(), None, set()
        for p in range(0, len(pages), 2):
            def draws(page):
                c = page.get_contents()
                data = c.get_data().decode("latin1") if c is not None else ""
                xo = page["/Resources"].get("/XObject", {}) if "/Resources" in page else {}
                out = []
                for m in _DRAW.finditer(data):
                    x, y, w, h, iw, ih = (float(v) for v in m.groups()[:6])
                    ref = xo.raw_get("/" + m[9])
                    img = ref.get_object()
                    if str(img.get("/Filter")) != "/DCTDecode":
                        return None
                    # the image is a card image (0.12 in bleed around a Standard card), stretched to a card size
                    cw, ch = iw * CARD_W / (CARD_W + 2 * IMAGE_BLEED), ih * CARD_H / (CARD_H + 2 * IMAGE_BLEED)
                    if cw < 20 or ch < 20 or abs((w - cw) - (h - ch)) > 1:
                        return None
                    cards.add((round(cw, 1), round(ch, 1)))
                    bleeds.add(round((w - cw) / 2, 2))
                    key = (str(path), getattr(ref, "idnum", None) or id(img))   # unique across PDFs
                    out.append((x, y, w, h, key, ref))
                return out
            front, back = draws(pages[p]), draws(pages[p + 1])
            if front is None or back is None:
                return None
            for x, y, w, h, key, ref in front:
                xs.add(round(x, 2)); ys.add(round(y, 2)); cell = (w, h)
                mirrored = float(pages[p].mediabox.width) - x - w
                b = next(((bk, br) for bx, by, bw, bh, bk, br in back if abs(bx - mirrored) < 0.5 and abs(by - y) < 0.5),
                         (None, None))
                front_img = ref.get_object()
                copies.append({"front": (key, ref), "back": b if b[0] else None,
                               "name": str(front_img["/FleemName"]) if "/FleemName" in front_img else None,
                               "sort": str(front_img["/FleemSort"]) if "/FleemSort" in front_img else None})
                for src in ((key, ref), b if b[0] else None):     # MPCFill art the PDF remembers (XML export)
                    if src is not None:
                        img = src[1].get_object()
                        if "/FleemArt" in img:
                            remember_mpc_art(src, str(img["/FleemArt"]), str(img.get("/FleemArtName") or ""),
                                             str(img.get("/FleemArtQuery") or ""))
        if not copies or len(bleeds) != 1 or len(cards) != 1:
            return None
        card = card_size_named(*cards.pop())
        bleed_mm = bleeds.pop() / PT_PER_MM
        spacing_mm = None
        if len(xs) > 1:
            sx = sorted(xs)
            spacing_mm = (min(b - a for a, b in zip(sx, sx[1:])) - cell[0]) / PT_PER_MM
        elif len(ys) > 1:
            sy = sorted(ys)
            spacing_mm = (min(b - a for a, b in zip(sy, sy[1:])) - cell[1]) / PT_PER_MM
        first = reader.pages[0].mediabox
        try:
            header = reader.metadata.get("/FleemHeader") if reader.metadata else None
            numbering = "selected" if reader.metadata and str(reader.metadata.get("/FleemNumbering")) == "/selected" else "all"
            marks = str(reader.metadata.get("/FleemMarks") or "").lstrip("/") if reader.metadata else ""
        except Exception:
            header, numbering, marks = None, "all", ""
        return {"bleed_mm": round(bleed_mm, 2) + 0.0,
                "spacing_mm": None if spacing_mm is None else round(spacing_mm, 2) + 0.0,
                "paper": paper_named(float(first.width), float(first.height)),
                "orientation": "landscape" if float(first.width) > float(first.height) + 1 else "portrait",
                "header": str(header) if header else None, "numbering": numbering, "card": card,
                "marks": marks if marks in CUT_MARKS else "none", "copies": copies}
    except Exception as e:
        log.info("Couldn't read the layout of %s: %s", Path(path).name, e)
        return None


def info_layout(info) -> tuple:
    """read_layout's answer as a layout (spacing None if the PDF doesn't show it)."""
    return (info["bleed_mm"], info["spacing_mm"], info.get("paper"), info.get("orientation") or "portrait",
            info.get("header"), info.get("numbering") or "all", info.get("card") or DEFAULT_CARD,
            info.get("marks") or "none")


def same_layout(info, layout) -> bool:
    """Is the PDF (read_layout) already laid out with this bleed edge and card spacing?"""
    return (abs(info["bleed_mm"] - layout[0]) < 0.02 and
            (info["spacing_mm"] is None or abs(info["spacing_mm"] - layout[1]) < 0.02))


def export_folder_of(settings) -> Path | None:
    """Where Export as PDF saves without asking (Advanced settings > Set PDF Export Location), or None."""
    value = (getattr(settings, "export_folder", "") or "").strip()
    return Path(os.path.expandvars(value)).expanduser() if value else None


def set_export_folder(settings, folder: Path) -> None:
    """Make Export as PDF save to `folder` from now on, without asking."""
    save_setting("export_folder", str(folder))
    if settings is not None:
        settings.export_folder = str(folder)
    log.info("PDF export location: %s", folder)


def free_path(path: Path) -> Path:
    """`path`, or "name (2).pdf", "name (3).pdf"... if it's taken."""
    path = Path(path)
    base, n = path.stem, 2
    while path.exists():
        path = path.with_name(f"{base} ({n}){path.suffix}")
        n += 1
    return path


def export_pdf_pages(src: Path, indexes: list[int], target: Path) -> None:
    """Save some pages (0 = first page) of a PDF as a PDF of their own."""
    import io
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(io.BytesIO(Path(src).read_bytes()))        # (so `target` may even be `src`)
    writer = PdfWriter()
    for i in indexes:
        writer.add_page(reader.pages[i])
    Path(target).parent.mkdir(parents=True, exist_ok=True)
    tmp = Path(target).with_name(Path(target).name + ".part")
    with open(tmp, "wb") as f:
        writer.write(f)
    try:
        os.replace(tmp, target)
    except PermissionError:
        tmp.unlink(missing_ok=True)
        raise RuntimeError(f"{Path(target).name} is open in another program (e.g. a PDF viewer), so it can't be "
                           "saved. Close it there, then try again.") from None


def write_deck_pdf(cards: list, layout, out_path: Path, numbered=None) -> None:
    """Write a PDF from a card list (the in-memory deck): each card has a "front" and maybe a "back"
    image - ("file", Path) or an image inside an open PDF (key, pypdf reference) - plus its name and
    sort text. Same layout as every Fleem PDF. Writes to a temporary file first, then replaces out_path."""
    work = WORK_DIR / ("write-" + re.sub(r"[^\w-]+", "_", Path(out_path).stem)[:40])
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)
    try:
        written = {}

        def image_file(src):
            kind, ref = src
            key = ("file", str(ref)) if kind == "file" else kind
            if key not in written:
                name = f"{len(written)}.jpg"
                if kind == "file":
                    shutil.copy(ref, work / name)
                else:
                    (work / name).write_bytes(ref.get_object()._data)
                written[key] = name
            return written[key]
        groups = []
        for c in cards:
            g = {"front_id": str(c["front"][0]), "back_id": "back" if c["back"] else None, "name": c.get("name") or "",
                 "_sort": c.get("sort"), "count": 1, "front_file": image_file(c["front"]),
                 "_art": {"front": mpc_art_of(c["front"]), "back": mpc_art_of(c["back"])}}
            if c["back"]:
                g["back_file"] = image_file(c["back"])
            groups.append(g)
        tmp = work / "deck.pdf"
        write_cards_pdf(groups, work, tmp, layout, keep_order=True, header_name=Path(out_path).stem, numbered=numbered)
        Path(out_path).parent.mkdir(parents=True, exist_ok=True)
        try:
            os.replace(tmp, out_path)
        except PermissionError:
            raise RuntimeError(f"{Path(out_path).name} is open in another program (e.g. a PDF viewer), so it can't be "
                               "saved. Close it there, then try again.") from None
        except OSError:                    # e.g. another drive
            try:
                shutil.move(str(tmp), str(out_path))
            except PermissionError:
                raise RuntimeError(f"{Path(out_path).name} is open in another program (e.g. a PDF viewer), so it "
                                   "can't be saved. Close it there, then try again.") from None
    finally:
        shutil.rmtree(work, ignore_errors=True)


# ============================================================ PDFs made elsewhere
# Proxy PDFs from other programs (Proxy PDF Maker with other settings, MPC Autofill, ...) almost always
# place every card as its own image. read_foreign() follows each page's drawing commands to find
# those images and turns them into a normal Fleem deck, so they can be edited, combined and saved as a
# Fleem PDF. PDFs whose pages are one flat picture can't be taken apart (they're printed as they are).
FOREIGN_MAX_BLEED = 4.5 * PT_PER_MM      # a card's visible area may have up to this much bleed
CARD_TOLERANCE = 1.6                     # pt: 63 x 88 mm and 2.48 x 3.46 in both count as a card


def _matmul(m, n):
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E, e * B + f * D + F)


def _rect_box(x, y, w, h, ctm):
    """A rectangle in user space -> its bounding box (x0, y0, x1, y1) on the page."""
    a, b, c, d, e, f = ctm
    pts = [(px * a + py * c + e, px * b + py * d + f) for px, py in ((x, y), (x + w, y), (x, y + h), (x + w, y + h))]
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


def _intersect(a, b):
    if a is None:
        return b
    if b is None:
        return a
    return max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])


def _page_images(reader, page):
    """Every image drawn on a page: [{"box": (x0, y0, x1, y1) where the image is drawn, "clip": visible box
    or None, "ref", "obj", "rotated"}], following q/Q, cm, clipping and nested Form XObjects."""
    from pypdf.generic import ContentStream
    found = []

    def walk(content, resources, ctm, depth, clip=None):
        try:
            resources = resources.get_object() if resources is not None else {}
            xobjects = resources.get("/XObject")
            xobjects = xobjects.get_object() if xobjects is not None else {}
        except Exception:
            xobjects = {}
        stack, path_box, pending = [], None, False
        for operands, op in content.operations:
            if op == b"q":
                stack.append((ctm, clip))
            elif op == b"Q":
                if stack:
                    ctm, clip = stack.pop()
            elif op == b"cm" and len(operands) == 6:
                ctm = _matmul(tuple(float(v) for v in operands), ctm)
            elif op == b"re" and len(operands) == 4:
                box = _rect_box(*(float(v) for v in operands), ctm)
                path_box = box if path_box is None else (min(path_box[0], box[0]), min(path_box[1], box[1]),
                                                         max(path_box[2], box[2]), max(path_box[3], box[3]))
            elif op in (b"W", b"W*"):
                pending = True
            elif op in (b"n", b"f", b"F", b"f*", b"S", b"s", b"B", b"B*", b"b", b"b*"):
                if pending and path_box is not None:
                    clip = _intersect(clip, path_box)
                path_box, pending = None, False
            elif op in (b"m", b"l", b"c", b"v", b"y", b"h"):
                if op == b"m" and not pending:
                    path_box = None               # a drawn line: not a clipping rectangle
            elif op == b"Do" and operands:
                name = operands[0]
                if name not in xobjects:
                    continue
                ref = xobjects.raw_get(name)
                obj = ref.get_object()
                kind = obj.get("/Subtype")
                if kind == "/Image":
                    a, b, c, d = ctm[:4]
                    found.append({"box": _rect_box(0, 0, 1, 1, ctm), "clip": clip, "ref": ref, "obj": obj,
                                  "rotated": abs(b) > 1e-6 or abs(c) > 1e-6 or a <= 0 or d <= 0})
                elif kind == "/Form" and depth < 4:
                    m = tuple(float(v) for v in obj.get("/Matrix", [1, 0, 0, 1, 0, 0]))
                    sub = ContentStream(obj, reader)
                    walk(sub, obj.get("/Resources", resources), _matmul(m, ctm), depth + 1, clip)
                    # (a form's own clipping ends with it, like a q/Q pair)

    content = page.get_contents()
    if content is not None:
        walk(content, page.get("/Resources"), (1, 0, 0, 1, 0, 0), 0)
    return found


def _paper_name(w, h):
    for name, (pw, ph) in {"A4": (595.3, 841.9), "Letter": (612, 792), "Legal": (612, 1008), "A3": (841.9, 1190.6)}.items():
        if abs(w - pw) < 3 and abs(h - ph) < 3 or abs(w - ph) < 3 and abs(h - pw) < 3:
            return name
    return f"{w / PT_PER_MM:.0f} x {h / PT_PER_MM:.0f} mm"


def read_foreign(path: Path, work: Path | None = None, pages_mode: str | None = None):
    """Take the cards out of a proxy PDF made by another program. pages_mode: None = work out which pages
    are backs, "pairs" = every second page is the back of the one before, "single" = no back pages.
    With `work`, card images that aren't already MPC-shaped are cut out and saved there (in parallel);
    without it only the layout is read (fast, for counting). Returns {"copies", "bleed_mm",
    "spacing_mm", "summary", "has_backs"} like read_layout, or None if the cards can't be taken out."""
    from pypdf import PdfReader
    try:
        reader = PdfReader(str(path))
        pages = []
        for page in reader.pages:
            if int(page.get("/Rotate", 0) or 0) % 360:
                log.info("%s: rotated pages aren't supported", Path(path).name)
                return None
            mb = page.mediabox
            left, bottom = float(mb.left), float(mb.bottom)
            pw, ph = float(mb.width), float(mb.height)
            cards = []
            for im in _page_images(reader, page):
                x0, y0, x1, y1 = im["box"]
                if x1 - x0 < 40 and y1 - y0 < 40:
                    continue                           # small icons, logos
                if im["rotated"]:
                    log.info("%s: a card is drawn turned sideways - not supported", Path(path).name)
                    return None
                vis = _intersect(im["box"], im["clip"])
                vw, vh = vis[2] - vis[0], vis[3] - vis[1]
                bx, by = (vw - CARD_W) / 2, (vh - CARD_H) / 2
                if not (-CARD_TOLERANCE <= bx <= FOREIGN_MAX_BLEED and -CARD_TOLERANCE <= by <= FOREIGN_MAX_BLEED
                        and abs(bx - by) <= CARD_TOLERANCE):
                    if vw > 0 and vh > 0 and vw * vh < CARD_W * CARD_H / 4:
                        continue                       # something small that isn't a card
                    log.info("%s: page %d has an image that isn't a card (%.0f x %.0f pt)", Path(path).name,
                             len(pages) + 1, vw, vh)
                    return None
                cx = (vis[0] + vis[2]) / 2 - left       # the card's centre, on the page
                cy = (vis[1] + vis[3]) / 2 - bottom
                cards.append({"cx": cx, "cy": cy, "bleed": max(0.0, (bx + by) / 2), "vw": vw, "vh": vh,
                              "box": (x0 - left, y0 - bottom, x1 - left, y1 - bottom), "ref": im["ref"], "obj": im["obj"]})
            pages.append({"w": pw, "h": ph, "cards": cards})
        if not any(p["cards"] for p in pages):
            return None

        # which pages are backs
        def mirrored_match(front, back):
            """For each card on `back`, the index of the front card it's the back of (mirrored), or None."""
            out = []
            for c in back["cards"]:
                mx = back["w"] - c["cx"]
                out.append(next((i for i, f in enumerate(front["cards"])
                                 if abs(f["cx"] - mx) < 2 and abs(f["cy"] - c["cy"]) < 2), None))
            return out

        def looks_like_back(front, back):
            if not back["cards"]:
                return True                            # an empty page after a front: blank backs
            if len(back["cards"]) > len(front["cards"]) or None in mirrored_match(front, back):
                return False
            rows = {}
            for c in back["cards"]:
                rows.setdefault(round(c["cy"] / 5), []).append(c["cx"])
            multi = [xs for xs in rows.values() if len(xs) > 1]
            if multi:                                  # backs are drawn right to left (mirrored order)
                return all(all(a > b for a, b in zip(xs, xs[1:])) for xs in multi)
            same = all(any(abs(f["cx"] - c["cx"]) < 2 and abs(f["cy"] - c["cy"]) < 2 for f in front["cards"])
                       for c in back["cards"])
            return not same

        kinds, i = [], 0
        while i < len(pages):
            nxt = pages[i + 1] if i + 1 < len(pages) else None
            if pages_mode == "single" or nxt is None or not pages[i]["cards"]:
                pair = False
            elif pages_mode == "pairs":
                pair = True
            else:
                pair = looks_like_back(pages[i], nxt)
            kinds.append((i, i + 1 if pair else None))
            i += 2 if pair else 1

        # card images: used as they are when already MPC-shaped (card + 0.12 in), otherwise cut out
        jobs = {}

        def src_for(c):
            x0, y0, x1, y1 = c["box"]
            tx0, ty0 = c["cx"] - CARD_W / 2 - IMAGE_BLEED, c["cy"] - CARD_H / 2 - IMAGE_BLEED
            tx1, ty1 = tx0 + CARD_W + 2 * IMAGE_BLEED, ty0 + CARD_H + 2 * IMAGE_BLEED
            obj = c["obj"]
            plain_jpeg = (str(obj.get("/Filter")) == "/DCTDecode" and "/SMask" not in obj and "/Decode" not in obj
                          and "/Mask" not in obj and not obj.get("/ImageMask"))
            idnum = getattr(c["ref"], "idnum", None) or id(obj)
            if plain_jpeg and max(abs(x0 - tx0), abs(y0 - ty0), abs(x1 - tx1), abs(y1 - ty1)) < 0.8:
                return ((str(path), idnum), c["ref"])
            # where the MPC-shaped area lies in the image, as fractions of the image (top-left origin)
            iw, ih = x1 - x0, y1 - y0
            frac = (round((tx0 - x0) / iw, 4), round((y1 - ty1) / ih, 4), round((tx1 - x0) / iw, 4),
                    round((y1 - ty0) / ih, 4))
            key = (idnum, frac)
            if key not in jobs:
                jobs[key] = {"obj": obj, "frac": frac, "file": None if work is None else
                             Path(work) / f"{Path(path).stem[:30]}-{len(jobs)}.jpg"}
            return ("file", jobs[key]["file"]) if work is not None else ("pending", key)

        copies, bleeds, xs, ys, cells = [], [], set(), set(), []
        for fi, bi in kinds:
            front, back = pages[fi], pages[bi] if bi is not None else None
            backs = {}
            if back:
                for c, idx in zip(back["cards"], mirrored_match(front, back)):
                    if idx is not None:
                        backs[idx] = c
            for idx, c in enumerate(front["cards"]):
                bleeds.append(c["bleed"])
                xs.add(round(c["cx"], 1)); ys.add(round(c["cy"], 1)); cells.append((c["vw"], c["vh"]))
                copies.append({"front": src_for(c), "back": src_for(backs[idx]) if idx in backs else None,
                               "name": None, "sort": None})

        if work is not None and jobs:
            Path(work).mkdir(parents=True, exist_ok=True)
            with ThreadPoolExecutor(max_workers=4) as pool:
                list(pool.map(lambda j: _cut_card_image(j["obj"], j["frac"], j["file"]), jobs.values()))

        bleeds.sort()
        bleed_mm = bleeds[len(bleeds) // 2] / PT_PER_MM
        cw = sum(w for w, _ in cells) / len(cells)
        ch = sum(h for _, h in cells) / len(cells)
        spacing_mm = None
        sx, sy = sorted(xs), sorted(ys)
        if len(sx) > 1:
            spacing_mm = max(0.0, (min(b - a for a, b in zip(sx, sx[1:])) - cw) / PT_PER_MM)
        elif len(sy) > 1:
            spacing_mm = max(0.0, (min(b - a for a, b in zip(sy, sy[1:])) - ch) / PT_PER_MM)
        has_backs = any(bi is not None for _, bi in kinds)
        n_pages = len(pages)
        summary = (f"{len(copies)} card{'s' * (len(copies) != 1)} on {n_pages} {_paper_name(pages[0]['w'], pages[0]['h'])} "
                   f"page{'s' * (n_pages != 1)}, " + ("with back pages" if has_backs else "no back pages"))
        return {"copies": copies, "bleed_mm": round(bleed_mm, 2) + 0.0,
                "paper": paper_named(pages[0]["w"], pages[0]["h"]),
                "spacing_mm": None if spacing_mm is None else round(spacing_mm, 2) + 0.0,
                "summary": summary, "has_backs": has_backs, "foreign": True}
    except Exception as e:
        log.info("Couldn't take the cards out of %s: %s", Path(path).name, e)
        return None


_DECODE_LOCK = threading.Lock()


def _extend_edges(im, left: int, top: int, right: int, bottom: int):
    """Grow an image by stretching its outermost pixels outwards (sides, then top/bottom incl. corners)."""
    from PIL import Image
    if not (left or top or right or bottom):
        return im
    w, h = im.size
    out = Image.new("RGB", (w + left + right, h + top + bottom), "white")
    out.paste(im, (left, top))
    if left:
        out.paste(im.crop((0, 0, 1, h)).resize((left, h)), (0, top))
    if right:
        out.paste(im.crop((w - 1, 0, w, h)).resize((right, h)), (left + w, top))
    W = out.width
    if top:
        out.paste(out.crop((0, top, W, top + 1)).resize((W, top)), (0, 0))
    if bottom:
        out.paste(out.crop((0, top + h - 1, W, top + h)).resize((W, bottom)), (0, top + h))
    return out


def _cut_card_image(obj, frac, out: Path) -> None:
    """Save the MPC-shaped part of a PDF image (frac = left, top, right, bottom as fractions of the image,
    may reach outside it: that part gets stretched-edge bleed) as a JPEG, at most MAX_DPI."""
    from PIL import Image
    import io
    import math
    if str(obj.get("/Filter")) == "/DCTDecode" and "/SMask" not in obj and "/Decode" not in obj:
        im = Image.open(io.BytesIO(obj._data))
        im.load()
    else:
        with _DECODE_LOCK:                    # pypdf decoding isn't thread-safe
            im = obj.decode_as_image()
    if im.mode in ("RGBA", "LA", "P", "PA"):
        im = im.convert("RGBA")
        bg = Image.new("RGBA", im.size, "white")
        bg.alpha_composite(im)
        im = bg
    im = im.convert("RGB")
    w, h = im.size
    l, t, r, b = frac[0] * w, frac[1] * h, frac[2] * w, frac[3] * h
    pl, pt = math.ceil(max(0, -l)), math.ceil(max(0, -t))
    pr, pb = math.ceil(max(0, r - w)), math.ceil(max(0, b - h))
    im = _pad_smart(im, pl, pt, pr, pb)       # border colour, or mirrored picture, past the image's own bleed
    im = im.crop((round(l + pl), round(t + pt), round(r + pl), round(b + pt)))
    max_w = round((CARD_W + 2 * IMAGE_BLEED) / 72 * MAX_DPI)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.LANCZOS)
    im.save(out, "JPEG", quality=JPEG_QUALITY, subsampling=0)


def read_any_pdf(path: Path, work: Path | None = None, pages_mode: str | None = None):
    """Fleem's own layout (read_layout), else the cards taken out of a PDF made elsewhere, else None."""
    return read_layout(path) or read_foreign(path, work, pages_mode)


# ============================================================ reading card names off the cards
# PDFs made elsewhere (and older Fleem PDFs) don't say which card is which. Fleem reads the name off
# each card with Windows' built-in text recognition (the same as the Snipping Tool's "copy text"):
# the title line at the top of the card - or, when there's a smaller line under the title (custom
# cards: "Willow, Inspiring Leader" over "Rick, Steadfast Leader"), that line, since it's the real
# card's name. What it reads is matched against Scryfall's list of real card names to fix misreads.
CARD_NAMES_FILE = HERE / "card_names.json"
SCRYFALL_NAMES_URL = "https://api.scryfall.com/catalog/card-names"
_catalog_cache = {}


def card_name_catalog() -> dict:
    """Real card names (from Scryfall, kept next to Fleem and refreshed monthly): normalised -> name.
    Double-faced cards are listed by each face too. {} if it can't be had (then names are used as read)."""
    if "names" in _catalog_cache:
        return _catalog_cache["names"]
    names = []
    try:
        saved = json.loads(CARD_NAMES_FILE.read_text(encoding="utf-8"))
        names, fetched = saved.get("names", []), saved.get("fetched", 0)
    except Exception:
        fetched = 0
    if not names or time.time() - fetched > 30 * 86400:
        try:
            req = urllib.request.Request(SCRYFALL_NAMES_URL, headers={"User-Agent": "FleemPdfMaker/1.0",
                                                                        "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=20) as r:
                names = json.loads(r.read().decode("utf-8")).get("data", []) or names
            CARD_NAMES_FILE.write_text(json.dumps({"fetched": time.time(), "names": names}), encoding="utf-8")
        except Exception as e:
            log.info("Couldn't get the card name list from Scryfall (%s) - names are used as read", e)
    catalog = {}
    for full in names:
        for face in [full] + (full.split(" // ") if " // " in full else []):
            catalog.setdefault(_norm_name(face), face)
    _catalog_cache["names"] = catalog
    return catalog


def _norm_name(text: str) -> str:
    import unicodedata
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii").lower()
    text = re.sub(r"[^a-z0-9 ]+", " ", text.replace("'", ""))
    return re.sub(r"\s+", " ", text).strip()


def rescue_card_name(text: str, catalog: dict):
    """A partly read title ("azikar, the Gye"): when its long words (6+ letters) appear in exactly one real
    card name, that's the card. None if they fit no name or several."""
    words = [w for w in _norm_name(text).split() if len(w) >= 4]
    if not catalog or not any(len(w) >= 6 for w in words):
        return None
    hits = [k for k in catalog if all(w in k for w in words)]
    return catalog[hits[0]] if len(hits) == 1 else None


def match_card_name(text: str, catalog: dict, cutoff: float = 0.84):
    """The real card name in `text` (a misread is fine, and so are stray bits read before or after it),
    or None if nothing is close enough."""
    tokens = text.split()
    windows = [" ".join(tokens[a:len(tokens) - b]) for a in range(0, min(3, len(tokens)))
               for b in range(0, min(4, len(tokens) - a))]
    windows = [w for w in windows if len(_norm_name(w)) >= 3]
    for w in windows:                                # exact (after tidying) first
        if _norm_name(w) in catalog:
            return catalog[_norm_name(w)]
    for w in windows[:8]:
        found = _match_one(w, catalog, cutoff)
        if found:
            return found
    return None


def _match_one(text: str, catalog: dict, cutoff: float):
    import difflib
    key = _norm_name(text)
    if len(key) < 5 or not catalog:                  # short bits only count when they match exactly
        return None
    if key in catalog:
        return catalog[key]
    bucket = _catalog_cache.setdefault("buckets", {})
    if not bucket:
        for k in catalog:
            bucket.setdefault((k[:1], len(k) // 3), []).append(k)
    best, best_ratio = None, cutoff
    sizes = range(len(key) // 3 - 1, len(key) // 3 + 2)
    for firsts in ([key[:1]], "abcdefghijklmnopqrstuvwxyz0123456789"):   # same first letter, then any (misread)
        for first in firsts:
            for size in sizes:
                for k in bucket.get((first, size), ()):
                    m = difflib.SequenceMatcher(None, key, k)
                    if m.real_quick_ratio() <= best_ratio or m.quick_ratio() <= best_ratio:
                        continue
                    r = m.ratio()
                    if r > best_ratio:
                        best, best_ratio = k, r
        if best:
            break
    return catalog[best] if best else None


def windows_ocr_lines(image):
    """Windows 10/11's built-in text recognition: [(text, x, y, width, height)] for each line of text in
    a PIL image, or None where it isn't available."""
    if os.name != "nt":
        return None
    try:
        import asyncio
        import ctypes
        import io
        try:
            from winrt.windows.media.ocr import OcrEngine
            from winrt.windows.graphics.imaging import BitmapDecoder
            from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter
            from winrt.windows.globalization import Language
        except ImportError:
            from winsdk.windows.media.ocr import OcrEngine
            from winsdk.windows.graphics.imaging import BitmapDecoder
            from winsdk.windows.storage.streams import InMemoryRandomAccessStream, DataWriter
            from winsdk.windows.globalization import Language
        try:
            ctypes.windll.combase.RoInitialize(1)        # this thread may call Windows' new-style APIs
        except Exception:
            pass
        buf = io.BytesIO()
        image.convert("RGB").save(buf, "BMP")
        data = buf.getvalue()

        async def run():
            stream = InMemoryRandomAccessStream()
            writer = DataWriter(stream)
            writer.write_bytes(data)
            await writer.store_async()
            writer.detach_stream()
            stream.seek(0)
            bitmap = await (await BitmapDecoder.create_async(stream)).get_software_bitmap_async()
            engine = OcrEngine.try_create_from_language(Language("en-US")) or \
                OcrEngine.try_create_from_user_profile_languages()
            if engine is None:
                return None
            result = await engine.recognize_async(bitmap)
            lines = []
            for line in result.lines:
                rects = [w.bounding_rect for w in line.words]
                if not rects:
                    continue
                x0, y0 = min(r.x for r in rects), min(r.y for r in rects)
                x1, y1 = max(r.x + r.width for r in rects), max(r.y + r.height for r in rects)
                lines.append((line.text, x0, y0, x1 - x0, y1 - y0))
            return lines
        return asyncio.run(run())
    except Exception as e:
        if not _catalog_cache.get("ocr_warned"):
            _catalog_cache["ocr_warned"] = True
            log.info("Windows text recognition isn't available (%s) - card names can't be read", e)
        return None


OCR_BACKEND = windows_ocr_lines                    # (tests swap in another reader)
BASIC_LANDS = {"plains": "Plains", "island": "Island", "swamp": "Swamp", "mountain": "Mountain", "forest": "Forest",
               "wastes": "Wastes"}


def read_card_name(image_bytes: bytes, catalog: dict | None = None):
    """Read a card's name off its picture (an MPC-shaped image: card + 0.12 in bleed). The smaller line
    under the title wins when there is one (custom cards put the real name there). Returns the name
    (a real card name when one is close enough, else the text as read), or None."""
    from PIL import Image, ImageOps
    import io
    catalog = card_name_catalog() if catalog is None else catalog
    im = Image.open(io.BytesIO(image_bytes))
    im.draft("RGB", (1200, 1600))
    im = im.convert("RGB")
    W, H = im.size
    bx = W * IMAGE_BLEED / (CARD_W + 2 * IMAGE_BLEED)    # where the card itself starts in the image
    by = H * IMAGE_BLEED / (CARD_H + 2 * IMAGE_BLEED)
    cw, ch = W - 2 * bx, H - 2 * by
    band = im.crop((round(bx + cw * 0.04), round(by + ch * 0.015), round(bx + cw * 0.80), round(by + ch * 0.20)))
    scale = 1500 / band.width
    band = band.resize((1500, max(1, round(band.height * scale))), Image.LANCZOS)
    name = _name_from_band(band, catalog)
    if name:
        return name
    # nothing at the top: full-art lands and some frames have the name at the bottom (real names only)
    low = im.crop((round(bx + cw * 0.04), round(by + ch * 0.70), round(bx + cw * 0.96), round(by + ch * 0.96)))
    low = low.resize((1500, max(1, round(low.height * 1500 / low.width))), Image.LANCZOS)
    for pic in (low, ImageOps.invert(ImageOps.autocontrast(low.convert("L"))).convert("RGB")):
        for line in sorted(OCR_BACKEND(pic) or [], key=lambda l: l[2]):
            # the whole line must be a name (rules text is full of words that are also card names)
            words = [w for w in line[0].split() if _letters(w) >= 2]
            if not words or len(words) > 4:
                continue
            key = _norm_name(" ".join(words))
            name = catalog.get(key) if catalog else None
            real_words = [w for w in _norm_name(" ".join(words)).split() if len(w) >= 3]
            if not name and len(real_words) <= 3:    # full-art basic lands: "Mountain" (plus specks)
                name = next((BASIC_LANDS[w] for w in real_words if w in BASIC_LANDS), None)
            if name:
                return name
    return None


def _name_from_band(band, catalog: dict):
    """The card's name from the strip at the top of the card (title, and the smaller line under it);
    if that finds nothing, just the top half of the strip (the title bar, without the artwork)."""
    return _name_from_strip(band, catalog) or _name_from_strip(band.crop((0, 0, band.width, band.height // 2)), catalog)


def _name_from_strip(band, catalog: dict):
    from PIL import Image, ImageOps
    # read it as it is, with the colours turned around (light text on dark bars), and by whiteness
    # (white text on gold/coloured bars), and merge what's found
    lines = list(OCR_BACKEND(band) or [])
    for pic in _ocr_variants(band):
        lines += [l for l in (OCR_BACKEND(pic) or [])
                  if not any(_overlaps(l, m) and _letters(m[0]) >= _letters(l[0]) for m in lines)]
    lines = [l for l in lines if not any(m is not l and _overlaps(l, m) and _letters(m[0]) > _letters(l[0])
                                         for m in lines)]
    if not lines:
        return None
    lines = [(re.sub(r"\s+", " ", t).strip(" .,:;|-_~*'\"`!"), x, y, w, h) for t, x, y, w, h in lines]
    lines = [l for l in lines if sum(c.isalpha() for c in l[0]) >= 3 and l[4] > band.height * 0.04]
    if not lines:
        return None
    lines.sort(key=lambda l: l[2])
    upper = [l for l in lines if l[2] < band.height * 0.62] or lines
    wordy = [l for l in upper if _wordiness(l[0]) >= 5] or upper       # real words, not specks of artwork
    title = max(wordy, key=lambda l: l[4])              # the biggest such text near the top is the title
    below = [l for l in lines if l[2] > title[2] + title[4] * 0.6
             and l[2] - (title[2] + title[4]) < title[4] * 2.4 and l[4] <= title[4] * 0.92]
    subtitle = below[0] if below else None
    if subtitle:
        name = match_card_name(subtitle[0], catalog)
        if name:
            return name
    if catalog and not match_card_name(title[0], catalog):
        # the title isn't a real card (a custom name): look closely just under it for the real name
        y0 = title[2] + title[4] * 0.8
        y1 = min(band.height, title[2] + title[4] * 3.4)
        if y1 - y0 > 10:
            under = band.crop((0, round(y0), band.width, round(y1)))
            under = under.resize((under.width * 2, under.height * 2), Image.LANCZOS)
            for im in [under] + _ocr_variants(under):
                for line in sorted(OCR_BACKEND(im) or [], key=lambda l: l[2]):
                    name = match_card_name(re.sub(r"\s+", " ", line[0]), catalog, cutoff=0.76)
                    if name:
                        return name
    name = match_card_name(title[0], catalog)
    if name:
        return name
    for line in ([subtitle] if subtitle else []) + [title]:                 # partly read titles
        name = rescue_card_name(line[0], catalog)
        if name:
            return name
    for line in [l for l in lines if l is not title and l is not subtitle]:   # top to bottom, stricter
        name = match_card_name(line[0], catalog, cutoff=0.9)
        if name:
            return name
    for line in ([subtitle] if subtitle else []) + [title]:
        raw = _tidy_read_name(line[0])
        if raw:
            return raw
    return None


def _ocr_variants(im):
    """Extra versions of a strip of card to read: the colours turned around (light text on dark), and
    'whiteness' (the darkest of R, G, B - white text stands out from gold or coloured bars), as dark
    text on light."""
    from PIL import ImageOps, ImageChops
    inverted = ImageOps.invert(ImageOps.autocontrast(im.convert("L"))).convert("RGB")
    r, g, b = im.convert("RGB").split()
    white = ImageChops.darker(ImageChops.darker(r, g), b)
    whiteness = ImageOps.invert(ImageOps.autocontrast(white, cutoff=2)).convert("RGB")
    return [inverted, whiteness]


def _wordiness(text: str) -> int:
    """Letters in the words of 3+ letters - how much of a read line looks like real words."""
    return sum(len(w) for w in re.findall(r"[A-Za-z']{3,}", text))


def _letters(text: str) -> int:
    return sum(c.isalpha() for c in text)


def _overlaps(a, b) -> bool:
    """Two read lines (text, x, y, w, h) covering mostly the same spot."""
    ox = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    oy = min(a[2] + a[4], b[2] + b[4]) - max(a[2], b[2])
    return ox > 0.5 * min(a[3], b[3]) and oy > 0.5 * min(a[4], b[4])


def _tidy_read_name(text: str):
    """Text read off a card, if it looks like a name (not specks of artwork): stray bits dropped."""
    words = [w for w in text.split() if _letters(w) >= 2 or w.lower() in ("a", "i", "o", "of")]
    while words and _letters(words[0]) < 2:
        words.pop(0)
    while words and _letters(words[-1]) < 2:
        words.pop()
    small = {"of", "the", "a", "an", "to", "in", "and", "from", "with", "for", "on", "at", "into", "by", "as", "or"}
    while len(words) > 1 and len(words[-1]) <= 3 and words[-1].lower() not in small and (
            words[-1].isupper() or not words[-1][:1].isupper()):
        words.pop()                                  # stray bits at the end ("Grave Pact OC")
    name = " ".join(words).strip(" ,.-")
    if name.isupper() and _letters(name) >= 4:       # ALL-CAPS titles (tokens): "GOBLIN" -> "Goblin"
        name = " ".join(w if w.lower() not in small else w.lower() for w in name.title().split())
    if any(re.search(r"[a-z][A-Z]", w) and not re.match(r"(Mc|Mac|O')[A-Z]", w) for w in name.split()):
        return None                                  # "GosLIN": a misread, not a name
    if re.search(r"[^A-Za-z0-9 ,'’\-]", name):
        return None                                  # names don't have ! ? : ; | ( ) and the like
    if not name.isupper() and any(w.isupper() and len(w) > 1 for w in name.split()):
        return None                                  # an ALL-CAPS word among normal ones: specks
    words = name.replace(",", " ").split()
    if _letters(name) < 5 or not name[:1].isupper() or not words:
        return None
    if sum(c.isalpha() or c in " ,'-" for c in name) < 0.9 * len(name):
        return None
    capitals = sum(1 for w in words if w[:1].isupper() and not w.isupper())
    if any(not (w[:1].isupper() or w.lower() in small) for w in words) or capitals < 0.6 * len(words):
        return None
    if sum(len(w) for w in words) / len(words) < 3.2:
        return None
    return name


# ============================================================ look: the RPG ("Adventurer") theme art
# Everything is drawn here as tiny pixel art and scaled up with hard edges, so no image files are needed:
# scroll buttons (parchment between two wooden rollers), wax-seal checkboxes, gem radio buttons, wooden
# scrollbars, a potion progress bar, a ribbon banner and the parchment grain behind windows.
RPG_INK = (74, 28, 30)                  # outlines: deep maroon, like the reference pixel art
RPG_PAPER = {"paper": (243, 205, 152), "hi": (251, 226, 186), "lo": (222, 170, 116)}
RPG_ROLLERS = {                         # (wood, wood highlight, knob, knob highlight)
    "plain": ((150, 88, 52), (196, 128, 78), (178, 72, 60), (220, 124, 100)),
    "royal": ((140, 40, 34), (192, 72, 54), (214, 160, 46), (250, 214, 112)),
    "charred": ((66, 50, 44), (102, 78, 68), (204, 86, 44), (240, 140, 80)),
}


def _rgb_shift(c, amount):
    return tuple(max(0, min(255, round(v + amount))) for v in c)


def _rgb_mix(c, other, t):
    return tuple(round(a + (b - a) * t) for a, b in zip(c, other))


def _pixels(w, h):
    from PIL import Image
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    return im, im.load()


def rpg_scroll(variant="plain", state="normal", fade_to=(241, 223, 180)):
    """A scroll button: parchment between two wooden rollers with knobs. 24 x 16 pixels; the middle
    stretches (see RPG_SCROLL_BORDER)."""
    W, H = 24, 16
    im, px = _pixels(W, H)
    wood, wood_hi, knob, knob_hi = RPG_ROLLERS[variant]
    paper, hi, lo = RPG_PAPER["paper"], RPG_PAPER["hi"], RPG_PAPER["lo"]
    if variant == "royal":
        paper, hi = _rgb_shift(paper, 6), _rgb_shift(hi, 4)
    shift = {"normal": 0, "active": 12, "pressed": -16, "disabled": 0}[state]
    paper, hi, lo = (_rgb_shift(c, shift) for c in (paper, hi, lo))
    wood, wood_hi, knob, knob_hi = (_rgb_shift(c, shift * 0.6) for c in (wood, wood_hi, knob, knob_hi))
    ink = RPG_INK

    def put(x, y, c):
        if state == "disabled":
            c = _rgb_mix(c, fade_to, 0.55)
        px[x, y] = c + (255,)
    for x in range(4, W - 4):                         # the parchment
        put(x, 2, ink)
        put(x, 13, ink)
        put(x, 3, hi)
        for y in range(4, 12):
            put(x, y, paper)
        put(x, 12, lo)
    for y in range(4, 12):                            # its stitched edges by the rollers
        for x in (4, 5, W - 6, W - 5):
            if (x + y) % 2 == 0:
                put(x, y, lo)
    for a in (0, W - 4):                              # the two rollers
        for x in (a + 1, a + 2):
            put(x, 0, ink)
            put(x, 15, ink)
        for y in (1, 14):
            put(a, y, ink)
            put(a + 1, y, knob_hi)
            put(a + 2, y, knob)
            put(a + 3, y, ink)
        for x in range(a, a + 4):
            put(x, 2, ink)
            put(x, 13, ink)
        for y in range(3, 13):
            put(a, y, ink)
            put(a + 1, y, wood_hi)
            put(a + 2, y, wood)
            put(a + 3, y, ink)
    return im


RPG_SCROLL_BORDER = (7, 4, 7, 4)                     # left, top, right, bottom (unscaled pixels)


def rpg_tab(on: bool, state="normal", fade_to=(241, 223, 180)):
    """A tab (Fronts | Backs): a little banner, parchment when on, leather when off. 20 x 14, middle stretches."""
    W, H = 20, 14
    im, px = _pixels(W, H)
    face = RPG_PAPER["paper"] if on else (196, 150, 100)
    face_hi = RPG_PAPER["hi"] if on else (214, 172, 122)
    edge = (214, 160, 46) if on else (150, 104, 66)
    if state == "active":
        face, face_hi = _rgb_shift(face, 10), _rgb_shift(face_hi, 8)

    def put(x, y, c):
        if state == "disabled":
            c = _rgb_mix(c, fade_to, 0.55)
        px[x, y] = c + (255,)
    for x in range(1, W - 1):
        put(x, 0, RPG_INK)
        put(x, H - 1, RPG_INK)
        put(x, 1, edge)
        put(x, H - 2, edge)
        put(x, 2, face_hi)
        for y in range(3, H - 2):
            put(x, y, face)
    for y in range(1, H - 1):
        put(0, y, RPG_INK)
        put(W - 1, y, RPG_INK)
        put(1, y, edge)
        put(W - 2, y, edge)
    return im


RPG_TAB_BORDER = (3, 3, 3, 3)


def rpg_checkbox(checked: bool, disabled=False, fade_to=(241, 223, 180)):
    """A checkbox: a little wooden frame; ticked, it carries a red wax seal. 11 x 11."""
    im, px = _pixels(11, 11)
    wood, field = (150, 88, 52), (251, 236, 205)
    seal, seal_hi, seal_lo = (176, 44, 36), (222, 92, 72), (122, 26, 24)

    def put(x, y, c):
        if disabled:
            c = _rgb_mix(c, fade_to, 0.55)
        px[x, y] = c + (255,)
    for i in range(1, 10):
        for j in (0, 10):
            put(i, j, RPG_INK)
            put(j, i, RPG_INK)
    for i in range(1, 10):
        for j in (1, 9):
            put(i, j, wood)
            put(j, i, wood)
    for x in range(2, 9):
        for y in range(2, 9):
            put(x, y, field)
    if checked:
        blob = [(4, 2), (5, 2), (6, 2), (3, 3), (4, 3), (5, 3), (6, 3), (7, 3)] + \
               [(x, y) for y in (4, 5, 6) for x in range(2, 9)] + \
               [(3, 7), (4, 7), (5, 7), (6, 7), (7, 7), (4, 8), (5, 8), (6, 8)]
        for x, y in blob:
            put(x, y, seal)
        for x, y in ((3, 7), (4, 8), (7, 6), (6, 8), (7, 7), (8, 6)):
            put(x, y, seal_lo)
        for x, y in ((4, 3), (3, 4), (5, 3)):
            put(x, y, seal_hi)
        for x, y in ((3, 5), (4, 6), (5, 5), (6, 4), (7, 3)):    # a tick pressed into the wax
            put(x, y, (250, 214, 112))
    return im


def rpg_radio(selected: bool, disabled=False, fade_to=(241, 223, 180)):
    """A radio button: a gem socket; picked, it holds a glowing emerald. 11 x 11."""
    im, px = _pixels(11, 11)
    rim, socket = (150, 88, 52), (84, 52, 36)
    gem, gem_hi, gem_lo = (64, 168, 92), (170, 236, 170), (30, 104, 60)

    def put(x, y, c):
        if disabled:
            c = _rgb_mix(c, fade_to, 0.55)
        px[x, y] = c + (255,)
    outline = [(3, 0), (4, 0), (5, 0), (6, 0), (7, 0), (2, 1), (8, 1), (1, 2), (9, 2), (0, 3), (10, 3), (0, 4), (10, 4),
               (0, 5), (10, 5), (0, 6), (10, 6), (0, 7), (10, 7), (1, 8), (9, 8), (2, 9), (8, 9), (3, 10), (4, 10),
               (5, 10), (6, 10), (7, 10)]
    for x, y in outline:
        put(x, y, RPG_INK)
    for y in range(1, 10):
        for x in range(1, 10):
            if (x, y) in outline:
                continue
            inside = (x - 5) ** 2 + (y - 5) ** 2
            if inside <= 26:
                put(x, y, rim if inside > 12 else socket)
    if selected:
        for y in range(3, 8):
            for x in range(3, 8):
                if (x - 5) ** 2 + (y - 5) ** 2 <= 5:
                    put(x, y, gem)
        for x, y in ((6, 6), (6, 7), (7, 6), (5, 7), (7, 5)):
            put(x, y, gem_lo)
        for x, y in ((4, 4), (4, 3), (3, 4)):
            put(x, y, gem_hi)
    return im


def rpg_rod(state="normal"):
    """A scrollbar thumb: a wooden rod with knobs at both ends. 7 x 12, the middle stretches."""
    W, H = 7, 12
    im, px = _pixels(W, H)
    wood, wood_hi, wood_lo = (150, 88, 52), (196, 128, 78), (112, 62, 38)
    knob, knob_hi = (178, 72, 60), (220, 124, 100)
    if state == "active":
        wood, wood_hi, knob = _rgb_shift(wood, 18), _rgb_shift(wood_hi, 14), _rgb_shift(knob, 14)

    def put(x, y, c):
        px[x, y] = c + (255,)
    for y in (0, H - 1):
        for x in range(1, W - 1):
            put(x, y, RPG_INK)
    for y in (1, H - 2):
        put(0, y, RPG_INK)
        put(W - 1, y, RPG_INK)
        for x in range(1, W - 1):
            put(x, y, knob_hi if x < 3 else knob)
    for x in range(W):
        put(x, 2, RPG_INK)
        put(x, H - 3, RPG_INK)
    for y in range(3, H - 3):
        put(0, y, RPG_INK)
        put(W - 1, y, RPG_INK)
        put(1, y, wood_hi)
        for x in (2, 3, 4):
            put(x, y, wood)
        put(5, y, wood_lo)
    return im


RPG_ROD_BORDER = (0, 3, 0, 3)


def rpg_potion_bar():
    """The progress bar's filling: green potion with a shine on top. 6 x 6, stretches."""
    im, px = _pixels(6, 6)
    rows = [(150, 226, 120), (96, 196, 84), (72, 170, 70), (72, 170, 70), (54, 140, 58), (40, 112, 48)]
    for y, c in enumerate(rows):
        for x in range(6):
            px[x, y] = c + (255,)
    return im


def rpg_potion_trough():
    """The progress bar's glass tube in a wooden frame. 8 x 10, stretches."""
    im, px = _pixels(8, 10)
    for x in range(8):
        for y in range(10):
            edge = x in (0, 7) or y in (0, 9)
            frame = x in (1, 6) or y in (1, 8)
            px[x, y] = (RPG_INK if edge else (150, 88, 52) if frame else (70, 44, 36)) + (255,)
    return im


RPG_TROUGH_BORDER = (2, 2, 2, 2)


def rpg_banner(width_px: int):
    """A ribbon banner with notched tails, `width_px` wide (unscaled), 24 high - for the home window title.
    The tails hang 4 pixels below the front, so there's as much empty space above it: text centred on the
    image is centred on the ribbon."""
    W, H, top = width_px, 24, 4
    im, px = _pixels(W, H)
    face, face_hi, face_lo = (243, 205, 152), (251, 226, 186), (214, 160, 104)
    tail, tail_lo = (196, 128, 78), (150, 88, 52)

    def put(x, y, c):
        y += top
        if 0 <= x < W and 0 <= y < H:
            px[x, y] = c + (255,)

    def shape(cells, fill, shade_row=None):
        """Fill a set of cells and draw the ink outline around it."""
        for x, y in cells:
            put(x, y, tail_lo if shade_row is not None and y >= shade_row else fill)
        for x, y in cells:
            if any((x + dx, y + dy) not in cells for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1))):
                put(x, y, RPG_INK)
    for side in (0, 1):                                     # the tails, behind the ribbon
        cells = set()
        for y in range(5, 19):
            notch = max(0, 3 - abs(y - 11.5))               # the V cut into the outer end
            for i in range(10):
                if i >= notch:
                    cells.add((i if side == 0 else W - 1 - i, y))
        shape(cells, tail, shade_row=16)
    front = {(x, y) for x in range(6, W - 6) for y in range(0, 16)}
    shape(front, face)
    for x in range(7, W - 7):
        put(x, 1, face_hi)
        put(x, 14, face_lo)
    for x in (7, 8, W - 9, W - 8):                          # folds at the ends
        for y in range(2, 14):
            if (x + y) % 2 == 0:
                put(x, y, face_lo)
    for x, y in ((7, 16), (8, 16), (8, 17), (W - 8, 16), (W - 9, 16), (W - 9, 17)):    # where it turns under
        put(x, y, RPG_INK)
    return im


def rpg_grain(size=(256, 256), base=(199, 160, 108), strength=10, seed=7):
    """Parchment grain to tile behind windows: chunky 2 x 2 noise with a few fibres."""
    import random
    from PIL import Image, ImageDraw
    rnd = random.Random(seed)
    w, h = size
    small = Image.effect_noise((w // 2, h // 2), 40).resize((w, h), Image.NEAREST)
    noise = small.load()
    im = Image.new("RGB", (w, h), base)
    px = im.load()
    for y in range(h):
        for x in range(w):
            d = (noise[x, y] - 128) / 128 * strength
            px[x, y] = tuple(max(0, min(255, round(c + d))) for c in base)
    draw = ImageDraw.Draw(im)
    for _ in range(14):                                      # fibres
        x, y = rnd.randrange(w), rnd.randrange(h)
        length = rnd.randrange(12, 46)
        shade = _rgb_shift(base, -rnd.randrange(8, 16))
        draw.line([(x, y), (x + length, y + rnd.choice((-1, 0, 0, 1)))], fill=shade, width=1)
        draw.line([(x - w, y), (x + length - w, y)], fill=shade, width=1)      # wraps around
    return im


# ============================================================ look: light / dark / RPG theme
# Every Fleem window is drawn from one of these palettes. Plain tk widgets remember which palette
# colour ("token") each of their colours is, so switching between light and dark repaints them; ttk
# widgets (buttons, boxes, ...) follow the ttk styles set up in apply_theme().
THEMES = {
    "light": {
        "bg": "#eef0f4", "surface": "#ffffff", "surface2": "#f5f6f8", "border": "#e1e4ea", "border_strong": "#c6ccd6",
        "text": "#1c2230", "muted": "#5f6b7d", "faint": "#98a2b3",
        "accent": "#3d6ef5", "accent_hover": "#2f5ee0", "accent_press": "#284fc0", "accent_text": "#ffffff",
        "accent_soft": "#e8eeff", "accent_disabled": "#c3d1fb", "accent_disabled_text": "#f4f7ff",
        "button": "#f2f3f6", "button_hover": "#e7e9ee", "button_press": "#dcdfe6",
        "field": "#ffffff", "field_border": "#cbd1db",
        "danger": "#d92d20", "danger_soft": "#fdeceb", "ok": "#12805c", "ok_soft": "#e4f5ed",
        "warn": "#b54708", "warn_soft": "#fdf1e3", "tile": "#d9dde5", "tile_hover": "#aeb6c4", "paper": "#f4f5f7",
    },
    "dark": {
        "bg": "#111318", "surface": "#1a1d24", "surface2": "#21252d", "border": "#2c313b", "border_strong": "#414856",
        "text": "#e6e8ec", "muted": "#a3abb8", "faint": "#6f7787",
        "accent": "#4f7fff", "accent_hover": "#6690ff", "accent_press": "#3f6de6", "accent_text": "#ffffff",
        "accent_soft": "#1d2843", "accent_disabled": "#283350", "accent_disabled_text": "#7b88a6",
        "button": "#252932", "button_hover": "#2d323c", "button_press": "#353b47",
        "field": "#14161b", "field_border": "#3a404c",
        "danger": "#ff6b5f", "danger_soft": "#3a1f1f", "ok": "#4cc38a", "ok_soft": "#163327",
        "warn": "#f0a54a", "warn_soft": "#3a2b16", "tile": "#2a2f38", "tile_hover": "#4a5262", "paper": "#2a2e36",
    },
    "rpg": {                            # "Adventurer": parchment, leather, wood and wax
        "bg": "#c7a06c", "surface": "#f1dfb4", "surface2": "#e4c890", "border": "#b48552", "border_strong": "#6b3f22",
        "text": "#3b2314", "muted": "#6e4a2c", "faint": "#9b7650",
        "accent": "#9a3324", "accent_hover": "#b3412e", "accent_press": "#7c281c", "accent_text": "#fbeed0",
        "accent_soft": "#eec48f", "accent_disabled": "#c9a07c", "accent_disabled_text": "#f3e3c4",
        "button": "#dcb67c", "button_hover": "#e8c891", "button_press": "#c49c62",
        "field": "#fbf0d6", "field_border": "#8b5e36",
        "danger": "#a52a1c", "danger_soft": "#f1c9a8", "ok": "#3d7a2b", "ok_soft": "#dbe4ad",
        "warn": "#a35b0c", "warn_soft": "#f5d99c", "tile": "#d2ab72", "tile_hover": "#8b5a2b", "paper": "#f6e6c0",
    },
}
THEMES["rpg_dark"] = {                  # the same adventure by candlelight: dark wood and leather, gold
    "bg": "#24170f", "surface": "#36231a", "surface2": "#452d1f", "border": "#654128", "border_strong": "#9a6a3c",
    "text": "#f1dcb0", "muted": "#d2b083", "faint": "#9d7d58",
    "accent": "#e2a640", "accent_hover": "#f1b955", "accent_press": "#c58c2b", "accent_text": "#24170f",
    "accent_soft": "#5c3d1e", "accent_disabled": "#5e4630", "accent_disabled_text": "#9a7f60",
    "button": "#5a3a24", "button_hover": "#6b4630", "button_press": "#4a2e1c",
    "field": "#21140d", "field_border": "#8a5e36",
    "danger": "#ff7c5e", "danger_soft": "#4a2018", "ok": "#86c972", "ok_soft": "#26381d",
    "warn": "#f2b354", "warn_soft": "#4a3418", "tile": "#4a3020", "tile_hover": "#8a5e36", "paper": "#3a2a1e",
}
THEME_LABELS = {"light": "Basic - Light", "dark": "Basic - Dark", "rpg": "RPG - Light", "rpg_dark": "RPG - Dark"}
RPG_INK_TEXT = "#3b2314"                # text written on the parchment art (in both RPG themes)
RPG_FONTS = ("Palatino Linotype", "Book Antiqua", "Georgia", "DejaVu Serif", "Liberation Serif")
# The pixel font for what's written on the RPG scrolls (buttons, the banner): Pixeloid Sans (and Pixeloid
# Sans Bold) by GGBotNet, SIL Open Font License - it comes with Fleem in its "fonts" folder. settings.ini
# rpg_button_font can name another font (installed, or put in a "fonts" folder next to Fleem).
RPG_BUTTON_FONT = "Pixeloid Sans"
FONTS_DIR = HERE / "fonts"


def load_private_fonts() -> None:
    """Make Fleem's own fonts (packed into the .exe) and the ones in a "fonts" folder next to Fleem usable
    by Fleem only - nothing is installed in Windows."""
    if os.name != "nt":
        return
    folders = [Path(getattr(sys, "_MEIPASS", HERE)) / "fonts", FONTS_DIR]
    loaded = set()
    try:
        import ctypes
        for folder in folders:
            if not folder.is_dir():
                continue
            for f in sorted(folder.iterdir()):
                if f.suffix.lower() in (".ttf", ".otf", ".fon") and f.name.lower() not in loaded:
                    if ctypes.windll.gdi32.AddFontResourceExW(str(f), 0x10, 0):     # FR_PRIVATE
                        loaded.add(f.name.lower())
    except Exception as e:
        log.info("Couldn't load Fleem's fonts: %s", e)


def rpg_button_family(win) -> str | None:
    """The pixel font for the scrolls, if there is one on this PC."""
    if not RPG_BUTTON_FONT:
        return None
    app = win._root()
    cache = app.__dict__.setdefault("_fleem_button_family", {})
    if RPG_BUTTON_FONT not in cache:
        from tkinter import font as tkfont
        have = {f.lower(): f for f in tkfont.families(app)}
        cache[RPG_BUTTON_FONT] = have.get(RPG_BUTTON_FONT.lower())
    return cache[RPG_BUTTON_FONT]


def rpg_scroll_font(win, big: bool = False, bold: bool = False):
    """The font for text on a scroll: the pixel font at a size it's crisp at (Pixeloid's pixels are
    9 px apart, so 18 px at normal scaling), else the storybook serif."""
    k = rpg_scale(win)
    family = rpg_button_family(win)
    if family:
        return (family, -9 * (round(1.5 * k) if big else k)) + (("bold",) if bold else ())
    return (rpg_font_family(win), 12 if big else 10, "bold")
# colours the older windows were written with -> the palette colour they stand for
LEGACY_COLOURS = {"#ffffff": "surface", "white": "surface", "systembuttonface": "surface", "systemwindow": "surface",
                  "#d9d9d9": "surface", "#f0f0f0": "surface", "#666666": "muted", "#777777": "muted",
                  "#999999": "faint", "#c0392b": "danger", "#f2f2f2": "surface2", "#f6f7f9": "surface2",
                  "#e2e4e8": "border", "#2f6fd6": "accent", "systembuttontext": "text", "systemwindowtext": "text",
                  "black": "text", "#000000": "text"}
_TOKEN_OPTIONS = {"bg": "background", "fg": "foreground", "hl": "highlightbackground", "hlc": "highlightcolor",
                  "abg": "activebackground", "afg": "activeforeground", "insert": "insertbackground",
                  "dfg": "disabledforeground", "sel": "selectbackground"}
APPEARANCE_FILE = HERE / "appearance.json"
DEFAULT_THEME = "rpg_dark"                 # until you pick another (Theme button / Fleem's menu > Appearance)


def appearance_setting() -> str:
    """"auto" (follow Windows), "light", "dark", "rpg" or "rpg_dark" (DEFAULT_THEME until one is picked)."""
    try:
        value = json.loads(APPEARANCE_FILE.read_text(encoding="utf-8")).get("theme", DEFAULT_THEME)
        return value if value in THEME_LABELS or value == "auto" else DEFAULT_THEME
    except Exception:
        return DEFAULT_THEME


def set_appearance(value: str) -> None:
    try:
        APPEARANCE_FILE.write_text(json.dumps({"theme": value}), encoding="utf-8")
    except Exception as e:
        log.info("Couldn't save the appearance: %s", e)


def windows_uses_dark() -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except Exception:
        return False


def theme_name() -> str:
    value = appearance_setting()
    if value == "auto":
        return "dark" if windows_uses_dark() else "light"
    return value


def is_rpg() -> bool:
    return theme_name().startswith("rpg")


def is_dark() -> bool:
    return theme_name() in ("dark", "rpg_dark")


def mpcfill_module():
    """The MPCFill card search (fleem_mpcfill.py), or None if that file isn't there - then the "Find
    cards…" button isn't shown. Delete the file to remove the feature."""
    try:
        import fleem_mpcfill
        return fleem_mpcfill
    except Exception as e:                   # missing, or broken by an edit
        if not isinstance(e, ImportError):
            log.warning("The MPCFill search couldn't be loaded: %s", e)
        return None


def palette() -> dict:
    return THEMES[theme_name()]


def T(widget, **tokens):
    """Colour a plain tk widget from the palette (bg="surface", fg="muted", ...) and remember the tokens,
    so a theme switch repaints it. Returns the widget."""
    P = palette()
    keep = getattr(widget, "_fleem_tokens", {})
    for short, token in tokens.items():
        option = _TOKEN_OPTIONS.get(short, short)
        keep[option] = token
        try:
            widget.configure(**{option: P[token]})
        except Exception:
            pass
    widget._fleem_tokens = keep
    if "background" in keep:
        _texture(widget)
    return widget


def rpg_scale(win) -> int:
    """How big one pixel of the RPG art is on this screen (2 at 100 % scaling)."""
    try:
        return max(2, round(win.winfo_fpixels("1i") / 96 * 2))
    except Exception:
        return 2


def _rpg_photo(win, image, scale=None):
    """PIL pixel art -> a Tk image, scaled up with hard edges."""
    from PIL import Image, ImageTk
    k = scale or rpg_scale(win)
    return ImageTk.PhotoImage(image.resize((image.width * k, image.height * k), Image.NEAREST), master=win)


GRAIN_STRENGTH = {"bg": 16, "surface": 9, "surface2": 10}


def _grain_tile(win, token="bg"):
    """Parchment grain in a palette colour (one tile per colour per Tk interpreter, made once)."""
    app = win._root()
    tiles = app.__dict__.setdefault("_fleem_grain", {})
    key = (theme_name(), token)
    if key not in tiles:
        from PIL import ImageTk
        base = tuple(int(palette()[token][i:i + 2], 16) for i in (1, 3, 5))
        tiles[key] = ImageTk.PhotoImage(rpg_grain((256, 256), base, GRAIN_STRENGTH[token], seed=7 + len(tiles)),
                                        master=app)
    return tiles[key]


def _wants_grain(w, token) -> bool:
    """Which frames get grain: anything in the window colour, and the big parchment panels (frames
    straight inside a window or a window-coloured frame; not tiles, which have a frame line)."""
    if token == "bg":
        return True
    if token not in ("surface", "surface2"):
        return False
    try:
        if int(w.cget("highlightthickness")) > 0:
            return False
        parent = w.nametowidget(w.winfo_parent())
        return parent.winfo_class() in ("Tk", "Toplevel") or getattr(parent, "_fleem_tokens", {}).get("background") == "bg"
    except Exception:
        return False


GRAIN_MAX_PIXELS = 3_000_000          # a very tall scrolling area only gets grain at the top


def _texture(w) -> None:
    """RPG theme: parchment grain behind a window or frame painted with the "bg" colour (a picture
    placed underneath its contents). Other themes: take it away again."""
    try:
        tokens = getattr(w, "_fleem_tokens", {})
        token = tokens.get("background")
        want = (is_rpg() and w.winfo_class() in ("Frame", "Tk", "Toplevel")
                and _wants_grain(w, token))
        lab = getattr(w, "_fleem_tex", None)
        if lab is not None and not lab.winfo_exists():
            lab = w._fleem_tex = None
        if lab is not None and (not want or lab._fleem_token != (theme_name(), token)):
            w._fleem_tex = None
            lab.destroy()
            lab = None
        if not want or lab is not None:
            return
        import tkinter as tk
        tile = _grain_tile(w, token)
        lab = tk.Label(w, bd=0, highlightthickness=0, padx=0, pady=0, anchor="nw", bg=palette()[token])
        lab._fleem_tokens = {}
        lab._fleem_token = (theme_name(), token)
        lab._fleem_size = (0, 0)
        lab.place(x=0, y=0, relwidth=1, relheight=1)
        lab.lower()
        try:                                        # clicks and keys on it count as on the frame
            lab.configure(cursor=w.cget("cursor"))
            lab.bindtags((str(lab), str(w), str(lab.winfo_toplevel()), "all"))
        except Exception:
            pass
        w._fleem_tex = lab

        def fill(_e=None):
            width, height = lab.winfo_width(), lab.winfo_height()
            height = min(height, max(256, GRAIN_MAX_PIXELS // max(1, width)))
            have = lab._fleem_size
            if width <= have[0] and height <= have[1]:
                return
            size = (-(-max(width, have[0]) // 256) * 256, -(-max(height, have[1]) // 256) * 256)
            img = tk.PhotoImage(master=lab, width=size[0], height=size[1])
            img.tk.call(img, "copy", tile, "-to", 0, 0, size[0], size[1])      # tiles the grain
            lab._img, lab._fleem_size = img, size      # (kept first: setting it can run this again)
            lab.configure(image=img)

        def gone(_e=None):                          # (a window clearing its frame removes it too)
            if getattr(w, "_fleem_tex", None) is lab:
                w._fleem_tex = None
                try:
                    w.after_idle(lambda: w.winfo_exists() and _texture(w))
                except Exception:
                    pass
        lab.bind("<Configure>", fill)
        lab.bind("<Destroy>", gone)
    except Exception as e:
        log.info("No texture: %s", e)


def _all_widgets(w):
    out = [w]
    for child in w.winfo_children():
        out += _all_widgets(child)
    return out


def paint(win) -> None:
    """Repaint every plain tk widget in a window from the current palette (older windows' fixed colours
    are mapped to palette colours the first time)."""
    P = palette()
    for w in _all_widgets(win):
        try:
            if not w.winfo_exists() or getattr(w, "_fleem_token", None) is not None:
                continue                              # gone already, or the grain behind a frame
            if w.winfo_class().startswith("T") and w.winfo_class() not in ("Tk", "Toplevel", "Text"):
                continue                              # ttk widgets: styled by apply_theme
        except Exception:
            continue
        tokens = getattr(w, "_fleem_tokens", None)
        if tokens is None:
            tokens = {}
            for option in ("background", "foreground", "highlightbackground", "activebackground",
                           "activeforeground", "insertbackground", "disabledforeground"):
                try:
                    value = str(w.cget(option)).lower()
                except Exception:
                    continue
                if value in LEGACY_COLOURS:
                    tokens[option] = LEGACY_COLOURS[value]
            if "background" in tokens and w.winfo_class() in ("Label", "Frame", "Canvas", "Tk", "Toplevel") \
                    and "highlightbackground" not in tokens:
                tokens["highlightbackground"] = tokens["background"]
            w._fleem_tokens = tokens
        for option, token in tokens.items():
            try:
                w.configure(**{option: P[token]})
            except Exception:
                pass
        if "background" in tokens:
            _texture(w)
        _swap_widget_font(w)


def rpg_font_family(win) -> str:
    """The storybook serif the RPG theme uses (the first one this PC has)."""
    app = win._root()
    family = getattr(app, "_fleem_rpg_family", None)
    if family is None:
        from tkinter import font as tkfont
        have = set(tkfont.families(app))
        family = app._fleem_rpg_family = next((f for f in RPG_FONTS if f in have), "TkDefaultFont")
    return family


def _swap_spec(win, spec):
    """A font spec in the RPG serif (Segoe UI Semibold -> bold), or None if it isn't a Segoe UI font."""
    try:
        parts = list(win.tk.splitlist(spec))
    except Exception:
        return None
    if not parts or not str(parts[0]).startswith("Segoe UI"):
        return None
    styles = [str(p) for p in parts[2:]]
    if "Semibold" in str(parts[0]) and "bold" not in styles:
        styles.append("bold")
    return (rpg_font_family(win), parts[1] if len(parts) > 1 else 9, *styles)


def _swap_widget_font(w) -> None:
    """RPG theme: a plain widget's Segoe UI font becomes the storybook serif (and back afterwards).
    Named fonts are swapped once for the whole window in swap_named_fonts()."""
    try:
        from tkinter import font as tkfont
        spec = w.cget("font")
    except Exception:
        return
    try:
        if str(spec) in tkfont.names(w._root()):
            return
        original = getattr(w, "_fleem_font_spec", None)
        if is_rpg():
            if original is None:
                new = _swap_spec(w, spec)
                if new is not None:
                    w._fleem_font_spec = spec
                    w.configure(font=new)
        elif original is not None:
            w.configure(font=original)
            w._fleem_font_spec = None
    except Exception:
        pass


def swap_named_fonts(win) -> None:
    """RPG theme: every named Segoe UI font of this Tk interpreter becomes the storybook serif; other
    themes: back to what it was."""
    from tkinter import font as tkfont
    app = win._root()
    originals = app.__dict__.setdefault("_fleem_font_families", {})
    rpg = is_rpg()
    for name in tkfont.names(app):
        try:
            f = tkfont.nametofont(name, root=app)
        except Exception:
            continue
        if rpg:
            family = originals.get(name, (f.cget("family"), f.cget("weight")))
            if str(family[0]).startswith("Segoe UI"):
                originals[name] = family
                f.configure(family=rpg_font_family(app), weight="bold" if "Semibold" in family[0] else family[1])
        elif name in originals:
            family, weight = originals.pop(name)
            f.configure(family=family, weight=weight)


def set_dark_title_bar(win, dark: bool) -> None:
    """Windows 10/11: dark or light title bar for a Tk window."""
    if os.name != "nt":
        return
    try:
        import ctypes
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id()) or win.winfo_id()
        value = ctypes.c_int(1 if dark else 0)
        for attribute in (20, 19):                    # DWMWA_USE_IMMERSIVE_DARK_MODE (newer, older builds)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(value), 4) == 0:
                break
        rpg = is_rpg()                                # Windows 11: a wooden title bar with parchment text

        def colorref(hex_colour):
            r, g, b = (int(hex_colour[i:i + 2], 16) for i in (1, 3, 5))
            return ctypes.c_uint(r | g << 8 | b << 16)
        default = ctypes.c_uint(0xFFFFFFFF)           # DWMWA_COLOR_DEFAULT
        wood = "#2e1d12" if theme_name() == "rpg_dark" else "#4a2a17"
        for attribute, colour in ((35, wood), (36, "#f3dcae"), (34, "#6b3f22")):   # caption, text, border
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(colorref(colour) if rpg else default), 4)
    except Exception:
        pass


def _rpg_ttk_theme(win, s, name: str) -> None:
    """The ttk theme for an RPG palette ("fleem_rpg" / "fleem_rpg_dark", made once per Tk interpreter):
    clam, with pixel-art buttons (scrolls), checkboxes (wax seals), radio buttons (gems), scrollbars
    (wooden rods) and progress bars (potion)."""
    ttk_name = "fleem_" + name
    if ttk_name in s.theme_names():
        return
    app = win._root()
    s.theme_use("clam")
    plain_button = s.layout("TButton")
    s.theme_create(ttk_name, parent="clam")
    s.theme_use(ttk_name)
    k = rpg_scale(app)
    keep = app.__dict__.setdefault("_fleem_rpg_images", {})
    fade = tuple(int(THEMES[name]["surface"][i:i + 2], 16) for i in (1, 3, 5))

    def photo(key, image):
        keep[(name, key)] = _rpg_photo(app, image, k)
        return keep[(name, key)]

    def border(b):
        return tuple(v * k for v in b)
    for variant, element in (("plain", "Rpg.scroll"), ("royal", "Rpg.royal"), ("charred", "Rpg.charred")):
        s.element_create(element, "image", photo(element, rpg_scroll(variant, "normal", fade)),
                         ("disabled", photo(element + "d", rpg_scroll(variant, "disabled", fade))),
                         ("pressed", photo(element + "p", rpg_scroll(variant, "pressed", fade))),
                         ("active", photo(element + "a", rpg_scroll(variant, "active", fade))),
                         border=border(RPG_SCROLL_BORDER), sticky="nsew")
    for on, element in ((True, "Rpg.tabon"), (False, "Rpg.taboff")):
        s.element_create(element, "image", photo(element, rpg_tab(on, "normal", fade)),
                         ("disabled", photo(element + "d", rpg_tab(on, "disabled", fade))),
                         ("active", photo(element + "a", rpg_tab(on, "active", fade))),
                         border=border(RPG_TAB_BORDER), sticky="nsew")

    def button(element):
        return [(element, {"sticky": "nsew", "children": [
            ("Button.padding", {"sticky": "nsew", "children": [("Button.label", {"sticky": "nsew"})]})]})]
    s.layout("TButton", button("Rpg.scroll"))
    s.layout("Accent.TButton", button("Rpg.royal"))
    s.layout("Danger.TButton", button("Rpg.charred"))
    s.layout("Seg.TButton", button("Rpg.taboff"))
    s.layout("SegOn.TButton", button("Rpg.tabon"))
    s.layout("Ghost.TButton", plain_button)                # a link-like button: no scroll
    s.element_create("Rpg.check", "image", photo("check", rpg_checkbox(False, False, fade)),
                     ("disabled", "selected", photo("checkds", rpg_checkbox(True, True, fade))),
                     ("disabled", photo("checkd", rpg_checkbox(False, True, fade))),
                     ("selected", photo("checks", rpg_checkbox(True, False, fade))), sticky="")
    s.element_create("Rpg.radio", "image", photo("radio", rpg_radio(False, False, fade)),
                     ("disabled", "selected", photo("radiods", rpg_radio(True, True, fade))),
                     ("disabled", photo("radiod", rpg_radio(False, True, fade))),
                     ("selected", photo("radios", rpg_radio(True, False, fade))), sticky="")
    for kind, element in (("Checkbutton", "Rpg.check"), ("Radiobutton", "Rpg.radio")):
        s.layout("T" + kind, [(f"{kind}.padding", {"sticky": "nswe", "children": [
            (element, {"side": "left", "sticky": ""}),
            (f"{kind}.focus", {"side": "left", "sticky": "w", "children": [(f"{kind}.label", {"sticky": "nswe"})]})]})])
    s.element_create("Vertical.Scrollbar.thumb", "image", photo("rod", rpg_rod()),
                     ("pressed", photo("roda", rpg_rod("active"))), ("active", keep[(name, "roda")]),
                     border=border(RPG_ROD_BORDER), sticky="nsew")
    s.layout("Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
        ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
    s.element_create("Horizontal.Progressbar.pbar", "image", photo("potion", rpg_potion_bar()),
                     border=(k, k, k, k), sticky="nsew")
    s.element_create("Horizontal.Progressbar.trough", "image", photo("tube", rpg_potion_trough()),
                     border=border(RPG_TROUGH_BORDER), sticky="nsew")


def rpg_banner_photo(win, width: int):
    """The ribbon banner, about `width` px wide on screen."""
    k = rpg_scale(win)
    keep = win._root().__dict__.setdefault("_fleem_rpg_images", {})
    key = ("banner", width)
    if key not in keep:
        keep[key] = _rpg_photo(win, rpg_banner(max(30, width // k)), k)
    return keep[key]


def apply_theme(win, font=None, bold=None) -> dict:
    """Set up Fleem's look on a window: ttk styles, plain widgets, title bar. Returns the palette.
    The ttk styles are shared by every window of the app, so a pop-up (no `font` given) only repaints
    itself and leaves the styles - and the main window's fonts - alone."""
    from tkinter import ttk
    P = palette()
    app = win._root()
    if font is None and getattr(app, "_fleem_styled", None) == theme_name():
        paint(win)
        set_dark_title_bar(win, is_dark() or is_rpg())
        return P
    rpg = is_rpg()
    swap_named_fonts(win)                       # (the RPG theme's storybook serif, or back to Segoe UI)
    font = font or getattr(app, "_fleem_font", None) or ("Segoe UI", 9)
    if isinstance(font, tuple) and rpg:
        font = _swap_spec(win, font) or font
    elif isinstance(font, tuple) and not str(font[0]).startswith("Segoe"):
        font = ("Segoe UI",) + tuple(font[1:2])
    if bold is None or app.__dict__.get("_fleem_bold_auto"):
        bold = (font.actual("family"), font.cget("size"), "bold") if hasattr(font, "actual") else (font[0], font[1], "bold")
        app._fleem_bold_auto = True
    app._fleem_styled, app._fleem_font, app._fleem_bold = theme_name(), font, bold
    s = ttk.Style(win)
    try:
        if rpg:
            _rpg_ttk_theme(win, s, theme_name())
        s.theme_use("fleem_" + theme_name() if rpg else "clam")
    except Exception as e:
        log.info("Theme: %s", e)
        s.theme_use("clam")
    s.configure(".", background=P["surface"], foreground=P["text"], fieldbackground=P["field"],
                bordercolor=P["border"], lightcolor=P["surface"], darkcolor=P["surface"], troughcolor=P["surface2"],
                selectbackground=P["accent"], selectforeground=P["accent_text"], insertcolor=P["text"],
                focuscolor=P["accent"], font=font)
    button = dict(background=P["button"], foreground=P["text"], bordercolor=P["border_strong"],
                  lightcolor=P["button"], darkcolor=P["button"], padding=(12, 5), relief="flat",
                  focusthickness=1, focuscolor=P["border_strong"], anchor="center")
    s.configure("TButton", **button)
    s.map("TButton", background=[("disabled", P["button"]), ("pressed", P["button_press"]), ("active", P["button_hover"])],
          lightcolor=[("pressed", P["button_press"]), ("active", P["button_hover"])],
          darkcolor=[("pressed", P["button_press"]), ("active", P["button_hover"])],
          foreground=[("disabled", P["faint"])], bordercolor=[("disabled", P["border"])])
    s.configure("Accent.TButton", background=P["accent"], foreground=P["accent_text"], bordercolor=P["accent"],
                lightcolor=P["accent"], darkcolor=P["accent"], focuscolor=P["accent_text"], font=bold)
    s.map("Accent.TButton",
          background=[("disabled", P["accent_disabled"]), ("pressed", P["accent_press"]), ("active", P["accent_hover"])],
          lightcolor=[("disabled", P["accent_disabled"]), ("pressed", P["accent_press"]), ("active", P["accent_hover"])],
          darkcolor=[("disabled", P["accent_disabled"]), ("pressed", P["accent_press"]), ("active", P["accent_hover"])],
          bordercolor=[("disabled", P["accent_disabled"]), ("active", P["accent_hover"])],
          foreground=[("disabled", P["accent_disabled_text"])])
    s.configure("Danger.TButton", foreground=P["danger"])
    s.map("Danger.TButton", background=[("pressed", P["danger_soft"]), ("active", P["danger_soft"])],
          lightcolor=[("active", P["danger_soft"])], darkcolor=[("active", P["danger_soft"])],
          bordercolor=[("active", P["danger"])])
    s.configure("Ghost.TButton", background=P["surface"], foreground=P["accent"], bordercolor=P["surface"],
                lightcolor=P["surface"], darkcolor=P["surface"], padding=(6, 4))
    s.map("Ghost.TButton", background=[("pressed", P["accent_soft"]), ("active", P["accent_soft"])],
          lightcolor=[("active", P["accent_soft"])], darkcolor=[("active", P["accent_soft"])],
          bordercolor=[("active", P["accent_soft"])], foreground=[("disabled", P["faint"])])
    for name, on in (("Seg.TButton", False), ("SegOn.TButton", True)):
        bg = P["accent_soft"] if on else P["surface"]
        s.configure(name, background=bg, foreground=P["accent"] if on else P["muted"],
                    bordercolor=P["accent"] if on else P["border_strong"], lightcolor=bg, darkcolor=bg,
                    padding=(14, 5), font=bold)
        s.map(name, background=[("disabled", P["surface"]), ("active", P["accent_soft"] if on else P["button_hover"])],
              lightcolor=[("active", P["accent_soft"] if on else P["button_hover"])],
              darkcolor=[("active", P["accent_soft"] if on else P["button_hover"])],
              foreground=[("disabled", P["faint"])])
    field = dict(fieldbackground=P["field"], foreground=P["text"], bordercolor=P["field_border"],
                 lightcolor=P["field"], darkcolor=P["field"], insertcolor=P["text"], padding=4)
    for name in ("TEntry", "TSpinbox", "TCombobox"):
        s.configure(name, **field)
        s.map(name, bordercolor=[("focus", P["accent"])], lightcolor=[("focus", P["accent"])],
              fieldbackground=[("disabled", P["surface2"]), ("readonly", P["field"])],
              foreground=[("disabled", P["faint"])],
              selectbackground=[("readonly", P["field"])] if name == "TCombobox" else [],
              selectforeground=[("readonly", P["text"])] if name == "TCombobox" else [])
    s.configure("TSpinbox", arrowcolor=P["muted"], background=P["button"], arrowsize=11)
    s.configure("TCombobox", arrowcolor=P["muted"], background=P["button"], arrowsize=13)
    s.map("TCombobox", background=[("active", P["button_hover"])], arrowcolor=[("disabled", P["faint"])])
    win.option_add("*TCombobox*Listbox.background", P["field"])
    win.option_add("*TCombobox*Listbox.foreground", P["text"])
    win.option_add("*TCombobox*Listbox.selectBackground", P["accent"])
    win.option_add("*TCombobox*Listbox.selectForeground", P["accent_text"])
    win.option_add("*TCombobox*Listbox.borderWidth", 0)
    for bg, prefix in ((P["surface"], ""), (P["surface2"], "Panel.")):
        for kind in ("TCheckbutton", "TRadiobutton"):
            s.configure(prefix + kind, background=bg, foreground=P["text"], indicatorbackground=P["field"],
                        indicatorforeground=P["accent_text"], upperbordercolor=P["field_border"],
                        lowerbordercolor=P["field_border"], focuscolor=bg, padding=(2, 3))
            s.map(prefix + kind, background=[("active", bg)],
                  indicatorbackground=[("disabled", P["surface2"]), ("selected", P["accent"]), ("pressed", P["accent_soft"])],
                  foreground=[("disabled", P["faint"])])
    s.configure("Vertical.TScrollbar", background=P["button"], troughcolor=P["bg"], bordercolor=P["bg"],
                lightcolor=P["button"], darkcolor=P["button"], arrowcolor=P["muted"], gripcount=0, arrowsize=12)
    s.map("Vertical.TScrollbar", background=[("active", P["button_hover"])])
    s.configure("Horizontal.TProgressbar", background=P["accent"], troughcolor=P["surface2"], bordercolor=P["surface2"],
                lightcolor=P["accent"], darkcolor=P["accent"], thickness=8)
    if rpg:                                     # the pixel-art pieces: colours for what's around them
        k = rpg_scale(win)
        flat = [("disabled", P["surface"]), ("pressed", P["surface"]), ("active", P["surface"])]
        for name, colour in (("TButton", RPG_INK_TEXT), ("Accent.TButton", "#7c1d12"), ("Danger.TButton", "#a52a1c"),
                             ("SegOn.TButton", RPG_INK_TEXT), ("Seg.TButton", "#4a2c18")):
            s.configure(name, background=P["surface"], foreground=colour, padding=(9 * k, 3, 9 * k, 3))
            s.map(name, background=flat, foreground=[("disabled", "#e3cc9f" if is_dark() else "#9b7650")])
        pixel = rpg_button_family(win)
        on_scroll = rpg_scroll_font(win) if pixel else font
        strong = rpg_scroll_font(win, bold=True) if pixel else bold
        for name in ("TButton", "Danger.TButton"):
            s.configure(name, font=on_scroll)
        s.configure("Accent.TButton", font=strong)
        s.configure("SegOn.TButton", padding=(5 * k, 2, 5 * k, 2), font=strong)
        s.configure("Seg.TButton", padding=(5 * k, 2, 5 * k, 2), font=strong)
        for prefix in ("", "Panel."):
            for kind in ("TCheckbutton", "TRadiobutton"):
                s.configure(prefix + kind, padding=(2, 3), indicatormargin=0)
        s.configure("Vertical.TScrollbar", troughcolor=P["surface2"], background=P["surface2"])
    paint(win)
    set_dark_title_bar(win, is_dark() or is_rpg())
    return P


def theme_menu(button, on_change) -> None:
    """Pop up the theme list (Match Windows / Light / Dark / Adventurer) under `button`; picking one
    saves it and calls on_change()."""
    import tkinter as tk
    P = palette()
    menu = tk.Menu(button, tearoff=0, bg=P["surface"], fg=P["text"], activebackground=P["accent_soft"],
                   activeforeground=P["text"], selectcolor=P["accent"], bd=1, relief="solid",
                   font=(rpg_font_family(button), 10) if is_rpg() else ("Segoe UI", 9))
    choice = tk.StringVar(master=button, value=appearance_setting())

    def pick(value):
        set_appearance(value)
        on_change()
    for value, label in THEME_LABELS.items():
        menu.add_radiobutton(label=label, value=value, variable=choice, command=lambda v=value: pick(v))
    button._fleem_menu = (menu, choice)
    menu.tk_popup(button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())


def refresh_title_bar(win) -> None:
    """Make Windows redraw a window's title bar in the new colours."""
    try:
        w, h = win.winfo_width(), win.winfo_height()
        win.geometry(f"{w + 1}x{h}")
        win.after(30, lambda: win.geometry(f"{w}x{h}"))
    except Exception:
        pass


def center_on(win, parent) -> None:
    try:
        win.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_reqwidth()) // 2
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - win.winfo_reqheight()) // 3)
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass


class Tooltip:
    """A small hint that appears when the mouse rests on a widget."""
    def __init__(self, widget, text):
        self.widget, self.text, self.tip, self.job = widget, text, None, None
        widget.bind("<Enter>", self.schedule, add="+")
        widget.bind("<Leave>", self.hide, add="+")
        widget.bind("<ButtonPress>", self.hide, add="+")

    def schedule(self, _e=None):
        self.hide()
        self.job = self.widget.after(550, self.show)

    def show(self):
        import tkinter as tk
        if self.tip or not self.widget.winfo_exists():
            return
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{self.widget.winfo_rootx() + 8}+{self.widget.winfo_rooty() + self.widget.winfo_height() + 4}")
        T(tk.Label(self.tip, text=self.text, padx=8, pady=4, justify="left", wraplength=320, bd=0,
                   highlightthickness=1), bg="surface2", fg="text", hl="border_strong").pack()

    def hide(self, _e=None):
        if self.job:
            try:
                self.widget.after_cancel(self.job)
            except Exception:
                pass
            self.job = None
        if self.tip:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None


class ThemedMessages:
    """Stand-in for tkinter.messagebox inside Fleem's windows: the same calls, drawn in Fleem's look
    (so dark mode doesn't flash a white box). Without a parent window it falls back to the normal ones."""
    ICONS = {"info": ("i", "accent"), "question": ("?", "accent"), "warning": ("!", "warn"), "error": ("✕", "danger")}

    def _box(self, title, message, kind, buttons, parent, default, cancel_value):
        import tkinter as tk
        from tkinter import ttk, messagebox as mb
        if parent is None:
            native = {"info": mb.showinfo, "warning": mb.showwarning, "error": mb.showerror}
            if kind in native and len(buttons) == 1:
                native[kind](title, message)
                return True
            return mb.askyesno(title, message)
        result = {"value": cancel_value}
        top = tk.Toplevel(parent)
        top.withdraw()
        top.title(title)
        top.transient(parent)
        top.resizable(False, False)
        T(top, bg="surface")
        body = T(tk.Frame(top, padx=22, pady=18), bg="surface")
        body.pack(fill="both")
        glyph, tone = self.ICONS.get(kind, self.ICONS["info"])
        T(tk.Label(body, text=glyph, font=("Segoe UI", 13, "bold"), width=2, pady=2), bg=tone + "_soft" if
          tone + "_soft" in THEMES["light"] else "accent_soft", fg=tone).grid(row=0, column=0, sticky="n", padx=(0, 14))
        T(tk.Label(body, text=message, justify="left", wraplength=470, anchor="w", font=("Segoe UI", 9)),
          bg="surface", fg="text").grid(row=0, column=1, sticky="w")
        foot = T(tk.Frame(top, padx=22, pady=12), bg="surface2")
        foot.pack(fill="x")
        focus = None

        def choose(value):
            result["value"] = value
            top.destroy()
        for text, value in reversed(buttons):
            b = ttk.Button(foot, text=text, style="Accent.TButton" if value == default else "TButton",
                           command=lambda v=value: choose(v))
            b.pack(side="right", padx=(8, 0))
            if value == default:
                focus = b
        top.bind("<Return>", lambda e: choose(default))
        top.bind("<Escape>", lambda e: choose(cancel_value))
        top.protocol("WM_DELETE_WINDOW", lambda: choose(cancel_value))
        apply_theme(top)
        center_on(top, parent)
        top.deiconify()
        top.grab_set()
        (focus or top).focus_set()
        parent.wait_window(top)
        return result["value"]

    def showinfo(self, title, message, parent=None, **_):
        return self._box(title, message, "info", [("OK", True)], parent, True, True)

    def showwarning(self, title, message, parent=None, **_):
        return self._box(title, message, "warning", [("OK", True)], parent, True, True)

    def showerror(self, title, message, parent=None, **_):
        return self._box(title, message, "error", [("OK", True)], parent, True, True)

    def askyesno(self, title, message, parent=None, **_):
        return self._box(title, message, "question", [("No", False), ("Yes", True)], parent, True, False)

    def askokcancel(self, title, message, parent=None, icon=None, default=None, **_):
        return self._box(title, message, "warning" if icon == "warning" else "question",
                         [("Cancel", False), ("OK", True)], parent, False if default == "cancel" else True, False)


def run_with_progress(text: str, fn, parent=None):
    """Run fn() in the background with a small "please wait" window (on top of `parent` if given).
    Returns (result, error text or None)."""
    import tkinter as tk
    from tkinter import ttk
    outcome = {}

    def job():
        try:
            outcome["result"] = fn()
        except BaseException as e:
            outcome["error"] = str(e) or e.__class__.__name__

    win = tk.Toplevel(parent) if parent is not None else tk.Tk()
    win.title(APP_NAME)
    win.resizable(False, False)
    if parent is not None:
        win.transient(parent)
    else:
        win.attributes("-topmost", True)
    tk.Label(win, text=text, justify="left", padx=18, pady=14).pack()
    bar = ttk.Progressbar(win, mode="indeterminate", length=320)
    bar.pack(padx=18, pady=(0, 16))
    bar.start(12)
    apply_theme(win)
    if parent is not None:
        center_on(win, parent)
    win.protocol("WM_DELETE_WINDOW", lambda: None)
    thread = threading.Thread(target=job, daemon=True)
    thread.start()

    def wait():
        if thread.is_alive():
            win.after(150, wait)
        else:
            bar.stop()
            win.destroy()
    win.after(150, wait)
    if parent is not None:
        win.grab_set()
        parent.wait_window(win)
    else:
        win.mainloop()
    return outcome.get("result"), outcome.get("error")


def relayout_pdf(path: Path, layout, info=None, copies=None) -> bool:
    """Remake a PDF with a different bleed edge / card spacing (and/or another card list), using the
    card images inside it. The PDF is replaced. False if it can't be done."""
    info = info or read_layout(path)
    if info is None:
        return False
    write_deck_pdf(copies if copies is not None else info["copies"], layout, path)
    log.info("Remade %s with %s", path.name, describe_layout(layout))
    return True


def deck_from_groups(groups, images_dir: Path) -> list:
    """An order's cards (from its XML, images downloaded to images_dir) as an in-memory card list, in
    print order: MDFCs first, then A-Z, one entry per copy."""
    cards = []
    for g in print_order(groups):
        entry = {"front": ("file", images_dir / g["front_file"]),
                 "back": ("file", images_dir / g["back_file"]) if g.get("back_file") else None,
                 "name": g.get("name") or g.get("query") or "", "sort": sort_key(g)[0],
                 "missing": bool(g.get("_missing"))}
        if not g.get("_missing"):              # (placeholders aren't MPCFill art)
            remember_mpc_art(entry["front"], g.get("front_id"), g.get("name"), g.get("query"))
            if entry["back"]:
                remember_mpc_art(entry["back"], g.get("back_id"), g.get("back_name") or g.get("name"),
                                 g.get("back_query") or "")
        cards += [dict(entry) for _ in range(g["count"])]
    return cards


class PageRenderer:
    """Draws page previews straight from the card images, with exactly the PDF's geometry, so the
    print window can show changes (bleed, spacing, added cards) without writing a PDF."""
    def __init__(self):
        self.cache = {}
        self.lock = threading.RLock()     # previews are drawn on a background thread: one user at a time

    def card_image(self, src, width: int, height: int | None = None):
        """A card's image (with its bleed) at width x height px (Standard proportions if no height)."""
        from PIL import Image
        import io
        kind, ref = src
        if height is None:
            height = round(width * (CARD_H + 2 * IMAGE_BLEED) / (CARD_W + 2 * IMAGE_BLEED))
        key = (("file", str(ref)) if kind == "file" else ("pdf", kind), width, height)
        with self.lock:
            if key not in self.cache:
                data = Path(ref).read_bytes() if kind == "file" else ref.get_object()._data
                im = Image.open(io.BytesIO(data))
                im.draft("RGB", (width, height))          # JPEGs decode straight at a smaller size: fast
                self.cache[key] = im.convert("RGB").resize((width, height), Image.BILINEAR)
            return self.cache[key]

    def page(self, cards: list, back: bool, layout, width: int, header: str | None = None):
        """One page (front page, or the back page for the same cards) as a PIL image `width` px wide.
        `header`: the header's text for this page (with its page number), if the layout has one."""
        with self.lock:
            return self._page(cards, back, layout, width, header)

    def _page(self, cards: list, back: bool, layout, width: int, header: str | None = None):
        from PIL import Image
        PAGE_W, PAGE_H = page_size(layout)
        scale = width / PAGE_W
        page = Image.new("RGB", (width, round(PAGE_H * scale)), "white")
        bleed = layout[0] * PT_PER_MM
        card_w, card_h, bleed_x, bleed_y = card_geometry(layout)
        cell_w, cell_h = card_w + 2 * bleed, card_h + 2 * bleed
        img_w = max(2, round((card_w + 2 * bleed_x) * scale))
        img_h = max(2, round((card_h + 2 * bleed_y) * scale))
        off_x, off_y = (bleed_x - bleed) * scale, (bleed_y - bleed) * scale   # the image starts this far outside
        cricut = layout_marks(layout) == "cricut"
        mask = None
        if cricut:                                     # Cricut cards are cut with their rounded corners
            from PIL import ImageDraw
            mw, mh = max(1, round(cell_w * scale)), max(1, round(cell_h * scale))
            mask = Image.new("L", (mw, mh), 0)
            ImageDraw.Draw(mask).rounded_rectangle((0, 0, mw - 1, mh - 1), fill=255,
                                                   radius=max(1, round(corner_radius(layout) * scale)))
        for (x, y), card in zip(card_slots(layout), cards):
            src = card["back"] if back else card["front"]
            if not src:
                continue
            if back:
                x = PAGE_W - x - cell_w                # backs: columns mirrored
            im = self.card_image(src, img_w, img_h)
            box = (round(off_x), round(off_y), round(off_x + cell_w * scale), round(off_y + cell_h * scale))
            piece = im.crop(box)
            at = (round(x * scale), round((PAGE_H - y - cell_h) * scale))
            if mask is not None and mask.size == piece.size:
                page.paste(piece, at, mask)
            else:
                page.paste(piece, at)
        if not back:                                   # cut marks, and where a Cricut sheet ends
            from PIL import ImageDraw
            draw = ImageDraw.Draw(page)
            for mx, my, mw, mh in mark_shapes(layout, len(cards)):
                x0p, y0p = mx * scale, (PAGE_H - my - mh) * scale
                draw.rectangle((round(x0p), round(y0p), max(round(x0p), round(x0p + mw * scale) - 1),
                                max(round(y0p), round(y0p + mh * scale) - 1)), fill="black")
            if cricut:
                ax0, ay0, ax1, ay1 = print_area(layout)
                box = (round(ax0 * scale), round((PAGE_H - ay1) * scale), round(ax1 * scale) - 1,
                       round((PAGE_H - ay0) * scale) - 1)
                for i in range(box[0], box[2], 8):                 # dashed outline (preview only)
                    draw.line((i, box[1], min(i + 4, box[2]), box[1]), fill="#9aa4b2")
                    draw.line((i, box[3], min(i + 4, box[2]), box[3]), fill="#9aa4b2")
                for i in range(box[1], box[3], 8):
                    draw.line((box[0], i, box[0], min(i + 4, box[3])), fill="#9aa4b2")
                    draw.line((box[2], i, box[2], min(i + 4, box[3])), fill="#9aa4b2")
        if header and shows_header(layout):
            from PIL import ImageDraw, ImageFont
            hx, hy, size = header_place(layout)
            px = max(6, round(size * scale))
            try:
                font = ImageFont.truetype("arial.ttf", px)
            except OSError:
                try:
                    font = ImageFont.load_default(size=px)
                except TypeError:
                    font = ImageFont.load_default()
            ImageDraw.Draw(page).text((round(hx * scale), round((PAGE_H - hy) * scale)), header, fill="black",
                                      font=font, anchor="ls")
        return page


def deck_pages(cards: list, layout) -> list[list]:
    """The deck split into pages (one list of cards per sheet)."""
    per = len(card_slots(layout))
    return [cards[i:i + per] for i in range(0, len(cards), per)] or [[]]


def ask_custom_card(parent, current=None):
    """Ask for a custom card size (width and height in mm, without bleed). Returns "WxH" or None."""
    import tkinter as tk
    from tkinter import ttk
    w0, h0 = custom_card_mm(current) or tuple(round(v / PT_PER_MM, 1)
                                              for v in card_dims((0, 0, None, None, None, None, current)))
    result = {}
    win = tk.Toplevel(parent)
    win.title("Custom card size")
    win.transient(parent)
    win.resizable(False, False)
    body = ttk.Frame(win, padding=16)
    body.pack(fill="both", expand=True)
    ttk.Label(body, text="The card's size without bleed, in mm.\nCard images are stretched to fit it.",
              justify="left").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
    w_var, h_var = tk.StringVar(value=f"{w0:g}"), tk.StringVar(value=f"{h0:g}")
    ttk.Label(body, text="Width").grid(row=1, column=0, sticky="w")
    w_entry = ttk.Entry(body, textvariable=w_var, width=7, justify="center")
    w_entry.grid(row=1, column=1, padx=(6, 14))
    ttk.Label(body, text="Height").grid(row=1, column=2, sticky="w")
    ttk.Entry(body, textvariable=h_var, width=7, justify="center").grid(row=1, column=3, padx=(6, 0))
    error = ttk.Label(body, text="", foreground="#c0392b")
    error.grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))
    buttons = ttk.Frame(body)
    buttons.grid(row=3, column=0, columnspan=4, sticky="e", pady=(10, 0))

    def ok(_event=None):
        value = f"{w_var.get().replace(',', '.').strip()}x{h_var.get().replace(',', '.').strip()}"
        mm = custom_card_mm(value)
        if not mm:
            error.configure(text="Width and height must be numbers between 10 and 400 mm.")
            return
        result["value"] = f"{mm[0]:g}x{mm[1]:g}"
        win.destroy()
    ttk.Button(buttons, text="OK", style="Accent.TButton", command=ok).pack(side="right")
    ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="right", padx=(0, 6))
    win.bind("<Return>", ok)
    win.bind("<Escape>", lambda e: win.destroy())
    try:
        apply_theme(win)
    except Exception:
        pass
    center_on(win, parent)
    w_entry.focus_set()
    w_entry.select_range(0, "end")
    win.grab_set()
    parent.wait_window(win)
    return result.get("value")


def ask_layout(names: list[str], layout):
    """Small window with the bleed edge, card spacing and paper boxes, shown before making a PDF from
    XMLs. Returns the (bleed, spacing, paper) to use, or None if cancelled."""
    import tkinter as tk
    from tkinter import ttk
    result = {}
    root = tk.Tk()
    root.title(APP_NAME)
    root.attributes("-topmost", True)
    root.resizable(False, False)
    listed = "\n".join(names[:6]) + ("\n…" if len(names) > 6 else "")
    intro = (f"{UNNAMED_LAYOUT_NOTE}, so these are the values you used last time." if UNNAMED_LAYOUT_NOTE else
             "The bleed edge, card spacing and paper for this PDF (what you used last time).")
    tk.Label(root, text=f"{listed}\n\n{intro}\nCheck them, then press Make PDF.", justify="left", padx=16,
             pady=12).pack(anchor="w")
    box = tk.Frame(root, padx=16)
    box.pack(fill="x")
    bleed_var, spacing_var = tk.StringVar(value=f"{layout[0]:.2f}"), tk.StringVar(value=f"{layout[1]:.2f}")
    tk.Label(box, text="Bleed edge (mm):").grid(row=0, column=0, sticky="w", pady=2)
    ttk.Entry(box, textvariable=bleed_var, width=8).grid(row=0, column=1, sticky="w", padx=6)
    tk.Label(box, text="Card spacing (mm):").grid(row=1, column=0, sticky="w", pady=2)
    ttk.Entry(box, textvariable=spacing_var, width=8).grid(row=1, column=1, sticky="w", padx=6)
    tk.Label(box, text="Paper:").grid(row=2, column=0, sticky="w", pady=2)
    paper_var = tk.StringVar(value=layout_paper(layout))
    ttk.Combobox(box, textvariable=paper_var, values=list(PAPER_SIZES), width=10, state="readonly").grid(
        row=2, column=1, sticky="w", padx=6)
    tk.Label(box, text="Card size:").grid(row=3, column=0, sticky="w", pady=2)
    card_choices = {card_size_label(n): n for n in CARD_SIZES}
    card_choices.setdefault(card_size_label(layout_card(layout)), layout_card(layout))   # a custom size
    card_var = tk.StringVar(value=card_size_label(layout_card(layout)))
    ttk.Combobox(box, textvariable=card_var, values=list(card_choices), width=24, state="readonly").grid(
        row=3, column=1, sticky="w", padx=6)
    tk.Label(box, text=LAYOUT_HINT, fg="#666666").grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))
    error = tk.Label(root, fg="#c0392b", padx=16)
    error.pack(anchor="w")
    buttons = tk.Frame(root, padx=16, pady=12)
    buttons.pack(fill="x")

    def ok(_event=None):
        value = parse_layout(bleed_var.get(), spacing_var.get())
        if isinstance(value, str):
            error.configure(text=value)
            return
        result["layout"] = value + (paper_var.get(), layout_orientation(layout), layout_header(layout),
                                    layout_numbering(layout), card_choices.get(card_var.get(), DEFAULT_CARD),
                                    layout_marks(layout))
        root.destroy()
    ttk.Button(buttons, text="Make PDF", command=ok, style="Accent.TButton").pack(side="right")
    ttk.Button(buttons, text="Cancel", command=root.destroy).pack(side="right", padx=(0, 6))
    apply_theme(root)
    root.bind("<Return>", ok)
    root.bind("<Escape>", lambda e: root.destroy())
    root.mainloop()
    return result.get("layout")


def installed_printers() -> set[str]:
    try:
        import ctypes
        from ctypes import wintypes

        class PRINTER_INFO_4W(ctypes.Structure):
            _fields_ = [("pPrinterName", wintypes.LPWSTR), ("pServerName", wintypes.LPWSTR),
                        ("Attributes", wintypes.DWORD)]
        winspool = ctypes.WinDLL("winspool.drv")
        flags = 0x2 | 0x4   # PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS
        needed, count = wintypes.DWORD(), wintypes.DWORD()
        winspool.EnumPrintersW(flags, None, 4, None, 0, ctypes.byref(needed), ctypes.byref(count))
        buf = ctypes.create_string_buffer(needed.value or 1)
        if not winspool.EnumPrintersW(flags, None, 4, buf, needed, ctypes.byref(needed), ctypes.byref(count)):
            return set()
        infos = ctypes.cast(buf, ctypes.POINTER(PRINTER_INFO_4W))
        return {infos[i].pPrinterName for i in range(count.value)}
    except Exception as e:
        log.warning("Couldn't list printers: %s", e)
        return set()




def default_printer() -> str | None:
    """The Windows default printer's name."""
    try:
        import ctypes
        from ctypes import wintypes
        winspool = ctypes.WinDLL("winspool.drv")
        size = wintypes.DWORD(0)
        winspool.GetDefaultPrinterW(None, ctypes.byref(size))
        buf = ctypes.create_unicode_buffer(size.value or 1)
        if winspool.GetDefaultPrinterW(buf, ctypes.byref(size)):
            return buf.value or None
    except Exception:
        pass
    return None


LAST_PRINTER_FILE = HERE / "last_printer.json"


def printer_for(available: set[str], profile: str | None = None) -> str | None:
    """The printer to suggest: the one you printed on last time, if it's still installed; otherwise
    the Windows default printer."""
    try:
        last = json.loads(LAST_PRINTER_FILE.read_text(encoding="utf-8")).get("printer")
        if last and (not available or last in available):
            return last
    except Exception:
        pass
    return default_printer() or (sorted(available)[0] if available else None)


def remember_printer(printer: str) -> None:
    try:
        LAST_PRINTER_FILE.write_text(json.dumps({"printer": printer}), encoding="utf-8")
    except OSError:
        pass


PRINTER_SETTINGS_DIR = HERE / "printer_settings"


def _document_properties(printer: str, saved: bytes | None, hwnd=None):
    """Windows' DocumentProperties: the printer's settings for this user, merged with `saved` (an
    earlier copy) if given; with `hwnd`, shows the printer's own settings window first (the
    window with Printing Presets). Returns the settings, or None if cancelled / unavailable."""
    import ctypes
    from ctypes import wintypes
    winspool = ctypes.WinDLL("winspool.drv")
    winspool.OpenPrinterW.argtypes = (wintypes.LPWSTR, ctypes.POINTER(wintypes.HANDLE), ctypes.c_void_p)
    winspool.DocumentPropertiesW.argtypes = (wintypes.HWND, wintypes.HANDLE, wintypes.LPWSTR,
                                             ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD)
    winspool.DocumentPropertiesW.restype = ctypes.c_long
    winspool.ClosePrinter.argtypes = (wintypes.HANDLE,)
    DM_OUT_BUFFER, DM_IN_PROMPT, DM_IN_BUFFER, IDOK = 2, 4, 8, 1
    handle = wintypes.HANDLE()
    if not winspool.OpenPrinterW(printer, ctypes.byref(handle), None):
        return None
    try:
        size = winspool.DocumentPropertiesW(None, handle, printer, None, None, 0)
        if size <= 0:
            return None
        out = ctypes.create_string_buffer(max(size, len(saved or b"")) + 1024)
        flags = DM_OUT_BUFFER
        given = None
        if saved:
            given = ctypes.create_string_buffer(saved, len(saved))
            flags |= DM_IN_BUFFER
        if hwnd is not None:
            flags |= DM_IN_PROMPT
        result = winspool.DocumentPropertiesW(hwnd, handle, printer, out, given, flags)
        if (hwnd is not None and result != IDOK) or result < 0:
            return None
        return out
    finally:
        winspool.ClosePrinter(handle)


def icm_method(devmode) -> int | None:
    """dmICMMethod of a DEVMODEW (1 = none, 2 = Windows does colour matching, 3 = driver, 4 = device)."""
    import ctypes
    from ctypes import wintypes

    U16, U32, I16 = ctypes.c_uint16, ctypes.c_uint32, ctypes.c_int16    # Windows sizes on any system

    class DEVMODEW(ctypes.Structure):
        _fields_ = [("dmDeviceName", U16 * 32), ("dmSpecVersion", U16), ("dmDriverVersion", U16),
                    ("dmSize", U16), ("dmDriverExtra", U16), ("dmFields", U32), ("dmUnion", ctypes.c_byte * 16),
                    ("dmColor", I16), ("dmDuplex", I16), ("dmYResolution", I16), ("dmTTOption", I16),
                    ("dmCollate", I16), ("dmFormName", U16 * 32), ("dmLogPixels", U16), ("dmBitsPerPel", U32),
                    ("dmPelsWidth", U32), ("dmPelsHeight", U32), ("dmNup", U32), ("dmDisplayFrequency", U32),
                    ("dmICMMethod", U32), ("dmICMIntent", U32), ("dmMediaType", U32), ("dmDitherType", U32)]
    DM_ICMMETHOD = 0x00800000
    try:
        if devmode is None or len(devmode) < ctypes.sizeof(DEVMODEW):
            return None
        dm = DEVMODEW.from_buffer_copy(bytes(devmode)[:ctypes.sizeof(DEVMODEW)])
        return dm.dmICMMethod if dm.dmFields & DM_ICMMETHOD else None
    except Exception:
        return None


def saved_settings_file(printer: str, profile: str | None = None) -> Path:
    """The settings picked with Printer settings... are remembered per printer (an add-on can keep
    them per profile too)."""
    return PRINTER_SETTINGS_DIR / f"{safe_name(printer)}.devmode"


def has_saved_settings(printer: str, profile: str | None = None) -> bool:
    return saved_settings_file(printer, profile).exists()


def job_settings(printer: str, profile: str | None = None):
    """The printer settings for a print job: the ones picked with Printer settings... for this printer
    if there are any, otherwise the printer's own defaults."""
    file = saved_settings_file(printer, profile)
    saved = file.read_bytes() if file.exists() else None
    try:
        devmode = _document_properties(printer, saved)
    except Exception as e:
        log.warning("Couldn't read %s's settings: %s", printer, e)
        return None
    if saved and devmode is None:
        log.warning("Saved printer settings %s couldn't be used; using the printer's defaults", file.name)
        devmode = _document_properties(printer, None)
    return devmode


def choose_printer_settings(printer: str, window, profile: str | None = None) -> bool:
    """Open the printer's own settings window (paper, quality, a saved preset, colour...) starting from
    the saved settings, and remember the result for this printer."""
    import ctypes
    file = saved_settings_file(printer, profile)
    saved = file.read_bytes() if file.exists() else None
    try:
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id()) or window.winfo_id()
        devmode = _document_properties(printer, saved, hwnd=hwnd)
    except Exception as e:
        show_error(f"Couldn't open {printer}'s settings: {e}")
        return False
    if devmode is None:
        return False
    PRINTER_SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
    file.write_bytes(devmode.raw)
    log.info("Saved printer settings for %s%s", printer, f" ({profile})" if profile else "")
    return True


def page_placement(page_w: float, page_h: float, paper_w: int, paper_h: int, dpi_x: int, dpi_y: int,
                   rotate: str = "counterclockwise") -> tuple[int, int, int, int, int]:
    """Where a PDF page (points) goes on the paper (printer pixels), like a PDF reader's Actual size +
    Auto-rotate + Auto-center: a portrait page on paper loaded sideways (landscape) is turned, and the
    page is centred on the paper. Returns (rotation for pypdfium2 in degrees clockwise,
    drawn width, drawn height, left, top) - sizes and offsets in printer pixels from the paper's corner."""
    turn = (page_w > page_h) != (paper_w > paper_h) and abs(page_w - page_h) > 1
    rotation = (90 if rotate == "clockwise" else 270) if turn else 0
    w_pt, h_pt = (page_h, page_w) if turn else (page_w, page_h)
    draw_w, draw_h = round(w_pt / 72 * dpi_x), round(h_pt / 72 * dpi_y)
    return rotation, draw_w, draw_h, round((paper_w - draw_w) / 2), round((paper_h - draw_h) / 2)


def offset_on_paper(offset_mm, rotation: int):
    """A back alignment offset (X right, Y down, as you look at the printed back page) turned into the
    printer's directions when the page is printed turned (sideways paper)."""
    x, y = offset_mm
    if rotation == 90:          # page turned clockwise: page right -> paper down, page down -> paper left
        return (-y, x)
    if rotation == 270:         # turned counter-clockwise: page right -> paper up, page down -> paper right
        return (y, -x)
    return (x, y)


ALIGNMENT_FILE = HERE / "back_alignment.json"


def load_alignment(printer: str) -> dict:
    """{"enabled", "x_mm", "y_mm"} for a printer (off and 0 if never set)."""
    try:
        d = json.loads(ALIGNMENT_FILE.read_text(encoding="utf-8")).get(printer, {})
        return {"enabled": bool(d.get("enabled")), "x_mm": float(d.get("x_mm", 0)), "y_mm": float(d.get("y_mm", 0))}
    except Exception:
        return {"enabled": False, "x_mm": 0.0, "y_mm": 0.0}


def save_alignment(printer: str, enabled: bool, x_mm: float, y_mm: float) -> None:
    try:
        data = json.loads(ALIGNMENT_FILE.read_text(encoding="utf-8")) if ALIGNMENT_FILE.exists() else {}
    except Exception:
        data = {}
    data[printer] = {"enabled": enabled, "x_mm": x_mm if enabled else 0.0, "y_mm": y_mm if enabled else 0.0}
    try:
        ALIGNMENT_FILE.write_text(json.dumps(data, indent=1), encoding="utf-8")
    except OSError as e:
        log.warning("Couldn't save the back alignment: %s", e)


def back_offset_for(printer: str):
    a = load_alignment(printer)
    return (a["x_mm"], a["y_mm"]) if a["enabled"] else (0.0, 0.0)


def write_alignment_test(out: Path, layout=None) -> None:
    """A 2-page test sheet: thin crosses on the front (at every card position and the middle), bold
    crosses on the back at the matching (mirrored) spots, with a ruler (0.5 mm marks) and +X / +Y arrows
    around the middle one. Printed front, then back on the same sheet; held up to a light it shows how
    far the backs land from the fronts."""
    PAGE_W, PAGE_H = page_size(layout)
    import zlib
    mm = PT_PER_MM
    card_w, card_h = card_dims(layout)
    centres = [(x + (card_w + 2 * mm) / 2, y + (card_h + 2 * mm) / 2)
               for x, y in card_slots((1.0, 0.0, layout_paper(layout), layout_orientation(layout), None, "all", layout_card(layout)))]
    mid = (PAGE_W / 2, PAGE_H / 2)

    def cross(x, y, arm, width):
        return f"{width} w {x - arm:.2f} {y:.2f} m {x + arm:.2f} {y:.2f} l {x:.2f} {y - arm:.2f} m {x:.2f} {y + arm:.2f} l S\n"

    def text(x, y, size, words):
        safe = words.replace("\\", "").replace("(", "[").replace(")", "]")
        return f"BT /F1 {size} Tf {x:.2f} {y:.2f} Td ({safe}) Tj ET\n"
    front = "0 0 0 RG 0 0 0 rg\n" + "".join(cross(x, y, 10 * mm, 0.25) for x, y in centres + [mid])
    front += text(40, PAGE_H - 50, 14, "FRONT - Fleem back alignment test")
    front += text(40, PAGE_H - 68, 9, "Put this sheet back in the printer the way you do for card backs, then press OK in Fleem.")
    back = "0 0 0 RG 0 0 0 rg\n" + "".join(cross(PAGE_W - x, y, 10 * mm, 1.2) for x, y in centres + [mid])
    cx, cy = mid
    for i in range(-10, 11):                      # ruler on the middle cross's arms: 0.5 mm marks, -5 .. +5 mm
        if not i:
            continue
        d = i * 0.5 * mm
        half = (1.6 if i % 2 == 0 else 0.9) * mm
        back += f"0.35 w {cx + d:.2f} {cy - half:.2f} m {cx + d:.2f} {cy + half:.2f} l S\n"     # marks on the X arm
        back += f"0.35 w {cx - half:.2f} {cy - d:.2f} m {cx + half:.2f} {cy - d:.2f} l S\n"     # marks on the Y arm
        if i % 4 == 0:                            # numbers at 2 and 4 mm (count the marks between)
            back += text(cx + d - 3, cy - 4.5 * mm, 5, f"{i // 2:+d}")                       # X: + = right
            back += text(cx + 7 * mm, cy - d - 1.7, 5, f"{i // 2:+d}")                       # Y: + = down
    back += text(cx + 11 * mm, cy - 1.5, 7, "+X ->")
    back += text(cx - 3.5, cy - 12.5 * mm, 7, "+Y")
    back += text(40, PAGE_H - 50, 14, "BACK - hold the sheet up to a light")
    back += text(40, PAGE_H - 68, 9, "The thin crosses (front) show through. How far must the bold cross move to sit on the thin one?")
    back += text(40, PAGE_H - 82, 9, "Right = +X, down = +Y (use the ruler). Add that to the offsets in Fleem, then print the test again.")
    body = bytearray(b"%PDF-1.4\n")
    offsets = []

    def add(obj: bytes):
        offsets.append(len(body))
        body.extend(f"{len(offsets)} 0 obj\n".encode() + obj + b"\nendobj\n")
    add(b"<< /Type /Catalog /Pages 2 0 R >>")
    add(b"<< /Type /Pages /Kids [3 0 R 5 0 R] /Count 2 >>")
    for i, content in enumerate((front, back)):
        packed = zlib.compress(content.encode("latin-1"))
        add(f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PAGE_W:.4f} {PAGE_H:.4f}] /Resources << /Font << /F1 7 0 R >> >> "
            f"/Contents {4 + 2 * i} 0 R >>".encode())
        add(b"<< /Filter /FlateDecode /Length %d >>\nstream\n" % len(packed) + packed + b"\nendstream")
    add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    xref = len(body)
    body.extend(b"xref\n0 %d\n0000000000 65535 f \n" % (len(offsets) + 1))
    body.extend(b"".join(b"%010d 00000 n \n" % o for o in offsets))
    body.extend(b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(offsets) + 1, xref))
    Path(out).write_bytes(bytes(body))


def gdi_print(path: Path, printer: str, indexes: list[int], job_name: str, progress=None,
              rotate: str = "counterclockwise", devmode=None, back_pages=(), back_offset=(0.0, 0.0)) -> None:
    """Print pages of a PDF (0 = first) on a Windows printer at exactly 100% size, with that printer's
    saved settings (its preset: paper size, quality...). Each page is centred on the paper and turned
    if the paper is loaded the other way round. `progress(done, total)` is called per page."""
    import ctypes
    from ctypes import wintypes
    import pypdfium2 as pdfium
    gdi32 = ctypes.windll.gdi32
    gdi32.CreateDCW.restype = wintypes.HDC
    gdi32.CreateDCW.argtypes = (wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p)

    class DOCINFOW(ctypes.Structure):
        _fields_ = [("cbSize", ctypes.c_int), ("lpszDocName", wintypes.LPCWSTR),
                    ("lpszOutput", wintypes.LPCWSTR), ("lpszDatatype", wintypes.LPCWSTR),
                    ("fwType", wintypes.DWORD)]

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
                    ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
                    ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
                    ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]

    LOGPIXELSX, LOGPIXELSY, PHYSICALWIDTH, PHYSICALHEIGHT, PHYSICALOFFSETX, PHYSICALOFFSETY = 88, 90, 110, 111, 112, 113
    if devmode is None:
        devmode = job_settings(printer)         # the saved settings, or the printer's own defaults
    hdc = gdi32.CreateDCW("WINSPOOL", printer, None, devmode)
    if not hdc:
        raise RuntimeError(f'Windows couldn\'t open the printer "{printer}".')
    H = wintypes.HDC(hdc)
    doc = pdfium.PdfDocument(str(path))
    started = False
    try:
        dpi_x, dpi_y = gdi32.GetDeviceCaps(H, LOGPIXELSX), gdi32.GetDeviceCaps(H, LOGPIXELSY)
        paper_w, paper_h = gdi32.GetDeviceCaps(H, PHYSICALWIDTH), gdi32.GetDeviceCaps(H, PHYSICALHEIGHT)
        off_x, off_y = gdi32.GetDeviceCaps(H, PHYSICALOFFSETX), gdi32.GetDeviceCaps(H, PHYSICALOFFSETY)
        log.info("Printer %s: paper %.1f x %.1f mm at %d x %d dpi", printer,
                 paper_w / dpi_x * 25.4, paper_h / dpi_y * 25.4, dpi_x, dpi_y)
        render_dpi = min(dpi_x, dpi_y, 1200)
        info = DOCINFOW(ctypes.sizeof(DOCINFOW), job_name[:200], None, None, 0)
        if gdi32.StartDocW(H, ctypes.byref(info)) <= 0:
            raise RuntimeError(f'The printer "{printer}" didn\'t accept the print job.')
        started = True
        # Colour: the printer settings decide. colour adjustments in the driver are applied by the
        # driver itself; only when the settings ask Windows to do colour matching (ICM handled by the
        # system) is Windows colour management switched on - never both.
        if icm_method(devmode) == 2:       # DMICMMETHOD_SYSTEM
            gdi32.SetICMMode(H, 2)         # ICM_ON
        # Pixels go 1:1 to the printer unless it's above 1200 dpi; only then smooth the scaling.
        gdi32.SetStretchBltMode(H, 3 if render_dpi == min(dpi_x, dpi_y) else 4)   # COLORONCOLOR / HALFTONE
        for done, index in enumerate(indexes, 1):
            page = doc[index]
            page_w, page_h = page.get_size()                  # points
            rotation, draw_w, draw_h, left, top_px = page_placement(page_w, page_h, paper_w, paper_h,
                                                                    dpi_x, dpi_y, rotate)
            if index in back_pages and any(back_offset):      # back alignment: move this back page
                sx, sy = offset_on_paper(back_offset, rotation)
                left += round(sx / 25.4 * dpi_x)
                top_px += round(sy / 25.4 * dpi_y)
            scale = render_dpi / 72
            out_w_pt, out_h_pt = (page_h, page_w) if rotation in (90, 270) else (page_w, page_h)
            full_w, full_h = round(out_w_pt * scale), round(out_h_pt * scale)
            if gdi32.StartPage(H) <= 0:
                raise RuntimeError("The printer stopped the job.")
            band = max(64, min(full_h, (48 * 1024 * 1024) // ((full_w + 1) * 4)))   # rows per strip (~48 MB)
            row = 0
            while row < full_h:
                rows = min(band, full_h - row)
                with PDFIUM_LOCK:
                    bmp = page.render(scale=scale, rotation=rotation,
                                      crop=(0, (full_h - row - rows) / scale, 0, row / scale),
                                      prefer_bgrx=True, fill_color=(255, 255, 255, 255))
                w, h = bmp.width, bmp.height
                if bmp.stride != w * 4:
                    raise RuntimeError("Unexpected image layout while printing.")
                header = BITMAPINFOHEADER(ctypes.sizeof(BITMAPINFOHEADER), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
                bits = (ctypes.c_char * (bmp.stride * h)).from_buffer_copy(bytes(bmp.buffer))
                dest_x = left - off_x
                dest_y = top_px + round(row * dpi_y / render_dpi) - off_y
                gdi32.StretchDIBits(H, dest_x, dest_y, round(w * dpi_x / render_dpi), round(h * dpi_y / render_dpi),
                                    0, 0, w, h, bits, ctypes.byref(header), 0, 0x00CC0020)   # SRCCOPY
                row += rows
            if gdi32.EndPage(H) <= 0:
                raise RuntimeError("The printer stopped the job.")
            if progress:
                progress(done, len(indexes))
        gdi32.EndDoc(H)
        started = False
    finally:
        if started:
            gdi32.AbortDoc(H)
        gdi32.DeleteDC(H)
        doc.close()


def send_to_printer(settings: "Settings", path: Path, printer: str, indexes: list[int], what: str,
                    back_pages=(), profile: str | None = None) -> bool:
    """Print pages of a PDF on a printer, showing a small "Printing…" window. True if it all went."""
    import tkinter as tk
    from tkinter import ttk
    available = installed_printers()
    if available and printer not in available:
        show_error(f'There\'s no printer called "{printer}" in Windows, so {path.name} wasn\'t printed.')
        return False
    log.info("Printing %s %s on %s with %s", path.name, what, printer,
             f"your saved {profile or 'printer'} settings" if has_saved_settings(printer, profile)
             else "the printer's default settings")
    remember_printer(printer)
    state = {"done": 0, "total": len(indexes)}

    def work():
        try:
            gdi_print(path, printer, indexes, f"{path.stem} ({what})",
                      progress=lambda d, t: state.update(done=d), rotate=settings.rotate_pages,
                      devmode=job_settings(printer, profile), back_pages=set(back_pages),
                      back_offset=back_offset_for(printer))
        except BaseException as e:
            state["error"] = str(e) or e.__class__.__name__

    root = tk.Tk()
    root.title(APP_NAME)
    root.resizable(False, False)
    root.attributes("-topmost", True)
    label = tk.Label(root, justify="left", padx=18, pady=14)
    label.pack()
    bar = ttk.Progressbar(root, mode="determinate", length=320, maximum=max(1, len(indexes)))
    bar.pack(padx=18, pady=(0, 16))
    apply_theme(root)
    thread = threading.Thread(target=work, daemon=True)
    thread.start()

    def tick():
        label.configure(text=f"Sending {path.name} to {printer}…\npage {min(state['done'] + 1, state['total'])} "
                             f"of {state['total']}")
        bar.configure(value=state["done"])
        if thread.is_alive():
            root.after(250, tick)
        else:
            root.destroy()
    root.protocol("WM_DELETE_WINDOW", lambda: None)   # let the job finish
    root.after(0, tick)
    root.mainloop()
    if "error" in state:
        log.error("Printing %s failed: %s", path.name, state["error"])
        show_error(f"Couldn't print {path.name}:\n{state['error']}")
        return False
    log.info("%s: sent %d page(s) to %s", path.name, len(indexes), printer)
    return True


def parse_fronts(text: str, count: int) -> set[int] | None:
    """'1, 5-8, 10-11' -> {1, 5, 6, 7, 8, 10, 11}. None if it isn't a valid list of fronts 1..count."""
    text = re.sub(r"\s*(?:-|–|—|\bto\b|\bthrough\b|\bthru\b)\s*", "-", text.lower().strip())
    chosen = set()
    for part in filter(None, re.split(r"[,;\s]+", text)):
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", part)
        if not m:
            return None
        a, b = int(m[1]), int(m[2] or m[1])
        if a > b:
            a, b = b, a
        if a < 1 or b > count:
            return None
        chosen.update(range(a, b + 1))
    return chosen


def describe_fronts(fronts) -> str:
    """{1, 5, 6, 7, 8, 10, 11} -> '1, 5-8, 10-11'."""
    out, run = [], []
    for n in sorted(fronts):
        if run and n == run[-1] + 1:
            run.append(n)
            continue
        if run:
            out.append(str(run[0]) if len(run) == 1 else f"{run[0]}-{run[-1]}")
        run = [n]
    if run:
        out.append(str(run[0]) if len(run) == 1 else f"{run[0]}-{run[-1]}")
    return ", ".join(out)


def pdf_page_count(path: Path) -> int:
    from pypdf import PdfReader
    return len(PdfReader(str(path)).pages)


def enable_file_drop(window, on_files) -> bool:
    """Let files be dragged from File Explorer onto a Tk window (Windows only). on_files(list of
    paths) runs in Tk's thread. Returns False where it isn't available."""
    if os.name != "nt":
        return False
    try:
        import ctypes
        from ctypes import wintypes
        user32, shell32 = ctypes.windll.user32, ctypes.windll.shell32
        LRESULT = ctypes.c_ssize_t
        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
        set_proc = getattr(user32, "SetWindowLongPtrW", user32.SetWindowLongW)
        set_proc.argtypes = (wintypes.HWND, ctypes.c_int, ctypes.c_void_p)
        set_proc.restype = ctypes.c_void_p
        user32.CallWindowProcW.argtypes = (ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM,
                                           wintypes.LPARAM)
        user32.CallWindowProcW.restype = LRESULT
        shell32.DragQueryFileW.argtypes = (ctypes.c_void_p, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT)
        shell32.DragQueryFileW.restype = wintypes.UINT
        shell32.DragFinish.argtypes = (ctypes.c_void_p,)
        shell32.DragAcceptFiles.argtypes = (wintypes.HWND, wintypes.BOOL)
        GWLP_WNDPROC, WM_DROPFILES = -4, 0x0233
        window.update_idletasks()
        hwnd = window.winfo_id()
        state = {}
        dropped = []          # filled by the window procedure, emptied by poll() in Tk's own loop

        def proc(h, msg, wparam, lparam):
            # Runs inside Windows' message delivery, i.e. in the middle of Tk's event loop: calling
            # anything in Tk from here (even window.after) corrupts tkinter's state and crashes Fleem
            # outright. So only note the file names; poll() hands them over.
            if msg == WM_DROPFILES:
                try:
                    for i in range(shell32.DragQueryFileW(wparam, 0xFFFFFFFF, None, 0)):
                        length = shell32.DragQueryFileW(wparam, i, None, 0)
                        buf = ctypes.create_unicode_buffer(length + 1)
                        shell32.DragQueryFileW(wparam, i, buf, length + 1)
                        dropped.append(buf.value)
                finally:
                    shell32.DragFinish(wparam)
                return 0
            return user32.CallWindowProcW(state["old"], h, msg, wparam, lparam)

        def poll():
            if not state.get("old"):
                return
            if dropped:
                files = dropped[:]
                del dropped[:]
                try:
                    on_files(files)
                except Exception:
                    log.exception("Couldn't add the dropped files")
            try:
                window.after(150, poll)
            except Exception:
                pass
        callback = WNDPROC(proc)
        state["old"] = set_proc(hwnd, GWLP_WNDPROC, ctypes.cast(callback, ctypes.c_void_p))
        shell32.DragAcceptFiles(hwnd, True)
        keep = getattr(window, "_fleem_drops", [])
        keep.append((callback, state))          # the callback must stay alive while the window does

        window._fleem_drops = keep

        def unhook(event):
            if event.widget is window and state.get("old"):
                try:
                    shell32.DragAcceptFiles(hwnd, False)
                    set_proc(hwnd, GWLP_WNDPROC, state["old"])
                except Exception:
                    pass
                state["old"] = None
        window.bind("<Destroy>", unhook, add="+")
        window.after(150, poll)
        return True
    except Exception as e:
        log.info("Drag and drop isn't available: %s", e)
        return False


# Print windows that are open: a PDF or XML opened with Fleem meanwhile (right-click > Print with Fleem,
# Open with) goes into the newest one, as if it had been dragged onto it.
_OPEN_PROJECTS = []
_OPEN_PROJECTS_LOCK = threading.Lock()


def hand_to_open_project(files) -> bool:
    """Give files to the print window that's open (it asks before adding them). False if none can take them."""
    with _OPEN_PROJECTS_LOCK:
        for inbox in reversed(_OPEN_PROJECTS):
            if inbox["editable"]:
                inbox["files"].extend(files)
                return True
    return False


def insert_added_cards(copies: list, added: list) -> list:
    """Put added cards into a PDF's card list: double-faced cards (with a back) go with the other MDFCs
    at the front, everything else alphabetically among the rest - or at the end of its group when the
    PDF doesn't know its cards' names (made before Fleem saved them). With Only Alphabetical, every card
    goes in its A-Z spot. The other cards stay where they are (even ones you moved)."""
    out = list(copies)
    for card in added:
        key = card.get("sort") or (sort_key({"name": card["name"]})[0] if card.get("name") else "")
        if mdfcs_first():
            mdfc_end = max((i + 1 for i, c in enumerate(out) if c["back"]), default=0)
            block = range(0, mdfc_end) if card["back"] else range(mdfc_end, len(out))
        else:
            block = range(0, len(out))
        pos = block.stop
        named = [i for i in block if out[i].get("sort")]
        if key and named:                            # A-Z among the named cards (unnamed ones are skipped)
            pos = next((i for i in named if out[i]["sort"] > key), named[-1] + 1)
        as_src = lambda x: x if isinstance(x, tuple) else ("file", x)     # a path, or an image already in use
        entries = [{"front": as_src(card["front"]), "back": as_src(card["back"]) if card["back"] else None,
                    "name": card["name"], "sort": key, "missing": card.get("missing", False)}
                   for _ in range(card["copies"])]
        out[pos:pos] = entries
    return out


def _src_id(src):
    return None if not src else (("file", str(src[1])) if src[0] == "file" else ("pdf", src[0]))


def group_cards(cards: list) -> list:
    """The deck as distinct cards (same front, back and name = one card) in order, with copies."""
    groups, index = [], {}
    for c in cards:
        key = (_src_id(c["front"]), _src_id(c["back"]), c.get("name") or "")
        if key not in index:
            index[key] = len(groups)
            groups.append({"front": c["front"], "back": c["back"], "name": c.get("name") or "",
                           "sort": c.get("sort"), "missing": bool(c.get("missing")), "copies": 0,
                           "was_mdfc": bool(c["back"])})
        groups[index[key]]["copies"] += 1
    return groups


def cards_from_groups(groups: list) -> list:
    """Back from distinct cards (after Edit cards) to one entry per copy. Cards that kept their
    front/back kind stay where they were; a card that got or lost a back moves (MDFCs go with the
    MDFCs, others alphabetically, like added cards); removed cards (0 copies) are dropped."""
    kept, moved = [], []
    for g in groups:
        if g["copies"] <= 0:
            continue
        if (bool(g["back"]) != g["was_mdfc"] and mdfcs_first()) or g.get("parked") or g.get("renamed"):
            moved.append(g)
            continue
        kept += [{"front": g["front"], "back": g["back"], "name": g["name"], "sort": g["sort"],
                  "missing": g["missing"]} for _ in range(g["copies"])]
    return insert_added_cards(kept, moved)


def card_key(card) -> tuple:
    """Which distinct card this is (the same front, back and name = the same card)."""
    return (_src_id(card["front"]), _src_id(card["back"]), card.get("name") or "")


def card_face(renderer, src, width: int):
    """A card as it looks once cut: its picture without the bleed, corners rounded (transparent), `width`
    px wide (RGBA)."""
    from PIL import Image, ImageDraw
    full_w = max(2, round(width * (CARD_W + 2 * IMAGE_BLEED) / CARD_W))
    im = renderer.card_image(src, full_w)
    bx = round(IMAGE_BLEED / (CARD_W + 2 * IMAGE_BLEED) * im.width)
    by = round(IMAGE_BLEED / (CARD_H + 2 * IMAGE_BLEED) * im.height)
    face = im.crop((bx, by, im.width - bx, im.height - by)).convert("RGBA")
    big = Image.new("L", (face.width * 4, face.height * 4), 0)         # drawn big, then shrunk: smooth corners
    ImageDraw.Draw(big).rounded_rectangle((0, 0, big.width - 1, big.height - 1), radius=round(big.width * 0.05),
                                          fill=255)
    face.putalpha(big.resize(face.size, Image.LANCZOS))
    return face


def face_with_back(front, back, border=(26, 29, 36), margin: int = 7):
    """A card face (card_face) with a small picture of its back in the bottom-left corner - one image,
    so it can't come apart when the Edit cards grid scrolls."""
    from PIL import Image, ImageDraw
    out = front.copy()
    pad = 2
    box = Image.new("RGBA", (back.width + 2 * pad, back.height + 2 * pad), (0, 0, 0, 0))
    ImageDraw.Draw(box).rounded_rectangle((0, 0, box.width - 1, box.height - 1), radius=round(box.width * 0.07),
                                          fill=border + (255,))
    box.alpha_composite(back, (pad, pad))
    out.alpha_composite(box, (margin, out.height - box.height - margin))
    return out


def tile_with_back(front, back, back_w: int, border=(26, 29, 36), margin: int = 6):
    """A card picture (PIL) with a small picture of its back drawn into the bottom-left corner -
    one image, so it can't come apart when the Edit cards grid scrolls."""
    from PIL import Image
    out = front.convert("RGB").copy()
    small = back.convert("RGB").resize((back_w, max(1, round(back.height * back_w / back.width))), Image.LANCZOS)
    pad = 2
    box = Image.new("RGB", (small.width + 2 * pad, small.height + 2 * pad), border)
    box.paste(small, (pad, pad))
    out.paste(box, (margin, out.height - box.height - margin))
    return out


def edit_cards(parent, cards: list, renderer, work: Path, parked: list | None = None, per_sheet: int = 9,
               focus=None):
    """The Edit cards window: every distinct card as a tile in a grid, with its copies underneath
    (0 = not printed, but kept so it can be brought back). Click a tile to pick it; the panel on the
    side shows that card and acts on it: Replace front, Add/Replace back, Remove back, Delete (gone).
    `parked`: cards set to 0 copies earlier in this print window. Returns (new card list, new parked
    list), or None if nothing changed / cancelled."""
    import tkinter as tk
    from tkinter import ttk, filedialog
    from tkinter import font as tkfont
    messages = ThemedMessages()
    groups = group_cards(cards) + [dict(p, copies=0, parked=True) for p in (parked or [])]
    top = tk.Toplevel(parent)
    top.withdraw()
    top.title(f"{APP_NAME} — edit cards")
    top.transient(parent)
    T(top, bg="bg")
    f_title = tkfont.Font(top, family="Segoe UI Semibold", size=12)
    f_name = tkfont.Font(top, family="Segoe UI Semibold", size=9)
    f_small = tkfont.Font(top, family="Segoe UI", size=8)
    f_chip = tkfont.Font(top, family="Segoe UI Semibold", size=8)
    f_tile = tkfont.Font(top, family="Segoe UI", size=9)
    TW = 208                                   # tile image width (px)
    BW = 74                                    # the back's little preview, in the tile's bottom-left corner
    TH = round(TW * CARD_H / CARD_W)            # (cards are shown as they look once cut: no bleed)
    PW = 128                                   # panel preview width
    PH = round(PW * CARD_H / CARD_W)
    result = {}
    changed = {"any": False}
    state = {"sel": groups[0] if groups else None, "cols": 0, "page": 0, "per_page": 10 ** 6}
    if focus is not None:                      # opened from a card's right-click menu: pick that card
        state["sel"] = next((g for g in groups if card_key(g) == focus), state["sel"])
    MAX_GRID_H = 28000                         # px: windows can't be taller than ~32000
    tiles = {}                                 # id(group) -> {"frame", "pic", "name", "chips", "var"}
    photos = {}
    image_types = [("Images", " ".join("*" + t for t in IMAGE_TYPES)), ("All files", "*.*")]

    # header: title, totals, search
    head = T(tk.Frame(top, padx=18, pady=12), bg="surface")
    head.pack(fill="x")
    left = T(tk.Frame(head), bg="surface")
    left.pack(side="left")
    T(tk.Label(left, text="Edit cards", font=f_title), bg="surface", fg="text").pack(anchor="w")
    totals = T(tk.Label(left, font=f_small), bg="surface", fg="muted")
    totals.pack(anchor="w")
    search_var = tk.StringVar()
    search_box = T(tk.Frame(head), bg="surface")
    search_box.pack(side="right")
    T(tk.Label(search_box, text="Find"), bg="surface", fg="muted").pack(side="left", padx=(0, 6))
    search = ttk.Entry(search_box, textvariable=search_var, width=24, font=f_tile)
    search.pack(side="left")
    all_backs_btn = ttk.Button(head, text="Add a back to all cards…", command=lambda: back_for_all())
    all_backs_btn.pack(side="right", padx=(0, 16))
    Tooltip(all_backs_btn, "Give every card that has no back the same back (e.g. a card back image) - from an image, "
                           "or from a card in this project. They'll print double-sided.")
    pager = T(tk.Frame(head), bg="surface")           # only shown for very big decks
    ttk.Button(pager, text="‹", width=3, command=lambda: turn(-1)).pack(side="left")
    page_lbl = T(tk.Label(pager, font=f_small, width=12), bg="surface", fg="muted")
    page_lbl.pack(side="left")
    ttk.Button(pager, text="›", width=3, command=lambda: turn(1)).pack(side="left")
    T(tk.Frame(top, height=1), bg="border").pack(fill="x")

    # footer
    foot = T(tk.Frame(top, padx=18, pady=10), bg="surface")
    foot.pack(side="bottom", fill="x")
    T(tk.Frame(top, height=1), bg="border").pack(side="bottom", fill="x")
    T(tk.Label(foot, font=f_small, anchor="w", justify="left",
               text="Click a card, then use the buttons on the right.  Copies 0 keeps a card here without printing "
                    "it; Delete removes it.  Saved to the PDF when you print."), bg="surface", fg="faint").pack(side="left")

    body = T(tk.Frame(top), bg="bg")
    body.pack(fill="both", expand=True)

    # side panel: the picked card and what to do with it
    panel = T(tk.Frame(body, padx=16, pady=14, width=300), bg="surface")
    panel.pack(side="right", fill="y")
    panel.pack_propagate(False)
    T(tk.Frame(body, width=1), bg="border").pack(side="right", fill="y")
    T(tk.Label(panel, text="SELECTED CARD", font=f_chip), bg="surface", fg="faint").pack(anchor="w")
    previews = T(tk.Frame(panel), bg="surface")
    previews.pack(anchor="w", pady=(8, 8))
    p_front = T(tk.Label(previews, bd=0, font=f_small), bg="surface", fg="faint")
    p_front.grid(row=0, column=0, padx=(0, 8))
    blank_back = tk.PhotoImage(master=top, width=PW, height=PH)     # the size of a card, for "No back"
    p_back = T(tk.Label(previews, bd=0, font=f_small, cursor="hand2", compound="center", justify="center",
                        highlightthickness=1), bg="surface", fg="faint", hl="border")
    p_back.grid(row=0, column=1)
    p_back.bind("<Button-1>", lambda e: replace_back())
    Tooltip(p_back, "Click to choose a back: another card in this project, or an image.")
    T(tk.Label(previews, text="Front", font=f_small), bg="surface", fg="muted").grid(row=1, column=0, pady=(2, 0))
    p_back_cap = T(tk.Label(previews, text="Back", font=f_small), bg="surface", fg="muted")
    p_back_cap.grid(row=1, column=1, pady=(2, 0))
    p_name = T(tk.Label(panel, font=f_name, anchor="w", justify="left", wraplength=264), bg="surface", fg="text")
    p_name.pack(anchor="w", fill="x")
    name_row = T(tk.Frame(panel), bg="surface")
    name_row.pack(fill="x", pady=(6, 0))
    T(tk.Label(name_row, text="Name", font=f_small), bg="surface", fg="muted").pack(side="left", padx=(0, 6))
    name_var = tk.StringVar()
    name_box = ttk.Entry(name_row, textvariable=name_var, font=f_tile)
    name_box.pack(side="left", fill="x", expand=True)
    Tooltip(name_box, "The card's name - used to sort the cards A-Z. Press Enter to apply.")
    naming = {"loading": False}
    p_chips = T(tk.Frame(panel), bg="surface")
    p_chips.pack(anchor="w", pady=(4, 12))
    btn_front = ttk.Button(panel, text="Replace front…", command=lambda: replace_front())
    btn_front.pack(fill="x", pady=(0, 6))
    btn_back = ttk.Button(panel, text="Add back…", command=lambda: replace_back())
    btn_back.pack(fill="x", pady=(0, 6))
    btn_noback = ttk.Button(panel, text="Remove back", command=lambda: remove_back())
    btn_noback.pack(fill="x", pady=(0, 6))
    btn_delete = ttk.Button(panel, text="Delete card", style="Danger.TButton", command=lambda: delete())
    btn_delete.pack(fill="x", pady=(12, 0))
    Tooltip(btn_delete, "Remove this card from the list completely (Cancel still undoes it). To skip it but keep "
                        "it here, set its copies to 0.  (Delete key)")
    T(tk.Label(panel, text="Back: makes it a double-faced card (MDFC)"
                           + (", printed with the other MDFCs." if mdfcs_first() else "."),
               font=f_small, justify="left", wraplength=264), bg="surface", fg="faint").pack(anchor="w", pady=(14, 0))

    # the grid of tiles
    outer = T(tk.Frame(body), bg="bg")
    outer.pack(side="left", fill="both", expand=True)
    canvas = T(tk.Canvas(outer, highlightthickness=0, bd=0, width=4 * (TW + 30), height=600), bg="bg")
    bar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=bar.set)
    bar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    grid = T(tk.Frame(canvas, padx=12, pady=12), bg="bg")
    canvas.create_window((0, 0), window=grid, anchor="nw")
    grid.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    for seq, step in (("<MouseWheel>", None), ("<Button-4>", -1), ("<Button-5>", 1)):
        top.bind(seq, lambda e, step=step: (canvas.yview_scroll(step or int(-e.delta / 120), "units"), "break")[1])

    def tile_photo(g):
        """The tile's picture: the front, with the back drawn small into its bottom-left corner."""
        if not g["back"]:
            return photo(g["front"], TW)
        key = (_src_id(g["front"]), _src_id(g["back"]), TW, theme_name())
        if key not in photos:
            try:
                from PIL import ImageTk
                border = tuple(int(palette()["surface"][i:i + 2], 16) for i in (1, 3, 5))
                photos[key] = ImageTk.PhotoImage(face_with_back(card_face(renderer, g["front"], TW),
                                                                card_face(renderer, g["back"], BW), border),
                                                 master=top)
            except Exception:
                photos[key] = photo(g["front"], TW)
        return photos[key]

    def photo(src, width):
        key = (_src_id(src), width)
        if key not in photos:
            try:
                from PIL import ImageTk
                photos[key] = ImageTk.PhotoImage(card_face(renderer, src, width), master=top)
            except Exception:
                photos[key] = None
        return photos[key]

    def card_name(g):
        return Path(g["name"]).stem if g["name"] else f"Card {groups.index(g) + 1}"

    def short_name(text):
        """At most two lines under a tile, cut with … (the full name is in the side panel)."""
        lines, line = [], ""
        for word in text.split():
            trial = f"{line} {word}".strip()
            if f_tile.measure(trial) <= TW or not line:
                line = trial
            else:
                lines.append(line)
                line = word
        lines.append(line)
        if len(lines) <= 2 and all(f_tile.measure(l) <= TW for l in lines):
            return "\n".join(lines)
        first = lines[0]
        while f_tile.measure(first) > TW and len(first) > 1:
            first = first[:-1]
        second = " ".join(lines[1:]) if first == lines[0] else text[len(first):].strip()
        while second and f_tile.measure(second + "…") > TW:
            second = second[:-1]
        return first + "\n" + second.rstrip(" ,(-") + "…"

    def chips_for(g, frame, font, wrap_after=2):
        for c in frame.winfo_children():
            c.destroy()
        items = []
        small = frame is not p_chips             # tiles get the short words
        if g["copies"] <= 0:
            items.append(("Skipped" if small else "Not printing (0 copies)", None))
        if g["back"]:
            items.append(("MDFC", "accent"))
        if g["missing"]:
            items.append(("Missing" if small else "Image couldn't be downloaded", "danger"))
        for k, (text, tone) in enumerate(items):
            bg, fg = (tone + "_soft", tone) if tone else ("surface2", "muted")
            T(tk.Label(frame, text=text, font=font, padx=5, pady=0), bg=bg, fg=fg).grid(
                row=k // wrap_after, column=k % wrap_after, padx=(0, 4), pady=(0, 2), sticky="w")

    def refresh_totals():
        printing = sum(g["copies"] for g in groups)
        zero = sum(1 for g in groups if g["copies"] <= 0)
        sheets = -(-printing // max(1, per_sheet))
        totals.configure(text=f"{len(groups)} different card{'s' * (len(groups) != 1)} · printing {printing} "
                              f"on {sheets} sheet{'s' * (sheets != 1)}" + (f" · {zero} set to 0 copies" if zero else ""))

    def style_tile(g):
        t = tiles.get(id(g))
        if not t:
            return
        picked = g is state["sel"]
        off = g["copies"] <= 0
        T(t["frame"], hl="accent" if picked else "border", bg="accent_soft" if picked else "surface")
        for w in (t["pic"], t["name"], t["chips"], t["copies_row"], t["copies_lbl"]):
            T(w, bg="accent_soft" if picked else "surface")
        T(t["name"], fg="faint" if off else "text")
        T(t["copies_lbl"], fg="faint" if off else "muted")
        chips_for(g, t["chips"], f_chip)

    def show_panel():
        g = state["sel"]
        has = g is not None
        for b in (btn_front, btn_back, btn_noback, btn_delete):
            b.state(["!disabled"] if has else ["disabled"])
        all_backs_btn.state(["!disabled"] if any(not x["back"] for x in groups) else ["disabled"])
        if not has:
            p_front.configure(image="", text="", width=0, height=0)
            p_back.configure(image="", text="", width=0, height=0)
            T(p_back, bg="surface", hl="surface")
            p_name.configure(text="No card picked")
            chips_for({"copies": 1, "back": None, "missing": False}, p_chips, f_small, 3)
            return
        img = photo(g["front"], PW)
        p_front.configure(image=img or "", text="" if img else "?", width=0 if img else 16, height=0 if img else 10)
        p_front._img = img
        if g["back"]:
            bimg = photo(g["back"], PW)
            p_back.configure(image=bimg or blank_back, text="" if bimg else "?", width=0, height=0)
            T(p_back, bg="surface", hl="surface")
            p_back._img = bimg
        else:
            p_back.configure(image=blank_back, text="No back\n\nClick to add one", width=0, height=0)
            T(p_back, bg="surface2", hl="border")
        p_name.configure(text=card_name(g) + (f"\n{g['copies']} cop{'y' if g['copies'] == 1 else 'ies'}"
                                              if g["copies"] > 0 else "\n0 copies - not printing"))
        naming["loading"] = True
        name_var.set(Path(g["name"]).stem if g["name"] else "")
        naming["loading"] = False
        chips_for(g, p_chips, f_small, 3)
        btn_back.configure(text="Replace back…" if g["back"] else "Add back…")
        btn_noback.state(["!disabled"] if g["back"] else ["disabled"])

    def rename(_event=None):
        g = state["sel"]
        if g is None or naming["loading"]:
            return
        new = name_var.get().strip()
        if not new or new == (Path(g["name"]).stem if g["name"] else ""):
            return
        g["name"], g["sort"], g["renamed"] = new, sort_key({"name": new})[0], True
        changed["any"] = True
        t = tiles.get(id(g))
        if t:
            t["name"].configure(text=short_name(new))
        show_panel()
        return "break"
    name_box.bind("<Return>", rename)
    name_box.bind("<FocusOut>", rename)

    def select(g):
        rename()                                 # (a name being typed is kept)
        old, state["sel"] = state["sel"], g
        if old is not None:
            style_tile(old)
        if g is not None:
            style_tile(g)
        show_panel()

    def visible_groups():
        words = search_var.get().strip().casefold()
        return [g for g in groups if not words or words in card_name(g).casefold()]

    def regrid(force=False):
        width = canvas.winfo_width()
        if width <= 50:                         # not on screen yet: the width it's going to get
            width = canvas.winfo_reqwidth()
        cols = max(1, (width - 24) // (TW + 30))
        if cols == state["cols"] and not force:
            return
        state["cols"] = cols
        row_h = TH + 110
        state["per_page"] = cols * max(1, MAX_GRID_H // row_h)
        everything = visible_groups()
        pages = max(1, -(-len(everything) // state["per_page"]))
        state["page"] = min(state["page"], pages - 1)
        if pages > 1:
            page_lbl.configure(text=f"page {state['page'] + 1} of {pages}")
            if not pager.winfo_ismapped():
                pager.pack(side="right", padx=(0, 16))
        elif pager.winfo_ismapped():
            pager.pack_forget()
        for t in tiles.values():
            t["frame"].grid_forget()
        shown = everything[state["page"] * state["per_page"]:(state["page"] + 1) * state["per_page"]]
        for c in range(max(cols, grid.grid_size()[0])):
            grid.columnconfigure(c, uniform="tile" if c < cols else "", minsize=0)
        for r in range(max((len(shown) + cols - 1) // cols, grid.grid_size()[1])):
            grid.rowconfigure(r, uniform="tile")
        for k, g in enumerate(shown):
            tiles[id(g)]["frame"].grid(row=k // cols, column=k % cols, padx=5, pady=5, sticky="nsew")
        grid.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
    canvas.bind("<Configure>", lambda e: regrid())

    def make_tile(g):
        frame = T(tk.Frame(grid, padx=6, pady=6, highlightthickness=2, cursor="hand2"), bg="surface", hl="border")
        img = tile_photo(g)
        pic = T(tk.Label(frame, image=img or "", text="" if img else "?", bd=0, width=0 if img else 26,
                         height=0 if img else 18, cursor="hand2"), bg="surface", fg="faint")
        pic._img = img
        pic.pack()
        name = T(tk.Label(frame, text=short_name(card_name(g)), font=f_tile, justify="center", height=2,
                          cursor="hand2"), bg="surface", fg="text")
        name.pack(fill="x", pady=(4, 0))
        chips = T(tk.Frame(frame, height=16), bg="surface")
        chips.pack()
        copies_row = T(tk.Frame(frame), bg="surface")
        copies_row.pack(pady=(3, 0))
        copies_lbl = T(tk.Label(copies_row, text="Copies", font=f_small), bg="surface", fg="muted")
        copies_lbl.pack(side="left", padx=(0, 4))
        var = tk.StringVar(value=str(g["copies"]))

        def set_copies(*_, g=g, v=var):
            try:
                n = max(0, min(99, int(v.get())))
            except ValueError:
                return
            if n != g["copies"]:
                g["copies"], changed["any"] = n, True
                style_tile(g)
                refresh_totals()
                if g is state["sel"]:
                    show_panel()
        var.trace_add("write", set_copies)
        spin = ttk.Spinbox(copies_row, from_=0, to=99, textvariable=var, width=3, justify="center", font=f_tile)
        spin.pack(side="left")
        spin.bind("<FocusIn>", lambda e, g=g: select(g) if state["sel"] is not g else None, add="+")
        for w in (frame, pic, name, chips, copies_row, copies_lbl):
            w.bind("<Button-1>", lambda e, g=g: select(g))
        tiles[id(g)] = {"frame": frame, "pic": pic, "name": name, "chips": chips, "var": var,
                        "copies_row": copies_row, "copies_lbl": copies_lbl}
        style_tile(g)

    def rebuild_tile(g):
        """After replacing an image: remake that one tile, in place."""
        t = tiles.pop(id(g), None)
        if t:
            t["frame"].destroy()
        make_tile(g)
        regrid(force=True)

    def draw():
        for t in tiles.values():
            t["frame"].destroy()
        tiles.clear()
        for g in groups:
            make_tile(g)
        regrid(force=True)
        refresh_totals()
        show_panel()

    def pick_image(title):
        chosen = filedialog.askopenfilename(parent=top, title=title, filetypes=image_types)
        if not chosen:
            return None
        work.mkdir(parents=True, exist_ok=True)
        out = work / f"edit-{time.time_ns()}.jpg"
        try:
            prepare_card_image(Path(chosen), out)
        except Exception as e:
            messages.showerror(APP_NAME, f"Couldn't use {Path(chosen).name}: {e}", parent=top)
            return None
        return ("file", out)

    def replace_front():
        g = state["sel"]
        if g is None:
            return
        src = pick_image(f"Pick the new front for {card_name(g)}")
        if src:
            g["front"], g["missing"], changed["any"] = src, False, True
            rebuild_tile(g)
            show_panel()

    def prepare_image(path):
        work.mkdir(parents=True, exist_ok=True)
        out = work / f"edit-{time.time_ns()}.jpg"
        try:
            prepare_card_image(Path(path), out)
        except Exception as e:
            messages.showerror(APP_NAME, f"Couldn't use {Path(path).name}: {e}", parent=top)
            return None
        return ("file", out)

    def choose_back(title, note, exclude):
        """The back chooser: one of this project's cards (taken out of the list and used as the back) or
        an image (dragged in, or chosen). Returns ("card", group), ("image", src) or None."""
        win = tk.Toplevel(top)
        win.withdraw()
        win.title(f"{APP_NAME} — choose a back")
        win.transient(top)
        T(win, bg="surface")
        picked = {"g": None, "out": None}
        CW = 112
        inner = T(tk.Frame(win, padx=18, pady=14), bg="surface")
        inner.pack(fill="both", expand=True)
        T(tk.Label(inner, text=title, font=f_title, anchor="w"), bg="surface", fg="text").pack(anchor="w")
        if note:
            T(tk.Label(inner, text=note, font=f_small, anchor="w", justify="left", wraplength=640), bg="surface",
              fg="muted").pack(anchor="w", pady=(2, 0))
        # an image: drag it in or choose it
        drop = T(tk.Frame(inner, padx=14, pady=12, highlightthickness=1), bg="surface2", hl="border")
        drop.pack(fill="x", pady=(12, 0))
        T(tk.Label(drop, text="Drag an image here", font=f_name), bg="surface2", fg="text").pack(side="left")
        T(tk.Label(drop, text="  (PNG, JPG, WEBP…)", font=f_small), bg="surface2", fg="muted").pack(side="left")

        def image_chosen(path):
            src = prepare_image(path)
            if src:
                picked["out"] = ("image", src)
                win.destroy()

        def choose_file():
            chosen = filedialog.askopenfilename(parent=win, title=title, filetypes=image_types)
            if chosen:
                image_chosen(chosen)
        ttk.Button(drop, text="Choose image…", command=choose_file).pack(side="right")
        if not enable_file_drop(win, lambda files: image_chosen(files[0]) if files else None):
            pass
        # or a card from this project
        candidates = [g for g in groups if not g["back"] and g not in exclude]
        cards_area = None
        if candidates:
            find_row = T(tk.Frame(inner), bg="surface")
            find_row.pack(fill="x", pady=(14, 6))
            T(tk.Label(find_row, text="Or use one of this project's cards (it's taken out of the list and becomes the back):",
                       font=f_small, anchor="w"), bg="surface", fg="muted").pack(side="left")
            find_var = tk.StringVar()
            find_box = ttk.Entry(find_row, textvariable=find_var, width=20, font=f_tile)
            find_box.pack(side="right")
            T(tk.Label(find_row, text="Find", font=f_small), bg="surface", fg="muted").pack(side="right", padx=(12, 6))
            cards_area = T(tk.Frame(inner, highlightthickness=1), bg="bg", hl="border")
            cards_area.pack(fill="both", expand=True)
            cv = T(tk.Canvas(cards_area, highlightthickness=0, bd=0, width=5 * (CW + 18) + 10,
                             height=min(2, -(-len(candidates) // 5)) * (round(CW * CARD_H / CARD_W) + 42) + 10),
                   bg="bg")
            sb = ttk.Scrollbar(cards_area, orient="vertical", command=cv.yview)
            cv.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            cv.pack(side="left", fill="both", expand=True)
            holder = T(tk.Frame(cv, padx=6, pady=6), bg="bg")
            cv.create_window((0, 0), window=holder, anchor="nw")
            holder.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
            for seq, step in (("<MouseWheel>", None), ("<Button-4>", -1), ("<Button-5>", 1)):
                win.bind(seq, lambda e, step=step: (cv.yview_scroll(step or int(-e.delta / 120), "units"), "break")[1])
            cells = {}

            def layout_cells(*_):
                words = find_var.get().strip().casefold()
                shown = [g for g in candidates if not words or words in card_name(g).casefold()]
                for c in cells.values():
                    c[0].grid_forget()
                for k, g in enumerate(shown):
                    cells[id(g)][0].grid(row=k // 5, column=k % 5, padx=4, pady=4, sticky="n")
                cv.yview_moveto(0)
            find_var.trace_add("write", layout_cells)

            def pick(g):
                old, picked["g"] = picked["g"], g
                for cg in (old, g):
                    if cg is not None:
                        on = cg is picked["g"]
                        for w in cells[id(cg)]:
                            T(w, bg="accent_soft" if on else "surface")
                        T(cells[id(cg)][0], hl="accent" if on else "surface")
                use_btn.state(["!disabled"])
            for k, g in enumerate(candidates):
                cell = T(tk.Frame(holder, padx=4, pady=4, highlightthickness=2, cursor="hand2"), bg="surface", hl="surface")
                img = photo(g["front"], CW)
                pic = T(tk.Label(cell, image=img or "", text="" if img else "?", bd=0, cursor="hand2"), bg="surface",
                        fg="faint")
                pic._img = img
                pic.pack()
                lbl = T(tk.Label(cell, text=card_name(g), font=f_small, cursor="hand2", width=1, anchor="w"),
                        bg="surface", fg="text")
                lbl.pack(pady=(2, 0), fill="x")
                cells[id(g)] = (cell, pic, lbl)
                for w in (cell, pic, lbl):
                    w.bind("<Button-1>", lambda e, g=g: pick(g))
                    w.bind("<Double-Button-1>", lambda e, g=g: (pick(g), use()))
            layout_cells()
            holder.update_idletasks()
            cv.configure(width=holder.winfo_reqwidth())
        foot2 = T(tk.Frame(win, padx=18, pady=10), bg="surface2")
        foot2.pack(fill="x")

        def use():
            if picked["g"] is not None:
                picked["out"] = ("card", picked["g"])
                win.destroy()
        use_btn = ttk.Button(foot2, text="Use this card as the back", style="Accent.TButton", command=use)
        use_btn.pack(side="right")
        use_btn.state(["disabled"])
        if not candidates:
            use_btn.pack_forget()
        ttk.Button(foot2, text="Cancel", command=win.destroy).pack(side="right", padx=(0, 8))
        win.bind("<Escape>", lambda e: win.destroy())
        apply_theme(win)
        win.update_idletasks()
        x = top.winfo_rootx() + (top.winfo_width() - win.winfo_reqwidth()) // 2
        y = top.winfo_rooty() + max(0, (top.winfo_height() - win.winfo_reqheight()) // 3)
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
        win.deiconify()
        win.grab_set()
        win.focus_force()
        top.wait_window(win)
        top.grab_set()                         # (back to Edit cards)
        return picked["out"]

    def take_out(g):
        """A card that became a back: it's no longer a card of its own."""
        groups.remove(g)
        t = tiles.pop(id(g), None)
        if t:
            t["frame"].destroy()

    def replace_back():
        g = state["sel"]
        if g is None:
            return
        got = choose_back(f"{'Replace the' if g['back'] else 'Choose a'} back for {card_name(g)}",
                          "Use this for double-faced cards (MDFCs) Fleem didn't recognise: pick the card that's "
                          "really this one's back.", exclude=[g])
        if not got:
            return
        kind, value = got
        if kind == "card":
            g["back"] = value["front"]
            take_out(value)
        else:
            g["back"] = value
        changed["any"] = True
        rebuild_tile(g)
        refresh_totals()
        show_panel()

    def back_for_all():
        rename()
        without = [g for g in groups if not g["back"]]
        if not without:
            messages.showinfo(APP_NAME, "Every card already has a back.", parent=top)
            return
        got = choose_back(f"Choose a back for all {len(without)} card{'s' * (len(without) != 1)} without one",
                          "Cards that already have a back (MDFCs) keep theirs. The others all get this back and "
                          "print double-sided.", exclude=[])
        if not got:
            return
        kind, value = got
        if kind == "card":
            src = value["front"]
            take_out(value)
        else:
            src = value
        for g in groups:
            if not g["back"]:
                g["back"] = src
        changed["any"] = True
        if state["sel"] not in groups:
            state["sel"] = groups[0] if groups else None
        draw()

    def remove_back():
        g = state["sel"]
        if g is None or not g["back"]:
            return
        g["back"], changed["any"] = None, True
        rebuild_tile(g)
        show_panel()

    def delete(_event=None):
        g = state["sel"]
        if g is None:
            return
        shown = visible_groups()
        k = shown.index(g) if g in shown else 0
        groups.remove(g)
        t = tiles.pop(id(g), None)
        if t:
            t["frame"].destroy()
        changed["any"] = True
        shown = visible_groups()
        state["sel"] = shown[min(k, len(shown) - 1)] if shown else None
        if state["sel"] is not None:
            style_tile(state["sel"])
        regrid(force=True)
        refresh_totals()
        show_panel()

    def move(step):
        shown = visible_groups()
        if not shown:
            return "break"
        k = shown.index(state["sel"]) if state["sel"] in shown else -step
        k = max(0, min(len(shown) - 1, k + step))
        if k // state["per_page"] != state["page"]:
            state["page"] = k // state["per_page"]
            regrid(force=True)
        select(shown[k])
        return "break"

    def turn(step):
        state["page"] = max(0, state["page"] + step)
        regrid(force=True)
        canvas.yview_moveto(0)

    def key_ok():
        w = top.focus_get()
        return not isinstance(w, (ttk.Entry, ttk.Spinbox))
    top.bind("<Delete>", lambda e: delete() if key_ok() else None)
    top.bind("<Left>", lambda e: move(-1) if key_ok() else None)
    top.bind("<Right>", lambda e: move(1) if key_ok() else None)
    top.bind("<Up>", lambda e: move(-state["cols"]) if key_ok() else None)
    top.bind("<Down>", lambda e: move(state["cols"]) if key_ok() else None)
    search_var.trace_add("write", lambda *_: (state.update(page=0), regrid(force=True), canvas.yview_moveto(0)))

    def done(_event=None):
        rename()                                 # a name still being typed counts
        if changed["any"]:
            if not any(g["copies"] > 0 for g in groups):
                messages.showerror(APP_NAME, "Every card is set to 0 copies or deleted - keep at least one to print.",
                                   parent=top)
                return
            new_parked = [{k: v for k, v in g.items() if k not in ("copies", "parked", "renamed")}
                          for g in groups if g["copies"] <= 0]
            result["out"] = (cards_from_groups(groups), new_parked)
        top.destroy()
    ttk.Button(foot, text="Done", command=done, style="Accent.TButton").pack(side="right")
    ttk.Button(foot, text="Cancel", command=top.destroy).pack(side="right", padx=(0, 8))
    top.bind("<Escape>", lambda e: top.destroy())
    top.bind("<Control-f>", lambda e: (search.focus_set(), "break")[1])
    draw()
    apply_theme(top)
    try:
        sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
        w, h = min(1560, sw - 60), min(1000, sh - 80)
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - h) // 3)
        top.geometry(f"{w}x{h}+{max(0, min(x, sw - w))}+{max(0, min(y, sh - h))}")
    except Exception:
        pass
    top.deiconify()
    top.grab_set()
    top.focus_force()
    top.after(50, lambda: regrid(force=True))
    parent.wait_window(top)
    return result.get("out")


def ask_card(parent, front: Path):
    """The "Add card" pop-up for an image dragged into the print window: name, copies, and a place to
    drag (or choose) the back image to make it a double-faced card (MDFC).
    Returns {"front", "back", "name", "copies"} or None if skipped."""
    import tkinter as tk
    from tkinter import ttk, filedialog
    from tkinter import font as tkfont
    top = tk.Toplevel(parent)
    top.withdraw()
    top.title(f"{APP_NAME} — add a card")
    top.transient(parent)
    top.resizable(False, False)
    T(top, bg="surface")
    result, state = {}, {"back": None}
    f_title = tkfont.Font(top, family="Segoe UI Semibold", size=12)
    f_label = tkfont.Font(top, family="Segoe UI Semibold", size=9)
    f_small = tkfont.Font(top, family="Segoe UI", size=8)
    f_body = tkfont.Font(top, family="Segoe UI", size=9)
    BOX_W, BOX_H = 250, 340                   # front and back boxes: the same size (a card's shape)

    def thumb(path: Path):
        try:
            from PIL import Image, ImageTk
            with Image.open(path) as im:
                im = im.convert("RGB")
                if im.width > im.height:
                    im = im.rotate(90, expand=True)
                im.thumbnail((BOX_W, BOX_H), Image.LANCZOS)
                return ImageTk.PhotoImage(im, master=top)
        except Exception:
            return None

    def card_box(parent, column, outline):
        frame = T(tk.Frame(parent, width=BOX_W + 4, height=BOX_H + 4, highlightthickness=2), bg="surface2", hl=outline)
        frame.pack_propagate(False)
        frame.grid(row=1, column=column, padx=(0, 16) if column == 0 else 0)
        label = T(tk.Label(frame, bd=0, justify="center", wraplength=BOX_W - 24, font=f_body), bg="surface2", fg="muted")
        label.pack(fill="both", expand=True)
        return frame, label
    head = T(tk.Frame(top, padx=20, pady=14), bg="surface")
    head.pack(fill="x")
    T(tk.Label(head, text="Add a card", font=f_title), bg="surface", fg="text").pack(anchor="w")
    T(tk.Label(head, text=front.name, font=f_small), bg="surface", fg="muted").pack(anchor="w")
    body = T(tk.Frame(top, padx=20), bg="surface")
    body.pack()
    front_img = thumb(front)
    T(tk.Label(body, text="Front", font=f_label), bg="surface", fg="text").grid(row=0, column=0, sticky="w", pady=(0, 4))
    _, front_box = card_box(body, 0, "surface")
    if front_img:
        front_box.configure(image=front_img)
    else:   # no preview available: just the file name, in the card-sized box
        front_box.configure(text=front.name)
    T(tk.Label(body, text="Back  (double-faced cards only)", font=f_label), bg="surface", fg="text").grid(
        row=0, column=1, sticky="w", pady=(0, 4))
    empty_text = "Drag the back image here\n\nor use Choose back…\n\n(leave empty for a normal card)"
    back_frame, back_box = card_box(body, 1, "border_strong")
    back_box.configure(text=empty_text)
    form = T(tk.Frame(body), bg="surface")
    form.grid(row=2, column=0, columnspan=2, sticky="we", pady=(14, 0))
    stem = re.sub(r"\s*\(\d+\)$", "", front.stem).strip()
    name_var, copies_var = tk.StringVar(value=stem), tk.StringVar(value="1")
    T(tk.Label(form, text="Name", font=f_label), bg="surface", fg="text").grid(row=0, column=0, sticky="w")
    name_entry = ttk.Entry(form, textvariable=name_var, width=40, font=f_body)
    name_entry.grid(row=0, column=1, sticky="w", padx=(10, 0))
    T(tk.Label(form, text="Copies", font=f_label), bg="surface", fg="text").grid(row=1, column=0, sticky="w", pady=(8, 0))
    ttk.Spinbox(form, from_=1, to=99, textvariable=copies_var, width=5, justify="center", font=f_body).grid(
        row=1, column=1, sticky="w", padx=(10, 0), pady=(8, 0))
    note = T(tk.Label(top, padx=20, justify="left", font=f_small), bg="surface", fg="muted")
    note.pack(anchor="w", pady=(8, 0))

    def set_back(path):
        path = Path(path)
        if path.suffix.lower() not in IMAGE_TYPES:
            T(note, fg="danger").configure(text=f"{path.name} isn't an image.")
            return
        state["back"] = path
        img = thumb(path)
        back_box.configure(image=img or "", text="" if img else path.name)
        back_box._img = img
        T(back_frame, hl="accent")
        add_btn.configure(text="Add double-faced card")
        T(note, fg="muted").configure(text=f"Back: {path.name}")

    def clear_back():
        state["back"] = None
        back_box.configure(image="", text=empty_text)
        T(back_frame, hl="border_strong")
        add_btn.configure(text="Add card")
        note.configure(text="")

    def choose_back():
        chosen = filedialog.askopenfilename(parent=top, title="Pick the back image",
                                            filetypes=[("Images", " ".join("*" + t for t in IMAGE_TYPES)), ("All files", "*.*")])
        if chosen:
            set_back(chosen)
    T(tk.Frame(top, height=1), bg="border").pack(fill="x", pady=(12, 0))
    buttons = T(tk.Frame(top, padx=20, pady=12), bg="surface2")
    buttons.pack(fill="x")
    ttk.Button(buttons, text="Choose back…", command=choose_back).pack(side="left")
    ttk.Button(buttons, text="No back", command=clear_back).pack(side="left", padx=(8, 0))
    add_btn = ttk.Button(buttons, text="Add card", style="Accent.TButton")
    add_btn.pack(side="right")
    ttk.Button(buttons, text="Skip", command=top.destroy).pack(side="right", padx=(0, 8))

    def add(_event=None):
        try:
            copies = max(1, min(99, int(copies_var.get())))
        except ValueError:
            T(note, fg="danger").configure(text="Copies must be a number.")
            return
        result.update(front=front, back=state["back"], name=name_var.get().strip() or stem, copies=copies)
        top.destroy()
    add_btn.configure(command=add)
    top.bind("<Return>", add)
    top.bind("<Escape>", lambda e: top.destroy())
    top._front = front_img
    if not enable_file_drop(top, lambda files: set_back(files[0]) if files else None):
        back_box.configure(text="Use Choose back… to add a back\n\n(leave empty for a normal card)")

    # read the card's name off the image in the background; it replaces the file name unless you typed one
    reading = {"name": None, "done": False, "typed": False, "setting": False}
    name_var.trace_add("write", lambda *_: reading.update(typed=True) if not reading["setting"] else None)

    def read_name():
        tmp = WORK_DIR / f"name-{time.time_ns()}.jpg"
        try:
            WORK_DIR.mkdir(parents=True, exist_ok=True)
            prepare_card_image(front, tmp)                  # (the reader expects an MPC-shaped image)
            reading["name"] = read_card_name(tmp.read_bytes())
        except Exception as e:
            log.info("Couldn't read the name of %s: %s", front.name, e)
        finally:
            tmp.unlink(missing_ok=True)
            reading["done"] = True

    def show_name():
        if not top.winfo_exists():
            return
        if not reading["done"]:
            top.after(150, show_name)
            return
        found = reading["name"]
        if not found:
            T(note, fg="muted").configure(text="" if state["back"] else "Couldn't read the name off the card - "
                                                                       "check it below.")
            return
        if not reading["typed"] or name_var.get().strip() in ("", stem):
            reading["setting"] = True
            name_var.set(found)
            reading["setting"] = False
            name_entry.select_range(0, "end")
            T(note, fg="ok").configure(text="Name read from the card.")
        elif name_var.get().strip() != found:
            T(note, fg="muted").configure(text=f"Read from the card: {found}")
    T(note, fg="muted").configure(text="Reading the card's name…")
    threading.Thread(target=read_name, daemon=True).start()
    top.after(150, show_name)
    apply_theme(top)
    center_on(top, parent)
    top.deiconify()
    top.grab_set()
    name_entry.focus_set()
    name_entry.select_range(0, "end")
    parent.wait_window(top)
    return result or None


def choose_pages(settings: Settings, title: str, out_path: Path, profile: str | None = None,
                 deck: list | None = None, pdf_path: Path | None = None, pdf_layout=None, from_xml: bool = False,
                 foreign: dict | None = None, can_take_apart: bool = False, restore: dict | None = None):
    """The print window. Shows previews of the fronts (and backs) to pick from, the printer, and - for
    Fleem's own cards - bleed edge / card spacing boxes and Add cards. `deck` is the card list in memory
    (from a Fleem PDF or an order XML): every change only changes the deck and its previews; the PDF is
    written to `out_path` when you press Print (only if something changed, for an existing PDF).
    Without a deck (a PDF made elsewhere) the PDF's own pages are shown and printed as they are.
    `foreign` (the cards taken out of a PDF made elsewhere, see read_foreign): Print always saves them as
    a new Fleem PDF at out_path. can_take_apart: a PDF shown as it is whose cards could be taken out.
    Returns (printer, description, page indexes starting at 0, PDF to print, back pages),
    ("reopen", how) to open the PDF again another way, or None. `restore`: everything from the same
    window just before it reloaded itself for a new theme (cards, picks, layout, added images...)."""
    import tkinter as tk
    from tkinter import ttk, font as tkfont
    messagebox = ThemedMessages()            # pop-ups in Fleem's light/dark look
    try:
        from PIL import ImageTk
    except Exception:
        ImageTk = None
    editable = deck is not None
    restore = restore or {}
    cards = restore["cards"] if "cards" in restore else list(deck) if editable else []
    renderer = restore.get("renderer") or (PageRenderer() if editable else None)
    pdf = None
    pages_total = 0
    if not editable:
        try:
            import pypdfium2 as pdfium
            pdf = pdfium.PdfDocument(str(pdf_path))
            pages_total = len(pdf)
        except Exception as e:
            log.info("No page previews (%s)", e)
            pages_total = pdf_page_count(pdf_path)

    # bleed edge + card spacing: what the file's name calls for (an add-on's rule), else the PDF's own (so
    # nothing changes unless you change it), else last time's
    last = last_layout()
    from_name = layout_from_name(title, settings)
    if from_name is not None:
        start_layout = full_layout(from_name)
        if pdf_layout is not None:             # the name sets bleed and spacing; the PDF keeps its own page
            start_layout = start_layout[:2] + (pdf_layout[2] if len(pdf_layout) > 2 and pdf_layout[2] else
                                               start_layout[2], layout_orientation(pdf_layout), layout_header(pdf_layout),
                                               layout_numbering(pdf_layout), layout_card(pdf_layout),
                                               layout_marks(pdf_layout))
    elif pdf_layout is not None and pdf_layout[0] is not None:
        start_layout = (pdf_layout[0], pdf_layout[1] if pdf_layout[1] is not None else last[1],
                        pdf_layout[2] if len(pdf_layout) > 2 and pdf_layout[2] else last[2],
                        layout_orientation(pdf_layout), layout_header(pdf_layout), layout_numbering(pdf_layout),
                        layout_card(pdf_layout), layout_marks(pdf_layout))
    else:
        start_layout = last
    start_layout = full_layout(start_layout)
    lay = {"source": "name" if from_name is not None else "pdf" if pdf_layout is not None else "last",
           "touched": False, "job": None,
           "layout": start_layout, "added": [], "cards_changed": False}
    if "lay" in restore:                     # reloaded for a new theme: as it was
        lay = restore["lay"]
        lay["job"] = None
        start_layout = full_layout(lay["layout"])

    # Page layout: "pairs" = front, back, front, back... (Fleem); "single" = every page a front
    paging = {"mode": restore.get("mode") or ("pairs" if editable else "single")}
    counts = {"front": 0, "back": 0}

    def recount():
        if editable:
            n = len(deck_pages(cards, lay["layout"]))
            counts.update(front=n, back=n)
        elif paging["mode"] == "single":
            counts.update(front=pages_total, back=0)
        else:
            counts.update(front=(pages_total + 1) // 2, back=pages_total // 2)
    recount()

    def page_index(side, n):
        if paging["mode"] == "single":
            return n - 1
        return 2 * n - 2 if side == "front" else 2 * n - 1

    def noun(side):
        return "page" if paging["mode"] == "single" else side

    BASE_W, COLS = 132, 6                # starting preview width and columns (rescale fits them to the window)
    grid_cols = {"n": COLS}
    zoom = {"f": 1.0}                    # − / + on the toolbar
    aspect = page_size(start_layout)[1] / page_size(start_layout)[0]
    if pdf is not None and pages_total:
        w, h = pdf[0].get_size()
        aspect = h / w
    size = {"w": BASE_W, "h": int(BASE_W * aspect)}
    result = {}
    chosen = {"front": set(), "back": set()}       # nothing picked to start with
    for side, picked in restore.get("chosen", {}).items():
        chosen[side] = set(picked)
    blank_backs: set[int] = set()
    view = {"side": restore.get("side", "front"), "anchor": None, "gen": 0}
    zoom_from = restore.get("zoom")
    photos, cells = {}, {}
    resize = {"job": None, "width": None}
    pending = set()                      # preview-drawing steps waiting to run

    root = tk.Tk()
    root.title(f"{APP_NAME} — print {title}")
    made_theme = theme_name()                # (a new theme reloads the window)

    def later(ms, fn, *args):
        def run():
            pending.discard(job)
            fn(*args)
        job = root.after(ms, run)
        pending.add(job)

    def close():
        """Close the window, first stopping preview drawing and pending work."""
        view["gen"] += 1
        for job in (*pending, lay.get("job"), resize.get("job")):
            if job:
                try:
                    root.after_cancel(job)
                except Exception:
                    pass
        root.destroy()

    def cancel(_event=None):
        if lay["added"] and not messagebox.askyesno(
                APP_NAME, "Close without printing? The cards you added or edited won't be kept.", parent=root):
            return
        close()
    root.protocol("WM_DELETE_WINDOW", cancel)
    try:
        icon_png = Path(getattr(sys, "_MEIPASS", HERE)) / "fleem_icon.png"
        if icon_png.exists():
            root.iconphoto(True, tk.PhotoImage(master=root, file=str(icon_png)))
    except Exception:
        pass
    # every text in the window uses one of these; they grow a little with a big window (see rescale)
    fonts = {"body": tkfont.Font(root, family="Segoe UI", size=9),
             "bold": tkfont.Font(root, family="Segoe UI", size=9, weight="bold"),
             "small": tkfont.Font(root, family="Segoe UI", size=8),
             "title": tkfont.Font(root, family="Segoe UI Semibold", size=12),
             "label": tkfont.Font(root, family="Segoe UI Semibold", size=9),
             "chip": tkfont.Font(root, family="Segoe UI Semibold", size=8),
             "caption": tkfont.Font(root, family="Segoe UI Semibold", size=9),
             "badge": tkfont.Font(root, family="Segoe UI", size=10, weight="bold"),
             "big": tkfont.Font(root, family="Segoe UI", size=28)}
    BASE_FONT_SIZES = {k: f.cget("size") for k, f in fonts.items()}
    root.option_add("*Font", fonts["body"])
    apply_theme(root, font=fonts["body"], bold=fonts["bold"])
    style = ttk.Style(root)
    T(root, bg="bg")

    # --- header (one line): name, chips (where the cards come from, size), main buttons
    header = T(tk.Frame(root, padx=16, pady=8), bg="surface")
    header.pack(fill="x")
    title_box = T(tk.Frame(header), bg="surface")
    title_box.pack(side="left", fill="x", expand=True)
    T(tk.Label(title_box, text=title, font=fonts["title"], anchor="w"), bg="surface", fg="text").pack(side="left")
    chips = T(tk.Frame(title_box), bg="surface")
    chips.pack(side="left", padx=(12, 0))

    def chip(text, tone=None):
        bg, fg = (tone + "_soft", tone) if tone else ("surface2", "muted")
        w = T(tk.Label(chips, text=text, font=fonts["chip"], padx=8, pady=1), bg=bg, fg=fg)
        w.pack(side="left", padx=(0, 6))
        return w
    if profile:
        chip(profile, "accent")
    if foreign is not None:
        chip("Made elsewhere - cards taken out", "warn")
    elif from_xml:
        chip("From an order XML", "ok")
    elif editable:
        chip("Fleem PDF")
    else:
        chip("Printed as it is", "warn")
    size_chip = chip("")
    side_label = tk.Label(root)                 # (kept for older code paths; not shown)
    side_btn = None
    theme_btn = ttk.Button(header, style="Ghost.TButton", width=11)
    theme_btn.pack(side="right", padx=(8, 0))
    add_btn = ttk.Button(header, text="Add cards…", command=lambda: pick_cards())
    add_btn.pack(side="right", padx=(8, 0))
    edit_btn = ttk.Button(header, text="Edit cards…", command=lambda: open_edit_cards())
    edit_btn.pack(side="right")
    find_btn = None
    if editable and mpcfill_module() is not None:
        find_btn = ttk.Button(header, text="Find cards…", command=lambda: open_find_cards())
        find_btn.pack(side="right", padx=(0, 8))
        Tooltip(find_btn, "Search MPCFill by card name (or paste a decklist), pick the art for the front and back, "
                          "and add the cards to this print.")
    undo_box = T(tk.Frame(header), bg="surface")         # Undo / Redo: only shown when there's something to do
    undo_box.pack(side="right", padx=(0, 8))
    undo_btn = ttk.Button(undo_box, text="↶ Undo", style="Ghost.TButton", command=lambda: undo())
    redo_btn = ttk.Button(undo_box, text="↷ Redo", style="Ghost.TButton", command=lambda: redo())
    Tooltip(undo_btn, "Undo the last change to the cards or the layout.  (Ctrl+Z)")
    Tooltip(redo_btn, "Redo what you undid.  (Ctrl+Y)")
    Tooltip(add_btn, "Add card images, order XMLs or PDFs to this print (you can also drag them onto the window).")
    Tooltip(edit_btn, "Change copies, replace images, add or remove backs, delete cards.  (Ctrl+E)")
    T(tk.Frame(root, height=1), bg="border").pack(fill="x")

    # --- settings (one row): printer, layout, Advanced; one status line under it
    top = T(tk.Frame(root, padx=16, pady=7), bg="surface")
    top.pack(fill="x")
    form = T(tk.Frame(top), bg="surface")
    form.pack(fill="x")
    T(tk.Label(form, text="Printer", font=fonts["label"]), bg="surface", fg="text").pack(side="left", padx=(0, 8))
    available = installed_printers()
    printer_row = form
    printer_var = tk.StringVar(value=printer_for(available, profile) or "")
    root.option_add("*TCombobox*Listbox.font", fonts["body"])
    printer_box = ttk.Combobox(form, textvariable=printer_var, values=sorted(available), width=30,
                               state="readonly" if available else "normal", font=fonts["body"])
    printer_box.pack(side="left")
    printer_box.bind("<<ComboboxSelected>>", lambda e: update())

    def open_printer_settings():
        printer = printer_var.get().strip()
        if printer and choose_printer_settings(printer, root, profile):
            update()
    settings_btn = ttk.Button(form, text="Printer settings…", command=open_printer_settings)
    settings_btn.pack(side="left", padx=(6, 0))
    Tooltip(settings_btn, "Open the printer's own settings (paper size and type, quality, colour, presets…). "
                          + (f"Fleem remembers them for {profile_label(profile)} on this printer."
                             if profile_label(profile) else "Fleem remembers them for this printer."))
    T(tk.Frame(form, width=1, height=22), bg="border").pack(side="left", padx=16)
    T(tk.Label(form, text="Bleed", font=fonts["label"]), bg="surface", fg="text").pack(side="left", padx=(0, 6))
    layout_row = form
    bleed_var = tk.StringVar(value=f"{start_layout[0]:.2f}")
    spacing_var = tk.StringVar(value=f"{start_layout[1]:.2f}")
    bleed_entry = ttk.Entry(form, textvariable=bleed_var, width=6, font=fonts["body"], justify="center")
    bleed_entry.pack(side="left")
    T(tk.Label(form, text="mm"), bg="surface", fg="muted").pack(side="left", padx=(4, 14))
    T(tk.Label(form, text="Spacing", font=fonts["label"]), bg="surface", fg="text").pack(side="left", padx=(0, 6))
    spacing_entry = ttk.Entry(form, textvariable=spacing_var, width=6, font=fonts["body"], justify="center")
    spacing_entry.pack(side="left")
    T(tk.Label(form, text="mm"), bg="surface", fg="muted").pack(side="left", padx=(4, 14))
    T(tk.Label(form, text="Paper", font=fonts["label"]), bg="surface", fg="text").pack(side="left", padx=(0, 6))
    paper_var = tk.StringVar(value=layout_paper(start_layout))
    paper_box = ttk.Combobox(form, textvariable=paper_var, values=list(PAPER_SIZES), width=8, state="readonly",
                             font=fonts["body"])
    paper_box.pack(side="left")
    paper_box.bind("<<ComboboxSelected>>", lambda e: paper_changed())
    Tooltip(paper_box, "The paper size of the PDF (the cards are laid out to fit it).")
    T(tk.Label(form, text="Card", font=fonts["label"]), bg="surface", fg="text").pack(side="left", padx=(14, 6))
    card_size = {"value": layout_card(start_layout)}
    CUSTOM_CHOICE = "Custom…"

    def card_box_label(value):
        mm = custom_card_mm(value)
        return f"{mm[0]:g} x {mm[1]:g} mm" if mm else value
    card_var = tk.StringVar(value=card_box_label(card_size["value"]))
    card_box = ttk.Combobox(form, textvariable=card_var, values=list(CARD_SIZES) + [CUSTOM_CHOICE], width=12,
                            state="readonly", font=fonts["body"])
    card_box.pack(side="left")
    card_box.bind("<<ComboboxSelected>>", lambda e: card_changed())
    Tooltip(card_box, "The card size (the same sizes as Proxy PDF Maker). Card images are stretched to fit it.\n"
                      + "\n".join(f"{n}: {CARD_SIZE_NOTES[n]}" for n in CARD_SIZES)
                      + "\nCustom…: any width and height in mm.")
    Tooltip(bleed_entry, BLEED_TIP)
    Tooltip(spacing_entry, SPACING_TIP)
    adv_toggle = ttk.Button(form, text="Advanced settings ▸", style="Ghost.TButton")
    adv_toggle.pack(side="right")
    status = T(tk.Frame(top), bg="surface")            # shown only when there's a message
    printer_label = tk.Label(status)                    # (not shown)
    layout_label = T(tk.Label(status, anchor="w", justify="left", font=fonts["small"]), bg="surface", fg="muted")
    layout_label.pack(fill="x")
    layout_label.bind("<Configure>", lambda e: layout_label.configure(wraplength=max(200, e.width - 4)))

    def set_status(text, tone="muted"):
        T(layout_label, fg=tone).configure(text=text)
        if text and not status.winfo_ismapped():
            status.pack(fill="x", pady=(5, 0), after=form)
        elif not text and status.winfo_ismapped():
            status.pack_forget()

    # --- Advanced settings (fold-out): how the PDF's pages are arranged
    adv = T(tk.Frame(top, padx=12, pady=10, highlightthickness=1), bg="surface2", hl="border")
    mode_var = tk.StringVar(value=paging["mode"])
    tk.Label(adv, text="Page layout:", bg="#f6f7f9").grid(row=0, column=0, sticky="nw", padx=(0, 8))
    radios = [ttk.Radiobutton(adv, style="Panel.TRadiobutton", text="Fronts and backs - every other page is a back (how Fleem makes PDFs)",
                              value="pairs", variable=mode_var, command=lambda: set_mode()),
              ttk.Radiobutton(adv, style="Panel.TRadiobutton", text="Every page is a front - no back pages (e.g. PDFs made somewhere else)",
                              value="single", variable=mode_var, command=lambda: set_mode())]
    if foreign is not None:
        mode_var.set("pairs" if foreign["has_backs"] else "single")
    for i, r in enumerate(radios):
        r.grid(row=i, column=1, sticky="w")
        if editable and foreign is None:
            r.state(["disabled"])
    if foreign is not None:
        note_text = (f"Fleem took the cards out of this PDF ({foreign['summary']}). If it got the back pages "
                     "wrong, pick the other option and Fleem reads the PDF again.")
    elif editable:
        note_text = "Fleem's own PDFs always have a back page after every front page."
    else:
        note_text = ("This PDF wasn't made by Fleem, so Fleem can't tell whether it has back pages - check this is "
                     "right.")
    adv_note = tk.Label(adv, bg="#f6f7f9", fg="#666666", anchor="w", justify="left", text=note_text)
    adv_note.grid(row=2, column=1, sticky="w", pady=(4, 0))

    # paper orientation and the header (Fleem's own PDFs and XMLs: things that change the PDF)
    default_header = Path(out_path).stem
    orientation_var = tk.StringVar(value=layout_orientation(start_layout))
    header_on = tk.BooleanVar(value=layout_header(start_layout) is not None)
    header_var = tk.StringVar(value=layout_header(start_layout) or default_header)
    tk.Label(adv, text="Orientation:", bg="#f6f7f9").grid(row=6, column=0, sticky="nw", padx=(0, 8), pady=(10, 0))
    orient_row = tk.Frame(adv, bg="#f6f7f9")
    orient_row.grid(row=6, column=1, sticky="w", pady=(10, 0))
    orient_radios = [ttk.Radiobutton(orient_row, style="Panel.TRadiobutton", text=label, value=value,
                                     variable=orientation_var, command=lambda: paper_changed())
                     for label, value in (("Portrait", "portrait"), ("Landscape", "landscape"))]
    for r in orient_radios:
        r.pack(side="left", padx=(0, 14))
    tk.Label(adv, text="Header:", bg="#f6f7f9").grid(row=7, column=0, sticky="nw", padx=(0, 8), pady=(10, 0))
    header_row = tk.Frame(adv, bg="#f6f7f9")
    header_row.grid(row=7, column=1, sticky="w", pady=(10, 0))
    header_box = ttk.Checkbutton(header_row, style="Panel.TCheckbutton", text="Print the file's name and page number "
                                 "at the top of each page", variable=header_on, command=lambda: header_changed())
    header_box.pack(side="left")
    header_entry = ttk.Entry(header_row, textvariable=header_var, width=34, font=fonts["body"])
    Tooltip(header_entry, "The header's text (the page number is added after it).")
    number_selected = tk.BooleanVar(value=layout_numbering(start_layout) == "selected")
    number_box = ttk.Checkbutton(adv, style="Panel.TCheckbutton", text="Number only the pages being printed "
                                 "(1 of 3, 2 of 3…)", variable=number_selected, command=lambda: header_changed())
    Tooltip(number_box, "On: only the pages you print get a page number, counted from 1. Off: every page has its "
                        "page number in the PDF.")

    def header_changed(*_):
        if header_on.get():
            if not header_entry.winfo_ismapped():
                header_entry.pack(side="left", padx=(10, 0))
            number_box.grid(row=8, column=1, sticky="w", pady=(4, 0))
        else:
            header_entry.pack_forget()
            number_box.grid_remove()
        if editable:
            lay["touched"] = True
            if lay["job"]:
                root.after_cancel(lay["job"])
            lay["job"] = root.after(350, apply_layout)
    header_var.trace_add("write", lambda *_: header_changed() if header_on.get() else None)
    if header_on.get():
        header_entry.pack(side="left", padx=(10, 0))
        number_box.grid(row=8, column=1, sticky="w", pady=(4, 0))
    if not editable:
        for w in orient_radios + [header_box, number_box]:
            w.state(["disabled"])
        header_entry.state(["disabled"])

    # cut marks: crop marks for paper cutters, or a cutting machine's whole Print & Cut setup
    tk.Label(adv, text="Cut marks:", bg="#f6f7f9").grid(row=13, column=0, sticky="nw", padx=(0, 8), pady=(10, 0))
    marks_var = tk.StringVar(value=CUT_MARKS[layout_marks(start_layout)])
    marks_box = ttk.Combobox(adv, textvariable=marks_var, values=list(CUT_MARKS.values()), width=26, state="readonly")
    marks_box.grid(row=13, column=1, sticky="w", pady=(10, 0))
    marks_box.bind("<<ComboboxSelected>>", lambda e: marks_changed())
    Tooltip(marks_box, "Crop marks: lines in the margins in line with every card edge, for a paper cutter.\n"
                       "Silhouette Print & Cut: registration marks on the front pages, the cards kept clear of "
                       "them, and cut lines (SVG + DXF) saved next to the PDF for Silhouette Studio.\n"
                       "Cricut Print Then Cut: the cards laid out in Design Space's 6.75 x 9.25 in area and "
                       "exported as PNG sheets (Design Space prints its own marks).")
    sil_row = tk.Frame(adv, bg="#f6f7f9")
    sil_vars = {}
    for key, label in (("inset", "Inset"), ("length", "Length"), ("thickness", "Thickness")):
        tk.Label(sil_row, text=label, bg="#f6f7f9").pack(side="left", padx=(0, 4))
        v = tk.StringVar(value=f"{SILHOUETTE_MARKS[key]:g}")
        e = ttk.Entry(sil_row, textvariable=v, width=5, justify="center")
        e.pack(side="left", padx=(0, 4))
        tk.Label(sil_row, text="mm", bg="#f6f7f9", fg="#666666").pack(side="left", padx=(0, 12))
        e.bind("<Return>", lambda _e: sil_changed())
        e.bind("<FocusOut>", lambda _e: sil_changed())
        sil_vars[key] = v
    Tooltip(sil_row, "The registration marks' size. Use the same values in Silhouette Studio: Page Setup > "
                     "Registration Marks (Type 1). Saved for every project.")
    marks_note = tk.Label(adv, text="", bg="#f6f7f9", fg="#666666", justify="left", wraplength=560)

    def marks_key():
        return next((k for k, v in CUT_MARKS.items() if v == marks_var.get()), "none")

    def show_marks_rows():
        key = marks_key()
        if key == "silhouette":
            sil_row.grid(row=14, column=1, sticky="w", pady=(4, 0))
        else:
            sil_row.grid_remove()
        tips = {"silhouette": "Print from Fleem, then cut in Silhouette Studio with the cut lines Fleem saves next "
                              "to the PDF. Landscape usually fits more cards.",
                "cricut": "Print and Export save PNG sheets for Design Space (Upload > Print Then Cut image). No "
                          "bleed and at least 6.5 mm between cards: Design Space adds its own bleed.",
                "crop": "Cut along the marks with a paper cutter (front pages only)."}
        if key in tips:
            marks_note.configure(text=tips[key])
            marks_note.grid(row=15, column=1, sticky="w", pady=(2, 0))
        else:
            marks_note.grid_remove()

    def marks_changed():
        marks_box.selection_clear()
        show_marks_rows()
        paper_changed()                        # (apply_layout sets Cricut's no bleed / wider gaps)
        if editable and layout_marks(lay["layout"]) != marks_key():
            log.warning("Cut marks: the layout didn't change to %s (%s)", marks_key(), describe_layout(lay["layout"]))
        photos.clear()                         # always draw the pages again with the new marks
        show(view["side"], keep_scroll=True)

    def sil_changed():
        values = {k: v.get().replace(",", ".").strip() for k, v in sil_vars.items()}
        before = dict(SILHOUETTE_MARKS)
        set_silhouette_marks(values["inset"], values["length"], values["thickness"])
        for k, v in sil_vars.items():
            v.set(f"{SILHOUETTE_MARKS[k]:g}")
        if SILHOUETTE_MARKS != before:
            for k in SILHOUETTE_MARKS:
                try:
                    save_setting(f"silhouette_mark_{k}", f"{SILHOUETTE_MARKS[k]:g}")
                except OSError as e:
                    log.warning("Couldn't save the Silhouette mark size: %s", e)
            if editable and layout_marks(lay["layout"]) == "silhouette":    # same layout, new marks: redraw
                counts_before = dict(counts)
                recount()
                if counts != counts_before:
                    chosen["front"].clear()
                    chosen["back"].clear()
                photos.clear()
                show(view["side"], keep_scroll=True)
                apply_layout()
    show_marks_rows()
    if not editable:
        marks_box.state(["disabled"])

    # where downloaded card images and temporary files are kept (for every project)
    tk.Label(adv, text="Cache folder:", bg="#f6f7f9").grid(row=9, column=0, sticky="nw", padx=(0, 8), pady=(10, 0))
    cache_row = tk.Frame(adv, bg="#f6f7f9")
    cache_row.grid(row=9, column=1, sticky="w", pady=(10, 0))
    cache_label = tk.Label(cache_row, bg="#f6f7f9", fg="#666666", text=str(cache_root(settings)))
    cache_label.pack(side="left")

    def change_cache():
        from tkinter import filedialog
        chosen = filedialog.askdirectory(parent=root, title="Where should Fleem keep its cache (downloaded card images)?",
                                         initialdir=str(cache_root(settings)), mustexist=True)
        if not chosen:
            return
        try:
            set_cache_folder(settings, Path(chosen))
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Couldn't use that folder: {e}", parent=root)
            return
        cache_label.configure(text=str(cache_root(settings)))

    def reset_cache():
        set_cache_folder(settings, None)
        cache_label.configure(text=str(cache_root(settings)))
    ttk.Button(cache_row, text="Change…", command=change_cache).pack(side="left", padx=(10, 0))
    ttk.Button(cache_row, text="Default", style="Ghost.TButton", command=reset_cache).pack(side="left", padx=(6, 0))
    Tooltip(cache_label, "Fleem keeps downloaded card images and temporary files in a \"Fleem cache\" folder here. "
                         "It's emptied whenever Fleem starts and quits.")

    # card order (for every project; changing it sorts this project again, undoing cards you moved)
    tk.Label(adv, text="Card order:", bg="#f6f7f9").grid(row=10, column=0, sticky="w", padx=(0, 8), pady=(10, 0))
    sort_var = tk.StringVar(value=CARD_SORTS[CARD_SORT])
    sort_box = ttk.Combobox(adv, textvariable=sort_var, values=list(CARD_SORTS.values()), width=24,
                            state="readonly")
    sort_box.grid(row=10, column=1, sticky="w", pady=(10, 0))
    sort_box.bind("<<ComboboxSelected>>", lambda e: card_sort_changed())
    Tooltip(sort_box, "MDFC then Alphabetical: double-faced cards first, so their backs share as few back pages "
                      "as possible, then the rest A-Z.\nOnly Alphabetical: every card A-Z.\n\nChanging it sorts "
                      "the cards again (cards you dragged to a new spot go back to theirs). New projects use it "
                      "too.")

    def card_sort_changed():
        key = next(k for k, v in CARD_SORTS.items() if v == sort_var.get())
        sort_box.selection_clear()
        if key == CARD_SORT:
            return
        try:
            save_setting("card_sort", key)
        except OSError as e:
            log.warning("Couldn't save the card order: %s", e)
        settings.card_sort = set_card_sort(key)
        log.info("Card order: %s", CARD_SORTS[key])
        if not editable:
            layout_note(f"Card order: {CARD_SORTS[key]} (for projects Fleem can lay out).")
            return
        resorted = sorted_cards(cards)
        if all(a is b for a, b in zip(resorted, cards)):
            layout_note(f"Card order: {CARD_SORTS[key]}. The cards were already in that order.")
            return
        moved = lay.pop("moved", False)
        cards[:] = resorted
        cards_changed_now(f"Cards sorted: {CARD_SORTS[key]}"
                          + (" (cards you moved went back to their spots)." if moved else "."))

    # where Export as PDF saves (for every project)
    tk.Label(adv, text="PDF export:", bg="#f6f7f9").grid(row=11, column=0, sticky="nw", padx=(0, 8), pady=(10, 0))
    export_row = tk.Frame(adv, bg="#f6f7f9")
    export_row.grid(row=11, column=1, sticky="w", pady=(10, 0))
    ttk.Button(export_row, text="Set PDF Export Location", command=lambda: set_export_location()).pack(side="left")
    export_label = tk.Label(export_row, bg="#f6f7f9", fg="#666666")
    export_label.pack(side="left", padx=(10, 0))

    def show_export_location():
        folder = export_folder_of(settings)
        export_label.configure(text=str(folder) if folder else "Fleem asks where to save each time")
    show_export_location()

    def set_export_location():
        from tkinter import filedialog
        start = export_folder_of(settings) or getattr(settings, "output_folder", None) or Path.home()
        chosen = filedialog.askdirectory(parent=root, title="Where should Export as PDF save?",
                                         initialdir=str(start), mustexist=True)
        if not chosen:
            return
        try:
            set_export_folder(settings, Path(chosen))
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Couldn't save the setting: {e}", parent=root)
            return
        show_export_location()
        messagebox.showinfo(APP_NAME, f"Saved. Export as PDF will save to this folder from now on:\n\n{Path(chosen)}",
                            parent=root)

    def reopen(how):
        if lay["added"] and not messagebox.askyesno(
                APP_NAME, "The cards you added or edited won't be kept. Open the PDF again?", parent=root):
            if foreign is not None:
                mode_var.set("pairs" if foreign["has_backs"] else "single")
            return
        result["choice"] = ("reopen", how)
        close()
    if foreign is not None or can_take_apart:
        tk.Label(adv, text="Original PDF:", bg="#f6f7f9").grid(row=20, column=0, sticky="nw", padx=(0, 8), pady=(10, 0))
        if foreign is not None:
            ttk.Button(adv, text="Print the original PDF as it is instead", command=lambda: reopen("plain")).grid(
                row=20, column=1, sticky="w", pady=(10, 0))
        else:
            ttk.Button(adv, text="Take the cards out of it (edit, change bleed/spacing, combine)",
                       command=lambda: reopen(None)).grid(row=20, column=1, sticky="w", pady=(10, 0))

    # back alignment (per printer): offsets only while it's switched on; 0 when off
    align_on = tk.BooleanVar(value=False)
    align_x, align_y = tk.StringVar(value="0.00"), tk.StringVar(value="0.00")
    align_state = {"loading": False}
    tk.Label(adv, text="Back alignment:", bg="#f6f7f9").grid(row=3, column=0, sticky="nw", padx=(0, 8), pady=(10, 0))
    ttk.Checkbutton(adv, style="Panel.TCheckbutton", text="Move the backs to line up with the fronts (for this printer)", variable=align_on,
                    command=lambda: align_toggled()).grid(row=3, column=1, sticky="w", pady=(10, 0))
    align_row = tk.Frame(adv, bg="#f6f7f9")
    tk.Label(align_row, text="X:", bg="#f6f7f9").pack(side="left")
    ttk.Entry(align_row, textvariable=align_x, width=6, font=fonts["body"]).pack(side="left", padx=(4, 2))
    tk.Label(align_row, text="mm (+ = right)     Y:", bg="#f6f7f9").pack(side="left")
    ttk.Entry(align_row, textvariable=align_y, width=6, font=fonts["body"]).pack(side="left", padx=(4, 2))
    tk.Label(align_row, text="mm (+ = down)", bg="#f6f7f9").pack(side="left")
    ttk.Button(align_row, text="Print test sheet", command=lambda: print_test_sheet()).pack(side="left", padx=(12, 0))
    align_note = tk.Label(adv, bg="#f6f7f9", fg="#666666", anchor="w", justify="left")

    def align_values():
        try:
            x, y = float(align_x.get().replace(",", ".")), float(align_y.get().replace(",", "."))
        except ValueError:
            return None
        return (x, y) if abs(x) <= 10 and abs(y) <= 10 else None

    def show_alignment_row():
        if align_on.get():
            align_row.grid(row=4, column=1, sticky="w", pady=(4, 0))
            align_note.grid(row=5, column=1, sticky="w", pady=(2, 0))
            T(align_note, fg="muted").configure(text="Measured on the back of the sheet as you look at it. Print the test sheet, "
                                      "hold it up to a light and read the ruler.")
        else:
            align_row.grid_remove()
            align_note.grid_remove()

    def load_printer_alignment(*_):
        printer = printer_var.get().strip()
        a = load_alignment(printer) if printer else {"enabled": False, "x_mm": 0.0, "y_mm": 0.0}
        align_state["loading"] = True
        align_on.set(a["enabled"])
        align_x.set(f"{a['x_mm']:.2f}")
        align_y.set(f"{a['y_mm']:.2f}")
        align_state["loading"] = False
        show_alignment_row()

    def save_printer_alignment(*_):
        if align_state["loading"]:
            return
        printer = printer_var.get().strip()
        if not printer:
            return
        value = align_values()
        if value is None:
            T(align_note, fg="danger").configure(text="X and Y must be numbers between -10 and 10 (mm).")
            return
        save_alignment(printer, align_on.get(), *value)
        if align_on.get():
            T(align_note, fg="muted").configure(text=f"Saved for {printer}: backs move {value[0]:+.2f} mm right, {value[1]:+.2f} mm down.")

    def align_toggled():
        if not align_on.get():                 # off: the offsets go back to 0
            align_state["loading"] = True
            align_x.set("0.00")
            align_y.set("0.00")
            align_state["loading"] = False
        save_printer_alignment()
        show_alignment_row()
    align_x.trace_add("write", save_printer_alignment)
    align_y.trace_add("write", save_printer_alignment)
    printer_box.bind("<<ComboboxSelected>>", load_printer_alignment, add="+")

    def print_test_sheet():
        printer = printer_var.get().strip()
        value = align_values()
        if not printer or value is None:
            T(align_note, fg="danger").configure(text="Pick a printer and enter valid offsets first.")
            return
        test = WORK_DIR / "alignment_test.pdf"
        WORK_DIR.mkdir(parents=True, exist_ok=True)
        try:
            write_alignment_test(test, lay["layout"])
            devmode = job_settings(printer, profile)
            root.configure(cursor="watch")
            T(align_note, fg="muted").configure(text="Printing the front of the test sheet…")
            root.update_idletasks()
            gdi_print(test, printer, [0], "Fleem alignment test (front)", rotate=settings.rotate_pages, devmode=devmode)
            root.configure(cursor="")
            if not messagebox.askokcancel(APP_NAME, "The front of the test sheet is printing.\n\nWhen it's done, put the "
                                          "same sheet back in the printer the way you do for card backs, then press OK "
                                          "to print the back.", parent=root):
                T(align_note, fg="muted").configure(text="Test sheet: back not printed.")
                return
            root.configure(cursor="watch")
            root.update_idletasks()
            gdi_print(test, printer, [1], "Fleem alignment test (back)", rotate=settings.rotate_pages, devmode=devmode,
                      back_pages={1}, back_offset=value)
            T(align_note, fg="muted").configure(text="Hold the sheet up to a light, looking at the back: how far must the bold cross "
                                      "move to sit on the thin one? Right = +X, down = +Y. Add it to X / Y, then print "
                                      "the test again to check.")
        except Exception as e:
            log.error("Alignment test failed: %s", e)
            T(align_note, fg="danger").configure(text=f"Couldn't print the test sheet: {e}")
        finally:
            root.configure(cursor="")
    load_printer_alignment()

    def toggle_advanced(show_it=None):
        open_now = adv.winfo_ismapped() if show_it is None else not show_it
        if open_now:
            adv.pack_forget()
            adv_toggle.configure(text="Advanced settings ▸")
        else:
            adv.pack(fill="x", pady=(8, 0))
            adv_toggle.configure(text="Advanced settings ▾")
    adv_toggle.configure(command=toggle_advanced)

    def set_mode():
        if foreign is not None:              # cards from a PDF made elsewhere: read it again that way
            if (mode_var.get() == "pairs") != foreign["has_backs"]:
                reopen(mode_var.get())
            return
        if mode_var.get() == paging["mode"]:
            return
        paging["mode"] = mode_var.get()
        recount()
        chosen["front"].clear()
        chosen["back"].clear()
        photos.clear()
        show("front")

    def changed() -> bool:
        """Does the PDF have to be written (new from an XML, or the deck/layout changed)?"""
        if not editable:
            return False
        if pdf_layout is not None and len(pdf_layout) > 2 and pdf_layout[2] != layout_paper(lay["layout"]):
            return True                          # another paper size
        if pdf_layout is not None and layout_card(pdf_layout) != layout_card(lay["layout"]):
            return True                          # another card size
        if pdf_layout is not None:
            header = layout_header(lay["layout"])
            if header == "":
                header = default_header
            if layout_orientation(pdf_layout) != layout_orientation(lay["layout"]) or layout_header(pdf_layout) != header:
                return True                      # turned, or the header changed
            if header is not None and layout_numbering(lay["layout"]) == "selected":
                return True                      # the page numbers depend on the pages picked
        return from_xml or foreign is not None or lay["cards_changed"] or pdf_layout is None or not (
            abs(pdf_layout[0] - lay["layout"][0]) < 0.02 and
            (pdf_layout[1] is None or abs(pdf_layout[1] - lay["layout"][1]) < 0.02))

    def layout_note(extra=""):
        if extra:
            bad = extra.startswith("!")
            set_status(extra[1:] if bad else extra, "danger" if bad else "muted")
            return
        # only things worth knowing - no routine hints
        if not editable and can_take_apart:
            text = "Printing the original PDF as it is (Advanced settings > Take the cards out of it, to edit it)."
        elif not editable:
            text = "This PDF can only be printed as it is - its bleed edge and card spacing can't be changed."
        elif unnamed_warning():
            text = f"{UNNAMED_LAYOUT_NOTE}, so these are last time's values (you'll be asked to confirm)."
        else:
            text = ""
        if foreign is not None:
            text = (text + "  " if text else "") + f"Print saves a new Fleem PDF: {Path(out_path).name}."
        set_status(text, "warn" if unnamed_warning() else "muted")

    def unnamed_warning() -> bool:
        """An add-on sets layouts from file names, this one's name doesn't say, and nothing was changed."""
        return bool(UNNAMED_LAYOUT_NOTE) and editable and lay["source"] == "last" and not lay["touched"]

    def page_extras():
        """Paper, orientation and header from the boxes."""
        header = None
        if header_on.get():
            text = header_var.get().strip()
            header = "" if text in ("", default_header) else text
        return (paper_var.get(), orientation_var.get(), header, "selected" if number_selected.get() else "all",
                card_size["value"], marks_key())

    def current_layout():
        value = parse_layout(bleed_var.get(), spacing_var.get())
        return None if isinstance(value, str) else value + page_extras()

    def card_changed():
        choice = card_var.get()
        if choice == CUSTOM_CHOICE:
            value = ask_custom_card(root, card_size["value"])
            if value is None:                  # cancelled: back to what it was
                card_var.set(card_box_label(card_size["value"]))
                return
        else:
            value = choice if choice in CARD_SIZES else card_size["value"]
        card_size["value"] = value
        card_var.set(card_box_label(value))
        root.after_idle(lambda: card_box.selection_clear())
        paper_changed()

    def paper_changed():
        nonlocal aspect
        if not editable:
            return
        lay["touched"] = True
        w, h = page_size((0, 0, paper_var.get(), orientation_var.get()))
        aspect = h / w
        apply_layout()
        rescale(force=True)

    def printed_pages():
        """The PDF pages being printed (1 = first page), in order."""
        return sorted([page_index("front", f) + 1 for f in chosen["front"]] +
                      [page_index("back", b) + 1 for b in chosen["back"]])

    def page_header(side, n, live=False):
        """The header (with its page number) for a preview, or None. With "number only the pages being
        printed", the small previews show just the text (the numbers change with every click) and the
        full-page view (`live`) numbers the picked pages."""
        text = layout_header(lay["layout"])
        if text is None:
            return None
        text = text or default_header
        if layout_numbering(lay["layout"]) == "selected":
            if not live:
                return text
            picked, page = printed_pages(), page_index(side, n) + 1
            return header_line(text, picked.index(page) + 1 if page in picked else None, len(picked))
        return header_line(text, page_index(side, n) + 1, 2 * max(1, counts["front"]))

    def apply_layout():
        """Use the boxes' values for the previews (nothing is written until Print)."""
        lay["job"] = None
        value = parse_layout(bleed_var.get(), spacing_var.get())
        if isinstance(value, str):
            layout_note("!" + value)
            return False
        value = machine_layout(value + page_extras())
        if value[:2] != parse_layout(bleed_var.get(), spacing_var.get()):   # (Cricut: no bleed, wider gaps)
            syncing["on"] = True
            try:
                bleed_var.set(f"{value[0]:.2f}")
                spacing_var.set(f"{value[1]:.2f}")
            finally:
                syncing["on"] = False
        if value != lay["layout"]:
            before = dict(counts)
            lay["layout"] = value
            recount()
            if counts != before:               # cards moved to other pages: pick again
                chosen["front"].clear()
                chosen["back"].clear()
            photos.clear()
            show(view["side"], keep_scroll=True)
        marks = layout_marks(value)
        if marks in ("silhouette", "cricut"):
            per = len(card_slots(value))
            turned = value[:3] + ("landscape" if layout_orientation(value) == "portrait" else "portrait",) + value[4:]
            better = len(card_slots(turned)) > per
            layout_note(f"{CUT_MARKS[marks]}: {per} card{'s' * (per != 1)} per sheet"
                        + (f" ({layout_orientation(turned)} fits {len(card_slots(turned))})" if better else "")
                        + (f", {cricut_size_text(value)}" if marks == "cricut" else "") + ".")
        else:
            layout_note()
        return True

    def layout_typed(*_):
        lay["touched"] = True
        if lay["job"]:
            root.after_cancel(lay["job"])
        value = parse_layout(bleed_var.get(), spacing_var.get())
        if isinstance(value, str):
            layout_note("!" + value)
            return
        lay["job"] = root.after(350, apply_layout)     # redraw once you stop typing

    def layout_enter(_event):
        if lay["job"]:
            root.after_cancel(lay["job"])
        apply_layout()
        return "break"                     # Enter here doesn't press Print
    bleed_var.trace_add("write", layout_typed)
    spacing_var.trace_add("write", layout_typed)
    for e in (bleed_entry, spacing_entry):
        e.bind("<Return>", layout_enter)
        if not editable:
            e.state(["disabled"])
    if not editable:
        paper_box.state(["disabled"])
        card_box.state(["disabled"])

    T(tk.Frame(root, height=1), bg="border").pack(fill="x")

    # --- toolbar: Fronts | Backs, selection count, Select all
    toolbar = T(tk.Frame(root, padx=18, pady=10), bg="bg")
    toolbar.pack(fill="x")
    seg = T(tk.Frame(toolbar), bg="bg")
    seg.pack(side="left")
    front_seg = ttk.Button(seg, style="SegOn.TButton", command=lambda: show("front"))
    front_seg.pack(side="left")
    back_seg = ttk.Button(seg, style="Seg.TButton", command=lambda: show("back"))
    back_seg.pack(side="left", padx=(4, 0))
    Tooltip(back_seg, "Show the back pages (Ctrl+Tab switches). Blank backs are labelled.")
    all_btn = ttk.Button(toolbar, width=12)
    all_btn.pack(side="right")
    sel_label = T(tk.Label(toolbar, font=fonts["label"]), bg="bg", fg="muted")
    sel_label.pack(side="right", padx=(0, 12))
    view_btn = ttk.Button(seg, text="View page", command=lambda: open_viewer(view["anchor"] or 1))
    view_btn.pack(side="left", padx=(16, 0))
    Tooltip(view_btn, "See a page full size (or double-click a page), and go through them with the arrow keys.")
    zoom_out = ttk.Button(seg, text="−", width=3, command=lambda: set_zoom(zoom["f"] / 1.2))
    zoom_out.pack(side="left", padx=(16, 0))
    zoom_in = ttk.Button(seg, text="+", width=3, command=lambda: set_zoom(zoom["f"] * 1.2))
    zoom_in.pack(side="left", padx=(4, 0))
    Tooltip(zoom_in, "Bigger previews (Ctrl + mouse wheel)")
    Tooltip(zoom_out, "Smaller previews (Ctrl + mouse wheel)")

    # --- middle: scrollable grid of pages
    mid = T(tk.Frame(root), bg="bg")
    mid.pack(fill="both", expand=True, padx=(12, 0))
    canvas = T(tk.Canvas(mid, highlightthickness=0, bd=0), bg="bg")
    bar = ttk.Scrollbar(mid, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=bar.set)
    bar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    grid = T(tk.Frame(canvas, pady=4), bg="bg")
    canvas.create_window((0, 0), window=grid, anchor="nw")
    grid.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    root.bind_all("<MouseWheel>", lambda e: (set_zoom(zoom["f"] * (1.1 if e.delta > 0 else 1 / 1.1))
                                             if e.state & 0x0004 else canvas.yview_scroll(int(-e.delta / 120), "units")))
    root.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
    root.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

    # --- footer: typed selection, Print / Cancel, hints
    T(tk.Frame(root, height=1), bg="border").pack(fill="x")
    footer = T(tk.Frame(root, padx=18, pady=10), bg="surface")
    footer.pack(fill="x")
    bottom = T(tk.Frame(footer), bg="surface")
    bottom.pack(fill="x")
    entry_label = T(tk.Label(bottom, font=fonts["label"]), bg="surface", fg="text")
    entry_label.pack(side="left")
    entry_var = tk.StringVar()
    entry = ttk.Entry(bottom, textvariable=entry_var, width=26, font=fonts["body"])
    entry.pack(side="left", padx=(8, 6))
    ttk.Button(bottom, text="Cancel", command=cancel).pack(side="right")
    print_btn = ttk.Button(bottom, text="Print", style="Accent.TButton")
    print_btn.pack(side="right", padx=(0, 8))
    export_btn = ttk.Button(bottom, text="Export as PDF")
    export_btn.pack(side="right", padx=(0, 8))
    Tooltip(export_btn, "Save the picked pages as a PDF of their own (nothing is printed).")
    xml_btn = ttk.Button(bottom, text="Export as XML", command=lambda: do_export_xml())   # shown when possible
    Tooltip(xml_btn, "Save the whole project as an MPCFill order XML (for MPCFill or MakePlayingCards). Only "
                     "when every card is MPCFill art - not added images, PDFs or replaced/edited art.")
    hint = tk.Label(bottom)                              # (not shown)

    hover = {"n": None}

    def draw_check(c):
        """The round ✓ badge on a picked page (drawn, so it doesn't depend on the font)."""
        P = palette()
        c.delete("all")
        c.configure(bg=P["paper"])
        c.create_oval(1, 1, 21, 21, fill=P["accent"], outline=P["accent"])
        c.create_line(6.5, 11.5, 9.8, 14.8, 15.8, 7.8, fill=P["accent_text"], width=2.4, capstyle="round",
                      joinstyle="round")

    def set_hover(n, leaving=None):
        if n is None and hover["n"] != leaving:
            return
        old, hover["n"] = hover["n"], n
        for k in {old, n} - {None}:
            if k in cells and k not in side_selected():
                T(cells[k][0], bg="tile_hover" if k == n else "tile")

    def side_selected() -> set:
        return chosen[view["side"]]

    syncing = {"on": False}

    def update(rewrite=False):
        side, sel = view["side"], side_selected()
        count = counts[side]
        for n, (frame, caption, badge) in cells.items():
            on = n in sel
            T(frame, bg="accent" if on else ("tile_hover" if hover.get("n") == n else "tile"))
            T(caption, fg="accent" if on else "muted")
            if on:
                draw_check(badge)
                badge.place(relx=1.0, x=-8, y=8, anchor="ne")
            else:
                badge.place_forget()
        sel_label.configure(text=f"{len(sel)} of {count} selected" if count else "")
        if rewrite:
            syncing["on"] = True
            entry_var.set(describe_fronts(sel))
            syncing["on"] = False
            entry.configure(foreground="")
        all_btn.configure(text="Deselect all" if count and len(sel) == count else "Select all")
        printer = printer_var.get().strip()
        if not printer:
            note, tone = "⚠  Pick a printer.", "warn"
        elif has_saved_settings(printer, profile):
            note, tone = (f"✓  Uses your saved printer settings for {profile_label(profile)}." if profile_label(profile)
                          else "✓  Uses your saved settings for this printer."), "ok"
        elif profile_label(profile):
            note, tone = (f"ⓘ  Uses the printer's default settings - click Printer settings… to pick the preset "
                          f"(Fleem remembers it for {profile_label(profile)})."), "warn"
        else:
            note, tone = "ⓘ  Uses the printer's default settings.", "warn"
        T(printer_label, fg=tone).configure(text=note)
        if editable:
            size_chip.configure(text=f"{len(cards)} card{'s' * (len(cards) != 1)} · {counts['front']} "
                                     f"sheet{'s' * (counts['front'] != 1)}")
        else:
            size_chip.configure(text=f"{pages_total} page{'s' * (pages_total != 1)}")
        nf, nb = len(chosen["front"]), len(chosen["back"])
        if nf and nb:
            label = f"Print {nf} front{'s' if nf != 1 else ''} + {nb} back{'s' if nb != 1 else ''}"
        elif nf or nb:
            which = noun("front") if nf else "back"
            n_sel, total = (nf, counts["front"]) if nf else (nb, counts["back"])
            label = f"Print {n_sel} of {total} {which}{'s' if total != 1 else ''}"
        else:
            label = "Print"
        if layout_marks(lay["layout"]) == "cricut":      # Design Space prints: Fleem makes the sheets
            label = f"Export {nf} Cricut sheet{'s' * (nf != 1)}" if nf else "Export Cricut sheets"
            print_btn.configure(text=label, state="normal" if nf else "disabled")
        else:
            print_btn.configure(text=label, state="normal" if (nf or nb) and printer else "disabled")
        export_btn.configure(state="normal" if nf or nb else "disabled")
        xml_ok = editable and can_export_xml(cards)
        if xml_ok and not xml_btn.winfo_ismapped():
            xml_btn.pack(side="right", padx=(0, 8), after=export_btn)
        elif not xml_ok and xml_btn.winfo_ismapped():
            xml_btn.pack_forget()

    def click(n, event):
        sel = side_selected()
        if event.state & 0x0001 and view["anchor"] is not None:          # Shift+click: a range
            a, b = sorted((view["anchor"], n))
            want = view["anchor"] in sel
            for k in range(a, b + 1):
                (sel.add if want else sel.discard)(k)
        else:
            sel.symmetric_difference_update({n})
            view["anchor"] = n
        update(rewrite=True)

    def to_photo(pil_image):
        if ImageTk is not None:
            return ImageTk.PhotoImage(pil_image, master=root)
        rgb = pil_image.convert("RGB")
        return tk.PhotoImage(master=root, data=b"P6 %d %d 255\n" % rgb.size + rgb.tobytes(), format="ppm")

    def pdf_page_image(index, width):
        """A page of the opened PDF as a PIL image (or PPM data if PIL can't), `width` px wide."""
        with PDFIUM_LOCK:
            page = pdf[index]
            scale = width / page.get_width()
            if ImageTk is not None:
                return page.render(scale=scale).to_pil()
            bmp = page.render(scale=scale, rev_byteorder=True)
            w, h, stride, buf, ch = bmp.width, bmp.height, bmp.stride, bytes(bmp.buffer), bmp.n_channels
        rows = b"".join(buf[y * stride:y * stride + w * ch] for y in range(h))
        if ch == 4:
            rows = bytes(b for i in range(0, len(rows), 4) for b in rows[i:i + 3])
        return b"P6 %d %d 255\n" % (w, h) + rows

    def preview_signature():
        """What the previews depend on (the layout, the cut marks' size, the cards, the width)."""
        if not editable:
            return (size["w"],)
        return (lay["layout"], tuple(sorted(SILHOUETTE_MARKS.items())), size["w"],
                tuple(id(c) for c in cards))

    def as_photo(image):
        return tk.PhotoImage(master=root, data=image, format="ppm") if isinstance(image, bytes) else to_photo(image)

    def load_preview(gen, side, _first=1):
        """Draw the page previews on a background thread (so the window stays smooth) and show each one
        as it's ready. A newer call (gen) makes an older one stop."""
        import queue
        if gen != view["gen"] or not root.winfo_exists():
            return
        total, width = counts[side], size["w"]
        made_for = preview_signature()
        if editable:                             # a snapshot: the drawing thread never touches live state
            layout = lay["layout"]
            sheets = [list(s) for s in deck_pages(cards, layout)]
            headers = {n: page_header(side, n) for n in range(1, total + 1)}
        else:
            indexes = {n: page_index(side, n) for n in range(1, total + 1)}
        done = queue.Queue()

        def work():
            for n in range(1, total + 1):
                if gen != view["gen"]:
                    return
                try:
                    if editable:
                        image = renderer.page(sheets[n - 1], side == "back", layout, width, headers[n])
                    else:
                        image = pdf_page_image(indexes[n], width)
                    done.put((n, image))
                except Exception as e:
                    log.info("Couldn't preview %s %d: %s", side, n, e)
            done.put(None)
        threading.Thread(target=work, daemon=True, name="previews").start()

        def drain():
            if gen != view["gen"] or not root.winfo_exists():
                return
            if preview_signature() != made_for:  # the layout changed under these previews: draw them again
                photos.clear()
                show(view["side"], keep_scroll=True)
                return
            try:
                for _ in range(8):               # a few at a time, so the window keeps responding
                    item = done.get_nowait()
                    if item is None:
                        return
                    n, image = item
                    if n in cells:
                        photo = as_photo(image)
                        photos[(side, n)] = photo
                        cells[n][0].winfo_children()[0].configure(image=photo)
            except queue.Empty:
                pass
            later(25, drain)
        later(25, drain)

    def find_blank_backs():
        blank_backs.clear()
        if editable:
            for n, sheet in enumerate(deck_pages(cards, lay["layout"]), 1):
                if not any(c["back"] for c in sheet):
                    blank_backs.add(n)
        elif pdf is not None and paging["mode"] == "pairs":
            for n in range(1, counts["back"] + 1):
                try:
                    with PDFIUM_LOCK:
                        bmp = pdf[page_index("back", n)].render(scale=24 / pdf[0].get_width())
                        data = bytes(bmp.buffer)
                    if not data or min(data) >= 250:
                        blank_backs.add(n)
                except Exception:
                    pass

    def show(side, keep_scroll=False):
        """Fill the grid with fronts or backs."""
        scrolled = canvas.yview()[0]
        view.update(side=side, anchor=None, gen=view["gen"] + 1)
        if side == "back":
            find_blank_backs()                 # label the blank backs so the MDFC ones stand out
        for child in grid.winfo_children():
            child.destroy()
        cells.clear()
        can_preview = editable or pdf is not None
        for n in range(1, counts[side] + 1):
            frame = T(tk.Frame(grid, padx=3, pady=3, cursor="hand2"), bg="tile")
            cols = grid_cols["n"]
            frame.grid(row=(n - 1) // cols * 2, column=(n - 1) % cols, padx=6, pady=(6, 0))
            image = photos.get((side, n))
            if image is None:   # a blank image sizes the tile until its preview is drawn
                if ("blank",) not in photos:
                    photos[("blank",)] = tk.PhotoImage(master=root, width=size["w"], height=size["h"])
                image = photos[("blank",)]
            pic = T(tk.Label(frame, image=image, bd=0, compound="center", text="" if can_preview else str(n),
                             font=fonts["big"]), bg="paper", fg="faint")
            pic.pack()
            badge = T(tk.Canvas(frame, width=22, height=22, highlightthickness=0, bd=0), bg="paper")
            blank = side == "back" and n in blank_backs
            caption = T(tk.Label(grid, text=f"{n}  ·  blank" if blank else str(n), font=fonts["caption"],
                                 cursor="hand2"), bg="bg", fg="muted")
            caption.grid(row=(n - 1) // cols * 2 + 1, column=(n - 1) % cols, pady=(2, 4))
            for wdg in (frame, pic, caption, badge):
                wdg.bind("<Button-1>", lambda e, n=n: press(n, e))
                wdg.bind("<B1-Motion>", lambda e, n=n: drag_motion(n, e))
                wdg.bind("<ButtonRelease-1>", lambda e, n=n: release(n, e))
                wdg.bind("<Double-Button-1>", lambda e, n=n: double_click(n))
                wdg.bind("<Button-3>", lambda e, n=n: card_menu(n, e))
                wdg.bind("<Enter>", lambda e, n=n: set_hover(n))
                wdg.bind("<Leave>", lambda e, n=n: set_hover(None, n))
            cells[n] = (frame, caption, badge)
        grid.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.yview_moveto(scrolled if keep_scroll else 0)
        word = noun(side).capitalize() + "s"
        front_word = noun("front").capitalize() + "s"
        front_seg.configure(text=f"{front_word}  {counts['front']}", style="SegOn.TButton" if side == "front" else "Seg.TButton")
        back_seg.configure(text=f"Backs  {counts['back']}", style="SegOn.TButton" if side == "back" else "Seg.TButton",
                           state="normal" if counts["back"] else "disabled")
        if paging["mode"] == "single":
            back_seg.pack_forget()             # no backs to switch to
        elif not back_seg.winfo_ismapped():
            back_seg.pack(side="left", padx=(4, 0))
        entry_label.configure(text=f"{word} to print")
        hint.configure(text="e.g. 1, 5-8, 10-11  ·  click to pick, Shift+click for a range, double-click to view"
                            + ("  ·  drag a card to move it  ·  drag in cards, XMLs or PDFs" if editable else ""))
        update(rewrite=True)
        if can_preview:
            later(20, load_preview, view["gen"], side, 1)
        checkpoint()                           # (anything that redraws the pages may have changed the project)

    def toggle_all():
        sel = side_selected()
        if len(sel) == counts[view["side"]]:
            sel.clear()
        else:
            sel.update(range(1, counts[view["side"]] + 1))
        update(rewrite=True)

    def typed(*_args):
        if syncing["on"]:
            return
        numbers = parse_fronts(entry_var.get(), counts[view["side"]])
        if numbers is None:
            entry.configure(foreground=palette()["danger"])
            return
        entry.configure(foreground="")
        sel = side_selected()
        sel.clear()
        sel.update(numbers)
        update()

    def do_print(_event=None):
        typed()
        if editable and layout_marks(lay["layout"]) == "cricut":
            return export_cricut()
        printer = printer_var.get().strip()
        if not (chosen["front"] or chosen["back"]) or not printer:
            return
        if lay["job"]:
            root.after_cancel(lay["job"])
            lay["job"] = None
        layout = lay["layout"]
        if editable:
            before = dict(counts)
            if not apply_layout():
                return
            layout = lay["layout"]
            if counts != before:
                messagebox.showinfo(APP_NAME, "With these values the cards moved to different pages, so your "
                                    "picks were cleared. Pick the pages again, then press Print.", parent=root)
                return
        fronts, backs = sorted(chosen["front"]), sorted(chosen["back"])
        checks = []
        if unnamed_warning():
            checks.append(f"{UNNAMED_LAYOUT_NOTE}, so Fleem used last time's values:\n"
                          f"    {describe_layout(layout)}")
        if fronts and backs:      # both sides: say exactly what will print
            checks.append("You've picked fronts AND backs:\n"
                          f"    Fronts: {describe_fronts(fronts)}\n"
                          f"    Backs:  {describe_fronts(backs)}\n"
                          f"That's {len(fronts) + len(backs)} sheets, printed one side each in page order "
                          "(front 1, back 1, front 2, back 2, …).")
        if checks:
            ok = messagebox.askokcancel(f"{APP_NAME} — is this right?", "\n\n".join(checks) + "\n\nPrint?",
                                        icon="warning", default="cancel", parent=root)
            if not ok:
                return
        print_path = pdf_path
        if editable:
            remember_layout(layout)
            if changed():                       # only now is the PDF written
                layout_note("Saving the PDF…")
                root.update_idletasks()
                try:
                    write_deck_pdf(cards, layout, out_path, numbered=printed_pages())
                except Exception as e:
                    log.error("Couldn't save %s: %s", out_path, e)
                    layout_note(f"!Couldn't save the PDF: {e}")
                    return
                log.info("Saved %s (%s, %d cards)", Path(out_path).name, describe_layout(layout), len(cards))
            print_path = out_path
        parts = ([f"{noun('front')}s {describe_fronts(fronts)}"] if fronts else []) + \
                ([f"backs {describe_fronts(backs)}"] if backs else [])
        indexes = sorted([page_index("front", n) for n in fronts] + [page_index("back", n) for n in backs])
        backs_idx = [page_index("back", n) for n in backs] if paging["mode"] == "pairs" else []
        if editable and layout_marks(layout) == "silhouette" and fronts:
            silhouette_files(layout, fronts, Path(print_path).parent, Path(print_path).stem)
        result["choice"] = (printer, " and ".join(parts), indexes, Path(print_path), backs_idx)
        close()

    def silhouette_files(layout, fronts, folder, stem):
        """Cut lines for Silhouette Studio next to the PDF (one set per number of cards on a sheet), and
        how to use them (once)."""
        sheets = deck_pages(cards, layout)
        per = len(card_slots(layout))
        written = []
        try:
            for count in sorted({len(sheets[n - 1]) for n in fronts if n - 1 < len(sheets)}, reverse=True):
                name = f"{stem} - cut lines" + ("" if count == per else f" ({count} cards)")
                written += write_cut_files(layout, count, folder, name)
        except OSError as e:
            log.error("Couldn't save the cut lines: %s", e)
            messagebox.showerror(APP_NAME, f"Couldn't save the Silhouette cut lines: {e}", parent=root)
            return
        log.info("Silhouette cut lines: %s", ", ".join(p.name for p in written))
        if settings.raw.get("silhouette_help_shown") == "yes":
            return
        mk = SILHOUETTE_MARKS
        x, y, _, h = trim_boxes(layout, 1)[0]
        messagebox.showinfo(f"{APP_NAME} — Silhouette Print & Cut", (
            f"Cut lines saved next to the PDF:\n{written[0].with_suffix('').name} (.svg and .dxf)\n\n"
            "1. Print the front pages from Fleem (they have the registration marks).\n"
            f"2. In Silhouette Studio: Page Setup → {layout_paper(layout)}, {layout_orientation(layout)}. "
            f"Registration Marks → Type 1, Length {mk['length']:g} mm, Thickness {mk['thickness']:g} mm, "
            f"Inset {mk['inset']:g} mm.\n"
            "3. Open the .dxf (Basic edition) or .svg (Designer edition and up). The first card should sit "
            f"{x / PT_PER_MM:.1f} mm from the left and {(page_size(layout)[1] - y - h) / PT_PER_MM:.1f} mm from "
            "the top - if Studio moved the lines, select them all and set that position.\n"
            "4. Put the printed sheet on the mat, then Send: the machine finds the marks and cuts.\n\n"
            "(This message shows once. The README has it too.)"), parent=root)
        try:
            save_setting("silhouette_help_shown", "yes")
            settings.raw["silhouette_help_shown"] = "yes"
        except OSError:
            pass

    def export_cricut():
        """Cricut Print Then Cut: the picked sheets as see-through PNGs for Design Space."""
        from tkinter import filedialog
        if lay["job"]:
            root.after_cancel(lay["job"])
            lay["job"] = None
        before = dict(counts)
        if not apply_layout():
            return
        if counts != before:
            messagebox.showinfo(APP_NAME, "With these values the cards moved to different sheets, so your picks "
                                "were cleared. Pick the sheets again.", parent=root)
            return
        fronts = sorted(chosen["front"])
        if not fronts:
            messagebox.showinfo(APP_NAME, "Pick the sheets (front pages) to export for the Cricut.", parent=root)
            return
        layout = lay["layout"]
        folder = export_folder_of(settings)
        if folder is None or not folder.is_dir():
            chosen_dir = filedialog.askdirectory(
                parent=root, title="Where should the Cricut sheets be saved?", mustexist=True,
                initialdir=str(getattr(settings, "output_folder", None) or Path.home()))
            if not chosen_dir:
                return
            folder = Path(chosen_dir)
        stem = safe_name(Path(out_path).stem)
        size_text = cricut_size_text(layout)
        sheets = deck_pages(cards, layout)
        layout_note("Making the Cricut sheets…")
        root.update_idletasks()
        written = []
        try:
            with renderer.lock:
                for n in fronts:
                    target = free_path(folder / f"{stem} - Cricut sheet {n} ({size_text}).png")
                    cricut_sheet(sheets[n - 1], layout, renderer).save(target, dpi=(300, 300))
                    written.append(target)
        except Exception as e:
            log.error("Couldn't make the Cricut sheets: %s", e)
            layout_note(f"!Couldn't make the Cricut sheets: {e}")
            return
        remember_layout(layout)
        log.info("Cricut sheets: %s", ", ".join(p.name for p in written))
        layout_note(f"Saved {len(written)} Cricut sheet{'s' * (len(written) != 1)} to {folder}")
        messagebox.showinfo(f"{APP_NAME} — Cricut Print Then Cut", (
            f"Saved {len(written)} sheet{'s' * (len(written) != 1)} to:\n{folder}\n\n"
            "In Cricut Design Space, for each sheet:\n"
            "1. Upload → Upload Image → pick the PNG → Complex → Continue (the background is already see-through) "
            "→ Print Then Cut image → Upload.\n"
            f"2. Add it to the canvas and set its size to exactly {size_text} (unlock the size lock if needed).\n"
            "3. Make It → turn Bleed on → Print, then load the sheet and cut. Design Space adds the sensor marks "
            "and cuts around each card.\n\n"
            "Double-sided cards: Print Then Cut prints one side, so the backs aren't included."), parent=root)

    def do_export_xml():
        """The whole project as an MPCFill XML (every card copy, in print order)."""
        from tkinter import filedialog
        if not can_export_xml(cards):
            messagebox.showinfo(APP_NAME, "Only projects whose cards are all MPCFill art can be saved as an XML "
                                "(added images, PDFs and replaced or edited art aren't on MPCFill).", parent=root)
            return
        name = safe_name(Path(out_path).stem) + ".xml"
        folder = export_folder_of(settings)
        if folder is not None and folder.is_dir():
            target = free_path(folder / name)
        else:
            chosen_path = filedialog.asksaveasfilename(
                parent=root, title="Export as MPCFill XML", initialfile=name, defaultextension=".xml",
                initialdir=str(getattr(settings, "output_folder", None) or Path.home()),
                filetypes=[("MPCFill order XML", "*.xml")])
            if not chosen_path:
                return
            target = Path(chosen_path)
        try:
            target.write_text(mpcfill_xml(cards), encoding="utf-8")
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Couldn't save the XML: {e}", parent=root)
            return
        n_back = sum(1 for c in cards if c["back"])
        log.info("Exported %s (%d cards) as an MPCFill XML", target, len(cards))
        layout_note(f"Saved {target.name}: {len(cards)} card{'s' * (len(cards) != 1)} as an MPCFill XML "
                    f"({target.parent})." + (" Cards without a back get MPCFill's project card back there."
                                             if n_back < len(cards) else ""))

    def do_export(_event=None):
        """Export as PDF: the picked pages as a PDF of their own (nothing is printed, the window stays)."""
        from tkinter import filedialog
        typed()
        if editable and layout_marks(lay["layout"]) == "cricut":
            return export_cricut()
        if not (chosen["front"] or chosen["back"]):
            return
        if lay["job"]:
            root.after_cancel(lay["job"])
            lay["job"] = None
        layout = lay["layout"]
        if editable:
            before = dict(counts)
            if not apply_layout():
                return
            layout = lay["layout"]
            if counts != before:
                messagebox.showinfo(APP_NAME, "With these values the cards moved to different pages, so your "
                                    "picks were cleared. Pick the pages again, then press Export as PDF.", parent=root)
                return
        fronts, backs = sorted(chosen["front"]), sorted(chosen["back"])
        indexes = sorted([page_index("front", n) for n in fronts] + [page_index("back", n) for n in backs])
        name = safe_name(f"{Path(out_path if editable else pdf_path).stem} - export") + ".pdf"
        folder = export_folder_of(settings)
        if folder is not None and folder.is_dir():
            target = free_path(folder / name)
        else:
            if folder is not None:
                log.info("The PDF export location %s isn't there any more: asking", folder)
            start = folder if folder is not None and folder.is_dir() else \
                Path(getattr(settings, "output_folder", None) or Path.home())
            chosen_path = filedialog.asksaveasfilename(
                parent=root, title="Export as PDF - where should it be saved?", initialdir=str(start),
                initialfile=name, defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if not chosen_path:
                return
            target = Path(chosen_path)
            if messagebox.askyesno(f"{APP_NAME} — export location",
                                   f"Always save exported PDFs to this folder?\n\n{target.parent}\n\n"
                                   "Yes: Fleem saves there from now on without asking.\nNo: Fleem asks each time.",
                                   parent=root):
                try:
                    set_export_folder(settings, target.parent)
                    show_export_location()
                    messagebox.showinfo(APP_NAME, "Saved - exported PDFs will go to this folder from now on.\n\n"
                                        "To change it at any time: Advanced settings > Set PDF Export Location.",
                                        parent=root)
                except OSError as e:
                    messagebox.showerror(APP_NAME, f"Couldn't save the setting: {e}", parent=root)
        layout_note("Exporting the PDF…")
        root.update_idletasks()
        temp = None
        try:
            src = pdf_path
            if editable:                       # the deck as it is now (with its edits), then those pages
                temp = WORK_DIR / f"export-{os.getpid()}-{int(time.time() * 1000)}.pdf"
                write_deck_pdf(cards, layout, temp, numbered=printed_pages())
                src = temp
            export_pdf_pages(src, indexes, target)
        except Exception as e:
            log.error("Couldn't export %s: %s", target, e)
            layout_note(f"!Couldn't export the PDF: {e}")
            return
        finally:
            if temp is not None:
                Path(temp).unlink(missing_ok=True)
        if editable:
            remember_layout(layout)
            if layout_marks(layout) == "silhouette" and fronts:
                silhouette_files(layout, fronts, target.parent, target.stem)
        log.info("Exported %d page(s) to %s", len(indexes), target)
        layout_note(f"Exported {len(indexes)} page{'s' * (len(indexes) != 1)} to {target}")
        open_pdfs(settings, [target])

    all_btn.configure(command=toggle_all)
    export_btn.configure(command=do_export)

    def switch_side(_event=None):
        if paging["mode"] == "pairs" and counts["back"]:
            show("back" if view["side"] == "front" else "front")
        return "break"

    def select_all_key(_event=None):
        if root.focus_get() in (entry, bleed_entry, spacing_entry):
            return None                          # Ctrl+A in a text box selects its text
        toggle_all()
        return "break"
    root.bind("<Control-Tab>", switch_side)
    root.bind("<Control-a>", select_all_key)
    root.bind("<Control-e>", lambda e: (open_edit_cards(), "break")[1] if editable else None)

    theme_btn.configure(text="Theme ▾", command=lambda: theme_menu(theme_btn, lambda: root.after(10, check_inbox)))
    Tooltip(theme_btn, "Light, dark, or Adventurer (an RPG look). Also in Fleem's menu > Appearance.")
    entry_var.trace_add("write", typed)
    print_btn.configure(command=do_print)
    root.bind("<Return>", do_print)
    root.bind("<Escape>", cancel)

    # resizing the window scales the previews and the text with it
    def rescale(force=False):
        """Previews as big as the window's height allows (one full row visible), times the zoom; text
        grows a little with a big window."""
        resize["job"] = None
        if not root.winfo_exists():
            return
        avail_w, avail_h = canvas.winfo_width(), canvas.winfo_height()
        if avail_w < 50 or avail_h < 50:
            return
        # the biggest size that shows every page at once (in 1, 2, 3... rows); then the zoom
        n = max(1, counts[view["side"]])
        best = 0
        for rows in range(1, n + 1):
            per_row = -(-n // rows)
            w_fit = min((avail_h / rows - 46) / aspect, (avail_w - 8) / per_row - 18)
            best = max(best, w_fit)
            if (avail_h / rows - 46) / aspect < 150:
                break
        new_w = int(max(90, min(max(best, 150) * zoom["f"], avail_w - 40, (avail_h - 46) / aspect * 1.6, 1600)))
        cols = max(1, (avail_w - 8) // (new_w + 18))
        while new_w > 90 and -(-n // cols) * (new_w * aspect + 60) > 28000:   # windows can't be taller
            new_w = int(new_w * 0.9)
            cols = max(1, (avail_w - 8) // (new_w + 18))
        factor = max(1.0, min(1.2, (root.winfo_width() / 1100) ** 0.5))
        for key, f in fonts.items():
            if key != "big":
                f.configure(size=max(7, round(BASE_FONT_SIZES[key] * factor)))
        if not force and abs(new_w - size["w"]) < 4 and cols == grid_cols["n"]:
            return
        size.update(w=new_w, h=int(new_w * aspect))
        grid_cols["n"] = cols
        photos.clear()                       # previews are redrawn at the new size
        show(view["side"], keep_scroll=True)

    def set_zoom(f):
        zoom["f"] = max(0.3, min(3.0, f))
        rescale(force=True)

    def on_configure(event):
        key = (canvas.winfo_width(), canvas.winfo_height())
        if resize["width"] == key:
            return
        resize["width"] = key
        if resize["job"]:
            root.after_cancel(resize["job"])
        resize["job"] = root.after(200, rescale)
    canvas.bind("<Configure>", on_configure)

    # --- full-page viewer
    def render_page(side, n, width):
        """One page as a Tk image `width` px wide."""
        if editable:
            return to_photo(renderer.page(deck_pages(cards, lay["layout"])[n - 1], side == "back", lay["layout"], width,
                                          page_header(side, n, live=True)))
        return as_photo(pdf_page_image(page_index(side, n), width))

    def open_viewer(n=1):
        if not (editable or pdf is not None) or not counts[view["side"]]:
            return
        at = {"side": view["side"], "n": max(1, min(n, counts[view["side"]]))}
        win = tk.Toplevel(root)
        win.withdraw()
        win.transient(root)
        T(win, bg="bg")
        bar_ = T(tk.Frame(win, padx=14, pady=8), bg="surface")
        bar_.pack(fill="x")
        prev_b = ttk.Button(bar_, text="‹ Prev", command=lambda: go(-1))
        prev_b.pack(side="left")
        where = T(tk.Label(bar_, font=fonts["label"], width=16), bg="surface", fg="text")
        where.pack(side="left", padx=8)
        next_b = ttk.Button(bar_, text="Next ›", command=lambda: go(1))
        next_b.pack(side="left")
        side_b = ttk.Button(bar_, command=lambda: flip())
        side_b.pack(side="left", padx=(16, 0))
        pick_var = tk.BooleanVar()
        ttk.Checkbutton(bar_, text="Picked for printing", variable=pick_var, command=lambda: toggle_pick()).pack(
            side="left", padx=(16, 0))
        ttk.Button(bar_, text="Close", command=win.destroy).pack(side="right")
        Tooltip(next_b, "Arrow keys or the mouse wheel go through the pages; Space picks; B shows the other side; "
                        "Esc closes.")
        pic = T(tk.Label(win, bd=0), bg="bg")
        pic.pack(fill="both", expand=True, padx=14, pady=10)
        state_ = {"img": None, "h": 0}

        def draw_page():
            side, n = at["side"], at["n"]
            win.title(f"{APP_NAME} — {noun(side)} {n} of {counts[side]}")
            where.configure(text=f"{noun(side).capitalize()} {n} of {counts[side]}")
            side_b.configure(text="Show the back" if side == "front" else "Show the front",
                             state="normal" if paging["mode"] == "pairs" and counts["back"] else "disabled")
            pick_var.set(n in chosen[side])
            h = pic.winfo_height() - 4
            if h < 200:
                h = state_["h"]
            width = int(min(h / aspect, max(300, pic.winfo_width() - 4)))
            try:
                state_["img"] = render_page(side, n, width)
                pic.configure(image=state_["img"], text="")
            except Exception as e:
                pic.configure(image="", text=f"Couldn't show this page: {e}")

        def go(step):
            at["n"] = max(1, min(counts[at["side"]], at["n"] + step))
            draw_page()

        def flip():
            if paging["mode"] == "pairs" and counts["back"]:
                at["side"] = "back" if at["side"] == "front" else "front"
                at["n"] = min(at["n"], counts[at["side"]])
                draw_page()

        def toggle_pick(_e=None):
            if _e is not None:
                pick_var.set(not pick_var.get())
            (chosen[at["side"]].add if pick_var.get() else chosen[at["side"]].discard)(at["n"])
            if at["side"] == view["side"]:
                update(rewrite=True)
            else:
                update()
            if layout_header(lay["layout"]) is not None and layout_numbering(lay["layout"]) == "selected":
                draw_page()                     # its page number changes with what's picked
        win.bind("<Left>", lambda e: go(-1))
        win.bind("<Right>", lambda e: go(1))
        win.bind("<space>", toggle_pick)
        win.bind("<b>", lambda e: flip())
        win.bind("<Escape>", lambda e: win.destroy())
        win.bind("<MouseWheel>", lambda e: (go(-1 if e.delta > 0 else 1), "break")[1])
        redraw = {"job": None, "size": None}

        def on_resize(_e):
            size_now = (pic.winfo_width(), pic.winfo_height())
            if size_now == redraw["size"]:
                return
            redraw["size"] = size_now
            if redraw["job"]:
                win.after_cancel(redraw["job"])
            redraw["job"] = win.after(150, draw_page)
        pic.bind("<Configure>", on_resize)
        sh = root.winfo_screenheight()
        hgt = int(sh * 0.86)
        state_["h"] = hgt - 70
        wid = int((hgt - 70) / aspect) + 60
        win.geometry(f"{max(wid, 760)}x{hgt}+{max(0, root.winfo_rootx() + 40)}+{max(0, (sh - hgt) // 3)}")
        apply_theme(win)
        win.deiconify()
        win.focus_force()
        win.after(60, draw_page)

    # --- adding cards: drag images onto the window, or Add cards…
    work = restore.get("work") or WORK_DIR / f"added-{time.time_ns()}"   # images of added cards (until the window closes)

    def add_cards(files):
        projects = [Path(f) for f in files if Path(f).suffix.lower() in (".xml", ".pdf")]
        images = [Path(f) for f in files if Path(f).suffix.lower() in IMAGE_TYPES]
        if not images and not projects:
            layout_note("!Only card images, order XMLs and Fleem PDFs can be added.")
            return
        if not editable:
            layout_note("!Cards can only be added to PDFs made by Fleem.")
            return
        if projects:
            combine_projects(projects)
        if images:
            add_images(images)

    def combine_projects(files):
        """Dropped order XMLs / Fleem PDFs: say what will be added, ask, then add their cards to the deck."""
        ready, problems = [], []
        for f in files:
            try:
                if f.suffix.lower() == ".pdf":
                    info = read_any_pdf(f)          # (cards from PDFs made elsewhere: layout only, for now)
                    if info is None:
                        problems.append(f"{f.name}: its cards aren't separate images, so they can't be taken out")
                        continue
                    n, mdfcs = len(info["copies"]), len({_src_id(c["front"]) for c in info["copies"] if c["back"]})
                    ready.append({"file": f, "cards": info["copies"], "foreign": bool(info.get("foreign"))})
                else:
                    groups = parse_order(f.read_bytes())
                    n, mdfcs = sum(g["count"] for g in groups), sum(1 for g in groups if g.get("back_id"))
                    ready.append({"file": f, "groups": groups})
            except Exception as e:
                problems.append(f"{f.name}: {e}")
                continue
            line = f"  •  {f.name} — {n} card{'s' * (n != 1)}" + (f" ({mdfcs} MDFC{'s' * (mdfcs != 1)})" if mdfcs else "")
            theirs = profile_of(f)
            if pdf_path and f.resolve() == Path(pdf_path).resolve():
                line += "\n      (this is the file that's already open: its cards will be in it twice)"
            elif theirs and profile and theirs != profile:
                line += f"\n      (made for {theirs}: they'll print with this project's {profile} settings)"
            ready[-1]["line"] = line
        if not ready:
            messagebox.showinfo(APP_NAME, "Nothing to combine:\n\n" + "\n".join(problems), parent=root)
            return
        text = ("Combine these with the open project?\n\n" + "\n".join(r["line"] for r in ready) +
                "\n\nThey'll be laid out like the rest (" + describe_layout(lay["layout"]) + "): "
                + ("double-faced cards with the other MDFCs, everything else in alphabetical order"
                   if mdfcs_first() else "each card in its alphabetical spot")
                + f". Saved into {Path(out_path).name} when you press Print.")
        if any("groups" in r for r in ready):
            text += "\n\nThe XML's card images will be downloaded first."
        if problems:
            text += "\n\nCan't be added:\n" + "\n".join("  •  " + p for p in problems)
        if not messagebox.askyesno(f"{APP_NAME} — combine?", text, parent=root):
            return
        incoming = []
        for r in ready:
            if "groups" in r:
                got = download_into_project(r)
                if got is None:
                    continue
                incoming.append((r["file"], got))
            elif r.get("foreign"):              # cut the card images out now
                got, err = run_with_progress(f"Taking the cards out of {r['file'].name}…",
                                             lambda r=r: read_foreign(r["file"], work / f"combine-{time.time_ns()}"),
                                             parent=root)
                if not got:
                    messagebox.showerror(APP_NAME, f"Couldn't take the cards out of {r['file'].name}"
                                         + (f":\n{err}" if err else "."), parent=root)
                    continue
                incoming.append((r["file"], got["copies"]))
            else:
                incoming.append((r["file"], r["cards"]))
        if not incoming:
            return
        for f, deck_cards in incoming:
            cards[:] = insert_added_cards(cards, group_cards(deck_cards))
            log.info("Combined %s into %s (%d cards)", f.name, title, len(deck_cards))
        lay["added"].append("combine")           # so Cancel asks before throwing it away
        lay["cards_changed"] = True
        recount()
        chosen["front"].clear()                  # cards moved: pick again
        chosen["back"].clear()
        photos.clear()
        show(view["side"], keep_scroll=True)
        layout_note(f"Combined {', '.join(f.name for f, _ in incoming)} ({len(cards)} cards in total). Pick the "
                    "pages to print again - the PDF is saved when you press Print.")
        root.after(300, start_name_reading)      # cards that came without names: read them

    def open_find_cards():
        """The MPCFill search window; the cards picked there are downloaded and added like an XML's."""
        module = mpcfill_module()
        if module is None:
            return
        try:
            module.open_search(root, sys.modules[__name__], add_found_cards)
        except Exception as e:
            log.error("The MPCFill search failed: %s", e)
            messagebox.showerror(APP_NAME, f"The MPCFill search didn't open: {e}", parent=root)

    def add_found_cards(groups):
        total = sum(g["count"] for g in groups)
        got = download_into_project({"groups": groups, "file": Path(f"{total} card{'s' * (total != 1)} from MPCFill")})
        if not got:
            return
        cards[:] = insert_added_cards(cards, group_cards(got))
        lay["added"].append("mpcfill")
        log.info("Added %d cards from MPCFill to %s", len(got), title)
        cards_changed_now(f"Added {len(got)} card{'s' * (len(got) != 1)} from MPCFill ({len(cards)} cards in total).")

    def download_into_project(r):
        """Download a dropped XML's card images (small progress window). Its card list, or None."""
        images = work / f"combine-{time.time_ns()}"
        images.mkdir(parents=True, exist_ok=True)
        outcome = {}

        def job():
            try:
                download_images(r["groups"], images)
                outcome["cards"] = deck_from_groups(r["groups"], images)
                outcome["problems"] = download_problems(r["groups"])
            except BaseException as e:
                outcome["error"] = str(e) or e.__class__.__name__

        top = tk.Toplevel(root)
        top.title(APP_NAME)
        top.resizable(False, False)
        top.transient(root)
        tk.Label(top, text=f"Getting the cards for {r['file'].name}…\nCard images are downloading; big orders "
                           "take a few minutes.", justify="left", padx=18, pady=14).pack()
        bar = ttk.Progressbar(top, mode="indeterminate", length=320)
        bar.pack(padx=18, pady=(0, 16))
        bar.start(12)
        top.protocol("WM_DELETE_WINDOW", lambda: None)
        thread = threading.Thread(target=job, daemon=True)
        thread.start()

        def wait():
            if thread.is_alive():
                top.after(200, wait)
            else:
                bar.stop()
                top.destroy()
        top.after(200, wait)
        top.grab_set()
        root.wait_window(top)
        if "error" in outcome:
            log.error("%s: %s", r["file"].name, outcome["error"])
            messagebox.showerror(APP_NAME, f"Couldn't get the cards for {r['file'].name}:\n{outcome['error']}",
                                 parent=root)
            return None
        if outcome.get("problems"):
            messagebox.showwarning(APP_NAME, f"{r['file'].name}:\n\n{outcome['problems']}", parent=root)
        return outcome["cards"]

    def add_images(images):
        work.mkdir(parents=True, exist_ok=True)
        added = []
        for image in images:
            card = ask_card(root, image)
            if card is None:
                continue
            n = len(lay["added"]) + len(added)
            try:
                front = work / f"{n}-front.jpg"
                prepare_card_image(card["front"], front)
                back = None
                if card["back"]:
                    back = work / f"{n}-back.jpg"
                    prepare_card_image(card["back"], back)
            except Exception as e:
                show_error(f"Couldn't use {image.name}: {e}")
                continue
            added.append(dict(card, front=front, back=back))
        if not added:
            return
        cards[:] = insert_added_cards(cards, added)
        lay["added"] += added
        lay["cards_changed"] = True
        recount()
        chosen["front"].clear()                 # cards moved: pick again
        chosen["back"].clear()
        photos.clear()
        show(view["side"], keep_scroll=True)
        names = ", ".join(f"{c['name']}" + (f" x{c['copies']}" if c["copies"] > 1 else "") +
                          (" (MDFC)" if c["back"] else "") for c in added)
        log.info("Added to %s: %s", title, names)
        layout_note(f"Added {names}. Pick the pages to print again - the PDF is saved when you press Print.")

    def pick_cards():
        from tkinter import filedialog
        files = filedialog.askopenfilenames(parent=root, title="Pick card images, order XMLs or Fleem PDFs to add",
                                            filetypes=[("Cards, XMLs or PDFs",
                                                        " ".join("*" + t for t in (*IMAGE_TYPES, ".xml", ".pdf"))),
                                                       ("Images", " ".join("*" + t for t in IMAGE_TYPES)),
                                                       ("Order XMLs or Fleem PDFs", "*.xml *.pdf"),
                                                       ("All files", "*.*")])
        if files:
            add_cards(list(files))

    def open_edit_cards(focus=None):
        if not editable:
            return
        new = edit_cards(root, cards, renderer, work, lay.setdefault("parked", []),
                         per_sheet=len(card_slots(lay["layout"])), focus=focus)
        if new is None:
            return
        cards[:], lay["parked"] = new
        cards_changed_now(f"Cards edited ({len(cards)} in total).")

    # --- undo / redo: snapshots of the cards and the layout after every change
    history = dict(restore.get("history") or {"undo": [], "redo": [], "now": None}, busy=False)

    def snapshot():
        return {"cards": [dict(c) for c in cards], "parked": [dict(c) for c in lay.get("parked", [])],
                "layout": lay["layout"]}

    def fingerprint(state):
        rows = lambda cs: tuple((_src_id(c["front"]), _src_id(c["back"]), c.get("name") or "", c.get("sort"),
                                 bool(c.get("missing"))) for c in cs)
        return rows(state["cards"]), rows(state["parked"]), state["layout"]

    def checkpoint(absorb=False):
        """Call after anything that may have changed the project: a real change becomes an undo step
        (`absorb`: take it in silently, e.g. card names read off the cards)."""
        if not editable or history["busy"]:
            return
        now = snapshot()
        if history["now"] is not None and fingerprint(now) == fingerprint(history["now"]):
            return
        if history["now"] is not None and not absorb:
            history["undo"].append(history["now"])
            del history["undo"][:-100]                 # the last 100 changes
            history["redo"].clear()
        history["now"] = now
        show_undo_buttons()

    def show_undo_buttons():
        undo_btn.pack_forget()
        redo_btn.pack_forget()
        if history["undo"]:
            undo_btn.pack(side="left", padx=(0, 4))
        if history["redo"]:
            redo_btn.pack(side="left", padx=(0, 4))

    def restore_state(state):
        """Put the cards and the layout back as they were in `state` (boxes included)."""
        nonlocal aspect
        history["busy"] = True
        try:
            if lay["job"]:
                root.after_cancel(lay["job"])
                lay["job"] = None
            cards[:] = [dict(c) for c in state["cards"]]
            lay["parked"] = [dict(c) for c in state["parked"]]
            layout = state["layout"]
            syncing["on"] = True
            try:
                bleed_var.set(f"{layout[0]:.2f}")
                spacing_var.set(f"{layout[1]:.2f}")
            finally:
                syncing["on"] = False
            paper_var.set(layout_paper(layout))
            orientation_var.set(layout_orientation(layout))
            header_on.set(layout_header(layout) is not None)
            if layout_header(layout):
                header_var.set(layout_header(layout))
            number_selected.set(layout_numbering(layout) == "selected")
            card_size["value"] = layout_card(layout)
            card_var.set(card_box_label(layout_card(layout)))
            marks_var.set(CUT_MARKS[layout_marks(layout)])
            show_marks_rows()
            header_changed()
            if lay["job"]:                               # (the boxes already match)
                root.after_cancel(lay["job"])
                lay["job"] = None
            before = dict(counts)
            lay["layout"] = layout
            w, h = page_size(layout)
            aspect = h / w
            recount()
            if counts != before:
                chosen["front"].clear()
                chosen["back"].clear()
            lay["cards_changed"] = True
            lay["added"].append("edit")
            photos.clear()
            rescale(force=True)
            show(view["side"], keep_scroll=True)
        finally:
            history["busy"] = False
        history["now"] = {"cards": [dict(c) for c in state["cards"]],
                          "parked": [dict(c) for c in state["parked"]], "layout": state["layout"]}
        show_undo_buttons()

    def undo(_event=None):
        if not history["undo"] or root.grab_current() not in (None, root):
            return "break"
        history["redo"].append(history["now"])
        restore_state(history["undo"].pop())
        layout_note("Undone." + ("" if (chosen["front"] or chosen["back"]) else " Pick the pages to print again."))
        return "break"

    def redo(_event=None):
        if not history["redo"] or root.grab_current() not in (None, root):
            return "break"
        history["undo"].append(history["now"])
        restore_state(history["redo"].pop())
        layout_note("Redone.")
        return "break"
    if history["undo"] or history["redo"]:       # (reloaded for a new theme)
        root.after(0, show_undo_buttons)
    if editable:
        for seq in ("<Control-z>", "<Control-Z>"):
            root.bind(seq, undo)
        for seq in ("<Control-y>", "<Control-Y>", "<Control-Shift-z>", "<Control-Shift-Z>"):
            root.bind(seq, redo)

    def cards_changed_now(message):
        """After the deck changed: redraw; picks stay unless the number of pages changed."""
        lay["cards_changed"] = True
        lay["added"].append("edit")              # so Cancel asks before throwing the changes away
        before = dict(counts)
        recount()
        if counts != before:                     # pages added/removed: pick again
            chosen["front"].clear()
            chosen["back"].clear()
            message += " Pick the pages to print again."
        photos.clear()
        show(view["side"], keep_scroll=True)
        layout_note(message + " Saved to the PDF when you press Print.")
        checkpoint()

    # --- cards without names (PDFs made elsewhere, older Fleem PDFs): read the names off the cards
    def start_name_reading():
        import queue
        if not editable or not root.winfo_exists():
            return
        todo = {}
        for c in cards + lay.get("parked", []):
            if not c.get("name"):
                key = _src_id(c["front"])
                if key not in todo:
                    kind, ref = c["front"]
                    todo[key] = ("file", ref) if kind == "file" else ("data", ref.get_object()._data)
        if not todo:
            return
        results = queue.Queue()

        def work():
            catalog = card_name_catalog()
            for key, (kind, value) in todo.items():
                try:
                    data = Path(value).read_bytes() if kind == "file" else value
                    name = read_card_name(data, catalog)
                except Exception as e:
                    log.info("Couldn't read a card name: %s", e)
                    name = None
                results.put((key, name))
                if name is None and OCR_BACKEND is windows_ocr_lines and _catalog_cache.get("ocr_warned"):
                    break                          # no text recognition on this PC: stop trying
            results.put(None)
        threading.Thread(target=work, daemon=True).start()
        found, done = {}, {"n": 0, "finished": False}
        set_status(f"Reading card names from the cards… 0 of {len(todo)}")

        def poll():
            if not root.winfo_exists():
                return
            busy = root.grab_current() not in (None, root)     # a pop-up is open: wait for it
            while not busy:
                try:
                    item = results.get_nowait()
                except queue.Empty:
                    break
                if item is None:
                    done["finished"] = True
                    break
                done["n"] += 1
                if item[1]:
                    found[item[0]] = item[1]
            if not done["finished"] or busy:
                if not busy:
                    set_status(f"Reading card names from the cards… {done['n']} of {len(todo)}")
                root.after(250, poll)
                return
            apply_names(found, len(todo))
        root.after(250, poll)

    def apply_names(found, asked):
        named = 0
        just_named = []
        for c in cards + lay.get("parked", []):
            name = found.get(_src_id(c["front"]))
            if name and not c.get("name"):
                c["name"], c["sort"] = name, sort_key({"name": name})[0]
                named += 1
                just_named.append(c)
        if not found:
            layout_note()
            if asked and _catalog_cache.get("ocr_warned"):
                layout_note("Card names couldn't be read on this PC (Windows text recognition isn't available).")
            return
        message = f"Read {len(found)} of {asked} card names from the cards."
        if foreign is not None:              # a PDF made elsewhere: Fleem's card order
            cards[:] = sorted_cards(cards)
            message += (" Sorted: double-faced cards first, then A-Z." if mdfcs_first() else " Sorted A-Z.")
            if chosen["front"] or chosen["back"]:
                chosen["front"].clear()
                chosen["back"].clear()
                message += " Pick the pages again."
        elif just_named:                     # (e.g. cards combined in from a PDF made elsewhere): A-Z spots
            moving = [c for c in cards if any(c is j for j in just_named)]
            rest = [c for c in cards if not any(c is j for j in just_named)]
            cards[:] = insert_added_cards(rest, group_cards(moving))
            if chosen["front"] or chosen["back"]:
                chosen["front"].clear()
                chosen["back"].clear()
                message += " Pick the pages again."
        if asked - len(found):
            message += " Name the rest in Edit cards."
        lay["cards_changed"] = True            # (the names are saved into the PDF when you print)
        log.info("%s: %s", title, message)
        photos.clear()
        history["busy"] = True                 # names read off the cards aren't an undo step
        try:
            show(view["side"], keep_scroll=True)
        finally:
            history["busy"] = False
        checkpoint(absorb=True)
        layout_note(message)
    root.after(600, start_name_reading)

    # --- right-click a card on a page: highlight it, and a menu of what to do with it
    def card_at(side, n, x_root, y_root):
        """(index in the deck, its box on the preview in px) for the card under the mouse, or None."""
        if not editable or n not in cells:
            return None
        pic = cells[n][0].winfo_children()[0]
        px, py = x_root - pic.winfo_rootx(), y_root - pic.winfo_rooty()
        PAGE_W, PAGE_H = page_size(lay["layout"])
        scale = size["w"] / PAGE_W
        bleed = lay["layout"][0] * PT_PER_MM
        card_w, card_h = card_dims(lay["layout"])
        cell_w, cell_h = card_w + 2 * bleed, card_h + 2 * bleed
        slots = card_slots(lay["layout"])
        first = (n - 1) * len(slots)
        for k, (x, y) in enumerate(slots):
            if first + k >= len(cards):
                break
            if side == "back":
                x = PAGE_W - x - cell_w
            x0, y0 = x * scale, (PAGE_H - y - cell_h) * scale
            x1, y1 = x0 + cell_w * scale, y0 + cell_h * scale
            if x0 <= px <= x1 and y0 <= py <= y1:
                return first + k, (round(x0), round(y0), round(x1), round(y1))
        return None

    marks = []

    def clear_mark():
        for m in marks:
            try:
                m.destroy()
            except Exception:
                pass
        marks.clear()

    def mark_card(n, box):
        """Outline the right-clicked card on its page preview."""
        clear_mark()
        pic = cells[n][0].winfo_children()[0]
        for inset, th, tone in ((0, 2, "accent_text"), (2, 3, "accent")):   # white outside, blue inside: shows on any card
            x0, y0, x1, y1 = box[0] + inset, box[1] + inset, box[2] - inset, box[3] - inset
            for (x, y, w, h) in ((x0, y0, x1 - x0, th), (x0, y1 - th, x1 - x0, th),
                                 (x0, y0, th, y1 - y0), (x1 - th, y0, th, y1 - y0)):
                bar = T(tk.Frame(pic, bd=0), bg=tone)
                bar.place(x=x, y=y, width=max(1, w), height=max(1, h))
                marks.append(bar)

    def card_menu(n, event):
        side = view["side"]
        menu = tk.Menu(root, tearoff=0)
        P = palette()
        menu.configure(bg=P["surface"], fg=P["text"], activebackground=P["accent"], activeforeground=P["accent_text"],
                       disabledforeground=P["muted"], bd=1, relief="solid", font=fonts["body"])
        hit = card_at(side, n, event.x_root, event.y_root)
        if hit is not None:
            index, box = hit
            card = cards[index]
            key = card_key(card)
            same = [i for i, c in enumerate(cards) if card_key(c) == key]
            name = Path(card["name"]).stem if card.get("name") else f"Card {index + 1}"
            mark_card(n, box)
            menu.add_command(label=f"{name}  ·  {len(same)} cop{'y' if len(same) == 1 else 'ies'}", state="disabled")
            menu.add_separator()
            menu.add_command(label="Add a copy", command=lambda: add_copy(index))
            if len(same) > 1:
                menu.add_command(label="Remove this copy", command=lambda: remove_copy(index))
            menu.add_command(label="Set to 0 copies", command=lambda: set_zero(index))
            menu.add_separator()
            menu.add_command(label="Make bleed (the picture has none)", command=lambda: make_bleed(index))
            menu.add_command(label="Edit this card…", command=lambda: open_edit_cards(key))
        menu.add_command(label="View this page", command=lambda: open_viewer(n))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

        def wait_closed():                   # the outline goes when the menu does
            try:
                if menu.winfo_exists() and menu.winfo_ismapped():
                    root.after(120, wait_closed)
                    return
            except Exception:
                pass
            clear_mark()
        root.after(120, wait_closed)

    # --- drag a card on a page preview to another spot (between other cards, on any page)
    drag = {}

    def press(n, event):
        """Mouse down on a page. On a card (in a project Fleem lays out) it may become a drag, so picking
        the page waits for the mouse to come up; anywhere else it picks the page straight away."""
        drag.clear()
        hit = card_at(view["side"], n, event.x_root, event.y_root) if editable else None
        if hit is None:
            click(n, event)
            return
        drag.update(n=n, index=hit[0], start=(event.x_root, event.y_root), state=event.state, active=False)

    def double_click(n):
        drag.clear()
        drag["skip"] = True                  # (the button coming up again isn't a click)
        open_viewer(n)

    def page_under(x_root, y_root):
        """The page tile number under the mouse, or None."""
        w = root.winfo_containing(x_root, y_root)
        while w is not None:
            for num, (frame, caption, badge) in cells.items():
                if w in (frame, caption, badge) or w is frame.winfo_children()[0]:
                    return num
            w = w.master if w is not root else None
        return None

    def drop_target(x_root, y_root):
        """(insert before this deck index, page, bar box on that page) for the mouse position, or None."""
        n = page_under(x_root, y_root)
        if n is None:
            return None
        side = view["side"]
        hit = card_at(side, n, x_root, y_root)
        if hit is None:                      # past the last card on the last page: the very end
            if n != counts[side] or not cards:
                return None
            last = card_at(side, n, *pic_point(n, len(cards) - 1))
            if last is None:
                return None
            box = last[1]
            edge = box[2] if side == "front" else box[0]
            return len(cards), n, (edge - 3, box[1], edge + 3, box[3])
        index, box = hit
        pic = cells[n][0].winfo_children()[0]
        right_half = x_root - pic.winfo_rootx() > (box[0] + box[2]) / 2
        after = right_half != (side == "back")   # backs are mirrored: the next card is to the left
        bar = (box[2] - 3, box[1], box[2] + 3, box[3]) if right_half else (box[0] - 3, box[1], box[0] + 3, box[3])
        return index + (1 if after else 0), n, bar

    def pic_point(n, index):
        """A screen point in the middle of card `index` on page tile n (for finding its box)."""
        pic = cells[n][0].winfo_children()[0]
        PAGE_W, PAGE_H = page_size(lay["layout"])
        scale = size["w"] / PAGE_W
        bleed = lay["layout"][0] * PT_PER_MM
        card_w, card_h = card_dims(lay["layout"])
        slots = card_slots(lay["layout"])
        x, y = slots[index % len(slots)]
        if view["side"] == "back":
            x = PAGE_W - x - (card_w + 2 * bleed)
        return (pic.winfo_rootx() + round((x + card_w / 2 + bleed) * scale),
                pic.winfo_rooty() + round((PAGE_H - y - card_h / 2 - bleed) * scale))

    def show_drop_bar(target):
        clear_mark()
        if target is None:
            return
        _, n, (x0, y0, x1, y1) = target
        pic = cells[n][0].winfo_children()[0]
        for inset, tone in ((0, "accent_text"), (1, "accent")):
            bar = T(tk.Frame(pic, bd=0), bg=tone)
            bar.place(x=x0 + inset, y=y0 + inset, width=max(2, x1 - x0 - 2 * inset), height=max(2, y1 - y0 - 2 * inset))
            marks.append(bar)

    def start_drag(event):
        drag["active"] = True
        card = cards[drag["index"]]
        ghost = tk.Toplevel(root)
        ghost.overrideredirect(True)
        try:
            ghost.attributes("-topmost", True)
            ghost.attributes("-alpha", 0.85)
        except tk.TclError:
            pass
        width = max(40, min(110, round(card_dims(lay["layout"])[0] * size["w"] / page_size(lay["layout"])[0])))
        try:
            drag["photo"] = to_photo(card_face(renderer, card["front"], width))
            tk.Label(ghost, image=drag["photo"], bd=0, bg="#000000").pack()
        except Exception:
            tk.Label(ghost, text=card_label(drag["index"]), padx=8, pady=4).pack()
        drag["ghost"] = ghost
        root.configure(cursor="fleur")
        mark_card(drag["n"], card_at(view["side"], drag["n"], *drag["start"])[1])   # where it comes from
        drag["from_marks"] = marks[:]
        marks.clear()

    def drag_motion(n, event):
        if not drag or "index" not in drag:
            return
        if not drag["active"]:
            if abs(event.x_root - drag["start"][0]) + abs(event.y_root - drag["start"][1]) < 8:
                return
            start_drag(event)
        drag["ghost"].geometry(f"+{event.x_root + 14}+{event.y_root + 10}")
        # scroll the pages when the mouse is near the top or bottom
        top, bottom = canvas.winfo_rooty(), canvas.winfo_rooty() + canvas.winfo_height()
        if event.y_root < top + 30:
            canvas.yview_scroll(-1, "units")
        elif event.y_root > bottom - 30:
            canvas.yview_scroll(1, "units")
        drag["target"] = drop_target(event.x_root, event.y_root)
        show_drop_bar(drag["target"])

    def end_drag():
        for w in drag.get("from_marks", []):
            try:
                w.destroy()
            except Exception:
                pass
        clear_mark()
        if drag.get("ghost") is not None:
            drag["ghost"].destroy()
        root.configure(cursor="")

    def release(n, event):
        if drag.pop("skip", False):
            return
        if not drag or "index" not in drag:
            return
        info = dict(drag)
        drag.clear()
        if not info["active"]:               # it was a click after all: pick the page
            click(info["n"], type("Click", (), {"state": info["state"]})())
            return
        drag.update(info)
        end_drag()
        drag.clear()
        target = drop_target(event.x_root, event.y_root)
        if target is None:
            return
        move_card(info["index"], target[0])

    def move_card(src, before):
        """Move card `src` so it's printed just before the card now at `before` (len = at the end)."""
        if before in (src, src + 1):
            return                           # dropped where it already is
        name = card_label(src)
        card = cards.pop(src)
        dest = before - 1 if before > src else before
        cards.insert(dest, card)
        lay["moved"] = True
        per_page = len(card_slots(lay["layout"]))
        cards_changed_now(f"Moved {name} to page {dest // per_page + 1}, spot {dest % per_page + 1}. Changing the "
                          "card order in Advanced settings puts moved cards back.")

    def make_bleed(index):
        """The card's picture is just the card (no bleed in it, even if it's MPC's shape): make bleed for
        it by continuing its border (or mirroring the picture), for every copy of the card."""
        key = card_key(cards[index])
        name = card_label(index)
        changed = {}
        try:
            for field in ("front", "back"):
                src = cards[index][field]
                if not src:
                    continue
                kind, ref = src
                work.mkdir(parents=True, exist_ok=True)
                raw = Path(ref) if kind == "file" else None
                if raw is None:                  # an image inside the PDF: write it out first
                    raw = work / f"bleed-src-{time.time_ns()}.img"
                    raw.write_bytes(ref.get_object()._data)
                out = free_path(work / f"{safe_name(name)} ({field}, bleed made).jpg")
                how = prepare_card_image(raw, out, as_card=True)
                changed[field] = ("file", out)
                log.info("%s (%s): %s", name, field, how)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Couldn't make bleed for {name}: {e}", parent=root)
            return
        for c in cards:
            if card_key(c) == key:
                c.update(changed)
        cards_changed_now(f"Made bleed for {name}.")

    def card_label(index):
        c = cards[index]
        return Path(c["name"]).stem if c.get("name") else f"card {index + 1}"

    def add_copy(index):
        name = card_label(index)
        cards.insert(index + 1, dict(cards[index]))
        cards_changed_now(f"Added a copy of {name} ({len(cards)} cards).")

    def remove_copy(index):
        name = card_label(index)
        del cards[index]
        cards_changed_now(f"Removed a copy of {name} ({len(cards)} cards).")

    def set_zero(index):
        key = card_key(cards[index])
        if all(card_key(c) == key for c in cards):
            messagebox.showinfo(APP_NAME, "That's the only card - keep at least one to print.", parent=root)
            return
        name = card_label(index)
        group = group_cards([c for c in cards if card_key(c) == key])[0]
        parked = lay.setdefault("parked", [])
        parked.append({k: v for k, v in group.items() if k != "copies"})
        cards[:] = [c for c in cards if card_key(c) != key]
        cards_changed_now(f"{name} set to 0 copies - bring it back in Edit cards ({len(cards)} cards).")
    if not editable:
        add_btn.state(["disabled"])
        edit_btn.state(["disabled"])
    enable_file_drop(root, add_cards)
    inbox = {"files": list(restore.get("inbox", [])), "editable": editable}   # files opened with Fleem meanwhile
    with _OPEN_PROJECTS_LOCK:
        _OPEN_PROJECTS.append(inbox)

    def reload_for_theme():
        """The theme changed: close and reopen the window in it, keeping everything as it is."""
        if lay.get("job"):
            root.after_cancel(lay["job"])
            lay["job"] = None
            if editable:
                apply_layout()
        with _OPEN_PROJECTS_LOCK:
            waiting, inbox["files"] = inbox["files"][:], []
        result["choice"] = ("retheme", {
            "cards": cards, "lay": lay, "mode": paging["mode"], "renderer": renderer, "work": work,
            "chosen": {k: set(v) for k, v in chosen.items()}, "side": view["side"], "zoom": zoom["f"],
            "geometry": root.geometry(), "advanced": adv.winfo_ismapped(), "printer": printer_var.get(),
            "typed": entry_var.get(), "inbox": waiting, "history": history})
        log.info("Reloading the print window for the %s theme", THEME_LABELS.get(theme_name(), theme_name()))
        close()

    def check_inbox():
        if not root.winfo_exists():
            return
        if theme_name() != made_theme and root.grab_current() is None:
            reload_for_theme()
            return
        if inbox["files"] and root.grab_current() is None:     # (not while Edit cards or a pop-up is open)
            with _OPEN_PROJECTS_LOCK:
                files, inbox["files"] = inbox["files"][:], []
            root.deiconify()
            root.lift()
            root.focus_force()
            add_cards(files)
        root.after(300, check_inbox)

    def leave_inbox(event):
        if event.widget is root:
            with _OPEN_PROJECTS_LOCK:
                if inbox in _OPEN_PROJECTS:
                    _OPEN_PROJECTS.remove(inbox)
            if inbox["files"]:                              # arrived just as it closed: open them normally
                threading.Thread(target=lambda f=inbox["files"][:]: message_running_fleem({"print": f}),
                                 daemon=True).start()
    root.bind("<Destroy>", leave_inbox, add="+")
    root.after(300, check_inbox)

    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{min(1500, sw - 80)}x{min(980, sh - 90)}+{max(0, (sw - min(1500, sw - 80)) // 2)}+20")
    root.minsize(760, 480)
    if restore:                              # reloaded for a new theme: as it was
        root.geometry(restore.get("geometry") or root.geometry())
        if restore.get("printer"):
            printer_var.set(restore["printer"])
        if zoom_from:
            zoom["f"] = zoom_from
        if restore.get("advanced"):
            root.after(10, lambda: toggle_advanced(True))
    paint(root)                              # older parts (Advanced settings) take the theme's colours
    show(view["side"])
    if restore.get("typed"):
        syncing["on"] = True
        entry_var.set(restore["typed"])
        syncing["on"] = False
    layout_note()
    if not editable and not restore:
        root.after(10, lambda: toggle_advanced(True))    # show the page layout choice for PDFs from elsewhere
    root.attributes("-topmost", True)
    root.after(600, lambda: root.attributes("-topmost", False))
    root.focus_force()
    try:
        root.mainloop()
    finally:
        photos.clear()                       # the previews only ever live in memory; drop them now
        view["gen"] += 1                     # the preview thread stops at its next page
        if pdf is not None:
            try:
                with PDFIUM_LOCK:            # (not in the middle of drawing one)
                    pdf.close()
            except Exception:
                pass
        rethemed = (result.get("choice") or ("",))[0] == "retheme"
        if not rethemed:                     # (reloading keeps the added cards' images)
            shutil.rmtree(work, ignore_errors=True)
        trim_memory()
    if rethemed:
        return choose_pages(settings, title, out_path, profile, deck=deck, pdf_path=pdf_path, pdf_layout=pdf_layout,
                            from_xml=from_xml, foreign=foreign, can_take_apart=can_take_apart,
                            restore=result["choice"][1])
    return result.get("choice")


def print_pdfs(settings: Settings, paths) -> int:
    """For each PDF: open the print window (pick fronts/backs, printer, and for Fleem's PDFs bleed,
    spacing and added cards - saved to the PDF only when you press Print), then print.
    Returns how many PDFs were sent to the printer."""
    paths = [Path(p).resolve() for p in paths if Path(p).suffix.lower() == ".pdf" and Path(p).exists()]
    if not paths:
        show_error("Pick a PDF to print.")
        return 0
    sent = 0
    for path in paths:
        how = None              # None: Fleem's layout or the cards taken out; "plain": as it is; "pairs"/"single"
        while True:
            work = WORK_DIR / f"pdf-{time.time_ns()}"   # card images cut out of a PDF made elsewhere
            try:
                info = read_layout(path) if how is None else None   # Fleem's own layout: cards can be edited
                if info is not None:
                    choice = choose_pages(settings, path.name, path, profile_of(path), deck=info["copies"], pdf_path=path,
                                          pdf_layout=info_layout(info))
                else:
                    found = None
                    if how != "plain":
                        found, err = run_with_progress(f"Reading the cards in {path.name}…",
                                                       lambda: read_foreign(path, work, how))
                    if found:
                        log.info("%s: made elsewhere - took out %s", path.name, found["summary"])
                        out = path.with_name(f"{path.stem} (Fleem).pdf")
                        choice = choose_pages(settings, path.name, out, profile_of(path), deck=found["copies"],
                                              pdf_path=path, pdf_layout=info_layout(found),
                                              foreign=found)
                    else:
                        choice = choose_pages(settings, path.name, path, profile_of(path), pdf_path=path,
                                              can_take_apart=how == "plain")
            except Exception as e:
                log.error("Couldn't open %s: %s", path.name, e)
                show_error(f"Couldn't open {path.name}: {e}")
                choice = None
            finally:
                shutil.rmtree(work, ignore_errors=True)
            if isinstance(choice, tuple) and len(choice) == 2 and choice[0] == "reopen":
                how = choice[1]
                continue
            break
        if choice is None:
            log.info("Printing cancelled for %s", path.name)
            continue
        printer, what, indexes, print_path, back_pages = choice
        if send_to_printer(settings, print_path, printer, indexes, what, back_pages, profile_of(path)):
            sent += 1
    return sent


def print_xmls(xml_paths: list[Path], settings: Settings | None = None) -> int:
    """Right-click an order XML > Print with Fleem: download its card images (with a small progress
    window), open the print window on those cards, and only when you press Print save the PDF (to the
    output folder, named after the XML) and print it. Cancel leaves nothing behind."""
    import tkinter as tk
    from tkinter import ttk
    settings = settings or Settings(SETTINGS_FILE)
    sent = 0
    for xml in xml_paths:
        title = f"{pdf_name(xml.name)}.pdf"
        work = WORK_DIR / ("print-" + re.sub(r"[^\w-]+", "_", xml.stem)[:40])
        shutil.rmtree(work, ignore_errors=True)
        images = work / "images"
        images.mkdir(parents=True)
        try:
            outcome = {}

            def download_all(xml=xml):
                try:
                    groups = parse_order(xml.read_bytes())
                    download_images(groups, images)
                    outcome["cards"] = deck_from_groups(groups, images)
                    outcome["problems"] = download_problems(groups)
                except BaseException as e:
                    outcome["error"] = str(e) or e.__class__.__name__

            root = tk.Tk()
            root.title(APP_NAME)
            root.resizable(False, False)
            root.attributes("-topmost", True)
            tk.Label(root, text=f"Getting the cards for {pdf_name(xml.name)}…\nCard images are downloading; big "
                                "orders take a few minutes.", justify="left", padx=18, pady=14).pack()
            bar = ttk.Progressbar(root, mode="indeterminate", length=320)
            bar.pack(padx=18, pady=(0, 16))
            bar.start(12)
            apply_theme(root)
            thread = threading.Thread(target=download_all, daemon=True)
            thread.start()

            def wait():
                if thread.is_alive():
                    root.after(200, wait)
                else:
                    bar.stop()                  # stop the bar's own animation before closing
                    root.destroy()
            root.after(200, wait)
            root.protocol("WM_DELETE_WINDOW", lambda: None)
            root.mainloop()
            if "error" in outcome:
                log.error("%s: %s", xml.name, outcome["error"])
                show_error(f"Couldn't get the cards for {xml.name}:\n{outcome['error']}")
                continue
            cards = outcome["cards"]
            log.info("%s: %d cards ready to print", xml.name, len(cards))
            if outcome.get("problems"):
                note = tk.Tk()
                note.withdraw()
                from tkinter import messagebox
                messagebox.showwarning(APP_NAME, outcome["problems"], parent=note)
                note.destroy()
            out = settings.output_folder / f"{safe_name(pdf_name(xml.name))}.pdf"
            choice = choose_pages(settings, title, out, profile_of(xml.name), deck=cards, from_xml=True)
            if choice is None:
                log.info("Printing cancelled for %s (no PDF made)", xml.name)
                continue
            printer, what, indexes, print_path, back_pages = choice
            if send_to_printer(settings, print_path, printer, indexes, what, back_pages, profile_of(xml.name)):
                sent += 1
        finally:
            shutil.rmtree(work, ignore_errors=True)
    return sent


def ask_ready(order_no, path: Path, sheets: int, settings: Settings) -> str | None:
    """'Fleem is waiting for you to load N sheets into the printer.' with I'm ready / Not yet.
    Returns the printer, or None for Not yet (or closing the window)."""
    import tkinter as tk
    from tkinter import ttk
    profile = profile_of(path)
    result = {}
    WHITE = "#ffffff"
    root = tk.Tk()
    root.title(f"{APP_NAME} — Order #{order_no}")
    root.configure(bg=WHITE)
    root.resizable(False, False)
    body = tk.Frame(root, bg=WHITE, padx=20, pady=18)
    body.pack(fill="both")
    try:
        icon_png = Path(getattr(sys, "_MEIPASS", HERE)) / "fleem_icon.png"
        from PIL import Image, ImageTk
        face = ImageTk.PhotoImage(Image.open(icon_png).resize((96, 96)), master=root)
        root.iconphoto(True, face)
        tk.Label(body, image=face, bg=WHITE).grid(row=0, column=0, rowspan=3, sticky="n", padx=(0, 18))
        root._face = face
    except Exception:
        pass
    sheet_word = "sheet" if sheets == 1 else "sheets"
    tk.Label(body, text=f"Fleem is waiting for you to load {sheets} {sheet_word}{loading_words(profile)} into the printer.",
             bg=WHITE, font=("Segoe UI", 13, "bold"), wraplength=420, justify="left").grid(
        row=0, column=1, sticky="w")
    details = tk.Label(body, bg=WHITE, fg="#666666", font=("Segoe UI", 9), wraplength=420, justify="left")
    details.grid(row=1, column=1, sticky="w", pady=(8, 0))
    available = installed_printers()
    printer_var = tk.StringVar(value=printer_for(available, profile) or "")
    prow = tk.Frame(body, bg=WHITE)
    prow.grid(row=3, column=1, sticky="w", pady=(8, 0))
    tk.Label(prow, text="Printer:", bg=WHITE, font=("Segoe UI", 9)).pack(side="left")
    pbox = ttk.Combobox(prow, textvariable=printer_var, values=sorted(available), width=40,
                        state="readonly" if available else "normal")
    pbox.pack(side="left", padx=(8, 0))
    pbox.bind("<<ComboboxSelected>>", lambda e: refresh())

    def open_printer_settings():
        printer = printer_var.get().strip()
        if printer and choose_printer_settings(printer, root, profile):
            refresh()
    ttk.Button(prow, text="Printer settings…", command=open_printer_settings).pack(side="left", padx=(8, 0))
    buttons = tk.Frame(root, bg=WHITE, padx=20)
    buttons.pack(fill="x", pady=(0, 16))
    ready_btn = ttk.Button(buttons, text="I'm ready", width=14, style="Accent.TButton")
    ready_btn.pack(side="right")
    ttk.Button(buttons, text="Not yet", width=14, command=root.destroy).pack(side="right", padx=(0, 8))

    layout_info = {}

    def refresh():
        printer = printer_var.get().strip()
        how = ""
        if printer:
            how = (" with your saved printer settings." if has_saved_settings(printer, profile)
                   else " with the printer's default settings (Printer settings… to pick a preset).")
        info = layout_info.get("info", "?")
        if info == "?":
            info = layout_info["info"] = read_layout(path)
        made = f"\nLayout: {describe_layout((info['bleed_mm'], info['spacing_mm'] or 0.0))}" if info else ""
        made += order_layout_note(path.name, settings) if made else ""
        details.configure(text=f"Order #{order_no} · {path.name}{made}\nPrints all {sheets} "
                               f"front{'s' if sheets != 1 else ''} (no backs)"
                               + (f" on {printer}" + how if printer else "."))
        ready_btn.configure(state="normal" if printer else "disabled")

    def ready(_event=None):
        if printer_var.get().strip():
            result["choice"] = printer_var.get().strip()
            root.destroy()
    ready_btn.configure(command=ready)
    root.bind("<Return>", ready)
    root.bind("<Escape>", lambda e: root.destroy())
    apply_theme(root)
    refresh()
    root.attributes("-topmost", True)
    root.after(800, lambda: root.attributes("-topmost", False))
    root.bell()
    root.focus_force()
    root.mainloop()
    return result.get("choice")


def offer_to_print_order(settings: Settings, order_no, pdf) -> None:
    """A new Discord order's PDF is ready: all fronts, no backs, once the paper is loaded."""
    path = Path(pdf)
    if not path.exists():
        return
    pages = pdf_page_count(path)
    sheets = (pages + 1) // 2
    choice = ask_ready(order_no, path, sheets, settings)
    if choice is None:
        log.info("Order #%s: not printed for now (Not yet) — the PDF is kept", order_no)
        return
    if not path.exists():
        show_error(f"{path.name} isn't there anymore (was the order completed or deleted?), so nothing was printed.")
        return
    send_to_printer(settings, path, choice, list(range(0, pages, 2)), f"all {sheets} fronts", profile=profile_of(path))


def print_menu_command() -> str:
    return f'{startup_command()} --print "%1"'


def register_print_menu() -> None:
    """Add "Print with Fleem" to the right-click menu of PDFs, and Fleem Proxy Companion to Open with.
    Only for this Windows user; it doesn't change which app opens PDFs normally."""
    try:
        import winreg
    except ImportError:
        return
    exe_icon = startup_command().split('" ', 1)[0].strip('"')
    command = print_menu_command()
    values = {
        r"Software\Classes\SystemFileAssociations\.pdf\shell\FleemPrint":
            {"": "Print with Fleem", "MUIVerb": "Print with Fleem", "Icon": exe_icon},
        r"Software\Classes\SystemFileAssociations\.pdf\shell\FleemPrint\command": {"": command},
        r"Software\Classes\SystemFileAssociations\.xml\shell\FleemPrint":
            {"": "Print with Fleem", "MUIVerb": "Print with Fleem", "Icon": exe_icon},
        r"Software\Classes\SystemFileAssociations\.xml\shell\FleemPrint\command": {"": command},
        r"Software\Classes\FleemPDFMaker.Print": {"": "PDF (print with Fleem)", "FriendlyAppName": APP_NAME},
        r"Software\Classes\FleemPDFMaker.Print\DefaultIcon": {"": exe_icon},
        r"Software\Classes\FleemPDFMaker.Print\shell\open\command": {"": command},
        r"Software\Classes\.pdf\OpenWithProgids": {"FleemPDFMaker.Print": ""},
        r"Software\Classes\.xml\OpenWithProgids": {"FleemPDFMaker.Print": ""},
    }
    # "Always open with" chosen for the .exe under its old name: point that at this one
    if getattr(sys, "frozen", False):
        for old in LEGACY_NAMES:
            key = rf"Software\Classes\Applications\{old}.exe\shell\open\command"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key):
                    values[key] = {"": f'"{sys.executable}" "%1"'}
            except OSError:
                pass
    try:
        changed = False
        for key, entries in values.items():
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key, 0,
                                    winreg.KEY_QUERY_VALUE | winreg.KEY_SET_VALUE) as k:
                for name, value in entries.items():
                    try:
                        if winreg.QueryValueEx(k, name)[0] == value:
                            continue
                    except OSError:
                        pass
                    winreg.SetValueEx(k, name, 0, winreg.REG_SZ, value)
                    changed = True
        if changed:
            import ctypes
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)   # refresh Explorer
            log.info('Added "Print with Fleem" to the right-click menu for PDFs and XMLs')
    except Exception as e:
        log.warning("Couldn't add Print with Fleem to the right-click menu: %s", e)


def merge_groups(group_lists) -> list[dict]:
    """Combine cards from several XMLs; the same card (and back) in two files has its copies added up."""
    merged: dict[tuple, dict] = {}
    for groups in group_lists:
        for g in groups:
            key = (g["front_id"], g["back_id"])
            if key in merged:
                merged[key]["count"] += g["count"]
            else:
                merged[key] = dict(g)
    return list(merged.values())


def merged_pdf_name(xml_names) -> str:
    """'kim - Order 1 + lee - Order 2' (shortened if very long)."""
    stems = [pdf_name(n) for n in xml_names]
    name = " + ".join(stems)
    if len(name) > 150:
        name = f"{stems[0][:100]} + {len(stems) - 1} more"
    return name


def make_pdf(settings: Settings, label: str, xmls: list[bytes], pdf_stem: str,
             work_name: str, layout=(1.0, 0.0)) -> list[str]:
    """One or more XMLs -> one PDF in the output folder. `label` is for the log."""
    with RENDER_LOCK:
        work = WORK_DIR / work_name
        shutil.rmtree(work, ignore_errors=True)
        images = work / "images"
        images.mkdir(parents=True)
        try:
            groups = merge_groups(parse_order(x) for x in xmls)
            total = sum(g["count"] for g in groups)
            mdfcs = sum(g["count"] for g in groups if g["back_id"])
            log.info("%s: %d cards (%d MDFCs), %s — downloading images…", label, total, mdfcs,
                     describe_layout(layout))
            download_images(groups, images)
            pdf = work / f"{safe_name(pdf_stem)}.pdf"
            write_cards_pdf(groups, images, pdf, layout)
            settings.output_folder.mkdir(parents=True, exist_ok=True)
            dest = settings.output_folder / pdf.name
            try:
                os.replace(pdf, dest)
            except PermissionError:
                raise RuntimeError(f"{dest.name} is open in another program (e.g. a PDF viewer), so it can't be "
                                   "replaced. Close it there; Fleem will try again.") from None
            except OSError:
                shutil.move(str(pdf), dest)
            other, gone = substituted_cards(groups), missing_cards(groups)
            if other:
                on_warning(f"{label}: the chosen art for {len(other)} card{'s' * (len(other) != 1)} couldn't be "
                           f"downloaded, so a different art of the same card was used "
                           f"({', '.join(other[:3])}{'…' if len(other) > 3 else ''}).")
            if gone:
                on_warning(f"{label}: {len(gone)} card image{'s' if len(gone) > 1 else ''} couldn't be downloaded "
                           f"({', '.join(gone[:3])}{'…' if len(gone) > 3 else ''}) - grey placeholders were used. "
                           "Replace them with Edit cards in the print window.")
            return [str(dest)]
        finally:
            shutil.rmtree(work, ignore_errors=True)


def process_order(settings: Settings, order_no: int, attachment: dict) -> list[str]:
    xml_bytes, _ = download(attachment["url"])
    name = attachment["filename"]
    return make_pdf(settings, f"Order #{order_no}", [xml_bytes], pdf_name(name),
                    f"order-{order_no}", layout=layout_from_name(name, settings) or last_layout())


def make_pdf_from_files(settings: Settings, xml_paths: list[Path], layout) -> list[str]:
    """For "Make PDF from XML…" (one file) and "Merge XMLs into one PDF…" (several)."""
    names = [p.name for p in xml_paths]
    label = names[0] if len(names) == 1 else f"Merge of {len(names)} XMLs"
    work = ("manual-" if len(names) == 1 else "merge-") + re.sub(r"[^\w-]+", "_", xml_paths[0].stem)[:40]
    return make_pdf(settings, label, [p.read_bytes() for p in xml_paths],
                    merged_pdf_name(names), work, layout=layout)


def recycle(path) -> bool:
    """Send a file to the Recycle Bin (falls back to deleting it). True if it's gone."""
    path = Path(path)
    if not path.exists():
        return True
    try:
        import ctypes
        from ctypes import wintypes

        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [("hwnd", wintypes.HWND), ("wFunc", wintypes.UINT), ("pFrom", wintypes.LPCWSTR),
                        ("pTo", wintypes.LPCWSTR), ("fFlags", ctypes.c_uint16),
                        ("fAnyOperationsAborted", wintypes.BOOL), ("hNameMappings", ctypes.c_void_p),
                        ("lpszProgressTitle", wintypes.LPCWSTR)]
        FO_DELETE, FOF_SILENT, FOF_NOCONFIRMATION, FOF_ALLOWUNDO, FOF_NOERRORUI = 3, 0x4, 0x10, 0x40, 0x400
        op = SHFILEOPSTRUCTW(None, FO_DELETE, str(path.resolve()) + "\0", None,
                             FOF_SILENT | FOF_NOCONFIRMATION | FOF_ALLOWUNDO | FOF_NOERRORUI, False, None, None)
        if ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op)) == 0 and not path.exists():
            return True
    except Exception:
        pass
    try:
        path.unlink()
        return True
    except OSError as e:
        log.warning("Couldn't delete %s: %s", path.name, e)
        return False


def remove_order_pdfs(entry: dict, why: str) -> None:
    """Delete (to the Recycle Bin) the PDFs Fleem made for an order."""
    for pdf in entry.get("pdfs", []):
        if Path(pdf).exists() and recycle(pdf):
            log.info("Order #%s %s — moved %s to the Recycle Bin", entry.get("order", "?"), why, Path(pdf).name)


def cleanup_finished_orders(discord: Discord, state: dict, open_channel_ids: set) -> None:
    """Orders whose channel was deleted (/orders delete) lose their PDF."""
    for key, entry in list(state.items()):
        if key in open_channel_ids or not entry.get("pdfs"):
            continue
        if discord.channel_gone(key):
            remove_order_pdfs(entry, "was deleted")
            state.pop(key)
            save_state(state)


def check_once(settings: Settings, discord: Discord, bot_id: str) -> None:
    state = load_state()
    channels = discord.order_channels(settings.guild_id)
    cleanup_finished_orders(discord, state, {ch["id"] for _, ch in channels})
    for order_no, ch in channels:
        if order_no < settings.skip_orders_before:
            continue
        key = ch["id"]
        done = state.get(key, {})
        if done.get("failed_at") and time.time() - done["failed_at"] < RETRY_FAILED_AFTER:
            continue
        try:
            att, status = discord.order_post(ch["id"], bot_id)
        except Exception as e:
            log.warning("Order #%s: couldn't read its channel: %s", order_no, e)
            continue
        if status and "ongoing" not in status.lower():
            if done.get("pdfs"):      # marked Completed in Discord: its PDF isn't needed anymore
                remove_order_pdfs(done, "was completed")
                state[key] = {**done, "pdfs": [], "completed": True}
                save_state(state)
            continue
        if att is None or (done.get("attachment_id") == att["id"] and not done.get("completed")):
            continue   # no XML yet, or this exact XML already has its PDF
        log.info("Order #%s: making PDF from %s", order_no, att["filename"])
        on_status(f"Making the PDF for order #{order_no}…")
        try:
            saved = process_order(settings, order_no, att)
        except SystemExit:
            raise
        except Exception as e:
            log.error("Order #%s failed: %s", order_no, e)
            state[key] = {**done, "failed_at": time.time(), "error": str(e)[:500]}
            save_state(state)
            continue
        for old in done.get("pdfs", []):   # the order's XML changed: remove the outdated PDF
            if old not in saved:
                recycle(old)
        state[key] = {"order": order_no, "attachment_id": att["id"], "xml": att["filename"],
                      "pdfs": saved, "made_at": time.strftime("%Y-%m-%d %H:%M:%S")}
        save_state(state)
        log.info("Order #%s: ✔ saved %s", order_no, ", ".join(Path(p).name for p in saved))
        on_pdf_ready(order_no, saved)
        if on_order_ready is not None:
            on_order_ready(order_no, saved)
        else:
            open_pdfs(settings, saved)


FLEEM_PORT = 47913
PRINT_QUEUE_FILE = HERE / "print_queue.json"     # Discord orders still waiting for "I'm ready"


# ============================================================ updates
# The version of this copy. Release tags on GitHub use the same number with a "v" in front (v0.4.0);
# build.bat copies it into the .exe's file properties. Small updates and fixes raise the last number
# (0.4.0 -> 0.4.1); bigger updates raise the middle one and reset the last (0.4.1 -> 0.5.0).
APP_VERSION = "0.7.3"
# The GitHub repository Fleem looks in for new releases ("owner/name"). If you fork Fleem, put your
# own repository here (or "" to switch the update check off).
GITHUB_REPO = "YOUR-GITHUB-NAME/fleem-proxy-companion"
UPDATE_STATE_FILE = HERE / "update_state.json"     # remembers a version you chose to skip
UPDATE_CHECK_EVERY = 24 * 60 * 60                  # seconds between automatic checks


def updates_configured() -> bool:
    """True when GITHUB_REPO names a real repository (not the placeholder)."""
    return bool(re.fullmatch(r"[\w.-]+/[\w.-]+", GITHUB_REPO or "")) and not GITHUB_REPO.startswith("YOUR-")


def version_tuple(text: str) -> tuple:
    """"v1.2" / "1.2.0" / "1.2.0-beta" -> (1, 2, 0, 0), so versions compare as numbers."""
    core = re.split(r"[-+ ]", (text or "").strip().lstrip("vV"), maxsplit=1)[0]
    nums = [int(n) for n in re.findall(r"\d+", core)[:4]]
    return tuple(nums + [0] * (4 - len(nums)))


def load_update_state() -> dict:
    try:
        return json.loads(UPDATE_STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_update_state(state: dict) -> None:
    try:
        UPDATE_STATE_FILE.write_text(json.dumps(state), encoding="utf-8")
    except OSError as e:
        log.warning("Couldn't save %s: %s", UPDATE_STATE_FILE.name, e)


def latest_release() -> dict | None:
    """The newest published release on GitHub (drafts and pre-releases are left out by GitHub), or
    None if the repository has no releases yet."""
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
        headers={"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28",
                 "User-Agent": f"FleemProxyCompanion/{APP_VERSION}"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise
    # a folder build updates from the release's .zip (the .exe and its _internal folder), a one-file
    # build from its .exe
    want = ".zip" if is_folder_build() else ".exe"
    exes = [a for a in data.get("assets", []) if str(a.get("name", "")).lower().endswith(want)]
    exes.sort(key=lambda a: "fleem" not in a["name"].lower())      # the one named after Fleem first
    asset = None
    if exes:
        digest = str(exes[0].get("digest") or "")
        asset = {"name": exes[0]["name"], "url": exes[0]["browser_download_url"],
                 "size": int(exes[0].get("size") or 0),
                 "sha256": digest.split(":", 1)[1].lower() if digest.lower().startswith("sha256:") else ""}
    tag = str(data.get("tag_name") or "")
    return {"version": tag.lstrip("vV"), "tag": tag, "title": data.get("name") or tag,
            "notes": (data.get("body") or "").strip(), "page": data.get("html_url") or
            f"https://github.com/{GITHUB_REPO}/releases/latest", "asset": asset}


def find_update(force: bool = False) -> dict | None:
    """The latest release if it's newer than this copy (and not a version you skipped, unless
    `force`: a check you asked for yourself). None if you're up to date."""
    if not updates_configured():
        return None
    release = latest_release()
    if not release or version_tuple(release["version"]) <= version_tuple(APP_VERSION):
        return None
    if not force and load_update_state().get("skipped") == release["version"]:
        return None
    return release


def is_folder_build() -> bool:
    """Is this the folder build (the .exe with an _internal folder next to it: starts fast, nothing to
    unpack) rather than the one-file .exe (unpacks itself into a temporary folder every time)?"""
    mei = getattr(sys, "_MEIPASS", None)
    return bool(getattr(sys, "frozen", False) and mei and
                Path(mei).resolve() == (Path(sys.executable).parent / "_internal").resolve())


UPDATE_FOLDER = "_update_new"          # a folder build's next version, unpacked, until Fleem has exited
UPDATE_SCRIPT = "_update.cmd"          # moves it into place once Fleem has exited, then starts it
_pending_swap = {}


def update_files(exe: Path) -> tuple[Path, Path, Path]:
    """(new download, part-downloaded file, the old copy kept until the next start) next to the exe."""
    if is_folder_build():
        return (exe.with_name(f"{exe.stem}.update.zip"), exe.with_name(f"{exe.stem}.update.zip.part"),
                exe.with_name(f"{exe.stem}.old.exe"))
    return (exe.with_name(f"{exe.stem}.new.exe"), exe.with_name(f"{exe.stem}.new.exe.part"),
            exe.with_name(f"{exe.stem}.old.exe"))


def download_update(release: dict, exe: Path, progress=lambda done, total: None) -> Path:
    """Download the release's .exe next to `exe` and check it's complete and untouched (size, and the
    SHA-256 GitHub lists for it). Returns the downloaded file."""
    import hashlib
    asset = release["asset"]
    new, part, _ = update_files(exe)
    req = urllib.request.Request(asset["url"], headers={"User-Agent": f"FleemProxyCompanion/{APP_VERSION}",
                                                        "Accept": "application/octet-stream"})
    sha, done = hashlib.sha256(), 0
    try:
        with urllib.request.urlopen(req, timeout=60) as resp, open(part, "wb") as f:
            total = asset["size"] or int(resp.headers.get("Content-Length") or 0)
            while True:
                chunk = resp.read(256 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                sha.update(chunk)
                done += len(chunk)
                progress(done, total)
        if asset["size"] and done != asset["size"]:
            raise RuntimeError(f"the download was incomplete ({done:,} of {asset['size']:,} bytes)")
        if asset["sha256"] and sha.hexdigest() != asset["sha256"]:
            raise RuntimeError("the download doesn't match the release (SHA-256 check failed)")
        with open(part, "rb") as f:
            magic = f.read(2)
        if is_folder_build():
            if magic != b"PK":
                raise RuntimeError("the download isn't the release's .zip")
        elif magic != b"MZ":
            raise RuntimeError("the download isn't a Windows program")
        os.replace(part, new)
        return new
    except BaseException:
        part.unlink(missing_ok=True)
        raise


def install_update(new: Path, exe: Path) -> None:
    """Put the downloaded copy in place of `exe`. Windows lets a running program be renamed (not
    deleted), so the running one becomes "<name>.old.exe", removed at the next start. A folder build
    unpacks the new version next to it instead; it's moved into place when Fleem exits (see
    start_new_copy)."""
    if is_folder_build():
        return _stage_folder_update(new, exe)
    _, _, old = update_files(exe)
    old.unlink(missing_ok=True)
    os.replace(exe, old)
    try:
        os.replace(new, exe)
    except BaseException:
        os.replace(old, exe)       # put things back as they were
        raise


def _stage_folder_update(zip_path: Path, exe: Path) -> None:
    """Unpack a folder build's release .zip next to it and write the script that swaps it in."""
    import zipfile
    folder = exe.parent
    staged = folder / UPDATE_FOLDER
    shutil.rmtree(staged, ignore_errors=True)
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        roots = {n.split("/", 1)[0] for n in names}
        prefix = ""
        if len(roots) == 1 and not any(n == next(iter(roots)) for n in names):   # everything in one top folder
            prefix = next(iter(roots)) + "/"
        if f"{prefix}{exe.name}" not in names or not any(n.startswith(f"{prefix}_internal/") for n in names):
            raise RuntimeError(f"the .zip doesn't have {exe.name} and its _internal folder")
        for n in names:
            if not n.startswith(prefix) or n.endswith("/"):
                continue
            rel = Path(n[len(prefix):])
            if rel.is_absolute() or ".." in rel.parts:
                raise RuntimeError("the .zip has an unsafe file path")
            target = staged / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            with z.open(n) as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
    zip_path.unlink(missing_ok=True)
    script = folder / UPDATE_SCRIPT
    pid = os.getpid()
    script.write_text("\r\n".join([
        "@echo off",
        "rem Written by Fleem Proxy Companion: puts its update in place once it has closed, then starts it.",
        'cd /d "%~dp0"',
        ":wait",
        f'tasklist /fi "PID eq {pid}" 2>nul | find "{pid}" >nul && (timeout /t 1 /nobreak >nul & goto wait)',
        f'robocopy "{UPDATE_FOLDER}" "." /E /MOVE /R:10 /W:1 /NFL /NDL /NJH /NJS /NP >nul',
        f'rmdir /s /q "{UPDATE_FOLDER}" 2>nul',
        f'start "" "{exe.name}" {RESTART_FLAG}',
        '(goto) 2>nul & del "%~f0"',
    ]) + "\r\n", encoding="utf-8")
    _pending_swap["script"] = script
    log.info("Update unpacked into %s; it goes in when Fleem closes", staged)


RESTART_FLAG = "--updated"       # what the new copy is told: "--updated" (from GitHub) or "--rebuilt"


def child_env() -> dict:
    """This environment without the bits a one-file build adds for itself (where it unpacked itself,
    which goes away when this copy exits)."""
    mei = getattr(sys, "_MEIPASS", None)
    env = {}
    for key, value in os.environ.items():
        if key.startswith("_PYI") or key == "_MEIPASS2":
            continue
        if mei and mei.lower() in value.lower():
            parts = [p for p in value.split(os.pathsep) if mei.lower() not in p.lower()]
            if not parts:
                continue        # e.g. TCL_LIBRARY pointing into this copy's unpacked files
            value = os.pathsep.join(parts)
        env[key] = value
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    return env


def start_new_copy(exe: Path, flag: str | None = None) -> None:
    """Start the freshly installed exe as its own program (not as part of this one: a one-file build
    otherwise passes on where it unpacked itself, and that folder goes away when this copy exits)."""
    flags = 0x00000008 | 0x00000200 if os.name == "nt" else 0   # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    script = _pending_swap.pop("script", None)
    if script is not None:              # a folder build's update: the script swaps it in after we exit
        subprocess.Popen(["cmd", "/c", str(script)], cwd=str(exe.parent), env=child_env(), close_fds=True,
                         creationflags=flags | 0x08000000)              # CREATE_NO_WINDOW
        return
    subprocess.Popen([str(exe), flag or RESTART_FLAG], cwd=str(exe.parent), env=child_env(), close_fds=True,
                     creationflags=flags)


# ------------------------------------------------------------ rebuilding from source
# A build made from the source in its own folder (the personal build) can rebuild itself when that
# source changes: the .exe carries a fingerprint of the files it was built from, and when the files next
# to it no longer match, Fleem runs the build script in the background and switches to the new .exe.
AUTO_REBUILD = False                       # an add-on turns this on
BUILD_SCRIPT = "build_exe.bat"             # run as "build_exe.bat --auto": no questions, leaves dist\
BUILD_SOURCES = ("fleem_pdf_maker.py", "fleem_personal.py", "fleem_mpcfill.py", "build_exe.bat", "requirements.txt")
BUILD_FINGERPRINT = "build_fingerprint.txt"
BUILD_LOG = "build.log"
REBUILD_CHECK_EVERY = 5 * 60               # seconds


def source_fingerprint(folder: Path | None = None) -> str | None:
    """A fingerprint of the source files in `folder` (Fleem's folder), or None if there are none."""
    import hashlib
    folder = folder or HERE
    sha, found = hashlib.sha256(), False
    for name in BUILD_SOURCES:
        f = folder / name
        if f.exists():
            found = True
            sha.update(name.encode() + b"\0" + f.read_bytes() + b"\0")
    return sha.hexdigest() if found else None


def write_build_fingerprint() -> None:
    """The build script: remember what this .exe is built from (packed into the .exe)."""
    (HERE / BUILD_FINGERPRINT).write_text(source_fingerprint() or "", encoding="utf-8")
    print("build_fingerprint.txt written")


def built_fingerprint() -> str | None:
    """The fingerprint packed into this .exe when it was built (None when running from source, or for
    a build without one)."""
    if not getattr(sys, "frozen", False):
        return None
    try:
        return (Path(sys._MEIPASS) / BUILD_FINGERPRINT).read_text(encoding="utf-8").strip() or None
    except (OSError, AttributeError):
        return None


def source_changed() -> str | None:
    """The source's fingerprint if Fleem should rebuild itself (it's a build from this folder and the
    source no longer matches it), else None."""
    if not (AUTO_REBUILD and getattr(sys, "frozen", False)):
        return None
    if not ((HERE / "fleem_pdf_maker.py").exists() and (HERE / BUILD_SCRIPT).exists()):
        return None                        # just the .exe (e.g. a friend's copy): nothing to build from
    built, current = built_fingerprint(), source_fingerprint()
    return current if built and current and current != built else None


def rebuild_from_source(exe: Path) -> Path:
    """Run the build script hidden in the background (what it prints goes to build.log) and put the new
    .exe next to this one as "<name>.new.exe". Returns that file."""
    import shutil
    with open(HERE / BUILD_LOG, "wb") as out:
        result = subprocess.run(["cmd", "/c", BUILD_SCRIPT, "--auto"], cwd=str(HERE), env=child_env(), stdout=out,
                                stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL, timeout=45 * 60,
                                creationflags=0x08000000 if os.name == "nt" else 0)    # CREATE_NO_WINDOW
    built = HERE / "dist" / f"{APP_NAME}.exe"
    if result.returncode != 0 or not built.exists():
        raise RuntimeError(f"the build didn't finish (see {BUILD_LOG} in Fleem's folder)")
    new, _, _ = update_files(exe)
    os.replace(built, new)
    for leftover in ("build", "dist"):
        shutil.rmtree(HERE / leftover, ignore_errors=True)
    (HERE / f"{APP_NAME}.spec").unlink(missing_ok=True)
    return new


def watch_source(stop: threading.Event, put, notify) -> None:
    """Personal build: check every few minutes whether the source changed; if so, rebuild and queue
    ("rebuilt", new exe) so Fleem switches over when no window is open."""
    if not (AUTO_REBUILD and getattr(sys, "frozen", False)):
        return
    tried = None
    stop.wait(30)
    while not stop.is_set():
        try:
            current = source_changed()
            if current and current != tried:
                stop.wait(20)              # files are often copied a few at a time: let them settle
                if source_fingerprint() != current:
                    continue
                tried = current            # (a failed build isn't tried again until the source changes)
                log.info("The source changed since this .exe was built: rebuilding")
                notify("Fleem's code was updated: rebuilding in the background (a minute or two).")
                new = rebuild_from_source(Path(sys.executable))
                log.info("Rebuilt: %s", new.name)
                put(("rebuilt", str(new)))
                return
        except Exception as e:
            log.error("Rebuild failed: %s", e)
            notify(f"Couldn't rebuild Fleem: {e}")
        stop.wait(REBUILD_CHECK_EVERY)


def install_rebuilt(new) -> None:
    """Put a rebuilt .exe in place of this one; the restart that follows tells it "--rebuilt"."""
    global RESTART_FLAG
    install_update(Path(new), Path(sys.executable))
    RESTART_FLAG = "--rebuilt"


def adopt_new_name() -> bool:
    """A .exe still called by Fleem's old name ("Fleem PDF Maker.exe", e.g. just rebuilt or updated by
    an older copy) copies itself to the new name, starts that and returns True (this copy then exits;
    the new one removes the old file). It can't rename itself: a one-file build keeps reading its own
    .exe while it runs."""
    if not getattr(sys, "frozen", False):
        return False
    exe = Path(sys.executable)
    if exe.stem not in LEGACY_NAMES:
        return False
    target = exe.with_name(f"{APP_NAME}.exe")
    try:
        part = target.with_name(target.name + ".part")
        shutil.copyfile(exe, part)
        os.replace(part, target)
        start_new_copy(target, "--rebuilt" if "--rebuilt" in sys.argv else "--renamed")
        log.info("Copied %s to %s and started it", exe.name, target.name)
        return True
    except OSError as e:
        log.warning("Couldn't move to the new name (%s): %s", target.name, e)
        return False


def clean_up_after_update() -> None:
    """Remove the copy an update replaced (it may take the old copy a moment to exit) and any
    half-finished download."""
    if not getattr(sys, "frozen", False):
        return
    new, part, old = update_files(Path(sys.executable))
    legacy = [Path(sys.executable).with_name(f"{n}{end}") for n in LEGACY_NAMES
              for end in (".exe", ".old.exe", ".new.exe")]      # the copy under Fleem's old name, once it has exited

    def job():
        part.unlink(missing_ok=True) if part.exists() else None
        if is_folder_build() and not (Path(sys.executable).parent / UPDATE_SCRIPT).exists():
            shutil.rmtree(Path(sys.executable).parent / UPDATE_FOLDER, ignore_errors=True)   # an update that didn't finish
        for _ in range(30):
            try:
                for f in (old, new, *legacy):
                    f.unlink(missing_ok=True)
                return
            except OSError:
                time.sleep(1)
    threading.Thread(target=job, daemon=True).start()


def write_version_info(path: Path = HERE / "version_info.txt") -> None:
    """build.bat: copy APP_VERSION into the .exe's file properties."""
    nums = version_tuple(APP_VERSION)
    text = path.read_bytes().decode("utf-8")        # bytes: keep its line endings as they are
    text = re.sub(r"filevers=\([^)]*\)", f"filevers={nums}", text)
    text = re.sub(r"prodvers=\([^)]*\)", f"prodvers={nums}", text)
    text = re.sub(r"StringStruct\('(FileVersion|ProductVersion)', '[^']*'\)",
                  lambda m: f"StringStruct('{m.group(1)}', '{APP_VERSION}')", text)
    text = re.sub(r"StringStruct\('(FileDescription|ProductName|InternalName)', '[^']*'\)",
                  lambda m: f"StringStruct('{m.group(1)}', '{APP_NAME}')", text)
    text = re.sub(r"StringStruct\('OriginalFilename', '[^']*'\)", f"StringStruct('OriginalFilename', '{APP_NAME}.exe')", text)
    path.write_bytes(text.encode("utf-8"))
    print(f"version_info.txt set to {APP_VERSION}")


def tidy_release_notes(text: str) -> str:
    """GitHub release notes (Markdown) as plain text for the update window."""
    lines = []
    for line in re.sub(r"\r\n?", "\n", text or "").split("\n"):
        line = re.sub(r"^\s{0,3}#{1,6}\s*", "", line)                     # headings
        line = re.sub(r"^(\s*)[-*+]\s+", lambda m: m.group(1) + "• ", line)  # bullets
        line = re.sub(r"!?\[([^\]]*)\]\([^)]*\)", r"\1", line)              # links and images: keep the words
        line = re.sub(r"(\*\*|__|`)", "", line)                             # bold, code
        lines.append(line.rstrip())
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


def update_window(release: dict) -> str:
    """Tell you about a new version, with its release notes. Returns "installed" (the new version is
    in place and Fleem should restart into it), "skip" (don't mention this version again) or "later"."""
    import tkinter as tk
    from tkinter import ttk
    import webbrowser
    frozen = getattr(sys, "frozen", False)
    can_install = frozen and release.get("asset") is not None
    exe = Path(sys.executable)
    result = {"value": "later"}
    busy = {"on": False}

    root = tk.Tk()
    root.withdraw()
    root.title(f"{APP_NAME} update")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    try:
        icon_png = Path(getattr(sys, "_MEIPASS", HERE)) / "fleem_icon.png"
        if icon_png.exists():
            root.iconphoto(True, tk.PhotoImage(master=root, file=str(icon_png)))
    except Exception:
        pass
    T(root, bg="surface")
    body = T(tk.Frame(root, padx=22, pady=18), bg="surface")
    body.pack(fill="both", expand=True)
    T(tk.Label(body, text=f"Version {release['version']} is available", font=("Segoe UI", 13, "bold"), anchor="w"),
      bg="surface", fg="text").pack(fill="x")
    how = ("Update now downloads it and restarts Fleem (orders waiting to print are kept)." if can_install else
           "Download it from the release page." if frozen else
           "You're running Fleem from its source code: download the new version from the release page.")
    T(tk.Label(body, text=f"You have {APP_VERSION}. {how}", font=("Segoe UI", 9), anchor="w", justify="left",
               wraplength=480), bg="surface", fg="muted").pack(fill="x", pady=(2, 10))
    if release.get("notes"):
        T(tk.Label(body, text="What's new", font=("Segoe UI", 9, "bold"), anchor="w"), bg="surface",
          fg="text").pack(fill="x")
        box = T(tk.Frame(body, highlightthickness=1), bg="field", highlightbackground="border",
                highlightcolor="border")
        box.pack(fill="both", expand=True, pady=(4, 0))
        notes = T(tk.Text(box, width=64, height=min(14, max(4, tidy_release_notes(release["notes"]).count("\n") + 2)), wrap="word",
                          relief="flat", borderwidth=0, padx=10, pady=8, font=("Segoe UI", 9)),
                  bg="field", fg="text", insertbackground="text")
        bar = ttk.Scrollbar(box, orient="vertical", command=notes.yview)
        notes.configure(yscrollcommand=bar.set)
        notes.insert("1.0", tidy_release_notes(release["notes"]))
        notes.configure(state="disabled")
        notes.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")
    status = T(tk.Label(body, text="", font=("Segoe UI", 9), anchor="w", justify="left", wraplength=480),
               bg="surface", fg="muted")
    status.pack(fill="x", pady=(10, 0))
    progress = ttk.Progressbar(body, mode="determinate", length=480, maximum=1000)

    foot = T(tk.Frame(root, padx=22, pady=12), bg="surface2")
    foot.pack(fill="x")
    page_btn = ttk.Button(foot, text="Release page", command=lambda: webbrowser.open(release["page"]))
    page_btn.pack(side="left")

    def finish(value):
        if busy["on"]:
            return
        result["value"] = value
        root.destroy()

    def open_page_and_close():
        webbrowser.open(release["page"])
        finish("later")

    def set_buttons(state):
        for b in (main_btn, later_btn, skip_btn):
            b.configure(state=state)

    def start_update():
        busy["on"] = True
        set_buttons("disabled")
        progress.pack(fill="x", pady=(6, 0))
        progress["value"] = 0
        status.configure(text="Downloading…")
        T(status, fg="muted")
        shared = {"done": 0, "total": 0}

        def job():
            try:
                new = download_update(release, exe, lambda d, t: shared.update(done=d, total=t))
                install_update(new, exe)
                shared["ok"] = True
            except BaseException as e:
                log.error("Update failed: %s", e)
                shared["error"] = str(e) or e.__class__.__name__
        threading.Thread(target=job, daemon=True).start()

        def poll():
            if "ok" in shared:
                busy["on"] = False
                log.info("Updated to %s; restarting", release["version"])
                finish("installed")
                return
            if "error" in shared:
                busy["on"] = False
                progress.pack_forget()
                status.configure(text=f"Couldn't update: {shared['error']}\nTry again, or download it from the "
                                      "release page.")
                T(status, fg="danger")
                set_buttons("normal")
                main_btn.configure(text="Try again")
                return
            done, total = shared["done"], shared["total"]
            if total:
                progress["value"] = 1000 * done / total
                status.configure(text=f"Downloading… {done / 1048576:.1f} of {total / 1048576:.1f} MB")
            root.after(120, poll)
        root.after(120, poll)

    main_btn = ttk.Button(foot, text="Update now" if can_install else "Open release page", style="Accent.TButton",
                          command=start_update if can_install else open_page_and_close)
    later_btn = ttk.Button(foot, text="Later", command=lambda: finish("later"))
    skip_btn = ttk.Button(foot, text="Skip this version", command=lambda: finish("skip"))
    main_btn.pack(side="right")
    later_btn.pack(side="right", padx=(0, 8))
    skip_btn.pack(side="right", padx=(0, 8))
    if not can_install:
        page_btn.pack_forget()           # the main button already opens it
    root.protocol("WM_DELETE_WINDOW", lambda: finish("later"))
    root.bind("<Escape>", lambda e: finish("later"))
    apply_theme(root)
    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_reqwidth()) // 2
    y = max(0, (root.winfo_screenheight() - root.winfo_reqheight()) // 3)
    root.geometry(f"+{x}+{y}")
    root.deiconify()
    main_btn.focus_set()
    root.mainloop()
    return result["value"]


def single_instance_lock(wait: float = 0.0):
    """Stop a second copy from running (e.g. started at login and again by hand). The same local
    port is where right-click > Print with Fleem hands its files to the running Fleem. `wait`: keep
    trying this many seconds (after an update, while the old copy closes)."""
    deadline = time.time() + wait
    while True:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock.bind(("127.0.0.1", FLEEM_PORT))
            break
        except OSError:
            sock.close()
            if time.time() >= deadline:
                raise SystemExit(f"{APP_NAME} is already running (look for Fleem in the taskbar corner).")
            time.sleep(0.5)
    sock.listen(8)
    return sock


def hand_to_running_fleem(paths) -> bool:
    """Right-click > Print with Fleem: if Fleem is already running, let it open the print window (it
    has the image cache and starts instantly). False if it isn't running."""
    try:
        with socket.create_connection(("127.0.0.1", FLEEM_PORT), timeout=2) as c:
            c.sendall(json.dumps({"print": [str(Path(p).resolve()) for p in paths]}).encode() + b"\n")
            c.settimeout(5)
            return c.recv(16).startswith(b"ok")
    except OSError:
        return False


def load_print_queue() -> list:
    try:
        return [e for e in json.loads(PRINT_QUEUE_FILE.read_text(encoding="utf-8")) if Path(e["pdf"]).exists()]
    except Exception:
        return []


def save_print_queue(entries: list) -> None:
    try:
        PRINT_QUEUE_FILE.write_text(json.dumps(entries, indent=1), encoding="utf-8")
    except OSError as e:
        log.warning("Couldn't save the print queue: %s", e)


def setup_logging(to_console: bool) -> None:
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s  %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = logging.handlers.RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=2, encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    if to_console and sys.stdout:
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(fmt)
        log.addHandler(ch)
    # a hard crash (Fleem just vanishing) writes what it was doing to crash.log
    global _crash_file
    try:
        import faulthandler
        _crash_file = open(HERE / "crash.log", "a", encoding="utf-8")
        faulthandler.enable(_crash_file)
    except Exception:
        pass
    # errors inside Fleem's windows go to the log (with no console they'd otherwise vanish)
    try:
        import tkinter
        tkinter.Tk.report_callback_exception = \
            lambda self, *exc: log.error("Error in a Fleem window", exc_info=exc)
    except Exception:
        pass


_crash_file = None


RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
_STARTUP_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "Microsoft" / "Windows" / "Start Menu" / \
    "Programs" / "Startup"
OLD_STARTUP_FILES = [_STARTUP_DIR / f"{APP_NAME}.vbs", _STARTUP_DIR / "Fleem PDF Helper.vbs"]


def startup_command() -> str:
    exe_file = HERE / f"{APP_NAME}.exe"
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    if exe_file.exists():
        return f'"{exe_file}"'
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe") if exe.name.lower() == "python.exe" else exe
    return f'"{pyw if pyw.exists() else exe}" "{Path(__file__).resolve()}"'


def login_command() -> str:
    """What Windows runs at login ("--startup": stay quietly in the background, no window)."""
    return f"{startup_command()} --startup"


def install_startup() -> None:
    """Start Fleem at every Windows login (listed in Task Manager > Startup apps). That only
    makes sense running in the background, so it turns on keep_running too."""
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
        winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, login_command())
    try:
        save_setting("keep_running", "yes")
    except OSError as e:
        log.warning("Couldn't save keep_running: %s", e)
    for f in OLD_STARTUP_FILES:          # the old method, replaced by this
        f.unlink(missing_ok=True)
    print(f"{APP_NAME} will now start when you log in (see Task Manager > Startup apps).")


def remove_startup() -> None:
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, APP_NAME)
    except OSError:
        pass
    for f in OLD_STARTUP_FILES:
        f.unlink(missing_ok=True)
    print(f"{APP_NAME} won't start at login anymore.")


def fix_startup_location() -> None:
    """If the folder was moved (or Fleem renamed), point the existing startup entry at where Fleem is now."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_QUERY_VALUE | winreg.KEY_SET_VALUE) as k:
            for old in LEGACY_NAMES:           # the entry under Fleem's old name becomes one under its new name
                try:
                    winreg.QueryValueEx(k, old)
                except OSError:
                    continue
                winreg.DeleteValue(k, old)
                winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, login_command())
                log.info("Moved the startup entry to the new name (%s)", APP_NAME)
            current, _ = winreg.QueryValueEx(k, APP_NAME)
            wanted = login_command()
            if current != wanted:
                winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, wanted)
                log.info("Updated the startup entry to this location: %s", wanted)
    except Exception:
        pass   # no startup entry (or not Windows): nothing to fix


def startup_enabled() -> bool:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, APP_NAME)
            return True
    except Exception:
        return False


def show_error(text: str) -> None:
    """Pop-up error box (there's no console window to print to)."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, text, APP_NAME, 0x10)
    except Exception:
        pass


def run_checks(settings: Settings, discord: Discord, bot_id: str, wake, stop) -> None:
    """Check for orders now, then every few minutes, until `stop` is set. `wake` forces an early check."""
    while not stop.is_set():
        on_status("Checking for new orders…")
        try:
            check_once(settings, discord, bot_id)
        except SystemExit as e:
            log.error(str(e))
            on_status(f"Error — see the log")
            show_error(str(e))
            stop.set()
            break
        except Exception as e:
            log.error("Check failed (will retry): %s", e)
        on_status(f"Watching for orders — last checked {time.strftime('%I:%M %p').lstrip('0')}")
        trim_memory()
        wake.wait(settings.check_every_minutes * 60)
        wake.clear()


def tray_icon_class(pystray):
    """pystray's icon, but a left-click opens the same menu as a right-click (Windows)."""
    try:
        from pystray._util import win32 as w
    except Exception:
        return pystray.Icon

    class FleemIcon(pystray.Icon):
        def _on_notify(self, wparam, lparam):
            if lparam == w.WM_LBUTTONUP:
                lparam = w.WM_RBUTTONUP
            return super()._on_notify(wparam, lparam)

    return FleemIcon


def save_setting(key: str, value: str, path: Path | None = None) -> None:
    """Change one line of settings.ini, leaving everything else in it (comments, the Discord token)
    exactly as it is."""
    path = path or SETTINGS_FILE
    raw = path.read_bytes().decode("utf-8") if path.exists() else "[helper]\n"
    nl = "\r\n" if "\r\n" in raw else "\n"
    line = f"{key} = {value}"
    pattern = re.compile(rf"^[ \t]*{re.escape(key)}[ \t]*[=:][^\r\n]*", re.M)
    if pattern.search(raw):
        raw = pattern.sub(lambda m: line, raw, count=1)
    else:
        raw = raw.rstrip("\r\n") + nl + line + nl
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_bytes(raw.encode("utf-8"))
    os.replace(tmp, path)


def set_keep_running(settings: "Settings | None", on: bool) -> None:
    """Keep running in the background (icon by the clock) or only while a window is open. Starting with
    Windows only makes sense in the background, so turning this off also turns that off."""
    try:
        save_setting("keep_running", "yes" if on else "no")
    except OSError as e:
        log.warning("Couldn't save keep_running: %s", e)
    if settings is not None:
        settings.keep_running = on
    if not on and startup_enabled():
        try:
            remove_startup()
        except Exception as e:
            log.warning("Couldn't remove the startup entry: %s", e)


def pick_xmls(merge: bool) -> list[Path]:
    """File picker for one order XML, or 2+ to merge. [] if cancelled."""
    import tkinter as tk
    from tkinter import filedialog, messagebox
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    types = [("MPCFill order XML", "*.xml"), ("All files", "*.*")]
    try:
        if merge:
            chosen = filedialog.askopenfilenames(
                parent=root, filetypes=types,
                title=f"{APP_NAME} — pick 2 or more XMLs to merge (hold Ctrl to pick several)")
            paths = [Path(c) for c in chosen]
            if len(paths) == 1:
                messagebox.showinfo(APP_NAME, "Pick at least 2 XMLs to merge — hold Ctrl and click each one.\n"
                                              "(For a single XML, use Make PDF from XML… instead.)", parent=root)
                return []
            return paths
        chosen = filedialog.askopenfilename(parent=root, title=f"{APP_NAME} — pick an MPCFill XML", filetypes=types)
        return [Path(chosen)] if chosen else []
    finally:
        root.destroy()


def pick_files_to_print(settings: "Settings") -> list[str]:
    """File picker for PDFs / order XMLs to print. [] if cancelled."""
    from tkinter import Tk, filedialog
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return list(filedialog.askopenfilenames(
            parent=root, title=f"{APP_NAME} — pick PDFs or order XMLs to print (hold Ctrl to pick several)",
            initialdir=str(settings.output_folder),
            filetypes=[("PDF or order XML", "*.pdf *.xml"), ("All files", "*.*")]))
    finally:
        root.destroy()


def print_files(settings: "Settings", files) -> int:
    """Open the print window for each PDF / XML in turn. Returns how many were sent to the printer."""
    files = [Path(f) for f in files]
    xmls = [f for f in files if f.suffix.lower() == ".xml"]
    pdfs = [f for f in files if f.suffix.lower() == ".pdf"]
    return (print_xmls(xmls, settings) if xmls else 0) + (print_pdfs(settings, pdfs) if pdfs else 0)


def run_with_status(text: str, fn):
    """Like run_with_progress, but the window also shows Fleem's progress messages (downloading card
    images…) as they come. Returns (result, error text or None)."""
    import tkinter as tk
    from tkinter import ttk
    global on_status
    outcome, latest = {}, {"text": ""}
    before = on_status

    def show(message):
        latest["text"] = message
        before(message)

    def job():
        try:
            outcome["result"] = fn()
        except SystemExit as e:
            outcome["error"] = str(e) or "Stopped"
        except BaseException as e:
            outcome["error"] = str(e) or e.__class__.__name__
    on_status = show
    try:
        win = tk.Tk()
        win.withdraw()
        win.title(APP_NAME)
        win.resizable(False, False)
        T(win, bg="surface")
        T(tk.Label(win, text=text, justify="left", anchor="w", padx=20, pady=14, font=("Segoe UI", 10, "bold"),
                   wraplength=420), bg="surface", fg="text").pack(fill="x")
        detail = T(tk.Label(win, text="", justify="left", anchor="w", padx=20, width=58, font=("Segoe UI", 9)),
                   bg="surface", fg="muted")
        detail.pack(fill="x")
        bar = ttk.Progressbar(win, mode="indeterminate", length=420)
        bar.pack(padx=20, pady=(8, 18), fill="x")
        bar.start(12)
        win.protocol("WM_DELETE_WINDOW", lambda: None)
        apply_theme(win)
        win.update_idletasks()
        win.geometry(f"+{(win.winfo_screenwidth() - win.winfo_reqwidth()) // 2}+"
                     f"{max(0, (win.winfo_screenheight() - win.winfo_reqheight()) // 3)}")
        win.deiconify()
        thread = threading.Thread(target=job, daemon=True)
        thread.start()

        def wait():
            detail.configure(text=latest["text"][:90])
            if thread.is_alive():
                win.after(200, wait)
            else:
                bar.stop()
                win.destroy()
        win.after(200, wait)
        win.mainloop()
    finally:
        on_status = before
    return outcome.get("result"), outcome.get("error")


def make_pdf_from_picked(settings: "Settings", merge: bool, with_window: bool = False, failed=None) -> list[str]:
    """Make PDF from XML… / Merge XMLs into one PDF…: pick the XMLs, check the layout, make the PDF and
    open it. `with_window`: show a progress window while it's made (the tray shows progress by its icon).
    `failed(what, error)` reports a failure (default: an error box). Returns the PDFs saved."""
    paths = pick_xmls(merge)
    if not paths:
        return []
    # what the names call for (an add-on's rule), else check the bleed edge / spacing / paper
    layout = (layout_from_name(merged_pdf_name([p.name for p in paths]) + ".pdf", settings)
              or ask_layout([p.name for p in paths], last_layout()))
    if layout is None:
        return []
    remember_layout(layout)
    what = paths[0].name if len(paths) == 1 else f"{len(paths)} merged XMLs"
    on_status(f"Making the PDF for {what}…"[:100])
    log.info("Making one PDF from %s", ", ".join(str(p) for p in paths))
    if with_window:
        saved, error = run_with_status(f"Making the PDF for {what}…", lambda: make_pdf_from_files(settings, paths, layout))
    else:
        try:
            saved, error = make_pdf_from_files(settings, paths, layout), None
        except SystemExit as e:
            saved, error = None, str(e)
        except Exception as e:
            saved, error = None, str(e) or e.__class__.__name__
    if error:
        log.error("%s failed: %s", what, error)
        (failed or (lambda w, e: show_error(f"Couldn't make the PDF for {w}:\n{e}")))(what, error)
        return []
    log.info("%s: ✔ saved %s", what, ", ".join(Path(s).name for s in saved))
    on_pdf_ready(what, saved)
    open_pdfs(settings, saved)
    return saved


# Discord orders waiting for "I'm ready" are saved to disk, so a restart doesn't lose them.
_QUEUE_LOCK = threading.Lock()


def queue_orders(order_no, saved, put) -> None:
    with _QUEUE_LOCK:
        waiting = load_print_queue()
        for pdf in saved:
            if all(e["pdf"] != str(pdf) for e in waiting):
                waiting.append({"order": order_no, "pdf": str(pdf)})
                put(("order", order_no, pdf))
        save_print_queue(waiting)


def forget_order(pdf) -> None:
    with _QUEUE_LOCK:
        save_print_queue([e for e in load_print_queue() if e["pdf"] != str(pdf)])


def requeue_waiting_orders(put) -> None:
    """Discord orders still waiting for "I'm ready" from before a restart."""
    with _QUEUE_LOCK:
        waiting = load_print_queue()
        save_print_queue(waiting)
    for e in waiting:
        put(("order", e["order"], e["pdf"]))
    if waiting:
        log.info("%d order(s) still waiting to print from last time", len(waiting))


def run_print_job(settings: "Settings", job) -> int:
    """A queued ("order", number, pdf) or ("print", [files]) job. Returns how many were printed."""
    if job[0] == "order":
        _, order_no, pdf = job
        try:
            offer_to_print_order(settings, order_no, pdf)
        finally:
            forget_order(pdf)
        return 0
    return print_files(settings, job[1])


def serve_requests(server_sock, stop, put) -> None:
    """Messages from other copies of Fleem: files from right-click > Print with Fleem, or "show"
    (Fleem was started again while running: open its window)."""
    if server_sock is None:
        return
    server_sock.settimeout(1)
    while not stop.is_set():
        try:
            conn, _ = server_sock.accept()
        except socket.timeout:
            continue
        except OSError:
            return
        with conn:
            try:
                conn.settimeout(5)
                data = b""
                while not data.endswith(b"\n") and len(data) < 65536:
                    chunk = conn.recv(4096)
                    if not chunk:
                        break
                    data += chunk
                message = json.loads(data.decode("utf-8"))
                files = [f for f in message.get("print", [])
                         if isinstance(f, str) and Path(f).suffix.lower() in (".pdf", ".xml") and Path(f).exists()]
                conn.sendall(b"ok\n")
                if files and hand_to_open_project(files):
                    log.info("Opened with Fleem, added to the open project: %s", ", ".join(Path(f).name for f in files))
                elif files:
                    log.info("Print with Fleem: %s", ", ".join(Path(f).name for f in files))
                    put(("print", files))
                elif message.get("show"):
                    put(("home",))
            except Exception as e:
                log.info("Ignored a message: %s", e)


def message_running_fleem(message: dict) -> bool:
    """Send a message to the Fleem that's already running. False if none is running."""
    try:
        with socket.create_connection(("127.0.0.1", FLEEM_PORT), timeout=2) as c:
            c.sendall(json.dumps(message).encode() + b"\n")
            c.settimeout(5)
            return c.recv(16).startswith(b"ok")
    except OSError:
        return False


_HOME_POSITION = {}


def home_window(settings: "Settings", ctx) -> tuple | None:
    """Fleem's main window. Returns what to do next — ("make",), ("merge",), ("print",), ("update",
    release) or ("job", queued job) — or None when it's closed. The things it can do on its own
    (settings, opening folders, checking for updates) don't close it.

    ctx: requests (queue of jobs: Print with Fleem files, orders, "home"), release ({"release": newer
    release or None}), status() -> text, check_orders (or None without Discord)."""
    import queue as queue_mod
    import tkinter as tk
    from tkinter import ttk
    result = {"choice": None}
    root = tk.Tk()
    root.withdraw()
    root.title(APP_NAME)
    root.resizable(False, False)
    images = []
    try:
        icon_png = Path(getattr(sys, "_MEIPASS", HERE)) / "fleem_icon.png"
        if icon_png.exists():
            big = tk.PhotoImage(master=root, file=str(icon_png))
            images += [big, big.subsample(4)]
            root.iconphoto(True, big)
    except Exception:
        pass

    def choose(*choice):
        result["choice"] = choice
        close()

    def close():
        try:
            _HOME_POSITION["xy"] = (root.winfo_x(), root.winfo_y())
        except Exception:
            pass
        root.destroy()

    T(root, bg="surface")
    # ---- header
    head = T(tk.Frame(root, padx=22, pady=16), bg="surface")
    head.pack(fill="x")
    if len(images) > 1:
        T(tk.Label(head, image=images[1], bd=0), bg="surface").pack(side="left", padx=(0, 14))
    names = T(tk.Frame(head), bg="surface")
    names.pack(side="left")
    title_lbl = T(tk.Label(names, text=APP_NAME, font=("Segoe UI", 16, "bold"), anchor="w", compound="center"),
                  bg="surface", fg="text")
    title_lbl.pack(anchor="w")
    T(tk.Label(names, text=f"Version {APP_VERSION}" + (f" · {EDITION} build" if EDITION != "public" else ""),
               font=("Segoe UI", 9), anchor="w"), bg="surface",
      fg="muted").pack(anchor="w")
    theme_btn = ttk.Button(head, text="")
    theme_btn.pack(side="right", anchor="n")
    rule = T(tk.Frame(root, height=1), bg="border")
    rule.pack(fill="x")
    # ---- a new version
    banner = T(tk.Frame(root, padx=22, pady=10), bg="accent_soft")
    banner_text = T(tk.Label(banner, text="", font=("Segoe UI", 9, "bold"), anchor="w"), bg="accent_soft", fg="text")
    banner_text.pack(side="left")
    ttk.Button(banner, text="See what's new", style="Accent.TButton",
               command=lambda: ctx.release["release"] and choose("update", ctx.release["release"])).pack(side="right")
    # ---- what Fleem does
    actions = T(tk.Frame(root, padx=22, pady=18), bg="surface")
    actions.pack(fill="x")
    tiles = [("Print PDFs or XMLs…", lambda: choose("print"), "Big.Accent.TButton"),
             ("Make PDF from XML…", lambda: choose("make"), "Big.TButton"),
             ("Merge XMLs into one PDF…", lambda: choose("merge"), "Big.TButton"),
             ("Open PDF folder", lambda: open_folder(), "Big.TButton")]
    for i, (text, command, style) in enumerate(tiles):
        ttk.Button(actions, text=text, command=command, style=style, width=26).grid(
            row=i // 2, column=i % 2, sticky="nsew", padx=(0 if i % 2 == 0 else 6, 0), pady=(0 if i < 2 else 6, 0))
    actions.columnconfigure(0, weight=1)
    actions.columnconfigure(1, weight=1)
    # ---- background
    options = T(tk.Frame(root, padx=22, pady=4), bg="surface")
    options.pack(fill="x", pady=(0, 14))
    keep_var, start_var = tk.BooleanVar(value=settings.keep_running), tk.BooleanVar(value=startup_enabled())
    keep_box = ttk.Checkbutton(options, text="Keep running in the background", variable=keep_var)
    keep_box.grid(row=0, column=0, sticky="w")
    start_box = ttk.Checkbutton(options, text="Start with Windows", variable=start_var)
    start_box.grid(row=1, column=0, sticky="w", padx=(24, 0))
    Tooltip(keep_box, "On: closing this window leaves Fleem by the clock, so Print with Fleem opens instantly"
                      + (" and Discord orders are caught" if ctx.check_orders else "") +
                      ".\nOff: Fleem stops when you close its windows.")
    if ctx.check_orders:
        ttk.Button(options, text="Check for orders now", command=ctx.check_orders).grid(row=0, column=1, rowspan=2,
                                                                                       sticky="e")
        options.columnconfigure(1, weight=1)

    def keep_changed():
        set_keep_running(settings, keep_var.get())
        start_var.set(startup_enabled())
        start_box.configure(state="normal" if keep_var.get() else "disabled")

    def start_changed():
        try:
            install_startup() if start_var.get() else remove_startup()
        except Exception as e:
            log.warning("Couldn't change the startup setting: %s", e)
        start_var.set(startup_enabled())
    keep_box.configure(command=keep_changed)
    start_box.configure(command=start_changed, state="normal" if settings.keep_running else "disabled")
    # ---- footer
    foot = T(tk.Frame(root, padx=22, pady=10), bg="surface2")
    foot.pack(fill="x")
    status = T(tk.Label(foot, text="", font=("Segoe UI", 9), anchor="w", width=34), bg="surface2", fg="muted")
    status.pack(side="left", fill="x", expand=True)
    ttk.Button(foot, text="Close", command=close).pack(side="right")
    checking = {"on": False, "note": ""}
    if updates_configured():
        def check_updates():
            if checking["on"]:
                return
            checking.update(on=True, note="Checking for updates…")

            def job():
                try:
                    release = find_update(force=True)
                    ctx.release["release"] = release
                    checking["note"] = "" if release else f"You have the latest version ({APP_VERSION})."
                except Exception as e:
                    log.warning("Couldn't check for updates: %s", e)
                    checking["note"] = "Couldn't check for updates (are you online?)"
                checking["on"] = False
            threading.Thread(target=job, daemon=True).start()
        ttk.Button(foot, text="Check for updates", command=check_updates).pack(side="right", padx=(0, 8))
    ttk.Button(foot, text="View log", command=lambda: open_file(LOG_FILE)).pack(side="right", padx=(0, 8))

    def open_file(path):
        try:
            os.startfile(str(path))
        except Exception as e:
            log.warning("Couldn't open %s: %s", path, e)

    def open_folder():
        settings.output_folder.mkdir(parents=True, exist_ok=True)
        open_file(settings.output_folder)

    def style():
        apply_theme(root)
        rpg = is_rpg()
        family = rpg_font_family(root) if rpg else "Segoe UI"
        s = ttk.Style(root)
        k = rpg_scale(root)
        big = rpg_scroll_font(root, bold=True) if rpg and rpg_button_family(root) else (family, 11 if rpg else 10, "bold")
        s.configure("Big.TButton", padding=(10 * k if rpg else 16, 14), font=big)
        s.configure("Big.Accent.TButton", padding=(10 * k if rpg else 16, 14), font=big)
        theme_btn.configure(text="Theme ▾")
        if rpg:                                # the name on a ribbon banner, centred on its front
            from tkinter import font as tkfont
            if rpg_button_family(root):
                title_lbl.configure(font=rpg_scroll_font(root, big=True, bold=True))
            measure = tkfont.Font(root, font=title_lbl.cget("font"))
            width = measure.measure(APP_NAME) + 26 * k
            title_lbl.configure(image=rpg_banner_photo(root, width), pady=0, padx=0, fg=RPG_INK_TEXT)
        else:
            title_lbl.configure(image="", pady=0, padx=1, font=("Segoe UI", 16, "bold"))
            title_lbl._fleem_font_spec = None

    theme_btn.configure(command=lambda: theme_menu(theme_btn, lambda: None))    # (poll() reopens it in the theme)

    shown = {"release": None}

    def poll():
        if not root.winfo_exists():
            return
        if theme_name() != made_theme and root.grab_current() is None:
            choose("retheme")                    # reopen in the new theme
            return
        release = ctx.release.get("release")
        if release is not shown["release"]:                     # show / hide the new version banner
            shown["release"] = release
            if release:
                banner_text.configure(text=f"Version {release['version']} is available.")
                banner.pack(fill="x", after=rule)
            else:
                banner.pack_forget()
        status.configure(text=(checking["note"] or ctx.status() or "")[:60])
        if ctx.requests is not None:
            try:
                job = ctx.requests.get_nowait()
            except queue_mod.Empty:
                job = None
            if job is not None:
                if job[0] == "home":                             # started again: come to the front
                    root.deiconify()
                    root.attributes("-topmost", True)
                    root.lift()
                    root.focus_force()
                    root.after(300, lambda: root.attributes("-topmost", False))
                else:
                    choose("job", job)
                    return
        root.after(300, poll)

    root.protocol("WM_DELETE_WINDOW", close)
    root.bind("<Escape>", lambda e: close())
    made_theme = theme_name()
    style()
    root.update_idletasks()
    if "xy" in _HOME_POSITION:
        x, y = _HOME_POSITION["xy"]
    else:
        x = (root.winfo_screenwidth() - root.winfo_reqwidth()) // 2
        y = max(0, (root.winfo_screenheight() - root.winfo_reqheight()) // 3)
    root.geometry(f"+{x}+{y}")
    root.deiconify()
    root.lift()
    root.focus_force()
    root.after(50, poll)
    root.mainloop()
    return result["choice"]


def home_session(settings: "Settings", ctx) -> str:
    """Show Fleem's window; whatever you pick runs, then the window comes back. Returns "closed" or
    "restart" (an update was installed)."""
    while True:
        choice = home_window(settings, ctx)
        if choice is None:
            return "closed"
        try:
            kind = choice[0]
            if kind == "retheme":
                continue
            if kind in ("make", "merge"):
                make_pdf_from_picked(settings, merge=kind == "merge", with_window=True)
            elif kind == "print":
                files = pick_files_to_print(settings)
                if files:
                    print_files(settings, files)
            elif kind == "job":
                if ctx.handle_job(choice[1]) == "restart":
                    return "restart"
            elif kind == "update":
                answer = update_window(choice[1])
                if answer == "skip":
                    save_update_state({**load_update_state(), "skipped": choice[1]["version"]})
                    ctx.release["release"] = None
                elif answer == "installed":
                    return "restart"
        except Exception as e:
            log.error("%s failed: %s", choice[0], e)
            show_error(f"Something went wrong: {e}")
        finally:
            trim_memory()


def run_window(settings: Settings, discord: "Discord | None", bot_id: str | None, server_sock=None) -> str | None:
    """Run only while Fleem's window (or one it opened) is open. Returns "tray" if you turned on Keep
    running in the background, else None (quit)."""
    import queue
    from types import SimpleNamespace
    global on_status, on_order_ready
    wake, stop = threading.Event(), threading.Event()
    requests = queue.Queue()
    status = {"text": "Ready"}
    on_status = lambda text: status.update(text=text)        # noqa: E731
    on_order_ready = lambda order_no, saved: queue_orders(order_no, saved, requests.put)   # noqa: E731
    threading.Thread(target=serve_requests, args=(server_sock, stop, requests.put), daemon=True).start()
    if discord is not None:
        requeue_waiting_orders(requests.put)
        threading.Thread(target=run_checks, args=(settings, discord, bot_id, wake, stop), daemon=True).start()
    release = {"release": None}
    if settings.check_for_updates and updates_configured():
        def check():
            try:
                release["release"] = find_update()
            except Exception as e:
                log.info("Couldn't check for updates: %s", e)
        threading.Thread(target=check, daemon=True).start()

    threading.Thread(target=watch_source, args=(stop, requests.put, lambda text: status.update(text=text)),
                     daemon=True).start()
    if "--rebuilt" in sys.argv:
        status["text"] = "Rebuilt with your latest changes"

    def handle_job(job):
        if job[0] == "rebuilt":
            install_rebuilt(job[1])
            return "restart"
        run_print_job(settings, job)
        return None
    ctx = SimpleNamespace(requests=requests, release=release, status=lambda: status["text"], handle_job=handle_job,
                          check_orders=wake.set if discord is not None else None)
    outcome = home_session(settings, ctx)
    stop.set()
    wake.set()
    if RENDER_LOCK.locked():          # an order's PDF is being made: let it finish
        run_with_progress("Finishing a PDF before closing…", lambda: (RENDER_LOCK.acquire(), RENDER_LOCK.release()))
    on_status, on_order_ready = _no_events, None
    if outcome == "restart":
        try:
            if server_sock is not None:
                server_sock.close()
            start_new_copy(Path(sys.executable))
        except Exception as e:
            log.error("Couldn't start the new version: %s", e)
            show_error(f"The update is installed, but {APP_NAME} couldn't restart itself ({e}).\n"
                       f"Start {APP_NAME} again to use the new version.")
        return None
    clean_work_folder()
    return "tray" if settings.keep_running else None


def run_tray(settings: Settings, discord: "Discord | None", bot_id: str | None, server_sock=None) -> bool:
    """Run in the background with Fleem's icon in the taskbar corner. Returns False if the tray library
    isn't installed."""
    try:
        import pystray
        from PIL import Image
    except ImportError:
        return False
    import queue
    from types import SimpleNamespace

    global on_status, on_pdf_ready, on_order_ready, on_warning
    wake, stop = threading.Event(), threading.Event()
    icon_file = Path(getattr(sys, "_MEIPASS", HERE)) / "fleem_icon.png"
    image = Image.open(icon_file) if icon_file.exists() else Image.new("RGB", (64, 64), (196, 90, 60))
    # One window at a time: Fleem's window, Discord orders waiting for "I'm ready" and Print with Fleem
    # requests share a queue, handled in order without holding up the order checks.
    jobs = queue.Queue()
    status = {"text": ""}

    def open_path(path: Path):
        try:
            os.startfile(str(path))   # Windows
        except Exception as e:
            log.warning("Couldn't open %s: %s", path, e)

    def idle_status():
        return "Watching for orders" if discord is not None else "Ready"

    def open_home(icon=None, item=None):
        jobs.put(("home",))

    def check_now(icon, item):
        wake.set()

    def pick_xml(icon, item):
        threading.Thread(target=picked_files_job, args=(False,), daemon=True).start()

    def pick_xmls_to_merge(icon, item):
        threading.Thread(target=picked_files_job, args=(True,), daemon=True).start()

    def pick_pdfs_to_print(icon, item):
        threading.Thread(target=print_picked_pdfs, daemon=True).start()

    def print_picked_pdfs():
        try:
            chosen = pick_files_to_print(settings)
            if chosen:
                jobs.put(("print", chosen))     # opens in turn, one window at a time
        except Exception as e:
            log.error("Printing failed: %s", e)
            show_error(f"Printing failed: {e}")
        finally:
            trim_memory()

    def picked_files_job(merge: bool):
        def failed(what, error):
            try:
                icon.notify(error[:200], f"Couldn't make the PDF for {what}"[:63])
            except Exception:
                pass
        try:
            saved = make_pdf_from_picked(settings, merge, failed=failed)
            on_status(f"{idle_status()} — made {Path(saved[0]).name}"[:100] if saved else idle_status())
        except Exception as e:
            log.error("Making the PDF failed: %s", e)
            show_error(f"Making the PDF failed: {e}")
        finally:
            trim_memory()

    def open_folder(icon, item):
        settings.output_folder.mkdir(parents=True, exist_ok=True)
        open_path(settings.output_folder)

    def open_log(icon, item):
        open_path(LOG_FILE)

    def toggle_startup(icon, item):
        try:
            remove_startup() if startup_enabled() else install_startup()
        except Exception as e:
            show_error(f"Couldn't change the startup setting: {e}")

    def toggle_keep_running(icon, item):
        set_keep_running(settings, not settings.keep_running)
        if not settings.keep_running:     # Fleem now stops when its window is closed: show the window
            open_home()

    def appearance_item(label, value):
        return pystray.MenuItem(label, lambda icon, item: set_appearance(value), radio=True,
                                checked=lambda item: appearance_setting() == value)

    def quit_app(icon, item):
        stop.set()
        wake.set()
        icon.stop()

    # ---- updates: a new release on GitHub shows "Update to vX" in the menu
    pending = {"release": None, "told": None}

    def update_label(item):
        return f"Update to version {pending['release']['version']}…" if pending["release"] else "Update…"

    def open_update(icon, item):
        if pending["release"]:
            jobs.put(("update", pending["release"]))

    def check_updates_now(icon, item):
        threading.Thread(target=check_updates, args=(True,), daemon=True).start()

    def check_updates(asked: bool):
        try:
            release = find_update(force=asked)
        except Exception as e:
            log.warning("Couldn't check for updates: %s", e)
            if asked:
                warn(f"Couldn't check for updates: {e}")
            return
        pending["release"] = release
        try:
            icon.update_menu()
        except Exception:
            pass
        if release is None:
            if asked:
                icon.notify(f"You have the latest version ({APP_VERSION}).", "No updates")
            return
        log.info("Version %s is available (this is %s)", release["version"], APP_VERSION)
        if asked:
            jobs.put(("update", release))
        elif pending.get("told") != release["version"]:      # mention each version once
            pending["told"] = release["version"]
            icon.notify(f"Version {release['version']} is available. Click Fleem's head > "
                        f"Update to version {release['version']}.", "Update available")

    def update_checks():
        if "--updated" in sys.argv or "--rebuilt" in sys.argv:
            stop.wait(3)
            try:
                icon.notify(f"Updated to version {APP_VERSION}." if "--updated" in sys.argv else
                            f"Rebuilt with your latest changes (version {APP_VERSION}).", APP_NAME)
            except Exception:
                pass
        if not (settings.check_for_updates and updates_configured()):
            return
        stop.wait(20)                       # let Fleem settle in first
        while not stop.is_set():
            check_updates(False)
            stop.wait(UPDATE_CHECK_EVERY)

    def restart_into_update():
        """The new version is in place: hand over to it."""
        RENDER_LOCK.acquire(timeout=15 * 60)     # let a PDF that's being made finish first
        try:
            if server_sock is not None:
                server_sock.close()          # the new copy takes over the port
        except OSError:
            pass
        try:
            start_new_copy(Path(sys.executable))
        except Exception as e:
            log.error("Couldn't start the new version: %s", e)
            show_error(f"The update is installed, but {APP_NAME} couldn't restart itself ({e}).\n"
                       f"Start {APP_NAME} again to use the new version.")
        quit_app(icon, None)

    icon = tray_icon_class(pystray)("FleemProxyCompanion", image, APP_NAME, menu=pystray.Menu(
        pystray.MenuItem(f"Open {APP_NAME}", open_home, default=True),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open PDF folder", open_folder),
        pystray.MenuItem("Make PDF from XML…", pick_xml),
        pystray.MenuItem("Merge XMLs into one PDF…", pick_xmls_to_merge),
        pystray.MenuItem("Print PDFs or XMLs…", pick_pdfs_to_print),
        pystray.MenuItem("Check for orders now", check_now, visible=discord is not None),
        pystray.MenuItem("View log", open_log),
        pystray.MenuItem(update_label, open_update, visible=lambda item: pending["release"] is not None),
        pystray.MenuItem("Check for updates", check_updates_now, visible=updates_configured()),
        pystray.MenuItem("Appearance", pystray.Menu(*(appearance_item(label, value)
                                                      for value, label in THEME_LABELS.items()))),
        pystray.MenuItem("Keep running in the background", toggle_keep_running,
                         checked=lambda item: settings.keep_running),
        pystray.MenuItem("Start with Windows", toggle_startup, checked=lambda item: startup_enabled()),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(f"Quit {APP_NAME}", quit_app),
    ))

    def set_status(text):
        status["text"] = text
        icon.title = f"{APP_NAME}\n{text}"[:127]   # tooltip (Windows allows 127 characters)

    def warn(text):
        log.warning(text)
        try:
            icon.notify(text[:250], APP_NAME)
        except Exception:
            pass
    on_warning = warn

    def pdf_ready(source, saved):
        names = ", ".join(Path(p).stem for p in saved)
        where = f"Order #{source}" if isinstance(source, int) else "From your XML"
        try:
            icon.notify(f"{where}: {names}", "PDF ready")
        except Exception:
            pass

    def handle_job(job):
        """One queued job. Returns "restart" when an update was installed."""
        if job[0] == "home":
            ctx = SimpleNamespace(requests=jobs, release=pending, status=lambda: status["text"],
                                  handle_job=handle_job, check_orders=wake.set if discord is not None else None)
            outcome = home_session(settings, ctx)
            if outcome == "restart":
                return "restart"
            if not settings.keep_running:       # turned off: Fleem stops with its window
                quit_app(icon, None)
            try:
                icon.update_menu()
            except Exception:
                pass
        elif job[0] == "rebuilt":
            install_rebuilt(job[1])
            return "restart"
        elif job[0] == "update":
            choice = update_window(job[1])
            if choice == "skip":
                save_update_state({**load_update_state(), "skipped": job[1]["version"]})
                pending["release"] = None
                icon.update_menu()
            elif choice == "installed":
                return "restart"
        else:
            sent = run_print_job(settings, job)
            if sent:
                icon.notify(f"Sent {sent} PDF{'s' if sent > 1 else ''} to the printer", "Printing")
        return None

    def gui_worker():
        while not stop.is_set():
            try:
                job = jobs.get(timeout=5)
            except queue.Empty:
                continue
            try:
                if handle_job(job) == "restart":
                    restart_into_update()
                    return
            except Exception as e:
                log.error("Printing failed: %s", e)
                show_error(f"Printing failed: {e}")
            finally:
                trim_memory()

    on_status, on_pdf_ready = set_status, pdf_ready
    on_order_ready = lambda order_no, saved: queue_orders(order_no, saved, jobs.put)   # noqa: E731

    def setup(icon):
        icon.visible = True
        if discord is not None:
            threading.Thread(target=run_checks, args=(settings, discord, bot_id, wake, stop), daemon=True).start()
        else:
            on_status("Ready — right-click a PDF or XML > Print with Fleem")
        threading.Thread(target=gui_worker, daemon=True).start()
        threading.Thread(target=serve_requests, args=(server_sock, stop, jobs.put), daemon=True).start()
        threading.Thread(target=update_checks, daemon=True).start()
        threading.Thread(target=watch_source, args=(stop, jobs.put, warn), daemon=True).start()
        if discord is not None:
            requeue_waiting_orders(jobs.put)

    icon.run(setup=setup)
    stop.set()
    clean_work_folder()                      # temporary files and the image cache go with Fleem
    return True


# ============================================================ first run
WELCOME_FILE = HERE / "welcome.json"     # there once the welcome message has been shown
WELCOME_TEXT = """Hi, thanks for trying Fleem. This is a personal application I designed that I decided other proxy makers may want to try it. For me, it is built around not needing to access many windows and has integrations with a discord bot (unreleased yet) to streamline the process of receiving an XML and turning it into a PDF. It is designed as a PDF Reader and Proxy Maker all in one.

Here's what he can do so far:

* Adds a right click option in Windows on all XMLs and PDFs to "Print with Fleem", which opens up the project window.

* Right Click an MPCFill XML (Also works for PDFs!), go to "Open with" Fleem (and set to "Always" if you want) and he will now automatically convert all of your XMLs into proxy projects.

* Automatically determine individual cards in a PDF and extract them all so they can be manipulated and cards can be added/removed.
* Merge any amount of PDFs/XMLs/Images of Cards into one cohesive sheet by dragging them into the project window.
* Will allow you to sort by fronts and backs and you can select whichever pages you want to print.
* Automatically print using your printers saved presets.

* Automatically detect MDFCs and place them all together, and then alphabetize the rest so you have to double print on as few sheets as possible.
* Automatically handles its own cache, deleting all MPC images every time it is restarted, but keeps them until then so if you quickly redo a project it is fast.
* And more!

This program was designed with constant running in mind but has an option to disable that. This has only been tested thoroughly on my own personal PC with an Epson 8550, so please let me know your feedback! Thank you proxxiers! :)"""


def show_welcome_once() -> None:
    """The welcome message, the very first time Fleem is opened (never again)."""
    if WELCOME_FILE.exists():
        return
    try:
        WELCOME_FILE.write_text(json.dumps({"shown": APP_VERSION}), encoding="utf-8")
    except OSError as e:
        log.warning("Couldn't save %s: %s", WELCOME_FILE.name, e)
        return                      # can't remember it: better not to show it every time
    try:
        welcome_window()
    except Exception as e:
        log.warning("Couldn't show the welcome message: %s", e)


def welcome_window() -> None:
    import tkinter as tk
    from tkinter import ttk
    root = tk.Tk()
    root.withdraw()
    root.title(f"Welcome to {APP_NAME}")
    root.resizable(False, False)
    images = []
    try:
        icon_png = Path(getattr(sys, "_MEIPASS", HERE)) / "fleem_icon.png"
        if icon_png.exists():
            big = tk.PhotoImage(master=root, file=str(icon_png))
            images += [big, big.subsample(4)]
            root.iconphoto(True, big)
    except Exception:
        pass
    T(root, bg="surface")
    head = T(tk.Frame(root, padx=24, pady=16), bg="surface")
    head.pack(fill="x")
    if len(images) > 1:
        T(tk.Label(head, image=images[1], bd=0), bg="surface").pack(side="left", padx=(0, 14))
    T(tk.Label(head, text=f"Welcome to {APP_NAME}", font=("Segoe UI", 16, "bold"), anchor="w"),
      bg="surface", fg="text").pack(side="left")
    T(tk.Frame(root, height=1), bg="border").pack(fill="x")
    box = T(tk.Frame(root, padx=24, pady=14), bg="surface")
    box.pack(fill="both", expand=True)
    text = T(tk.Text(box, width=78, height=24, wrap="word", relief="flat", borderwidth=0, padx=4, pady=2,
                     font=("Segoe UI", 10), spacing1=1, spacing3=1, cursor="arrow", highlightthickness=0),
             bg="surface", fg="text", insertbackground="surface")
    bar = ttk.Scrollbar(box, orient="vertical", command=text.yview)
    text.configure(yscrollcommand=bar.set)
    text.tag_configure("bullet", lmargin1=10, lmargin2=28, spacing1=3)
    text.tag_configure("gap", font=("Segoe UI", 4))
    for line in WELCOME_TEXT.split("\n"):
        if line.startswith("* "):
            text.insert("end", "•  " + line[2:] + "\n", "bullet")
        elif line.strip():
            text.insert("end", line + "\n")
        else:
            text.insert("end", "\n", "gap")
    text.configure(state="disabled")
    text.pack(side="left", fill="both", expand=True)
    bar.pack(side="right", fill="y")
    foot = T(tk.Frame(root, padx=24, pady=12), bg="surface2")
    foot.pack(fill="x")
    go = ttk.Button(foot, text="Get started", style="Accent.TButton", command=root.destroy)
    go.pack(side="right")
    root.bind("<Return>", lambda e: root.destroy())
    root.bind("<Escape>", lambda e: root.destroy())
    apply_theme(root)
    # tall enough for the whole message on most screens (the scroll bar covers the rest)
    text.configure(height=max(12, min(30, (root.winfo_screenheight() - 240) // 21)))
    root.update_idletasks()
    root.geometry(f"+{(root.winfo_screenwidth() - root.winfo_reqwidth()) // 2}+"
                  f"{max(0, (root.winfo_screenheight() - root.winfo_reqheight()) // 3)}")
    root.deiconify()
    root.lift()
    root.focus_force()
    go.focus_set()
    root.mainloop()


def main() -> None:
    if "--version" in sys.argv:
        print(APP_VERSION)
        return
    if "--write-version-info" in sys.argv:
        write_version_info()
        return
    if "--write-build-fingerprint" in sys.argv:
        write_build_fingerprint()
        return
    if "--install-startup" in sys.argv:
        install_startup()
        return
    if "--remove-startup" in sys.argv:
        remove_startup()
        return
    load_private_fonts()                 # the "fonts" folder (e.g. a pixel font for the RPG scrolls)
    pdf_args = [a for a in sys.argv[1:] if a.lower().endswith(".pdf")]
    xml_args = [Path(a) for a in sys.argv[1:] if a.lower().endswith(".xml") and Path(a).exists()]
    if "--print" in sys.argv or pdf_args or xml_args:      # right-click > Print with Fleem, or Open with
        setup_logging(to_console=False)
        files = [str(x) for x in xml_args] + pdf_args
        if files and hand_to_running_fleem(files):          # the running Fleem opens it
            return
        show_welcome_once()                                  # the very first time Fleem is opened
        # While this print window is open, more PDFs/XMLs opened with Fleem come here (and are added to it).
        try:
            lock = single_instance_lock()
        except SystemExit:
            lock = None
        stop, later = threading.Event(), []
        threading.Thread(target=serve_requests, args=(lock, stop, later.append), daemon=True).start()
        try:
            settings = Settings(SETTINGS_FILE, check=False)
            if xml_args:      # order XMLs: cards in the print window, PDF made only on Print
                print_xmls(xml_args, settings)
            if pdf_args:
                print_pdfs(settings, pdf_args)
            while any(job[0] == "print" for job in later):    # opened after the window had closed
                job = later.pop(0)
                if job[0] == "print":
                    print_files(settings, job[1])
        except SystemExit as e:
            show_error(str(e))
        except Exception as e:
            log.error("Printing failed: %s", e)
            show_error(f"Printing failed: {e}")
        finally:
            stop.set()
            if lock is not None:
                lock.close()
        if any(job[0] == "home" for job in later):              # Fleem was started meanwhile: start it now
            command = [sys.executable] if getattr(sys, "frozen", False) else [sys.executable, str(Path(__file__).resolve())]
            subprocess.Popen(command, cwd=str(HERE), env=child_env(), close_fds=True)
        return
    once = "--once" in sys.argv
    setup_logging(to_console=once)
    try:
        # kept open while running (also where Print with Fleem hands over files); after an update, wait
        # for the old copy to close
        restarted = any(f in sys.argv for f in ("--updated", "--rebuilt", "--renamed"))   # by the copy this one replaced
        if not once and not restarted and message_running_fleem({"show": True} if "--startup" not in sys.argv else {}):
            return                     # already running: it brings its window to the front
        lock = single_instance_lock(wait=30 if restarted else 0)
        settings = Settings(SETTINGS_FILE)
        if "--startup" in sys.argv and not settings.keep_running:
            log.info("Started at login, but keep_running is off: not staying in the background")
            return
        if adopt_new_name():           # still "Fleem PDF Maker.exe": the copy under the new name takes over
            return
        clean_up_after_update()        # the copy an update replaced
        clean_work_folder()            # leftovers from last time
        if not once:
            show_welcome_once()        # the very first time Fleem is opened
        discord = bot_id = None
        if settings.watch_discord:
            discord = Discord(settings.token)
            bot_id = discord.get("/users/@me")["id"]
        log.info("%s started — PDFs go to %s", APP_NAME, settings.output_folder)
        fix_startup_location()
        register_print_menu()
        if once:
            if discord:
                check_once(settings, discord, bot_id)
            else:
                print(f"Discord order watching is off (no discord_token in {SETTINGS_FILE.name}).")
            return
        # in the background with an icon by the clock, or only while Fleem's window is open
        background = settings.keep_running
        while True:
            if background and run_tray(settings, discord, bot_id, server_sock=lock) is not False:
                break                  # quit from the tray
            if background:
                log.warning("The tray library isn't installed: running only while the window is open")
            if run_window(settings, discord, bot_id, server_sock=lock) != "tray" or background:
                break                  # closed (or no tray library to go to)
            background = True          # Keep running in the background was turned on
    except SystemExit as e:
        if str(e):
            log.error(str(e))
            print(e) if once else show_error(str(e))
        sys.exit(1)


# ============================================================ add-on
try:
    import fleem_personal          # your own tools, never published (see the add-on hooks near the top)
except ModuleNotFoundError as e:
    if e.name != "fleem_personal":
        raise
    fleem_personal = None
if fleem_personal is not None:
    fleem_personal.setup(sys.modules[__name__])


if __name__ == "__main__":
    main()
