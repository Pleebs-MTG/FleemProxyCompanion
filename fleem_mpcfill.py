"""Fleem Proxy Companion - MPCFill card search (an optional add-on file).

"Find cards…" in the print window: type a card's name (or paste a whole decklist), see the artwork
MPCFill has for it, pick the front and back, choose how many copies, and add them to the project. The
picked artwork is downloaded the same way as an order XML's cards (same cache, bleed, retries).

How it's built, so it's easy to change or remove:
- Everything that talks to MPCFill is in the MPCFillAPI class below: the web address, the requests
  and how the answers are read. If MPCFill changes its API, that class is the only part to edit
  (the window only uses its methods and the plain CardArt records they return).
- MPCFill's API isn't an official, documented one: it's what the mpcfill.com website itself uses
  (github.com/chilli-axe/mpc-autofill, routes "2/sources/", "3/editorSearch/" (or the older
  "2/editorSearch/", like the website itself falls back to), "2/cards/", "2/cardbacks/" and
  "2/DFCPairs/"). Fleem asks politely: searches only when you press Search,
  caches answers, and loads at most a few thumbnails at a time.
- To remove the feature, delete this file (and its --hidden-import in build.bat): fleem_pdf_maker.py
  only shows the "Find cards…" button when this file can be imported.

The window talks to Fleem through two things only: `core` (fleem_pdf_maker, for downloads, theming and
logging) and `on_add(groups)`, which receives card groups in the same form an order XML gives
({"front_id", "back_id", "name", "query", "count"}).
"""
import json
import re
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass

MAX_QUERIES_PER_SEARCH = 300          # MPCFill's limit for one editorSearch request
MAX_CARDS_PER_REQUEST = 1000          # ... and for one cards request
RESULTS_SHOWN = 60                    # artworks shown for one card name


# ============================================================ the API (edit here if MPCFill changes)
class MPCFillError(RuntimeError):
    """MPCFill answered with an error (code: the HTTP status, 0 if it couldn't be reached)."""
    def __init__(self, message: str, code: int = 0):
        super().__init__(message)
        self.code = code



@dataclass
class CardArt:
    """One piece of artwork on MPCFill."""
    identifier: str                   # MPCFill's id; also the image's id for downloading
    name: str                         # the artwork's name, e.g. "Lightning Bolt (Kev Walker).png"
    source: str                       # who uploaded it
    dpi: int
    thumbnail: str                    # a small preview image's address ("" if none)


class MPCFillAPI:
    """The few calls Fleem makes to MPCFill. Every method returns plain Python data (CardArt, lists,
    dicts), so the window doesn't depend on how MPCFill's answers look."""
    BASE = "https://mpcfill.com"
    TIMEOUT = 30

    def __init__(self, user_agent: str = "FleemProxyCompanion"):
        self.user_agent = user_agent
        self._sources = None
        self._dfc = None
        self._cards = {}                         # identifier -> CardArt (cache)
        self._searches = {}                      # (query, type) -> [identifier] (cache)
        self._legacy_search = False              # True: this MPCFill only has the older 2/editorSearch/
        self._lock = threading.Lock()

    # --- plumbing
    def _call(self, path: str, body: dict | None = None) -> dict:
        data = None if body is None else json.dumps(body).encode("utf-8")
        req = urllib.request.Request(f"{self.BASE}/{path}", data=data, method="GET" if body is None else "POST",
                                     headers={"User-Agent": self.user_agent, "Accept": "application/json",
                                              "Content-Type": "application/json"})
        last = None
        for attempt in range(3):
            try:
                with urllib.request.urlopen(req, timeout=self.TIMEOUT) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except urllib.error.HTTPError as e:
                if e.code < 500 and e.code != 429:
                    detail = ""
                    try:                                  # MPCFill explains errors as {"name", "message"}
                        body = json.loads(e.read().decode("utf-8", "replace"))
                        detail = str(body.get("message") or body.get("name") or "")[:300]
                    except Exception:
                        pass
                    raise MPCFillError(f"MPCFill said no ({e.code}{': ' + detail if detail else ''}) "
                                       f"for {path}. Its API may have changed.", e.code) from None
                last = e
            except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
                last = e
            time.sleep(1.5 * (attempt + 1))
        raise MPCFillError(f"Couldn't reach MPCFill ({last})")

    def _search_settings(self) -> dict:
        """Every source, in MPCFill's own order; no DPI, size, language or tag limits."""
        return {"searchTypeSettings": {"fuzzySearch": False, "filterCardbacks": False},
                "sourceSettings": {"sources": [[pk, True] for pk in self.sources()]},
                "filterSettings": {"minimumDPI": 0, "maximumDPI": 1500, "maximumSize": 30,
                                   "languages": [], "includesTags": [], "excludesTags": []}}

    @staticmethod
    def clean(name: str) -> str:
        """A card name the way MPCFill searches it: lower case, no punctuation."""
        name = re.sub(r"[^\w\s/]", "", name.lower().replace("’", "'").replace("'", ""))
        return re.sub(r"\s+", " ", name).strip()

    # --- the calls
    def sources(self) -> list[int]:
        if self._sources is None:
            results = self._call("2/sources/").get("results", {})
            order = sorted(results.values(), key=lambda s: (s.get("ordinal", 0), s.get("pk", 0)))
            self._sources = [int(s["pk"]) for s in order if "pk" in s] or [int(k) for k in results]
        return self._sources

    def search(self, names: list[str], card_type: str = "CARD") -> dict[str, list[str]]:
        """Artwork ids for each name (best first). card_type: "CARD", "TOKEN" or "CARDBACK"."""
        out, todo = {}, []
        for n in names:
            key = (self.clean(n), card_type)
            if key in self._searches:
                out[n] = self._searches[key]
            elif key[0]:
                todo.append(n)
        for i in range(0, len(todo), MAX_QUERIES_PER_SEARCH):
            chunk = todo[i:i + MAX_QUERIES_PER_SEARCH]
            found = self._search_chunk([self.clean(n) for n in chunk], card_type)
            for n in chunk:
                ids = found.get(self.clean(n), [])
                self._searches[(self.clean(n), card_type)] = ids
                out[n] = ids
        return out

    def _search_chunk(self, queries: list[str], card_type: str) -> dict[str, list[str]]:
        """{cleaned query: [ids]}. Newer MPCFill: "3/editorSearch/" takes {key: query} and answers
        {key: [ids]}. Older: "2/editorSearch/" takes [query] and answers {query: {type: [ids]}}. Like the
        website, try the newer one and fall back to the older one."""
        settings = self._search_settings()
        if not self._legacy_search:
            try:
                body = {"queries": {str(j): {"query": q, "cardType": card_type} for j, q in enumerate(queries)},
                        "searchSettings": settings}
                results = self._call("3/editorSearch/", body).get("results", {})
                return {q: [str(x) for x in results.get(str(j), [])] for j, q in enumerate(queries)}
            except MPCFillError as e:
                if e.code not in (400, 404, 405):
                    raise
                self._legacy_search = True       # an older MPCFill: use 2/editorSearch/ from now on
        body = {"queries": [{"query": q, "cardType": card_type} for q in queries], "searchSettings": settings}
        results = self._call("2/editorSearch/", body).get("results", {})
        out = {}
        for q in queries:
            hits = results.get(q, {})
            hits = hits.get(card_type, []) if isinstance(hits, dict) else hits
            out[q] = [str(x) for x in hits]
        return out

    def cards(self, identifiers: list[str]) -> dict[str, CardArt]:
        """Details (name, source, DPI, thumbnail) for artwork ids."""
        todo = [i for i in dict.fromkeys(identifiers) if i not in self._cards]
        for k in range(0, len(todo), MAX_CARDS_PER_REQUEST):
            results = self._call("2/cards/", {"cardIdentifiers": todo[k:k + MAX_CARDS_PER_REQUEST]}).get("results", {})
            for ident, c in results.items():
                self._cards[ident] = CardArt(
                    identifier=str(c.get("identifier") or ident), name=str(c.get("name") or "card"),
                    source=str(c.get("sourceName") or c.get("sourceVerbose") or c.get("source") or ""),
                    dpi=int(c.get("dpi") or 0),
                    thumbnail=str(c.get("smallThumbnailUrl") or c.get("mediumThumbnailUrl") or ""))
        return {i: self._cards[i] for i in identifiers if i in self._cards}

    def cardbacks(self) -> list[str]:
        """Artwork ids of card backs."""
        return [str(x) for x in self._call("2/cardbacks/", {"searchSettings": self._search_settings()})
                .get("cardbacks", [])]

    def dfc_back_name(self, front_name: str) -> str | None:
        """The back face's name of a double-faced card, if MPCFill knows the card."""
        if self._dfc is None:
            pairs = self._call("2/DFCPairs/").get("dfcPairs", {})
            self._dfc = {self.clean(f): b for f, b in pairs.items()}
        return self._dfc.get(self.clean(front_name))


# ============================================================ decklists
_SECTION = re.compile(r"^(deck|main ?board|side ?board|sideboard|commander|companion|maybe ?board|tokens?|"
                      r"lands?|creatures?|spells?|artifacts?|enchantments?|instants?|sorcer(y|ies)|planeswalkers?)"
                      r"\s*(\(\d+\))?:?$", re.I)


def parse_decklist(text: str) -> list[tuple[int, str]]:
    """[(copies, card name)] from a pasted decklist: "4 Lightning Bolt", "1x Sol Ring", "Sol Ring",
    "2 Fable of the Mirror-Breaker (NEO) 141 *F*" ... Comments and section headings are skipped; the
    same card twice is added up."""
    out = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith(("#", "//")) or _SECTION.match(line):
            continue
        m = re.match(r"^(?:SB:\s*)?(\d+)\s*[xX]?\s+(.+)$", line)
        count, name = (int(m[1]), m[2]) if m else (1, line)
        name = re.sub(r"\s+\*[A-Z]+\*\s*$", "", name)                 # *F* (foil) and the like
        name = re.sub(r"\s+\([A-Za-z0-9]{2,6}\)(\s+[\w-]+)?\s*$", "", name)   # (SET) 123
        name = re.sub(r"\s+\[[^\]]+\]\s*$", "", name).strip()          # [SET]
        if " // " in name:                                             # a double-faced card: its front
            name = name.split(" // ", 1)[0].strip()
        if name and count > 0:
            out[name] = out.get(name, 0) + count
    return [(c, n) for n, c in out.items()]


# ============================================================ the window
def open_search(parent, core, on_add, api: MPCFillAPI | None = None):
    """Open the "Find cards" window over `parent`. `on_add(groups)` gets the picked cards."""
    import tkinter as tk
    from tkinter import ttk
    api = api or getattr(core, "_mpcfill_api", None) or MPCFillAPI(f"FleemProxyCompanion/{core.APP_VERSION}")
    core._mpcfill_api = api                          # (one cache for the whole session)
    T, P = core.T, core.palette()
    log = core.log

    win = tk.Toplevel(parent)
    win.title("Find cards on MPCFill")
    win.transient(parent)
    win.geometry("1040x720")
    win.minsize(760, 520)
    T(win, bg="bg")
    state = {"gen": 0, "mode": "single", "front": None, "back": None, "rows": [], "edit_row": None}
    photos = {}                                      # thumbnail id -> PhotoImage (kept alive)
    pool = ThreadPoolExecutor(max_workers=4)
    events = []                                      # (fn, args) to run on the window's thread
    lock = threading.Lock()

    def later(fn, *args):
        with lock:
            events.append((fn, args))

    def pump():
        if not win.winfo_exists():
            return
        with lock:
            todo, events[:] = events[:], []
        for fn, args in todo:
            try:
                fn(*args)
            except tk.TclError:
                pass
            except Exception as e:
                log.warning("MPCFill window: %s", e)
        win.after(60, pump)
    win.after(60, pump)

    def background(work, done, fail=None):
        def run():
            try:
                result = work()
            except Exception as e:
                log.info("MPCFill: %s", e)
                later(fail or show_error, str(e))
                return
            later(done, result)
        threading.Thread(target=run, daemon=True).start()

    # --- top: search box, decklist toggle
    top = T(tk.Frame(win, padx=16, pady=12), bg="surface")
    top.pack(fill="x")
    T(tk.Label(top, text="Card name"), bg="surface", fg="text").pack(side="left")
    query = tk.StringVar()
    entry = ttk.Entry(top, textvariable=query, width=40)
    entry.pack(side="left", padx=(8, 6))
    search_btn = ttk.Button(top, text="Search", style="Accent.TButton", command=lambda: do_search())
    search_btn.pack(side="left")
    kind = tk.StringVar(value="CARD")
    for label, value in (("Cards", "CARD"), ("Tokens", "TOKEN")):
        ttk.Radiobutton(top, text=label, value=value, variable=kind).pack(side="left", padx=(12, 0))
    deck_btn = ttk.Button(top, text="Paste a decklist…", command=lambda: toggle_decklist())
    deck_btn.pack(side="right")
    status = T(tk.Label(win, text="Type a card's name and press Search, or paste a whole decklist.", anchor="w",
                        padx=16, pady=4), bg="bg", fg="muted")
    status.pack(fill="x")

    # --- decklist box (hidden until asked for)
    deck_frame = T(tk.Frame(win, padx=16, pady=6), bg="bg")
    deck_text = tk.Text(deck_frame, height=7, wrap="none", relief="solid", bd=1)
    deck_text.pack(side="left", fill="x", expand=True)
    deck_side = T(tk.Frame(deck_frame, padx=10), bg="bg")
    deck_side.pack(side="left", fill="y")
    ttk.Button(deck_side, text="Find all", style="Accent.TButton", command=lambda: do_decklist()).pack(fill="x")
    T(tk.Label(deck_side, text="One card per line:\n4 Lightning Bolt\n1x Sol Ring\nSol Ring (CMM) 400",
               justify="left"), bg="bg", fg="muted").pack(anchor="w", pady=(8, 0))

    # --- middle: a scrolling area for artwork grids and the decklist rows
    mid = T(tk.Frame(win), bg="bg")
    mid.pack(fill="both", expand=True, padx=16, pady=(4, 0))
    canvas = tk.Canvas(mid, highlightthickness=0, bd=0)
    T(canvas, bg="bg")
    bar = ttk.Scrollbar(mid, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=bar.set)
    bar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    inner = T(tk.Frame(canvas), bg="bg")
    canvas.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    win.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))   # (only this window)

    # --- bottom: what's picked, copies, add
    bottom = T(tk.Frame(win, padx=16, pady=10), bg="surface")
    bottom.pack(fill="x", side="bottom")
    picked = T(tk.Label(bottom, text="", anchor="w", justify="left"), bg="surface", fg="text")
    picked.pack(side="left", fill="x", expand=True)
    ttk.Button(bottom, text="Close", command=lambda: close()).pack(side="right")
    add_btn = ttk.Button(bottom, text="Add to project", style="Accent.TButton", command=lambda: do_add())
    add_btn.pack(side="right", padx=(0, 8))
    copies = tk.StringVar(value="1")
    copies_box = ttk.Spinbox(bottom, from_=1, to=99, textvariable=copies, width=4, justify="center")
    copies_box.pack(side="right", padx=(0, 12))
    T(tk.Label(bottom, text="Copies"), bg="surface", fg="text").pack(side="right", padx=(0, 6))

    def to_photo(data: bytes, box):
        """Downloaded image bytes -> a Tk image at most `box` (w, h) big."""
        from PIL import Image
        import io
        im = Image.open(io.BytesIO(data)).convert("RGB")
        im.thumbnail(box)
        try:
            from PIL import ImageTk
            return ImageTk.PhotoImage(im, master=win)
        except Exception:                      # no ImageTk: hand Tk the pixels as PPM
            return tk.PhotoImage(master=win, data=b"P6 %d %d 255\n" % im.size + im.tobytes(), format="ppm")

    def show_error(text):
        status.configure(text=f"MPCFill: {text}")
        T(status, fg="danger")

    def say(text):
        status.configure(text=text)
        T(status, fg="muted")

    def clear_middle():
        state["gen"] += 1
        for w in inner.winfo_children():
            w.destroy()
        canvas.yview_moveto(0)

    def thumb(label, art: CardArt, gen):
        """Load an artwork's small preview into `label` (in the background)."""
        if art.identifier in photos:
            label.configure(image=photos[art.identifier])
            return
        if not art.thumbnail:
            return

        def work():
            data, _ = core.download(art.thumbnail, tries=2)
            return data

        def done(data):
            if gen != state["gen"] or not label.winfo_exists():
                return
            try:
                photos[art.identifier] = to_photo(data, (130, 182))
                label.configure(image=photos[art.identifier])
            except Exception as e:
                log.info("MPCFill thumbnail: %s", e)
        pool.submit(lambda: (lambda d: later(done, d))(work()) if gen == state["gen"] else None)

    NO_BACK = "__none__"                         # the "No back" tile's id

    def art_grid(parent_frame, arts: list[CardArt], chosen_id, on_pick, cols=6, no_back=False):
        """A grid of artwork tiles; clicking one calls on_pick(art). `no_back`: a first tile "No back"
        (on_pick(None)) - for double-faced cards you want printed with a blank back."""
        tiles = {}
        gen = state["gen"]
        blank = photos.setdefault("__blank", tk.PhotoImage(master=win, width=130, height=182))
        offset = 0
        if no_back:
            tile = T(tk.Frame(parent_frame, padx=3, pady=3, cursor="hand2"), bg="tile")
            tile.grid(row=0, column=0, padx=4, pady=4, sticky="n")
            face = T(tk.Label(tile, image=blank, bd=0, compound="center", text="No back\n(prints blank)",
                              font=("Segoe UI", 10, "bold")), bg="paper", fg="muted")
            face.pack()
            cap = T(tk.Label(tile, text="\n", justify="center", font=("Segoe UI", 8)), bg="tile", fg="muted")
            cap.pack()
            for w in (tile, face, cap):
                w.bind("<Button-1>", lambda e: pick(None))
            tiles[NO_BACK] = tile
            offset = 1
        for i, art in enumerate(arts[:RESULTS_SHOWN], offset):
            tile = T(tk.Frame(parent_frame, padx=3, pady=3, cursor="hand2"), bg="tile")
            tile.grid(row=i // cols, column=i % cols, padx=4, pady=4, sticky="n")
            pic = T(tk.Label(tile, image=blank, bd=0), bg="paper")
            pic.pack()
            caption = f"{art.source[:18]}\n{art.dpi} DPI" if art.dpi else art.source[:18]
            cap = T(tk.Label(tile, text=caption, justify="center", font=("Segoe UI", 8)), bg="tile", fg="muted")
            cap.pack()
            for w in (tile, pic, cap):
                w.bind("<Button-1>", lambda e, a=art: pick(a))
            tiles[art.identifier] = tile
            thumb(pic, art, gen)

        def pick(art):
            chosen = NO_BACK if art is None else art.identifier
            for ident, t in tiles.items():
                T(t, bg="accent" if ident == chosen else "tile")
            on_pick(art)
        if no_back and chosen_id is None:
            chosen_id = NO_BACK
        if chosen_id in tiles:
            T(tiles[chosen_id], bg="accent")
        return tiles

    def heading(text):
        T(tk.Label(inner, text=text, anchor="w", font=("Segoe UI", 10, "bold")), bg="bg", fg="text").pack(
            fill="x", pady=(10, 2))

    # --- one card: search, pick front and back
    def do_search(_event=None):
        name = query.get().strip()
        if not name:
            return
        state.update(mode="single", front=None, back=None)
        clear_middle()
        update_picked()
        say(f"Searching MPCFill for “{name}”…")
        card_type = kind.get()

        def work():
            ids = api.search([name], card_type).get(name, [])
            arts = api.cards(ids[:RESULTS_SHOWN])
            back_name = api.dfc_back_name(name) if card_type == "CARD" else None
            back_arts = []
            if back_name:
                back_ids = api.search([back_name], "CARD").get(back_name, [])
                back_arts = list(api.cards(back_ids[:RESULTS_SHOWN]).values())
            return [arts[i] for i in ids if i in arts], back_name, back_arts
        background(work, lambda r: show_single(name, *r))

    def show_single(name, arts, back_name, back_arts):
        clear_middle()
        if not arts:
            say(f"MPCFill has nothing for “{name}”. Check the spelling (the full card name works best).")
            return
        say(f"{len(arts)} artwork{'s' * (len(arts) != 1)} for “{name}”. Click one for the front"
            + (f"; its back ({back_name}) is picked too - change it below." if back_name else "."))
        heading("Front")
        front_box = T(tk.Frame(inner), bg="bg")
        front_box.pack(fill="x")
        state["front"] = arts[0]
        state["query"] = name
        art_grid(front_box, arts, arts[0].identifier, lambda a: (state.update(front=a), update_picked()))
        heading(f"Back: {back_name}" if back_name else "Back")
        back_box = T(tk.Frame(inner), bg="bg")
        back_box.pack(fill="x")
        if back_arts:                            # a double-faced card: its back, or no back at all
            state["back"] = back_arts[0]
            art_grid(back_box, back_arts, back_arts[0].identifier, lambda a: (state.update(back=a), update_picked()),
                     no_back=True)
        else:
            state["back"] = None
            row = T(tk.Frame(back_box), bg="bg")
            row.pack(fill="x")
            T(tk.Label(row, text="No back (it prints blank) - or pick a card back:"), bg="bg", fg="muted").pack(side="left")
            ttk.Button(row, text="Show card backs", command=lambda: show_cardbacks(back_box, row)).pack(side="left", padx=8)
        update_picked()

    def show_cardbacks(box, row):
        row.destroy()
        say("Loading card backs…")

        def work():
            ids = api.cardbacks()[:RESULTS_SHOWN]
            arts = api.cards(ids)
            return [arts[i] for i in ids if i in arts]

        def done(arts):
            say("Click a card back (click it again to go back to no back).")
            holder = {"tiles": None}

            def pick(a):
                if state["back"] is not None and state["back"].identifier == a.identifier:
                    state["back"] = None                   # clicked again: no back
                    for t in holder["tiles"].values():
                        T(t, bg="tile")
                else:
                    state["back"] = a
                update_picked()
            holder["tiles"] = art_grid(box, arts, None, pick)
        background(work, done)

    def update_picked():
        if state["mode"] == "single":
            f, b = state["front"], state["back"]
            if f is None:
                picked.configure(text="Nothing picked yet.")
                add_btn.state(["disabled"])
            else:
                picked.configure(text=f"Front: {f.name}" + (f"\nBack: {b.name}" if b else "\nBack: none (blank)"))
                add_btn.state(["!disabled"])
            copies_box.state(["!disabled"])
        else:
            rows = [r for r in state["rows"] if r["front"] is not None]
            total = sum(r["count"] for r in rows)
            missing = [r["name"] for r in state["rows"] if r["front"] is None]
            picked.configure(text=f"{len(rows)} card{'s' * (len(rows) != 1)}, {total} copies"
                             + (f"  ·  not found: {', '.join(missing[:6])}{'…' if len(missing) > 6 else ''}"
                                if missing else ""))
            add_btn.state(["!disabled"] if rows else ["disabled"])
            copies_box.state(["disabled"])

    # --- a decklist: every card at once, first artwork picked, each one changeable
    def toggle_decklist():
        if deck_frame.winfo_ismapped():
            deck_frame.pack_forget()
            deck_btn.configure(text="Paste a decklist…")
        else:
            deck_frame.pack(fill="x", before=mid)
            deck_btn.configure(text="Hide the decklist")
            deck_text.focus_set()

    def do_decklist():
        wanted = parse_decklist(deck_text.get("1.0", "end"))
        if not wanted:
            say("Paste a decklist first: one card per line, like “4 Lightning Bolt”.")
            return
        state.update(mode="decklist", rows=[], front=None, back=None)
        clear_middle()
        say(f"Searching MPCFill for {len(wanted)} card{'s' * (len(wanted) != 1)}…")
        card_type = kind.get()

        def work():
            names = [n for _, n in wanted]
            found = api.search(names, card_type)
            backs = {n: api.dfc_back_name(n) for n in names} if card_type == "CARD" else {}
            back_names = [b for b in backs.values() if b]
            found_backs = api.search(back_names, "CARD") if back_names else {}
            first = [ids[0] for ids in found.values() if ids] + [ids[0] for ids in found_backs.values() if ids]
            arts = api.cards(first)
            rows = []
            for count, n in wanted:
                ids = found.get(n, [])
                b = backs.get(n)
                b_ids = found_backs.get(b, []) if b else []
                rows.append({"name": n, "count": count, "ids": ids, "front": arts.get(ids[0]) if ids else None,
                             "back_name": b, "back_ids": b_ids,
                             "back": arts.get(b_ids[0]) if b_ids else None})
            return rows
        background(work, show_rows)

    def show_rows(rows):
        state["rows"] = rows
        clear_middle()
        found = sum(1 for r in rows if r["front"] is not None)
        say(f"Found {found} of {len(rows)} cards (each has MPCFill's first artwork). Click “Change art” to pick "
            "another one; change the copies in the box.")
        gen = state["gen"]
        blank = photos.setdefault("__small", tk.PhotoImage(master=win, width=65, height=91))
        for r in rows:
            line = T(tk.Frame(inner, pady=3), bg="bg")
            line.pack(fill="x")
            count = tk.StringVar(value=str(r["count"]))
            sb = ttk.Spinbox(line, from_=0, to=99, textvariable=count, width=4, justify="center")
            sb.pack(side="left")
            count.trace_add("write", lambda *_a, r=r, v=count: (r.update(count=int(v.get()) if v.get().isdigit()
                                                                          else r["count"]), update_picked()))
            pic = T(tk.Label(line, image=blank, bd=0), bg="paper")
            pic.pack(side="left", padx=8)
            if r["front"] is not None:
                small_thumb(pic, r["front"], gen)
            text = r["name"] + ((f"  //  {r['back_name']}" if r["back"] is not None else "  (no back)")
                                if r["back_name"] else "")
            detail = (f"{r['front'].source} · {r['front'].dpi} DPI" if r["front"] else "not found on MPCFill")
            info = T(tk.Frame(line), bg="bg")
            info.pack(side="left", fill="x", expand=True)
            T(tk.Label(info, text=text, anchor="w", font=("Segoe UI", 10, "bold")), bg="bg", fg="text").pack(fill="x")
            T(tk.Label(info, text=detail, anchor="w"), bg="bg", fg="danger" if r["front"] is None else "muted").pack(fill="x")
            if r["ids"]:
                ttk.Button(line, text="Change art", command=lambda r=r: change_art(r)).pack(side="right")
        update_picked()

    def small_thumb(label, art, gen):
        key = "s:" + art.identifier

        def done(data):
            if gen != state["gen"] or not label.winfo_exists():
                return
            try:
                photos[key] = to_photo(data, (65, 91))
                label.configure(image=photos[key])
            except Exception as e:
                log.info("MPCFill thumbnail: %s", e)
        if key in photos:
            label.configure(image=photos[key])
        elif art.thumbnail:
            pool.submit(lambda: later(done, core.download(art.thumbnail, tries=2)[0]) if gen == state["gen"] else None)

    def change_art(r):
        """Pick another front (and back) for one decklist card, then return to the list."""
        rows = state["rows"]
        clear_middle()
        say(f"Loading the artwork for {r['name']}…")

        def work():
            arts = api.cards(r["ids"][:RESULTS_SHOWN] + r["back_ids"][:RESULTS_SHOWN])
            return ([arts[i] for i in r["ids"][:RESULTS_SHOWN] if i in arts],
                    [arts[i] for i in r["back_ids"][:RESULTS_SHOWN] if i in arts])

        def done(result):
            fronts, backs = result
            clear_middle()
            say(f"Pick the artwork for {r['name']}, then press “Back to the list”.")
            ttk.Button(inner, text="← Back to the list", command=lambda: show_rows(rows)).pack(anchor="w", pady=(4, 0))
            heading("Front")
            fb = T(tk.Frame(inner), bg="bg")
            fb.pack(fill="x")
            art_grid(fb, fronts, r["front"].identifier if r["front"] else None, lambda a: r.update(front=a))
            if backs:
                heading(f"Back: {r['back_name']}")
                bb = T(tk.Frame(inner), bg="bg")
                bb.pack(fill="x")
                art_grid(bb, backs, r["back"].identifier if r["back"] else None, lambda a: r.update(back=a),
                         no_back=True)
        background(work, done)

    # --- add to the project
    def group(front: CardArt, back: CardArt | None, query_name: str, count: int) -> dict:
        return {"front_id": front.identifier, "back_id": back.identifier if back else None,
                "name": front.name, "query": query_name, "count": count}

    def do_add():
        if state["mode"] == "single":
            if state["front"] is None:
                return
            try:
                n = max(1, min(99, int(copies.get())))
            except ValueError:
                n = 1
            groups = [group(state["front"], state["back"], state.get("query") or state["front"].name, n)]
        else:
            groups = [group(r["front"], r["back"], r["name"], r["count"])
                      for r in state["rows"] if r["front"] is not None and r["count"] > 0]
        if not groups:
            return
        close()
        on_add(groups)

    def close():
        state["gen"] += 1
        pool.shutdown(wait=False, cancel_futures=True)
        win.destroy()

    entry.bind("<Return>", do_search)
    win.protocol("WM_DELETE_WINDOW", close)
    win.bind("<Escape>", lambda e: close())
    try:
        core.apply_theme(win)
    except Exception:
        pass
    update_picked()
    try:
        core.center_on(win, parent)
    except Exception:
        pass
    entry.focus_set()
    return win
