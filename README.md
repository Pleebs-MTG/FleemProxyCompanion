<p align="center">
  <img src="fleem_icon.png" width="128" alt="Fleem">
</p>

<h1 align="center">Fleem Proxy Companion</h1>

<p align="center">
  Turn <a href="https://mpcfill.com">MPCFill</a> order XMLs and proxy PDFs into print-ready card sheets,
  and print them at exact size on any printer.<br>
  Windows · free · no other programs needed
</p>

---

## What it does

- **Makes print-ready PDFs from MPCFill XMLs.** It downloads every card image and lays the cards out
  on the paper size you choose (A4, Letter, Legal, A3, Tabloid or A5), with your choice of card size
  (Standard, Poker, Japanese, Miniature, Oversized or your own), bleed edge and card spacing. Double-faced cards come
  first and get their real backs on the following page, mirrored so they line up when you flip the sheet.
  If an art can't be downloaded, Fleem uses a different MPCFill art of the same card and tells you which.
- **Prints them properly.** Prints at exactly 100% size, centred and turned to fit the paper, through
  any Windows printer, with that printer's own settings (paper, quality, colour, presets), remembered
  per printer.
- **Lets you choose exactly what to print.** Preview every page, pick fronts and backs, print just page 3,
  or pages 1, 5-8.
- **Lets you edit the cards before printing.** Change copies, replace art, add or remove backs, add new
  cards by dragging images in, combine several orders into one print, change bleed and spacing. It all
  updates live, and the PDF is only saved when you press Print.
- **Works with PDFs from other programs too.** Proxy PDFs made by Proxy PDF Maker, MPC Autofill and
  similar tools are taken apart into cards, so you can edit, combine and re-space them as well.
- **Reads card names off the cards** (Windows' built-in text recognition), so cards from PDFs and
  images get sorted A-Z.
- **Finds cards on MPCFill.** Type a card's name (or paste a whole decklist), pick the art for the
  front and back, and add it to the print.
- **Cut marks and cutting machines.** Crop marks for paper cutters, or the whole Print & Cut setup for
  Silhouette (registration marks + cut lines) and Cricut (Print Then Cut sheets).
- **Undo and redo** every change to the cards and the layout.
- **Lines up the backs.** Optional back alignment calibration with a printable test sheet.
- **Light and dark mode.**
- **Optional Discord integration** with the Fleem order bot (coming soon): new orders become PDFs
  automatically, and Fleem asks you to load the paper and print.

---

## Getting started

### Option A: download the app
1. Download the `.zip` from the [Releases](../../releases) page.
2. Unzip it into a folder of its own (e.g. `Documents\Fleem Proxy Companion`). It holds
   `Fleem Proxy Companion.exe` and an `_internal` folder: keep them together. Fleem keeps its settings
   next to them.
3. Double-click `Fleem Proxy Companion.exe`. If Windows warns about an unknown app, click **More info >
   Run anyway**.

### Option B: run it from source
1. Install [Python 3.10+](https://www.python.org/downloads/) (tick **Add Python to PATH**).
2. Download this repository and double-click **`build.bat`**. It installs what's needed and builds
   `Fleem Proxy Companion.exe` in the same folder.
   (Or run it directly: `pip install -r requirements.txt`, then `pythonw fleem_pdf_maker.py`.)

Fleem's window opens, with buttons for everything it does. The first time it runs, Fleem also adds
**Print with Fleem** to the right-click menu of PDF and XML files.

### Running in the background (optional)
Fleem only runs while one of its windows is open. Close them all and it stops.

To keep it running instead, tick **Keep running in the background** in Fleem's window. When you close
the window, Fleem's icon stays in the **taskbar corner next to the clock** (click the `^` if you don't
see it). Click the icon for the menu, or choose **Open Fleem Proxy Companion** to get the window back.
In the background, **Print with Fleem** opens instantly and Discord orders are caught all the time.
Tick **Start with Windows** too to have it start whenever you log in.

To stop running in the background, untick it in Fleem's window or in the icon's menu.

---

## How to use it

### Make a PDF from an XML
**Make PDF from XML…** in Fleem's window (or the icon's menu) and pick your MPCFill XML. Check the bleed edge, card spacing and
paper size, press **Make PDF**, and the card images download (big orders take a few minutes). The PDF is saved in
`Documents\Fleem PDFs` and opened.
To put several orders into one PDF, use **Merge XMLs into one PDF…** (hold Ctrl to pick several).

### Print
Right-click a PDF or XML > **Show more options** > **Print with Fleem**, or use **Print PDFs or XMLs…**
in Fleem's window. The print window opens:

1. **Pick the pages to print.** Click a page to pick it (click again to unpick). Shift+click picks a
   range. **Select all** (or Ctrl+A) picks every page. You can also type page numbers in the box at the
   bottom, like `1, 5-8, 10-11`.
2. **Backs (if you need them).** Click **Backs** at the top of the pages (or press Ctrl+Tab) and pick
   those too. If you pick fronts and backs, Fleem lists both and asks you to confirm.
3. **Choose the printer** from the list. Fleem remembers the one you used last.
4. **Printer settings (optional).** **Printer settings…** opens your printer's own settings window,
   whatever the brand. Choose the paper type, quality, colour or one of your saved presets, and press OK.
   Fleem remembers these settings for that printer.
5. **Bleed, spacing, paper and card size (optional).** Change **Bleed** (extra picture around each card,
   in mm), **Spacing** (gap between cards, in mm), **Paper** (the page size) or **Card** (the card size).
   The previews update straight away. The card sizes are the same as Proxy PDF Maker's:

   | Card size | Size | For |
   |---|---|---|
   | Japanese | 59 x 86 mm | e.g. Yu-Gi-Oh! |
   | Miniature | 1.24 x 1.73 in | half-size Magic cards (Proxy PDF Maker calls it Novelty) |
   | Oversized | 3.46 x 4.96 in | oversized Magic cards |
   | Poker | 2.5 x 3.5 in | poker cards |
   | Standard | 2.48 x 3.46 in | Magic, Pokemon and most other TCGs |
   | Custom… | your width x height in mm | anything else |

   Card images are stretched to fit the size you pick.

**Moving cards:** drag a card on a page preview (fronts or backs) and drop it between two other cards,
on any page; a blue bar shows where it will go. Drop it past the last card to put it at the end. The
cards stay where you put them (also when you add cards, edit them or change the layout) and are saved
that way in the PDF. Only changing **Advanced settings > Card order** sorts them again.

To look closer, click **View page** or double-click a page (the arrow keys go through the pages and
Space picks one). **−** and **+** (or Ctrl + mouse wheel) make the previews smaller or bigger.

Press **Print**. Nothing is saved until you do: **Cancel** leaves everything as it was. For an XML,
Print saves the PDF in your PDF folder first. For a PDF, your changes are saved into it.

**Export as PDF** (next to Print) saves just the pages you picked as a PDF of their own, without
printing. The first time, Fleem asks where to save it, then asks whether to always save exported PDFs
in that folder. Say yes and it stops asking. You can change the folder at any time with
**Advanced settings > Set PDF Export Location**.

**Export as XML** (shown when every card came from MPCFill: an order XML, Find cards, or a Fleem PDF
made from them) saves the project as an MPCFill order XML, so you can order it from MPCFill or share it.

If a print window is already open and you open another PDF or XML with Fleem (Print with Fleem, or
Open with), its cards are added to the open project, just like dragging it in. Fleem asks first.

**Advanced settings** (top right of the print window) also has:
- **Orientation:** portrait or landscape pages. In landscape, flip the sheet on its short edge when you
  print the backs.
- **Header:** tick it to print the file's name and the page number at the top of each page. Change the
  text in the box next to it if you want something else. Tick **Number only the pages being printed**
  to number just the pages you print, counted from 1 (print pages 3, 5 and 7 and they say 1 of 3,
  2 of 3, 3 of 3). The other pages show the name only.
- **Cache folder:** where Fleem keeps downloaded card images and temporary files. Fleem makes a
  `Fleem cache` folder inside the one you pick and empties only that. **Default** puts it back next to
  the app.
- **Card order:** **MDFC then Alphabetical** (the default: double-faced cards first, so their backs share
  as few back pages as possible, then the rest A-Z) or **Only Alphabetical** (every card A-Z). Changing it
  sorts the open project again (cards you moved go back to their spots), and new projects use it too.
- **PDF export:** **Set PDF Export Location** picks the folder **Export as PDF** saves to.

### Edit the cards
- **Right-click a card** on a page: Add a copy, Remove this copy, Set to 0 copies, Make bleed, Edit this
  card…, View this page. **Make bleed** is for art that fills MPC's shape but has no bleed in it (the
  card's edge would be cut off): Fleem treats the picture as the card itself and makes the bleed.
- **Undo / Redo** (Ctrl+Z / Ctrl+Y) appear at the top as soon as there's something to undo: adding,
  removing, moving and editing cards, the bleed, spacing, paper, card size, orientation, header, cut
  marks and card order. The last 100 changes are kept while the window is open.
- **Edit cards…** shows every card in a grid (as it looks once cut, without the bleed) with its number
  of copies underneath. Click a card, then use
  the buttons on the right:
  - **Replace front… / Replace back…** to use other art (any image; it's sized to fit)
  - **Add back…** (or click the **No back** box) to give a card a back: drag in an image, choose one,
    or pick another card from the project. A card you pick becomes the back and is taken out of the list.
    That's the quick fix for double-faced cards Fleem didn't recognise. **Remove back** takes it off.
  - **Add a back to all cards…** gives every card without a back the same back (an image, or a card from
    the project), for example a custom card back
  - **Delete card**, or set its **copies to 0** (it stays in the list so you can bring it back)
  - **Name** to rename it (cards are sorted A-Z by name)
  - **Find** (Ctrl+F) to search
- **Add cards:** drag card images (PNG, JPG, WEBP…) onto the print window, or click **Add cards…**.
  Fleem reads the card's name, asks how many copies you want, and lets you drag in a back image for
  double-faced cards. Images without bleed get some added automatically: the border colour is
  continued, or the art is mirrored for borderless cards.
- **Combine orders:** drag another XML or PDF onto the print window. Fleem shows what will be added
  and asks first.
- **Find cards…** searches [MPCFill](https://mpcfill.com) by card name. Pick the art for the front
  (and the back: a double-faced card's back is picked for you, or choose **No back** to leave it blank;
  other cards can have a card back or none), choose the number of copies and press **Add to project**. **Paste a decklist…** finds every
  card of a list at once (`4 Lightning Bolt`, `1x Sol Ring`, `Sol Ring (CMM) 400`…), with MPCFill's
  first art for each; **Change art** picks another. The art is downloaded like an order XML's.

Cards whose image couldn't be downloaded (even with another art) show as grey placeholders. Replace them in Edit cards.

### Cut marks and cutting machines
**Advanced settings > Cut marks** (saved in the PDF):

- **Crop marks (paper cutter):** short lines in the page margins, in line with every card's edge, on
  the front pages. Cut along them with a guillotine or trimmer.
- **Silhouette Print & Cut:** Silhouette's registration marks (type 1: a square top left, L shapes top
  right and bottom left) on the front pages, with the cards kept clear of them. When you print (or
  export), Fleem saves the **cut lines** next to the PDF: `… - cut lines.dxf` (Silhouette Studio's free
  Basic edition) and `.svg` (Designer edition and up, and other cutter software), each card's outline
  with its rounded corners, exactly where it's printed. Then:
  1. Print the front pages from Fleem.
  2. In Silhouette Studio: Page Setup → the same paper and orientation; Registration Marks → Type 1 with
     the **same Length, Thickness and Inset** as in Fleem (shown under Cut marks; 20 mm, 0.5 mm and
     10 mm unless you change them).
  3. Open the cut lines file. The first card's position is in Fleem's message, in case Studio moves the
     lines.
  4. Put the printed sheet on the mat and Send: the machine reads the marks and cuts each card.
  The marks take room, so fewer cards fit (landscape usually fits more; Fleem tells you).
- **Cricut Print Then Cut:** Design Space prints and cuts itself, so Fleem makes the sheets for it. The
  cards are laid out inside Design Space's biggest Print Then Cut size (6.75 x 9.25 in), with no bleed
  (Design Space adds its own) and at least 6.5 mm between them. **Print** and **Export** save each
  picked sheet as a PNG with a see-through background and rounded corners, so Design Space cuts around
  every card. In Design Space: Upload → the PNG → Complex → Continue → Print Then Cut image; add it to
  the canvas and set its size to exactly what the file name says (`6.75 x 9.25 in`); Make It with Bleed
  on. Print Then Cut is one-sided, so backs aren't included.

### PDFs from other programs
Print with Fleem on a proxy PDF made elsewhere takes its cards out (any paper size, grid, bleed or
spacing) and works out which pages are backs. Everything above works on them, and **Print** saves a new
`name (Fleem).pdf` next to the original (on the same paper size, unless you pick another). The original is
never changed.
In **Advanced settings**: if the back pages were guessed wrong, pick the other **Page layout**. To print
the original unchanged, use **Print the original PDF as it is instead**.
PDFs where each page is one flat picture can only be printed as they are.

### Card names
For cards with no name (PDFs from other programs, or images you add), Fleem reads the name off the card
using Windows' built-in text recognition. It reads the title, or the smaller line under it when there is
one, since custom cards put the real card's name there. The result is checked against
[Scryfall](https://scryfall.com)'s list of card names to fix misreads. Cards it can't read stay
unnamed: type the name in Edit cards.

### Line up the backs
If the backs come out shifted compared to the fronts: **Advanced settings** > tick **Back alignment** >
**Print test sheet**. Print the front, put the sheet back in the printer the way you do for backs, and
let it print the back. Hold the sheet up to a light, read how far the bold cross is from the thin one
(looking at the back: +X is right, +Y is down), type it into **X** and **Y**, and print another test sheet
to check. It's saved per printer, and only back pages move.

### Themes
Click **Theme** in Fleem's window or the print window (or the icon's menu > **Appearance**) and pick one:

- **Basic - Light** and **Basic - Dark**: clean and simple.
- **RPG - Light**: an adventurer's desk. Buttons are pixel-art scrolls between wooden rollers, checkboxes
  take a red wax seal, options are gem sockets that light up green, scrollbars are wooden rods, progress
  bars fill with potion, and everything sits on grainy parchment. The writing on the scrolls is in the
  pixel font Pixeloid Sans; the rest is in a storybook serif.
- **RPG - Dark** (the default): the same adventure by candlelight, on dark wood and leather with gold
  highlights.

Pixeloid Sans is packed into the app (`build.bat` puts the `fonts` folder inside the .exe), so nobody has
to download or install a font. Fleem loads it for itself only; nothing is added to Windows' fonts.

Every Fleem window follows your choice, pop-ups included. A window that's open when you switch reloads
itself in the new theme, and keeps everything you were doing. On Windows 11 the RPG themes give the title bar
a wooden colour too.

### Updates
Fleem checks for a new version when its window opens, and once a day in the background. When there is
one, Fleem's window shows **Version X is available** (in the background, you get a notification and the
icon's menu shows **Update to version X**). Click **See what's new**, then:

- **Update now** downloads the new version, checks it, and restarts Fleem. Your settings stay, and so do
  orders waiting to print.
- **Later** asks again next time.
- **Skip this version** stops reminding you about that version.

**Check for updates** checks right away. To turn the daily check off, set
`check_for_updates = no` in `settings.ini`. If you run Fleem from source, the update window opens the
download page instead of installing.

---

## Discord orders (optional)

> **The Fleem Discord order bot isn't available yet.** It will be released soon. Until then you can
> ignore this section: without a bot token, all of this is off and hidden.

Once you run the Fleem Discord order bot, Fleem can watch its order channels:

1. Open `settings.ini` (next to the app) and fill in `discord_token` (the bot's token) and `guild_id`
   (your server's ID).
2. Restart Fleem.

Fleem checks for orders while its window is open, or all the time with **Keep running in the
background** on. Every `check_every_minutes` (or **Check for orders now**), each new order's XML becomes a
PDF. Fleem then says *"Fleem is waiting for you to load 12 sheets into the printer."* **I'm ready**
prints all the fronts; **Not yet** keeps the PDF for later. Orders still waiting are remembered if Fleem
restarts. When an order is marked Completed or deleted in Discord, its PDF goes to the Recycle Bin.
`skip_orders_before` ignores older orders. `check_discord.bat` checks once and shows what happens.

Without a token, all of this is off and hidden.

---

## Settings

`settings.ini` is created next to the app the first time it runs:

- `output_folder`: where PDFs are saved (default: `Documents\Fleem PDFs`)
- `open_pdfs`: open PDFs made with Make PDF from XML, `yes` or `no` (default: `yes`)
- `rotate_pages`: which way to turn pages when the paper is set up sideways, `counterclockwise` or
  `clockwise` (default: `counterclockwise`)
- `discord_token` and `guild_id`: Discord orders, see above (default: empty, which means off)
- `check_every_minutes`: how often to check Discord (default: `120`)
- `skip_orders_before`: ignore Discord orders numbered lower than this (default: `0`)
- `keep_running`: keep running in the background after the window is closed, `yes` or `no`
  (default: `no`). The same as the **Keep running in the background** box.
- `check_for_updates`: check GitHub for a new version, `yes` or `no` (default: `yes`)
- `cache_folder`: where the `Fleem cache` folder goes (default: empty, which means the `work` folder next
  to the app). The same as **Advanced settings > Cache folder**.
- `card_sort`: `mdfc_first` (MDFC then Alphabetical, the default) or `alphabetical` (Only Alphabetical).
  The same as **Advanced settings > Card order**.
- `silhouette_mark_inset`, `silhouette_mark_length`, `silhouette_mark_thickness`: the Silhouette
  registration marks' size in mm (default 10, 20 and 0.5). The same as **Advanced settings > Cut marks**;
  use the same values in Silhouette Studio.
- `export_folder`: where **Export as PDF** saves (default: empty, which means Fleem asks each time). The
  same as **Advanced settings > Set PDF Export Location**.

The bleed edge, card spacing, paper, card size, orientation and header aren't settings: Fleem remembers what you
used last (the header's text belongs to each PDF).

> **Keep `settings.ini` private if it has a Discord token in it.** It's in `.gitignore`.

<details>
<summary>Files Fleem creates next to itself</summary>

- `settings.ini`: your settings
- `fleem.log`, `crash.log`: what Fleem did, and any errors
- `printer_settings\`: settings saved with Printer settings…, one file per printer (delete one to forget it)
- `last_layout.json`, `last_printer.json`: the bleed, spacing, paper and printer you used last
- `back_alignment.json`: back alignment offsets per printer
- `appearance.json`: light or dark mode
- `card_names.json`: Scryfall's card name list (refreshed monthly)
- `processed_orders.json`, `print_queue.json`: Discord orders already handled, and orders waiting to print
- `update_state.json`: a version you chose to skip
- `welcome.json`: the welcome message has been shown (delete it to see it again)
- `work\`: temporary files and the image cache (emptied when Fleem starts and quits), unless you chose
  another cache folder
- `_update_new\`, `_update.cmd`: an update waiting to be put in place (only while Fleem closes to update)
</details>

---

## How printing works

Fleem prints straight through Windows (no PDF reader needed) at **actual size**: the page is centred on
the paper and turned if the paper is set up sideways. It works the same as "Actual size, Auto-rotate,
Auto-centre" in a PDF reader. Colour is left to your printer's settings. Windows colour management is
only used when those settings choose it. Make the PDF on the same paper size you print on (the **Paper**
box). If they differ, the page is still printed at actual size, centred.

---

## Building

`build.bat` installs the requirements (plus the optional text recognition packages) and builds
`Fleem Proxy Companion.exe` with PyInstaller, as a **folder build**: the .exe plus an `_internal` folder
next to it. That starts much faster than a single-file .exe, which has to unpack itself every time it
opens. It also makes `release\Fleem Proxy Companion <version>.zip` for GitHub. Run it again after
updating `fleem_pdf_maker.py`.

`fleem_mpcfill.py` is the MPCFill card search (**Find cards…**). Everything that talks to MPCFill is in
its `MPCFillAPI` class, so if MPCFill changes its API, that's the only part to edit. To remove the
feature, delete the file (and its `--hidden-import` in `build.bat`); the button only shows when the file
is there.

Command line:
```
pythonw fleem_pdf_maker.py                  open Fleem's window (or the icon by the clock, with keep_running)
pythonw fleem_pdf_maker.py --print FILE     open the print window for a PDF or XML
python  fleem_pdf_maker.py --once           check Discord once and print what happens
python  fleem_pdf_maker.py --install-startup / --remove-startup
python  fleem_pdf_maker.py --version        show the version
```
Helper scripts: `start.bat`, `stop.bat`, `add_to_startup.bat`, `remove_from_startup.bat`, `check_discord.bat`.

### Adding your own tools
Tools only you need can go in a file called `fleem_personal.py` next to `fleem_pdf_maker.py`. It's in
`.gitignore`, so it's never uploaded. If it's there, Fleem loads it and calls its `setup(core)` function,
which can replace the hooks listed under "add-on hooks" near the top of `fleem_pdf_maker.py` (for example
picking the bleed and spacing from a file's name, or keeping printer settings per paper type).
`build.bat` includes it in the .exe. That way your own build gets every update to Fleem, plus your tools.

### Releasing a new version
1. In `fleem_pdf_maker.py`, raise `APP_VERSION`: the last number for small updates and fixes (`0.4.0` to
   `0.4.1`), the middle number for bigger ones (`0.4.1` to `0.5.0`). `build.bat` copies it into the
   .exe's file properties.
2. Run `build.bat`.
3. On GitHub, go to **Releases > Draft a new release**. Make a new tag that's the version with a `v` in
   front (`v0.4.1`), write what's new in the description, attach
   `release\Fleem Proxy Companion <version>.zip`, and publish. (Fleem's updater downloads the .zip,
   checks it, and swaps the new version in when Fleem closes.)

Everyone's Fleem picks it up within a day. Drafts and pre-releases are ignored, so you can test a build
first. If you fork Fleem, set `GITHUB_REPO` near the top of the updates section of
`fleem_pdf_maker.py` to your own repository.

---

## Troubleshooting

- **Fleem's icon isn't there:** with **Keep running in the background** off, there's no icon once the
  window is closed: start Fleem again to open it. With it on, click the `^` by the clock. Fleem runs only
  once at a time (starting it again brings its window to the front). If it's stuck, run `stop.bat`.
- **"…is open in another program":** close the PDF in your PDF viewer, then press Print again.
- **Card names aren't read:** the optional text recognition packages aren't installed (run `build.bat`
  again), or Windows has no OCR language installed (Settings > Time & language > Language).
- **An update didn't install:** download the new `Fleem Proxy Companion.exe` from the Releases page and put it
  in place of the old one (quit Fleem first). If the app is in a folder Windows protects, like
  Program Files, move it to a folder of your own (for example `Documents\Fleem Proxy Companion`).
- **Something went wrong:** **View log** in Fleem's window (or the icon's menu). If Fleem closed by itself, look in `crash.log`.

---

## Privacy

Fleem only goes online to download card images (from MPC Autofill's image server), Scryfall's list of
card names, GitHub (to check for a new version of Fleem, unless you turn that off), and, if
you set it up, your Discord bot's order channels. Right-click **Print with Fleem**
talks to the running Fleem over `127.0.0.1` (your own PC) only.

---

## License

[MIT](LICENSE). Fleem Proxy Companion isn't affiliated with Wizards of the Coast, MPCFill or Scryfall.

The RPG themes use the font Pixeloid Sans by GGBotNet, included in `fonts\` under the SIL Open Font
License 1.1 (see `fonts\Pixeloid-OFL.txt`).
The card sizes are the ones in [Proxy PDF Maker](https://github.com/Malacath-92/Proxy-PDF-Maker) (MIT).
Find cards uses the public search of [MPCFill](https://github.com/chilli-axe/mpc-autofill); Fleem isn't
affiliated with it.
Proxies are for personal, casual use. Card names and images belong to their owners.
