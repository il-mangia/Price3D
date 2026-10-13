# Price3D

<p align="center">
  <img src="https://raw.githubusercontent.com/il-mangia/Price3D/main/Logo.png" alt="Price3D logo" width="160">
</p>

<h1 align="center">Price3D</h1>

<p align="center">
  <b>Fast quotes and professional order summaries for your 3D printing business.</b>
</p>

<p align="center">
  <a href="https://github.com/il-mangia/Price3D/releases/latest">
    <img src="https://img.shields.io/github/v/release/il-mangia/Price3D?label=release&style=flat&color=blue" alt="Latest release">
  </a>
  <a href="https://github.com/il-mangia/Price3D/releases">
    <img src="https://img.shields.io/github/downloads/il-mangia/Price3D/total?label=downloads&style=flat&color=brightgreen" alt="Total downloads">
  </a>
  <a href="https://github.com/il-mangia/Price3D/stargazers">
    <img src="https://img.shields.io/github/stars/il-mangia/Price3D?label=stars&style=flat&color=yellow" alt="GitHub stars">
  </a>
  <a href="https://github.com/il-mangia/Price3D/blob/main/LICENSE">
    <img src="https://img.shields.io/github/license/il-mangia/Price3D?label=license&style=flat&color=orange" alt="License">
  </a>
  <a href="https://github.com/il-mangia/Price3D/issues">
    <img src="https://img.shields.io/github/issues/il-mangia/Price3D?label=issues&style=flat" alt="Issues">
  </a>
</p>

<p align="center">
  <a href="#features">Features</a> ·
  <a href="#requirements">Requirements</a> ·
  <a href="#download">Download</a> ·
  <a href="#usage">Usage</a> ·
  <a href="#printer-profiles">Profiles</a> ·
  <a href="#pdf-and-word-order-summary">Export</a> ·
  <a href="#file-structure">File structure</a> ·
  <a href="#adding-a-language">Languages</a> ·
  <a href="#license">License</a>
</p>

---

**Price3D** is a lightweight desktop app (PyQt6) that calculates the real price of a 3D print job from material, energy and design-time costs — and turns it into a clean, professional **order summary (PDF or Word)** that you can send to your customer for acceptance before you start production.

No subscriptions, no accounts, no cloud. Just your numbers, on your machine.

---

## Features

- **Printer profiles** — save as many printers as you like (energy cost, printer power draw, material cost, design rate, PC power draw while designing) as small JSON files, and switch between them from a dropdown before calculating.
- **Single-object and multi-object mode** — either quote a single print, or describe a *batch*: N prints × M total objects. Time and material are entered *per print*, while the total is correctly split across all objects.
- **Design-time toggle** — include or exclude 3D design hours with one checkbox. Design is billed **once per order, not per piece**: it gets its own line item and is added a single time to the total, regardless of how many pieces are produced.
- **Extra items** — add optional add-ons (keychains, packaging, post-processing, …) with a name and a price, billed once per order.
- **Round-up toggle** — round the total up to the next whole euro. The rounding difference is folded quietly into the *Energy* line, and every other figure on the document (per-piece cost, subtotal, total) is recomputed from that adjusted value — so nothing ever shows a rounded total next to line items that don't actually add up to it.
- **PDF order summary ("Order Summary and Acceptance")** — a formal, invoice-like but explicitly *non-invoice* PDF you can send to the client for approval before production starts. It includes:
  - Your company logo (square PNG, e.g. 500×500 px) shown in the header
  - Order date, print time, number of prints and total objects
  - Full cost breakdown: material, energy, per-print cost, prints subtotal, one-time design fee, extra items and the final total
  - An optional photo of the printed project (or a render of the 3D model)
  - A clear disclaimer stating that it is a summary/quote, not a tax invoice
- **Word (.docx) order summary** — the exact same document, same layout, same colors, same content as the PDF, but exported as an editable Word file — handy if you want to tweak the wording before sending it.
- **First-run wizard** — a 3-step setup on first launch: welcome, company details (name + logo) and your first printer profile. Nothing to configure by hand.
- **Multi-language UI** — English (default), Italian, Spanish and French. Add more by editing `languages.json`. The chosen language is remembered between runs.
- **Dark theme** — a clean, modern dark UI that scales down to 720p screens.
- **Built-in update check** — the splash screen quietly checks GitHub for a newer release; the About dialog (F1) shows the result and links to the project page.

---

## Requirements

Only needed if you're running **from source** (see [Usage](#usage)). If you're using a prebuilt release for your OS, you don't need Python at all — jump to [Download](#download).

- Python **3.10+**
- [PyQt6](https://pypi.org/project/PyQt6/)
- [reportlab](https://pypi.org/project/reportlab/) — for PDF export
- [python-docx](https://pypi.org/project/python-docx/) — for Word export
- [Pillow](https://pypi.org/project/Pillow/) — used to size the project photo correctly and to generate the checkbox tick icon

```bash
pip install PyQt6 reportlab python-docx Pillow
```
## Download

Grab the build for your operating system from the [**Releases**](https://github.com/il-mangia/Price3D/releases/latest) page.

### Windows

1. Download the latest `Price3D_vX.X.exe` file.
2. Double-click it — no installation needed, it's a portable executable.

> If Windows SmartScreen warns you about an unrecognized app, click **More info → Run anyway**. This happens because the executable isn't code-signed.

### macOS

1. Download the latest macOS build (`.dmg` or `.zip`, depending on the release).
2. Open it and drag **Price3D** into your **Applications** folder.

> On first launch, macOS Gatekeeper may block the app because it isn't notarized. Right-click the app → **Open** → confirm **Open** in the dialog to allow it once.

### Linux

1. Download the latest `.tar.gz` archive.
2. Extract it and make the binary executable:
```bash
tar -xzf Price3D-linux.tar.gz
cd Price3D
chmod +x Price3D
./Price3D
```

## Usage

Running from source instead of a release build:

```bash
python main.py
```
On first launch a short wizard guides you through:

1. A welcome screen with a link to the GitHub project.

2. Your company details — name and (optional) logo. These are used in the exported PDF/Word documents.

3. Your first printer profile — energy cost, printer power, material cost, design rate and PC power while designing.

After that, the main window is where you work:

1. Pick a printer profile from the dropdown.
2. Choose the job type: Single object or Multi-object (In multi-object mode, set the number of prints and the total number of objects; time and material are entered per print).
3. Enter print time (hours / minutes), material used (grams), design hours and number of pieces.
4. Toggle Include design hours and Round up to the next euro as needed.
5. Optionally add extra items in the "Extra items" box.
6. Click Calculate — the receipt preview appears immediately, ready to copy into WhatsApp or an email.
7. Click Export PDF summary or Export Word summary to generate the Order Summary and Acceptance document (identical layout in both formats). You'll be prompted to pick a square company logo (PNG), then optionally a photo of the project, then where to save the file.

## Printer profiles

Click **Settings** to open the profile manager. Each profile stores:

| Field | Description |
| --- | --- |
| Profile name | Shown in the dropdown, used to build the filename |
| Energy cost (€/kWh) | Your electricity rate |
| Printer power (W) | Rated power draw of the printer |
| Material cost (€/kg) | Filament / resin cost |
| Design rate (€/h) | Your hourly rate for 3D design work |
| PC power while designing (W) | Power draw of the computer used for modelling |

The same dialog also lets you edit your **company name and logo**, which are reused in every exported document.

Profiles are stored as individual JSON files in the app data folder (see [File structure](#file-structure)). The filename is derived from the profile name — a profile named **"Artillery Genius"** is saved as `artillery_genius.json`.

---

## File structure

The application stores its user data in a fixed per-OS folder, so your profiles survive updates and stay out of the program directory:

| OS | Data folder |
| --- | --- |
| Windows | `%APPDATA%\Price3D` |
| macOS | `~/Documents/Price3D` |
| Linux | `~/Documents/Price3D` |

Inside that folder you'll find:

    Price3D/                     # app data folder (per-OS, see table above)
    ├── config.json              # remembers the selected language
    ├── settings.json            # company name and logo filename
    ├── company_logo.png         # your company logo (copied here on selection)
    ├── tick.png                 # generated checkbox tick (Pillow)
    └── profiles/                # one JSON file per printer profile
        ├── artillery_genius.json
        └── default.json

The repository itself only contains the source files:

    .
    ├── main.py            # main application (single file)
    ├── languages.json     # UI translations (en, it, es, fr)
    ├── Logo.png           # app logo and default export logo
    ├── README.md
    └── LICENSE

> **Migrating from an older version?** On startup Price3D looks for a legacy `config.json` and `profiles/` folder next to the script and copies them into the new data folder automatically, so you don't lose anything.

---

## PDF and Word order summary

Clicking **Export PDF summary** or **Export Word summary** builds a formal document called *Order Summary and Acceptance* — in PDF or editable Word format, sharing the exact same layout, colors and content.

It is deliberately **not** a tax invoice: it's meant to be sent to the client so they can review and confirm the order before you start production. It includes:

1. Your company logo (top left) — pick a **square** PNG (e.g. 500×500 px) when prompted. If a company logo is already saved in Settings, it's used automatically.
2. Order date, print time, number of prints and total objects.
3. Full cost breakdown: material, energy, per-print cost, prints subtotal, one-time design fee and any extra items.
4. The final total — always a clean number consistent with the round-up setting; no line items that fail to add up to the total shown.
5. A short disclaimer clarifying it's a summary, not an invoice, plus a line asking the client to confirm acceptance.
6. An optional photo of the printed project (or a render of the 3D model), near the end of the document.

---

## Adding a language

1. Open `languages.json`.
2. Add a new top-level key with the two-letter language code (e.g. `"de"`).
3. Copy the full key set from the `"en"` block and translate each value.
4. The `_name` key is what shows up in the language dropdown.

The app falls back to English for any missing key, so a partial translation is fine to start with.

---

## Contributing

Bug reports, feature requests and pull requests are all welcome.

- Found a bug? [Open an issue](https://github.com/il-mangia/Price3D/issues).
- Want to contribute code? Fork the repo, create a feature branch and open a pull request.
- Translating to a new language? Edit `languages.json` and send a PR — that's it.

---

## License

Price3D is released under the **GNU General Public License v3.0**.  
See the full text here: [GPL-3.0 License](https://github.com/il-mangia/Price3D/blob/main/LICENSE).

You are free to use, study, share and modify the software, provided that any derivative work is distributed under the same license.

---

<p align="center">
  Made with ☕ and 🖨️ by <a href="https://github.com/il-mangia">il-mangia</a>
</p>
