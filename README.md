# sanitize_pptx.py

Masks confidential content in a PowerPoint deck so it can be sent out for formatting, then puts the originals back when the deck returns.

1. The client runs the redactor on their own machine. It writes a redacted deck and a restore key (.xlsx).
2. The key is an unencrypted Excel file holding every original value: it never leaves the client's machine.
3. When the formatted deck comes back, the client restores it with the key.
## Setup

Python 3.9 or later + no packages to install. 

Works on **.pptx** only: save **.ppt** files as **.pptx** first
## Commands

| Goal | Command | Output |
|---|---|---|
| *See what would be caught* | **python sanitize_pptx.py deck.pptx --preview** | **deck_preview.pptx** (yellow highlights), **deck_preview_report.csv** |
| *See what --all would mask* | **python sanitize_pptx.py deck.pptx --preview --all** | **deck_preview.pptx** (yellow highlights), **deck_preview_report.csv** |
| *Mask only your own terms* | **python sanitize_pptx.py deck.pptx --keywords keywords.txt** | **deck_redacted.pptx**, **deck_restore_key.xlsx** |
| *Mask all text and numbers* | **python sanitize_pptx.py deck.pptx --all** | **deck_redacted.pptx**, **deck_restore_key.xlsx** |
| *Full masking: all text, titles included* | **python sanitize_pptx.py deck.pptx --all --titles** | **deck_redacted.pptx**, **deck_restore_key.xlsx** |
| *Restore* | **python sanitize_pptx.py deck_redacted.pptx --restore deck_restore_key.xlsx** | **deck_redacted_restored.pptx** |

- Every run needs a mode: **--preview**, **--all**, **--keywords** or **--restore**. There is no mode that masks only the detected emails, phones and IDs: list the values to mask with **--keywords**, or mask all text with **--all** (add **--titles** to mask slide titles too).
- **--preview** highlights what a run would mask and removes nothing: the built-in detection on its own, only your terms with **--keywords**, all text except slide titles with **--all**. Chart data, links, alt text and file properties can't be highlighted; they are listed in the report.
- **--titles** goes with **--all** and masks slide titles too, on slides, layouts and the master: no text stays readable. Use it when the titles themselves are confidential; plain **--all** keeps them so designers can tell the slides apart. **--preview --all --titles** highlights the titles as well.
- **--keywords file.txt** masks only the terms in the file, one per line. The built-in detection (emails, phones, IDs, IBANs, cards) is turned off. Name a term in the key with **term | Label**. An empty file stops the run and no deck is written.
- **--key-dir folder** saves the key somewhere else.
- **--force** restores even when the key was made for a different deck. Avoid it: every deck numbers its markers the same way, so another deck's key puts its values into the wrong places without any warning.

In the key workbook, **Restore key** sheet set **Restore(Y/N)** field to N for anything that should stay masked. The **Review notes** sheet lists what needs a manual check.

## Desktop app

Open **Sanitize** from the desktop or the Start menu (see **Installing** below). A small window offers four choices:

- **Mask keywords**: type the terms, one per line, then choose the deck.
- **Mask all**: choose the deck (**--all**).
- **Full masking**: choose the deck. Like **Mask all**, and slide titles become **[X]** too (**--all --titles**).
- **Restore**: choose the deck that came back, then its restore key.

Before masking, it says how many matches it found and on which slides; **Preview in PowerPoint** opens a highlighted copy to check first. Masking saves two files to Downloads: **deck_redacted.pptx** to send and **deck_restore_key.xlsx** to keep. Restore saves **deck_redacted_restored.pptx** and lists the slides to check for layout. Nothing is overwritten: when a name is taken, the deck and its key both get **(2)**.

- Everything runs on the computer. Nothing is sent anywhere or logged, and typed terms are forgotten when the window closes. Highlighted previews are temporary copies, deleted when Sanitize starts or closes.
- A key made for a different deck is refused, with no way to force it.

### Installing

Run **Sanitize Setup.exe**. It needs no admin rights and nothing installed first: Python comes inside it. It installs for the current user into **%LOCALAPPDATA%\Programs\Sanitize** and adds **Sanitize** to the Start menu and, unless unticked, the desktop. Remove it from **Settings > Apps**. To update, run a newer installer: it replaces the installed copy.

- The installer is not code-signed, so Windows first shows **Windows protected your PC**: choose **More info**, then **Run anyway**. On computers with Smart App Control turned on, Windows blocks it outright until it is signed.
- IT can install it silently: **"Sanitize Setup.exe" /VERYSILENT** for the signed-in user, or **/VERYSILENT /ALLUSERS** from an administrator prompt for everyone on the computer.

### Building the installer

**python build_installer.py** writes **Sanitize Setup.exe** (about 14 MB) one level up, next to this folder: the project folder shows only the installer and **Source**, which holds everything else and is the repository, so the installer is never committed. Rebuild it after changing **sanitize_pptx.py**, **sanitize_desktop.py** or the icon.

- It needs **Inno Setup 6.7** or later (**winget install JRSoftware.InnoSetup**) and is run with 64-bit Python 3.9 or later from **python.org**, with **tcl/tk and IDLE** ticked. That Python goes into the installer without its site-packages, test suite and IDLE, so its programs keep the Python Software Foundation's signature.
- The installed copy uses only its own Python: other Pythons on the computer, and their packages, can't interfere.
- **Sanitize.iss** is the Inno Setup script it compiles. Never change its **AppId**: that is what lets a newer installer update an installed copy instead of adding a second one.

To try a change before building, run **pythonw sanitize_desktop.py**.

The icon, in light and dark mode alike, is **sanitize.ico**: the app's card with the Prezlab **pl:** logomark, in the brand colors. Its source is **sanitize.svg**. After changing the SVG, run **python build_icon.py** to rebuild the .ico (Microsoft Edge renders each size), then rebuild the installer, which also renders its wizard images from the SVG.

## Browser view

**python sanitize_app.py** opens the same tool as a page in the browser, laid out like the add-in and with slide previews: open a deck, then choose **Mask Keywords**, **Mask All** or **Restore**. It has no **Full masking** option yet: use the desktop app or **--all --titles** for that. Preview highlights what will be masked before you confirm. Masking gives two downloads, the redacted deck and the restore key: browsers save them to Downloads, so move the key somewhere safe and never send it. Restore lists the slides to check for layout.

- It runs on this computer only (127.0.0.1) and loads nothing from the internet. Close the window running the script to stop it.
- Slide previews are exported with PowerPoint, so they need PowerPoint on Windows; everything else works without it. Only the first 60 slides are shown.
- A key made for a different deck is refused, with no way to force it.

## Using it from other code

These functions do the same work as the commands on files held in memory: decks and keys go in and come out as bytes, and nothing is read from or written to disk or printed. **sanitize_desktop.py** and **sanitize_app.py** are built on them.

The PowerPoint add-in does not run this file. It uses a C# port that lives with the add-in's code, not in this repository, so a change to the masking or restore logic here needs the same change there. The port has no **--titles** yet.

| Call | Same as | Returns |
|---|---|---|
| **preview_deck(deck, keywords=None, full=False, titles=False)** | **--preview** (**full=True**: **--preview --all**, add **titles=True** for **--titles**) | **deck** (highlighted), **findings**, **notes**, **summary** |
| **mask_all(deck, name, keywords=None, titles=False)** | **--all** (**titles=True**: **--all --titles**) | **deck**, **key** (.xlsx), **notes**, **summary**, **messages** |
| **mask_keywords(deck, keywords, name)** | **--keywords** | **deck**, **key** (.xlsx), **notes**, **summary**, **messages** |
| **restore_deck(deck, key, force=False)** | **--restore** | **deck**, **summary**, **messages** |

- **keywords** is the keyword list as text, in the same format as keywords.txt. In **preview_deck** and **mask_all** it is optional and works like adding **--keywords** to the command: **None** (not an empty string) uses the built-in detection.
- **findings** are the rows of the preview report: location, type and value (**[X]** plus its last 4 characters).
- **notes** are what to check by hand (text in images, embedded objects and so on), the same as the key's **Review notes** sheet. Show them in any interface built on these functions: nobody opens the key to read them.
- **messages** are the warnings and notes the commands print (not found, will not restore cleanly, check layout and so on). Each has a **level** (warning or note), a **code**, the **text** and the affected **items** as data. Show them to the user: they are the safety net.
- A problem that stops the run raises **SanitizeError** with a **code**: **not_pptx**, **not_key**, **wrong_key**, **empty_keywords**, **keyword_label** (a label ending in a digit), **too_many**. On **wrong_key**, ask for the matching key. Don't offer a forced restore in an interface (see **--force** above).
- **mask_all** and **mask_keywords** take an optional **restore_hint**, the "To restore" line in the key's About sheet. The default doesn't mention the command line.

## What gets masked

**Built-in detection** (what **--preview** highlights on its own, and what **--all** masks in slide titles): emails, UAE/KSA/Jordan/Qatar phone numbers and IBANs, payment cards, Emirates ID, Saudi ID/Iqama, Qatar QID, Jordan national number and passport numbers.

**--keywords mode:** only the terms in your file are masked, each match becoming a single **[X]** whatever its length, so a slide or table full of values stays readable (**Project Falcon** becomes **[X]**). Everything else stays readable, including emails, phones, IDs, IBANs and cards. With **--preview** it highlights only your terms; with **--all** it limits what is masked in slide titles to your terms.

**--all mode:** all text except slide titles becomes **[X]**: one per paragraph, plus one for each line break and each formatting change, so a bold lead-in keeps its bold on restore (**Note:** driven by retail becomes **[X]** [X]). Chart labels become **[X]**, and chart and Excel numbers become random placeholders. Web links point to a placeholder address.

**Full masking (--all --titles):** the same as **--all**, and slide titles become **[X]** too, on slides, layouts and the master. Headers, footers, dates typed into footers, and text on layouts and the master are masked in both modes. Only automatic fields stay as they are (the slide number and an automatic date): PowerPoint fills those in itself every time it shows the slide.

**All modes** cover slides, notes, layouts, master, comments, alt text, SmartArt, charts and embedded Excel data. They also set author, last modified by, company and manager in file properties and comment author names to **[X]**, and blank the file preview image. **[X]** is the only mask used anywhere; the two exceptions are chart and Excel numbers in **--all** (charts need numbers to draw) and link addresses (they must stay valid web addresses).

## Limitations

### Not masked in any mode

- **Charts drawn with shapes.** The text is masked, but bar heights, pie angles and line paths keep the true proportions. No warning is given.
- **Newer chart types** (waterfall, treemap, sunburst, funnel, histogram, box and whisker, map). The labels and values shown on the chart stay original, even with **--all**.
- **Images.** Text in pictures, screenshots, logos and ink drawings is untouched. Slides with images are flagged, but images on layouts and the master (where logos usually sit) are not.
- **Embedded Excel or Word objects.** The slide shows a snapshot picture of the original; only the file behind it is masked.
- **Other embedded files.** PDFs and legacy Office files are flagged but not scanned. Audio and video are neither.
- **Linked file paths**, such as a linked chart's path.
- **Names that are not on the slide itself:** shape names in the Selection Pane, Excel sheet names and formulas, chart number formats (**"SAR "#,##0**).
- **The file name.** The output is **<original name>_redacted.pptx**. Rename it if the name identifies the client.

### Detection

- The built-in detection only finds the patterns listed above. Names, companies, addresses, amounts and project names are missed in previews and slide titles.
- **--keywords** finds only your terms, so the patterns above are not detected in the same run. The two cannot be combined.
- Phones and IBANs from other countries (Kuwait, Bahrain, Oman, Egypt and so on) are missed.
- Passport, Jordan national and Qatar QID numbers are only caught with a label such as "passport" or "national no" within 40 characters before them.
- Keywords match inside other words (**ADI** also masks part of "Tradition") and need the exact spacing. Each Arabic spelling or diacritics variant needs its own line.
- A value split across two text boxes or paragraphs is not detected.

### **--all** mode

- Slide titles stay readable; only detected values in them are masked (only your terms, with **--keywords**). A title typed in a regular text box is masked like body text. Add **--titles** to mask titles too.
- With **--titles**, designers see only **[X]** on every slide, titles included, so they have nothing to tell the slides apart. Send the slide order or a brief with the deck if they need it.
- Only the structure survives: paragraphs, line breaks and where the formatting changes. Designers cannot see how long each text is, so restored slides need a layout check (restore lists them).
- **[X]** is wider than a single digit. In small shapes such as numbered circles it wraps onto several lines.
- Runs that differ only in spelling language are shown as one **[X]** and come back as one run, in the first run's language.
- Chart placeholders (101 to 999) ignore fixed axis limits and number formats.
### Restore

- Each masked text ends in an invisible marker. Retyping the text, or deleting its last character, breaks the marker. That value is then reported as not found and has to be restored by hand from the key. For **[X]**s from **--keywords** and in **--all** slide titles, restore suggests likely matches by position for someone to confirm. Decks masked by an older version of this tool still restore with their keys, but each masked text comes back whole in one formatting, and the slides are listed to check formatting. Their comment authors have to be restored by hand from the key.
- If a masked text is split across several text boxes, the full original goes back into the box holding the marker and the others keep their **[X]**. The slide is listed as edited.
- Chart numbers restore by their placeholder values. A chart with retyped data cannot be restored, and a new value that happens to equal a placeholder gets swapped.
- Restored text can be longer than what the designer laid out. Affected slides are listed for a layout check. This is most slides with masked text, since **[X]** is shorter than almost anything it hides.
- If the slides are moved into a new file, the key cannot be matched to the deck automatically. Restore still works, with a warning.
