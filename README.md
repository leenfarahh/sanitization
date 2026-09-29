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
| *Mask personal details* | **python sanitize_pptx.py deck.pptx** | **deck_redacted.pptx**, **deck_restore_key.xlsx**
| *Mask all text and numbers* | **python sanitize_pptx.py deck.pptx --all** | **deck_redacted.pptx**, **deck_restore_key.xlsx** |
| *Restore* | **python sanitize_pptx.py deck_redacted.pptx --restore deck_restore_key.xlsx** | **deck_redacted_restored.pptx** |

- **--keywords file.txt** adds your own terms, one per line. Names the term in the key **term | Label**.
- **--key-dir folder** saves the key somewhere else.
- **--force** restores even when the key was made for a different deck.

In the key workbook, **Restore key** sheet set **Restore(Y/N)** field to N for anything that should stay masked. The **Review notes** sheet lists what needs a manual check.

## What gets masked

**Default mode:** each detected value becomes a single **[X]**, whatever its length, so a slide or table full of values stays readable (**+971 50 123 4567** becomes **[X]**). It detects emails, UAE/KSA/Jordan/Qatar phone numbers and IBANs, payment cards, Emirates ID, Saudi ID/Iqama, Qatar QID, Jordan national number, passport numbers, and your keywords. Everything else stays readable.

**--all mode:** all text except slide titles becomes **[X]**: one per paragraph, plus one for each line break and each formatting change, so a bold lead-in keeps its bold on restore (**Note:** driven by retail becomes **[X]** [X]). Chart labels become **[X]**, and chart and Excel numbers become random placeholders. Web links point to a placeholder address.

**Both modes** cover slides, notes, layouts, master, comments, alt text, SmartArt, charts and embedded Excel data. They also set author, last modified by, company and manager in file properties and comment author names to **[X]**, and blank the file preview image. **[X]** is the only mask used anywhere; the two exceptions are chart and Excel numbers in **--all** (charts need numbers to draw) and link addresses (they must stay valid web addresses).

## Limitations

### Not masked in either mode

- **Charts drawn with shapes.** The text is masked, but bar heights, pie angles and line paths keep the true proportions. No warning is given.
- **Newer chart types** (waterfall, treemap, sunburst, funnel, histogram, box and whisker, map). The labels and values shown on the chart stay original, even with **--all**.
- **Images.** Text in pictures, screenshots, logos and ink drawings is untouched. Slides with images are flagged, but images on layouts and the master (where logos usually sit) are not.
- **Embedded Excel or Word objects.** The slide shows a snapshot picture of the original; only the file behind it is masked.
- **Other embedded files.** PDFs and legacy Office files are flagged but not scanned. Audio and video are neither.
- **Linked file paths**, such as a linked chart's path.
- **Names that are not on the slide itself:** shape names in the Selection Pane, Excel sheet names and formulas, chart number formats (**"SAR "#,##0**).
- **The file name.** The output is **<original name>_redacted.pptx**. Rename it if the name identifies the client.

### Default mode detection

- It only finds the patterns listed above. Names, companies, addresses, amounts and project names are missed unless they are keywords.
- Phones and IBANs from other countries (Kuwait, Bahrain, Oman, Egypt and so on) are missed.
- Passport, Jordan national and Qatar QID numbers are only caught with a label such as "passport" or "national no" within 40 characters before them.
- Keywords match inside other words (**ADI** also masks part of "Tradition") and need the exact spacing. Each Arabic spelling or diacritics variant needs its own line.
- A value split across two text boxes or paragraphs is not detected.

### **--all** mode

- Slide titles stay readable; only detected values and keywords in them are masked. A title typed in a regular text box is masked like body text.
- Only the structure survives: paragraphs, line breaks and where the formatting changes. Designers cannot see how long each text is, so restored slides need a layout check (restore lists them).
- **[X]** is wider than a single digit. In small shapes such as numbered circles it wraps onto several lines.
- Runs that differ only in spelling language are shown as one **[X]** and come back as one run, in the first run's language.
- Chart placeholders (101 to 999) ignore fixed axis limits and number formats.
### Restore

- Each masked text ends in an invisible marker. Retyping the text, or deleting its last character, breaks the marker. That value is then reported as not found and has to be restored by hand from the key. For default-mode **[X]**s, restore suggests likely matches by position for someone to confirm. Decks masked by an older version of this tool still restore with their keys, but each masked text comes back whole in one formatting, and the slides are listed to check formatting. Their comment authors have to be restored by hand from the key.
- If a masked text is split across several text boxes, the full original goes back into the box holding the marker and the others keep their **[X]**. The slide is listed as edited.
- Chart numbers restore by their placeholder values. A chart with retyped data cannot be restored, and a new value that happens to equal a placeholder gets swapped.
- Restored text can be longer than what the designer laid out. Affected slides are listed for a layout check. This is most slides with masked text, since **[X]** is shorter than almost anything it hides.
- If the slides are moved into a new file, the key cannot be matched to the deck automatically. Restore still works, with a warning.
