"""Desktop app for sanitize_pptx.py: the "Sanitize" desktop shortcut opens it through pythonw.exe (no console window).

A small window offers Mask keywords, Mask all, Full masking (all text, slide titles included) and Restore; the new
files are saved to Downloads.
Everything runs on this computer: nothing is sent anywhere, logged or kept. Typed terms are forgotten when the
window closes, and highlighted previews are temporary copies, deleted when Sanitize starts or closes.

It is installed with "Sanitize Setup.exe", built by build_installer.py, which brings its own copy of Python. To try a
change before building, run pythonw sanitize_desktop.py.
"""

from __future__ import annotations

import ctypes
import os
import re
import subprocess
import tempfile
import threading
import time
from pathlib import Path
from tkinter import Button, Frame, Label, Text, Tk, Toplevel, filedialog, font, messagebox, ttk

import sanitize_pptx as sp

APP_NAME = "Sanitize"
HERE = Path(__file__).resolve().parent
ICON = HERE / "sanitize.ico"  # window, taskbar and shortcut icon, built from sanitize.svg by build_icon.py
PREVIEW_DIR = Path(tempfile.gettempdir()) / "Sanitize previews"
RESTORE_HINT = "Open Sanitize from the desktop, choose Restore, then pick the returned deck and this key."
FRIENDLY = {"wrong_key": "This key was made for a different deck. Pick the key saved when this deck was masked."}
DECKS = [("PowerPoint decks", "*.pptx")]
KEYS = [("Restore keys", "*.xlsx")]

# Prezlab corporate palette (Brand Guidelines 2026, p. 21), as in ICP Finder
DARK_TEAL = "#002528"
LIGHT_TEAL = "#b6edf3"
SLATE = "#62848c"
OFF_WHITE = "#f4f6f6"
WHITE = "#ffffff"
RULE = "#d0dbdd"  # Slate tint, for dividers and the input border
ERROR = "#b3261e"


def downloads_dir() -> Path:
    """The user's Downloads folder, following it if Windows has it relocated."""
    try:
        import winreg

        key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            value, _ = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")
        path = Path(os.path.expandvars(value))
        if path.is_dir():
            return path
    except (ImportError, OSError):
        pass
    return Path.home() / "Downloads"


try:  # own handle with 64-bit safe signatures, for the window icons
    _user32 = ctypes.WinDLL("user32")
    _user32.LoadImageW.restype = ctypes.c_void_p
    _user32.LoadImageW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint, ctypes.c_int, ctypes.c_int,
                                   ctypes.c_uint]
    _user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_void_p]
except AttributeError:  # not Windows
    _user32 = None


def window_icons() -> dict[int, int]:
    """Icon handles from sanitize.ico at the sizes Windows asks for: small for title bars, big for the taskbar and
    Alt+Tab. Empty when the file or the Windows calls aren't available: the windows then keep Tk's icon."""
    if _user32 is None or not ICON.exists():
        return {}
    icons = {}
    for which, metric in ((0, 49), (1, 11)):  # ICON_SMALL: SM_CXSMICON, ICON_BIG: SM_CXICON
        size = _user32.GetSystemMetrics(metric)
        icons[which] = _user32.LoadImageW(None, str(ICON), 1, size, size, 0x10)  # IMAGE_ICON, LR_LOADFROMFILE
    return icons if all(icons.values()) else {}


def slides_in(locations) -> list[int]:
    """Slide numbers named in report locations ("Slide 3", "Slide 3 chart"...), in order."""
    return sorted({int(m.group(1)) for loc in locations for m in [re.match(r"Slide (\d+)", loc or "")] if m})


def slide_list(numbers: list[int]) -> str:
    """'slide 4', 'slides 2, 5 and 7', or the first ten and how many more."""
    if len(numbers) == 1:
        return f"slide {numbers[0]}"
    shown = [str(n) for n in numbers[:10]]
    rest = len(numbers) - len(shown)
    return "slides " + (", ".join(shown) + f" and {rest} more" if rest else ", ".join(shown[:-1]) + " and " + shown[-1])


def free_paths(folder: Path, *names: str) -> list[Path]:
    """The names in folder, numbered together ("deck_redacted (2).pptx") when one is taken: nothing is overwritten,
    and a deck and its key keep matching names."""
    n = 1
    while True:
        paths = [folder / (name if n == 1 else f"{Path(name).stem} ({n}){Path(name).suffix}") for name in names]
        if not any(p.exists() for p in paths):
            return paths
        n += 1


def write_preview(name: str, data: bytes) -> Path:
    PREVIEW_DIR.mkdir(exist_ok=True)
    (path,) = free_paths(PREVIEW_DIR, Path(name).stem + "_preview.pptx")
    path.write_bytes(data)
    return path


def clear_previews() -> None:
    """Previews hold the deck's original text. One still open in PowerPoint can't be deleted: it goes next time."""
    for path in PREVIEW_DIR.glob("*.pptx"):
        try:
            path.unlink()
        except OSError:
            pass


# flows: ui is TkUI, or a stand-in with the same methods in tests
def mask_flow(ui, mode: str, terms: str, out_dir: Path) -> None:
    titles = mode == "full"  # Full masking: Mask all with slide titles masked too
    full = mode == "all" or titles
    deck = ui.pick_file("Choose the deck to mask", DECKS)
    if deck is None:
        return
    data = deck.read_bytes()
    preview = ui.busy(f"Checking {deck.name}",
                      lambda: sp.preview_deck(data, None if full else terms, full=full, titles=titles))
    found = [f for f in preview["findings"] if not f["type"].startswith("METADATA_")]  # file properties are always cleared
    if not found:
        ui.info(f"{deck.name} has no text to mask." if full else
                f"None of these terms appear in {deck.name}.\n\nCheck the spelling. Each Arabic spelling needs its own line.")
        return
    slides = slides_in(f["location"] for f in found)
    if full:
        charts = ", and chart numbers become placeholders" if any(f["type"] == "CHART_NUMBER" for f in found) else ""
        title = "Mask all text and titles?" if titles else "Mask all text?"
        text = ("All text, slide titles included, becomes [X]" if titles else "All text except slide titles becomes [X]")
        text += (f" on {slide_list(slides)}" if slides else "") + charts + "."
    else:
        title = f"Mask {len(found)} {'match' if len(found) == 1 else 'matches'}?"
        text = (f"Found on {slide_list(slides)}. " if slides else "") + "Each match becomes [X]."
    if not ui.confirm(title, text + "\n\nA restore key is saved so you can put the text back.",
                      preview=lambda: ui.open(write_preview(deck.name, preview["deck"]))):
        return

    if full:
        result = ui.busy(f"Masking {deck.name}",
                         lambda: sp.mask_all(data, deck.name, restore_hint=RESTORE_HINT, titles=titles))
    else:
        result = ui.busy(f"Masking {deck.name}",
                         lambda: sp.mask_keywords(data, terms, deck.name, restore_hint=RESTORE_HINT))
    out, key = free_paths(out_dir, deck.stem + "_redacted.pptx", deck.stem + "_restore_key.xlsx")
    out.write_bytes(result["deck"])
    key.write_bytes(result["key"])

    head = (f"All text, slide titles included, is masked in {deck.name}." if titles
            else f"All text except slide titles is masked in {deck.name}." if full
            else f"{len(found)} {'match' if len(found) == 1 else 'matches'} masked in {deck.name}.")
    lines = [head, "", f"Saved to {out_dir.name}:", "", out.name, "Send this one.", "",
             key.name, "Keep it on this computer and never send it: it holds the original text."]
    notes = len(result["notes"])
    if notes:
        lines += ["", f"Check {notes} {'place' if notes == 1 else 'places'} by hand, such as text in pictures. "
                      "The key's Review notes sheet lists them."]
    lines += [line for m in result["messages"] for line in ("", m["text"])]
    if ui.ask("\n".join(lines) + "\n\nShow the files?"):
        ui.reveal(out)


def restore_flow(ui, out_dir: Path) -> None:
    deck = ui.pick_file("Choose the deck that came back", DECKS)
    if deck is None:
        return
    key = ui.pick_file("Choose its restore key", KEYS)
    if key is None:
        return
    data, key_data = deck.read_bytes(), key.read_bytes()
    #never forced: another deck's key puts its values in the wrong places without any warning
    result = ui.busy(f"Restoring {deck.name}", lambda: sp.restore_deck(data, key_data))
    s = result["summary"]
    others = [line for m in result["messages"]
              if m["code"] != "check_layout" and (m["code"], m["level"]) != ("fields", "note")  # fields: all restored
              for line in ("", m["text"])]
    if s["to_restore"] and not s["restored"]:
        ui.error("\n".join([f"Nothing was restored in {deck.name}.", *others]))
        return
    (out,) = free_paths(out_dir, deck.stem + "_restored.pptx")
    out.write_bytes(result["deck"])

    lines = [f"Original text restored in {deck.name}." if s["restored"] == s["to_restore"]
             else f"{s['restored']} of {s['to_restore']} items restored in {deck.name}."]
    layout = slides_in(next((m["items"] for m in result["messages"] if m["code"] == "check_layout"), []))
    if layout:
        lines += ["", f"Check the layout on {slide_list(layout)}: the real text is usually longer than [X]."]
    lines += others + ["", f"Saved to {out_dir.name} as {out.name}."]
    if ui.ask("\n".join(lines) + "\n\nOpen it now?"):
        ui.open(out)


def run_app(ui, out_dir: Path) -> int:
    """The start window, then the chosen flow, until the window is closed."""
    while True:
        choice = ui.start()
        if choice is None:
            return 0
        mode, terms = choice
        try:
            if mode == "restore":
                restore_flow(ui, out_dir)
            else:
                mask_flow(ui, mode, terms, out_dir)
        except sp.SanitizeError as exc:
            ui.error(FRIENDLY.get(exc.code, str(exc)))
        except Exception as exc:  # shown, not lost: pythonw has no console
            ui.error(f"Something went wrong.\n\n{exc}")


def _brand_fonts(families: set[str]) -> dict[str, tuple]:
    """Figtree, the brand's functional typeface, with Segoe UI where it isn't installed."""

    def family(name: str, fallback: str) -> str:
        return name if name in families else fallback

    return {
        "title": (family("Figtree", "Segoe UI"), 20, "bold"),
        "heading": (family("Figtree SemiBold", "Segoe UI Semibold"), 12),
        "body": (family("Figtree Medium", "Segoe UI"), 10),
        "small": (family("Figtree Medium", "Segoe UI"), 9),
        "input": (family("Figtree Medium", "Segoe UI"), 11),
        "button": (family("Figtree SemiBold", "Segoe UI Semibold"), 10),
        "mark": (family("Figtree", "Segoe UI"), 15, "bold"),
    }


class TkUI:
    def __init__(self, start_dir: Path) -> None:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # sharp dialogs on high-DPI screens
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Prezlab.Sanitize")  # own taskbar icon
        except (AttributeError, OSError):
            pass
        self._root = Tk()
        self._root.withdraw()
        self._root.attributes("-topmost", True)  # open dialogs in front of other windows
        self._root.report_callback_exception = lambda _type, exc, _tb: self.error(f"Something went wrong.\n\n{exc}")
        self._fonts = _brand_fonts(set(font.families(self._root)))
        self._icons = window_icons()  # set on each window as it opens
        style = ttk.Style(self._root)
        style.theme_use("clam")  # the only ttk widget is the progress bar; clam lets it take the brand colors
        style.configure("Busy.Horizontal.TProgressbar", troughcolor=OFF_WHITE, background=DARK_TEAL,
                        bordercolor=OFF_WHITE, lightcolor=DARK_TEAL, darkcolor=DARK_TEAL, thickness=6)
        self._terms = ""  # kept while the app is open, never saved
        self._folders = {"deck": start_dir, "key": start_dir}

    def _button(self, parent: Frame | Toplevel, text: str, command, primary: bool = True) -> Button:
        bg, hover, fg = (DARK_TEAL, SLATE, OFF_WHITE) if primary else (OFF_WHITE, LIGHT_TEAL, DARK_TEAL)
        button = Button(
            parent, text=text, command=command, font=self._fonts["button"], bg=bg, fg=fg,
            activebackground=hover, activeforeground=fg, relief="flat", bd=0, padx=18, pady=7, cursor="hand2",
        )
        button.bind("<Enter>", lambda _event: button.config(bg=hover))
        button.bind("<Leave>", lambda _event: button.config(bg=bg))
        return button

    def _label(self, parent: Frame | Toplevel, value: str, style: str, color: str = DARK_TEAL,
               wrap: str = "4.3i", **pack) -> Label:
        label = Label(parent, text=value, font=self._fonts[style], fg=color, bg=WHITE, anchor="w", justify="left",
                      wraplength=wrap)
        label.pack(fill="x", **pack)
        return label

    def _window(self) -> Toplevel:
        window = Toplevel(self._root, bg=WHITE, padx=32, pady=28)
        window.title(APP_NAME)
        window.resizable(False, False)
        window.bind("<Escape>", lambda _event: window.destroy())
        return window

    def _show(self, window: Toplevel) -> None:
        """Center the window, bring it to the front without pinning it there, and wait until it closes."""
        self._root.eval(f"tk::PlaceWindow {window} center")
        window.update_idletasks()
        for which, icon in self._icons.items():
            _user32.SendMessageW(int(window.wm_frame(), 16), 0x0080, which, icon)  # WM_SETICON
        window.grab_set()
        window.attributes("-topmost", True)
        window.after_idle(window.attributes, "-topmost", False)
        window.focus_force()
        self._root.wait_window(window)

    def start(self) -> tuple[str, str] | None:
        """The start window: ("keywords", terms), ("all", ""), ("full", ""), ("restore", "") or None when closed."""
        window = self._window()
        chosen: list[tuple[str, str]] = []

        def rule() -> None:
            Frame(window, bg=RULE, height=1).pack(fill="x", pady=16)

        def section(title: str, about: str, button: str, command) -> None:
            row = Frame(window, bg=WHITE)
            row.pack(fill="x")
            words = Frame(row, bg=WHITE)
            words.pack(side="left", fill="x", expand=True)
            self._label(words, title, "heading", wrap="3i")
            self._label(words, about, "body", SLATE, wrap="3i", pady=(2, 0))
            self._button(row, button, command).pack(side="right", padx=(16, 0))

        def keep_terms() -> str:
            self._terms = box.get("1.0", "end").strip()
            return self._terms

        def choose(mode: str) -> None:
            terms = keep_terms()
            if mode == "keywords" and not terms:
                hint.config(text="Type at least one term first.", fg=ERROR)
                box.focus_set()
                return
            chosen.append((mode, terms))
            window.destroy()

        self._label(window, APP_NAME, "title")
        self._label(window, "Mask confidential text in a deck before it goes out, and put it back when it returns.",
                    "body", SLATE, pady=(4, 0))
        rule()
        self._label(window, "Mask keywords", "heading")
        self._label(window, "Only these terms become [X]. Type one per line.", "body", SLATE, pady=(2, 10))
        frame = Frame(window, bg=WHITE, highlightthickness=1, highlightbackground=RULE)
        frame.pack(fill="x")
        box = Text(
            frame, height=4, width=40, font=self._fonts["input"], relief="flat", bd=0, wrap="word", undo=True,
            bg=WHITE, fg=DARK_TEAL, insertbackground=DARK_TEAL, selectbackground=LIGHT_TEAL, selectforeground=DARK_TEAL,
        )
        box.pack(fill="x", padx=10, pady=7)
        box.insert("1.0", self._terms)
        box.bind("<FocusIn>", lambda _event: frame.config(highlightbackground=DARK_TEAL))
        box.bind("<FocusOut>", lambda _event: frame.config(highlightbackground=RULE))
        row = Frame(window, bg=WHITE)
        row.pack(fill="x", pady=(10, 0))
        hint = Label(row, text="Terms also match inside longer words.\nAdd each Arabic spelling on its own line.",
                     font=self._fonts["small"], fg=SLATE, bg=WHITE, anchor="w", justify="left")
        hint.pack(side="left", fill="x", expand=True)
        self._button(row, "Choose deck", lambda: choose("keywords")).pack(side="right", padx=(16, 0))
        rule()
        section("Mask all", "All text except slide titles becomes [X].", "Choose deck", lambda: choose("all"))
        rule()
        section("Full masking", "All text becomes [X], slide titles included.", "Choose deck", lambda: choose("full"))
        rule()
        section("Restore", "Put the original text back, using the deck's restore key.", "Choose returned deck",
                lambda: choose("restore"))
        rule()
        footer = Frame(window, bg=WHITE)
        footer.pack(fill="x")
        Label(footer, text="Saved to your Downloads folder. Nothing leaves this computer.", font=self._fonts["body"],
              fg=SLATE, bg=WHITE).pack(side="left")
        Label(footer, text="pl:", font=self._fonts["mark"], fg=DARK_TEAL, bg=WHITE).pack(side="right")

        window.protocol("WM_DELETE_WINDOW", lambda: (keep_terms(), window.destroy()))
        window.bind("<Escape>", lambda _event: (keep_terms(), window.destroy()))
        window.after_idle(box.focus_set)
        self._show(window)
        return chosen[0] if chosen else None

    def pick_file(self, title: str, types: list[tuple[str, str]]) -> Path | None:
        kind = "key" if types is KEYS else "deck"
        path = filedialog.askopenfilename(parent=self._root, title=title, initialdir=self._folders[kind], filetypes=types)
        if not path:
            return None
        self._folders[kind] = Path(path).parent
        return Path(path)

    def busy(self, text: str, work):
        """Run work in the background behind a small progress window; return its result or raise its error."""
        window = self._window()
        window.protocol("WM_DELETE_WINDOW", lambda: None)  # it closes itself when the work is done
        window.unbind("<Escape>")
        self._label(window, text, "heading", wrap="4i")
        self._label(window, "This takes a few seconds, longer for big decks.", "body", SLATE, pady=(2, 14))
        bar = ttk.Progressbar(window, mode="indeterminate", length=320, style="Busy.Horizontal.TProgressbar")
        bar.pack(fill="x")
        bar.start(12)
        outcome: dict = {}

        def run() -> None:
            try:
                outcome["value"] = work()
            except BaseException as exc:
                outcome["error"] = exc

        worker = threading.Thread(target=run, daemon=True)
        shown = time.monotonic()

        def poll() -> None:
            if worker.is_alive() or time.monotonic() - shown < 0.4:  # no flash on quick runs
                window.after(80, poll)
            else:
                window.destroy()

        worker.start()
        window.after(80, poll)
        self._show(window)
        if "error" in outcome:
            raise outcome["error"]
        return outcome["value"]

    def confirm(self, title: str, text: str, preview) -> bool:
        """Mask or cancel; Preview in PowerPoint opens the highlighted copy and leaves the question open."""
        window = self._window()
        chosen: list[bool] = []
        self._label(window, title, "heading", wrap="4i")
        self._label(window, text, "body", SLATE, wrap="4i", pady=(6, 18))

        def mask() -> None:
            chosen.append(True)
            window.destroy()

        def show() -> None:
            try:
                preview()
            except OSError as exc:
                self.error(f"Couldn't open the preview.\n\n{exc}", parent=window)

        buttons = Frame(window, bg=WHITE)
        buttons.pack(fill="x")
        self._button(buttons, "Cancel", window.destroy, primary=False).pack(side="right", padx=(8, 0))
        ok = self._button(buttons, "Mask", mask)
        ok.pack(side="right")
        self._button(buttons, "Preview in PowerPoint", show, primary=False).pack(side="left")
        window.after_idle(ok.focus_set)
        self._show(window)
        return bool(chosen)

    def info(self, message: str) -> None:
        messagebox.showinfo(APP_NAME, message, parent=self._root)

    def error(self, message: str, parent: Toplevel | None = None) -> None:
        messagebox.showerror(APP_NAME, message, parent=parent or self._root)

    def ask(self, message: str) -> bool:
        return messagebox.askyesno(APP_NAME, message, parent=self._root)

    def open(self, path: Path) -> None:
        os.startfile(path)

    def reveal(self, path: Path) -> None:
        subprocess.Popen(f'explorer /select,"{path}"')

    def close(self) -> None:
        self._root.destroy()


def main() -> int:
    out_dir = downloads_dir()
    clear_previews()
    ui = TkUI(out_dir)
    try:
        return run_app(ui, out_dir)
    finally:
        clear_previews()
        ui.close()


if __name__ == "__main__":
    raise SystemExit(main())
