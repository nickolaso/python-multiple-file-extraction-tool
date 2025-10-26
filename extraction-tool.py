#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import shutil
import tempfile
import threading
import subprocess
import platform
from pathlib import Path
from typing import Tuple
from zipfile import ZipFile, BadZipFile
import tarfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional pure-Python fallbacks (used if present)
try:
    import py7zr  # for .7z
except Exception:
    py7zr = None

try:
    import rarfile  # for .rar (often needs 'unrar' or 'bsdtar' installed)
except Exception:
    rarfile = None

APP_TITLE = "Py Extraction Tool"
UNARCHIVED_DIRNAME = "unarchived"
OVERWRITE_EXISTING = False      # set True to overwrite instead of auto-rename

# Optional manual overrides if tools aren’t on PATH:
SEVENZ_PATH_OVERRIDE  = r""     # e.g. r"C:\Program Files\7-Zip\7z.exe"  or "/opt/homebrew/bin/7z"
BSDTAR_PATH_OVERRIDE  = r""     # e.g. "/usr/bin/bsdtar" (macOS: /usr/bin/tar is usually bsdtar)
UNRAR_PATH_OVERRIDE   = r""     # e.g. "/usr/bin/unrar"
UNAR_PATH_OVERRIDE    = r""     # e.g. "/usr/local/bin/unar"

# -----------------------------
# Helpers: safety & naming
# -----------------------------
def safe_member_target(base: Path, member_name: str) -> Path:
    parts = [p for p in Path(member_name).parts if p not in ("", ".", "..")]
    target = base.joinpath(*parts) if parts else base
    base_resolved = base.resolve()
    try:
        target_resolved = target.resolve()
    except FileNotFoundError:
        target_resolved = (base / Path(*parts)).absolute()
    if not str(target_resolved).startswith(str(base_resolved)):
        target = base / Path(member_name).name
    return target

def unique_file(path: Path) -> Path:
    if OVERWRITE_EXISTING or not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 1
    while True:
        candidate = path.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1

def merge_tree_flat(src_root: Path, dest_root: Path) -> int:
    moved = 0
    for p in sorted(src_root.rglob("*")):
        rel = p.relative_to(src_root)
        target = dest_root / rel
        if p.is_dir():
            target.mkdir(parents=True, exist_ok=True)
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        final = unique_file(target)
        shutil.move(str(p), str(final))
        moved += 1
    return moved

# -----------------------------
# Tool detection
# -----------------------------
def find_7z_exe() -> str | None:
    cands = []
    if SEVENZ_PATH_OVERRIDE:
        cands.append(SEVENZ_PATH_OVERRIDE)
    for name in ("7z", "7za", "7zr"):
        exe = shutil.which(name)
        if exe: cands.append(exe)
    sysname = platform.system()
    if sysname == "Windows":
        cands += [r"C:\Program Files\7-Zip\7z.exe", r"C:\Program Files (x86)\7-Zip\7z.exe"]
    elif sysname == "Darwin":
        cands += ["/opt/homebrew/bin/7z", "/usr/local/bin/7z", "/usr/bin/7z"]
    else:
        cands += ["/usr/bin/7z", "/usr/local/bin/7z"]
    for c in cands:
        if c and Path(c).exists():
            return c
    return None

def find_bsdtar_exe() -> str | None:
    if BSDTAR_PATH_OVERRIDE and Path(BSDTAR_PATH_OVERRIDE).exists():
        return BSDTAR_PATH_OVERRIDE
    for name in ("bsdtar", "tar"):
        exe = shutil.which(name)
        if not exe:
            continue
        try:
            out = subprocess.run([exe, "--version"], stdout=subprocess.PIPE,
                                 stderr=subprocess.STDOUT, text=True)
            if "libarchive" in out.stdout.lower() or "bsdtar" in out.stdout.lower():
                return exe
            if platform.system() == "Darwin" and Path(exe).as_posix() == "/usr/bin/tar":
                return exe
        except Exception:
            pass
    return None

def find_unrar_exe() -> str | None:
    if UNRAR_PATH_OVERRIDE and Path(UNRAR_PATH_OVERRIDE).exists():
        return UNRAR_PATH_OVERRIDE
    return shutil.which("unrar")

def find_unar_exe() -> str | None:
    if UNAR_PATH_OVERRIDE and Path(UNAR_PATH_OVERRIDE).exists():
        return UNAR_PATH_OVERRIDE
    return shutil.which("unar")

# -----------------------------
# Format detection
# -----------------------------
def is_zip(p: Path) -> bool: return p.suffix.lower() == ".zip"
def is_7z(p: Path)  -> bool: return p.suffix.lower() == ".7z"
def is_rar(p: Path) -> bool: return p.suffix.lower() == ".rar"
def is_tar_like(path: Path) -> bool:
    name = path.name.lower()
    return (
        name.endswith(".tar") or name.endswith(".tar.gz") or name.endswith(".tgz") or
        name.endswith(".tar.bz2") or name.endswith(".tbz2") or
        name.endswith(".tar.xz") or name.endswith(".txz")
    )

def archive_list(root: Path) -> list[Path]:
    exts = (".zip", ".7z", ".rar", ".tar", ".tgz", ".tbz2", ".txz", ".tar.gz", ".tar.bz2", ".tar.xz")
    return sorted([p for p in root.iterdir() if p.is_file() and p.name.lower().endswith(exts)])

# -----------------------------
# Extractors (flat into dest)
# -----------------------------
def extract_zip_flat(archive: Path, dest: Path) -> int:
    with ZipFile(archive, "r") as zf:
        written = 0
        for info in zf.infolist():
            name = info.filename
            if not name:
                continue
            if name.endswith("/"):
                safe_dir = safe_member_target(dest, name)
                safe_dir.mkdir(parents=True, exist_ok=True)
                continue
            target = safe_member_target(dest, name)
            target.parent.mkdir(parents=True, exist_ok=True)
            target = unique_file(target)
            with zf.open(info, "r") as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out)
            written += 1
        return written

def extract_tar_flat(archive: Path, dest: Path) -> int:
    name = archive.name.lower()
    if name.endswith((".tar.gz", ".tgz")):
        mode = "r:gz"
    elif name.endswith((".tar.bz2", ".tbz2")):
        mode = "r:bz2"
    elif name.endswith((".tar.xz", ".txz")):
        mode = "r:xz"
    elif name.endswith(".tar"):
        mode = "r:"
    else:
        mode = "r:*"
    written = 0
    with tarfile.open(archive, mode) as tf:
        for m in tf.getmembers():
            if m.issym() or m.islnk():
                continue
            name = m.name
            if m.isdir():
                safe_dir = safe_member_target(dest, name)
                safe_dir.mkdir(parents=True, exist_ok=True)
                continue
            try:
                src_f = tf.extractfile(m)
            except Exception:
                continue
            if not src_f:
                continue
            target = safe_member_target(dest, name)
            target.parent.mkdir(parents=True, exist_ok=True)
            target = unique_file(target)
            with src_f, open(target, "wb") as out:
                shutil.copyfileobj(src_f, out)
            written += 1
    return written

def extract_via_7z_cli(archive: Path, dest: Path, sevenz: str) -> Tuple[int, str | None]:
    tmpdir = Path(tempfile.mkdtemp(prefix="unarch_7z_"))
    try:
        cmd = [sevenz, "x", "-y", f"-o{tmpdir}", str(archive)]
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        if proc.returncode != 0:
            return (0, f"7z failed (code {proc.returncode}). Output:\n{proc.stdout}")
        moved = merge_tree_flat(tmpdir, dest)
        return (moved, None)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

def extract_via_bsdtar_cli(archive: Path, dest: Path, bsdtar: str) -> Tuple[int, str | None]:
    tmpdir = Path(tempfile.mkdtemp(prefix="unarch_bsdtar_"))
    try:
        cmd = [bsdtar, "-x", "-f", str(archive), "-C", str(tmpdir)]
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        if proc.returncode != 0:
            return (0, f"bsdtar failed (code {proc.returncode}). Output:\n{proc.stdout}")
        moved = merge_tree_flat(tmpdir, dest)
        return (moved, None)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

def extract_via_unrar_cli(archive: Path, dest: Path, unrar: str) -> Tuple[int, str | None]:
    tmpdir = Path(tempfile.mkdtemp(prefix="unarch_unrar_"))
    try:
        cmd = [unrar, "x", "-o+", str(archive), str(tmpdir)]
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        if proc.returncode not in (0, 1):
            return (0, f"unrar failed (code {proc.returncode}). Output:\n{proc.stdout}")
        moved = merge_tree_flat(tmpdir, dest)
        return (moved, None)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

def extract_via_unar_cli(archive: Path, dest: Path, unar: str) -> Tuple[int, str | None]:
    tmpdir = Path(tempfile.mkdtemp(prefix="unarch_unar_"))
    try:
        cmd = [unar, "-quiet", "-force-overwrite", "-output-directory", str(tmpdir), str(archive)]
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        if proc.returncode != 0:
            return (0, f"unar failed (code {proc.returncode}). Output:\n{proc.stdout}")
        moved = merge_tree_flat(tmpdir, dest)
        return (moved, None)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

def extract_7z_py(archive: Path, dest: Path) -> int:
    if not py7zr:
        raise RuntimeError("py7zr not installed")
    tmpdir = Path(tempfile.mkdtemp(prefix="unarch7z_"))
    try:
        with py7zr.SevenZipFile(archive, mode="r") as z:
            z.extractall(path=tmpdir)
        return merge_tree_flat(tmpdir, dest)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

def extract_rar_py(archive: Path, dest: Path) -> int:
    if not rarfile:
        raise RuntimeError("rarfile not installed")
    tmpdir = Path(tempfile.mkdtemp(prefix="unarchrar_"))
    try:
        with rarfile.RarFile(archive) as rf:
            rf.extractall(tmpdir)
        return merge_tree_flat(tmpdir, dest)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

def extract_archive_flat(archive: Path, dest: Path) -> Tuple[int, str | None]:
    sevenz = find_7z_exe()
    bsdtar = find_bsdtar_exe()
    unrar = find_unrar_exe()
    unar  = find_unar_exe()

    try:
        if is_zip(archive):
            return (extract_zip_flat(archive, dest), None)
        if is_tar_like(archive):
            return (extract_tar_flat(archive, dest), None)
        if is_7z(archive) or is_rar(archive):
            # prefer robust CLIs first
            if sevenz:
                return extract_via_7z_cli(archive, dest, sevenz)
            if bsdtar:
                return extract_via_bsdtar_cli(archive, dest, bsdtar)
            if is_rar(archive) and unrar:
                return extract_via_unrar_cli(archive, dest, unrar)
            if is_rar(archive) and unar:
                return extract_via_unar_cli(archive, dest, unar)
            # fallbacks
            if is_7z(archive) and py7zr:
                return (extract_7z_py(archive, dest), None)
            if is_rar(archive) and rarfile:
                return (extract_rar_py(archive, dest), None)
            need = []
            if is_7z(archive): need.append("7-Zip (7z) or py7zr")
            if is_rar(archive): need.append("7-Zip (7z) / bsdtar / unrar / unar / rarfile")
            return (0, "No extractor available. Install " + " / ".join(need) + ".")
        # Unknown → try general tools
        if sevenz:
            return extract_via_7z_cli(archive, dest, sevenz)
        if bsdtar:
            return extract_via_bsdtar_cli(archive, dest, bsdtar)
        return (0, f"Unsupported archive type: {archive.name}")
    except BadZipFile:
        return (0, f"Corrupt/invalid archive: {archive.name}")
    except tarfile.TarError as e:
        return (0, f"Tar error on {archive.name}: {e}")
    except Exception as e:
        return (0, f"Error on {archive.name}: {e}")

# -----------------------------
# Core: extract all (flat)
# -----------------------------
def extract_all_in_folder_flat(root: Path, progress_cb, log_cb) -> Path:
    unarchived_dir = root / UNARCHIVED_DIRNAME
    unarchived_dir.mkdir(exist_ok=True)

    archives = archive_list(root)
    total = len(archives)
    # Log tool availability once
    log_cb(
        f"Tools → 7z: {find_7z_exe() or 'no'}, "
        f"bsdtar: {find_bsdtar_exe() or 'no'}, "
        f"unrar: {find_unrar_exe() or 'no'}, "
        f"unar: {find_unar_exe() or 'no'}, "
        f"py7zr: {'yes' if py7zr else 'no'}, "
        f"rarfile: {'yes' if rarfile else 'no'}"
    )

    if total == 0:
        log_cb("ℹ️ No archives found in the selected folder.")
        progress_cb(0, 0)
        return unarchived_dir

    log_cb(f"📦 Found {total} archive(s). Extracting to: {unarchived_dir}")
    success = 0
    failed = 0
    total_written = 0

    for idx, arc in enumerate(archives, start=1):
        written, err = extract_archive_flat(arc, unarchived_dir)
        if err:
            log_cb(f"❌ {arc.name}: {err}")
            failed += 1
        else:
            log_cb(f"✅ {arc.name} → unarchived ({written} file(s))")
            success += 1
            total_written += written
        progress_cb(idx, total)

    log_cb(f"\nDone. ✅ {success} succeeded, ❌ {failed} failed. Files written: {total_written}")
    return unarchived_dir

# -----------------------------
# Dark Theme
# -----------------------------
def apply_dark_theme(root: tk.Tk):
    root.configure(bg="#0f1115")
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    dark_bg = "#0f1115"
    panel_bg = "#151821"
    text_fg = "#e5e7eb"
    muted_fg = "#9aa4b2"
    accent = "#4f46e5"
    accent_dim = "#3b3bc7"

    style.configure(".", background=dark_bg, foreground=text_fg, fieldbackground=panel_bg)
    style.configure("TFrame", background=dark_bg)
    style.configure("Card.TFrame", background=panel_bg, relief="flat")
    style.configure("TLabel", background=dark_bg, foreground=text_fg)
    style.configure("Muted.TLabel", background=dark_bg, foreground=muted_fg)
    style.configure("Card.TLabel", background=panel_bg, foreground=text_fg)
    style.configure("TButton", background=accent, foreground="#ffffff", padding=8, borderwidth=0)
    style.map("TButton",
              background=[("active", accent_dim), ("disabled", "#2d2f39")],
              relief=[("pressed", "flat"), ("!pressed", "flat")])
    style.configure("Accent.TButton", background=accent)
    style.configure("TProgressbar", troughcolor="#0b0d12", background=accent)

# -----------------------------
# GUI
# -----------------------------
class App(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=16, style="TFrame")
        self.master = master
        self.pack(fill="both", expand=True)

        header = ttk.Frame(self, style="TFrame"); header.pack(fill="x", pady=(0, 12))
        ttk.Label(header, text="🗂️  Unzipper", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text=f"Select a folder and extract all archives (zip/rar/7z/tar*) into a single “{UNARCHIVED_DIRNAME}” directory.",
            style="Muted.TLabel",
        ).pack(anchor="w")

        card = ttk.Frame(self, padding=16, style="Card.TFrame"); card.pack(fill="x", pady=(0, 12))
        chooser = ttk.Frame(card, style="Card.TFrame"); chooser.pack(fill="x")
        self.path_var = tk.StringVar(value=str(Path.cwd()))
        ttk.Label(chooser, text="Destination Folder:", style="Card.TLabel").pack(anchor="w")
        row = ttk.Frame(chooser, style="Card.TFrame"); row.pack(fill="x", pady=(6, 0))
        self.entry = ttk.Entry(row, textvariable=self.path_var); self.entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Browse…", command=self.pick_folder).pack(side="left", padx=(8, 0))

        controls = ttk.Frame(card, style="Card.TFrame"); controls.pack(fill="x", pady=(12, 0))
        self.start_btn = ttk.Button(controls, text="Start Extraction", style="Accent.TButton", command=self.start)
        self.start_btn.pack(side="left")
        self.status_label = ttk.Label(controls, text="Ready.", style="Card.TLabel"); self.status_label.pack(side="right")

        prog = ttk.Frame(card, style="Card.TFrame"); prog.pack(fill="x", pady=(12, 0))
        self.progress = ttk.Progressbar(prog, mode="determinate"); self.progress.pack(fill="x")

        log_card = ttk.Frame(self, padding=16, style="Card.TFrame"); log_card.pack(fill="both", expand=True)
        ttk.Label(log_card, text="Activity Log", style="Card.TLabel").pack(anchor="w", pady=(0, 8))
        self.log = tk.Text(log_card, height=12, wrap="word", bd=0, highlightthickness=0)
        self.log.pack(fill="both", expand=True)
        self.log.configure(bg="#0c0f14", fg="#e5e7eb", insertbackground="#e5e7eb")
        self.log.tag_configure("muted", foreground="#9aa4b2")

        footer = ttk.Frame(self, style="TFrame"); footer.pack(fill="x", pady=(8, 0))
        ttk.Label(footer, text="Tip: Place your archives directly in the selected folder.",
                  style="Muted.TLabel").pack(anchor="w")

        self.worker = None

    def pick_folder(self):
        initial = self.path_var.get() or str(Path.home())
        chosen = filedialog.askdirectory(initialdir=initial, title="Choose a folder")
        if chosen:
            self.path_var.set(chosen)

    def set_progress(self, current, total):
        self.progress["maximum"] = max(total, 1)
        self.progress["value"] = current
        self.status_label.configure(text=f"Progress: {current}/{total}" if total else "Ready.")

    def log_line(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.master.update_idletasks()

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        folder = Path(self.path_var.get().strip() or ".").resolve()
        if not folder.exists() or not folder.is_dir():
            messagebox.showerror(APP_TITLE, f"Folder not found:\n{folder}")
            return

        self.log.delete("1.0", "end")
        self.set_progress(0, 1)
        self.status_label.configure(text="Working…")
        self.start_btn.state(["disabled"])

        def run():
            try:
                dest = extract_all_in_folder_flat(
                    folder,
                    progress_cb=lambda c, t: self.master.after(0, self.set_progress, c, t),
                    log_cb=lambda msg: self.master.after(0, self.log_line, msg),
                )
                self.master.after(0, self.on_done, dest)
            except Exception as e:
                self.master.after(0, self.on_error, str(e))

        self.worker = threading.Thread(target=run, daemon=True)
        self.worker.start()

    def on_done(self, dest: Path):
        self.status_label.configure(text="Complete.")
        self.start_btn.state(["!disabled"])
        try:
            if platform.system() == "Windows":
                os.startfile(dest)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", str(dest)])
            else:
                subprocess.Popen(["xdg-open", str(dest)])
        except Exception:
            pass
        messagebox.showinfo(APP_TITLE, f"All done!\nUnarchived folder:\n{dest}")

    def on_error(self, err: str):
        self.status_label.configure(text="Error.")
        self.start_btn.state(["!disabled"])
        messagebox.showerror(APP_TITLE, f"An error occurred:\n{err}")

def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("760x560"); root.minsize(640, 460)
    try:
        root.iconbitmap("")  # optional: path to .ico
    except Exception:
        pass
    apply_dark_theme(root)
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
