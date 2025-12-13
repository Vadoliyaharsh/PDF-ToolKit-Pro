"""
PDF TOOLKIT PRO — Single File (Professional Edition)
Features: SQLite Auth, Modern Enterprise UI, Full PDF Tools.
Update: Added Logout functionality.

Requirements:
 - Python 3.8+
 - PyPDF2
 - docx2pdf (optional)
 - tkinterdnd2 (optional)
"""

import os
import json
import threading
import queue
import time
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Authentication imports
import sqlite3
import hashlib
import binascii
import secrets
import datetime

# ----------------------------
# Dependency Checks
# ----------------------------

# Optional OS-level drag/drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# PDF libraries
try:
    from PyPDF2 import PdfMerger, PdfReader, PdfWriter
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# docx2pdf optional
try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False


# ----------------------------
# Config Persistence
# ----------------------------
CONFIG_PATH = os.path.expanduser("~/.pdf_toolkit_config.json")
DEFAULT_CONFIG = {
    "last_folder": os.path.expanduser("~"), 
    "theme": "light", 
    "window_size": (1200, 800)
}

def load_config():
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass

# Global config object
config = load_config()


# ----------------------------
# Authentication (SQLite + PBKDF2)
# ----------------------------
DB_PATH = os.path.expanduser("~/.pdf_toolkit_users.db")
PBKDF2_ITERATIONS = 150_000
HASH_NAME = 'sha256'

def init_db():
    conn = sqlite3.connect(DB_PATH)
    try:
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        ''')
        conn.commit()
    finally:
        conn.close()

def _hash_password(password: str, salt: bytes = None):
    if salt is None:
        salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac(HASH_NAME, password.encode('utf-8'), salt, PBKDF2_ITERATIONS)
    return binascii.hexlify(salt).decode('ascii'), binascii.hexlify(dk).decode('ascii')

def create_user(username: str, password: str):
    username = username.strip()
    if not username or not password:
        raise ValueError("Username and password are required")
    salt_hex, hash_hex = _hash_password(password)
    conn = sqlite3.connect(DB_PATH)
    try:
        c = conn.cursor()
        c.execute('INSERT INTO users (username, password_hash, salt, created_at) VALUES (?, ?, ?, ?)',
                  (username, hash_hex, salt_hex, datetime.datetime.utcnow().isoformat()))
        conn.commit()
    finally:
        conn.close()

def verify_user(username: str, password: str) -> bool:
    username = username.strip()
    conn = sqlite3.connect(DB_PATH)
    try:
        c = conn.cursor()
        c.execute('SELECT password_hash, salt FROM users WHERE username = ?', (username,))
        row = c.fetchone()
        if not row:
            return False
        hash_hex, salt_hex = row
        salt = binascii.unhexlify(salt_hex.encode('ascii'))
        _, attempt_hash = _hash_password(password, salt=salt)
        return secrets.compare_digest(attempt_hash, hash_hex)
    finally:
        conn.close()

init_db()


# ----------------------------
# PDF Modules
# ----------------------------
def module_merge_pdf(files, output, progress_callback=None):
    if not PDF_AVAILABLE:
        raise RuntimeError("PyPDF2 is not available. Please install it.")
    merger = PdfMerger(strict=False)
    total = max(1, len(files))
    for idx, pdf in enumerate(files, start=1):
        try:
            merger.append(pdf)
            if progress_callback:
                progress_callback(idx / total, f"Appending {os.path.basename(pdf)} ({idx}/{total})")
            time.sleep(0.01)
        except Exception as e:
            print(f"Error appending {pdf}: {e}")
    merger.write(output)
    merger.close()
    if progress_callback:
        progress_callback(1.0, "Merge complete")

def module_split_pdf(input_file, split_page, output_folder, progress_callback=None):
    if not PDF_AVAILABLE:
        raise RuntimeError("PyPDF2 is not available. Please install it.")
    reader = PdfReader(input_file)
    total_pages = len(reader.pages)
    
    if split_page >= total_pages:
        raise ValueError("Split page index out of range")

    writer1 = PdfWriter()
    writer2 = PdfWriter()
    
    for i in range(total_pages):
        if i < split_page:
            writer1.add_page(reader.pages[i])
        else:
            writer2.add_page(reader.pages[i])
        if progress_callback:
            progress_callback((i+1)/total_pages, f"Processing page {i+1}/{total_pages}")
    
    part1 = os.path.join(output_folder, "part1.pdf")
    part2 = os.path.join(output_folder, "part2.pdf")
    
    with open(part1, "wb") as f1:
        writer1.write(f1)
    with open(part2, "wb") as f2:
        writer2.write(f2)
        
    if progress_callback:
        progress_callback(1.0, f"Split complete: part1.pdf, part2.pdf")

def module_compress_pdf(input_file, output_file, progress_callback=None):
    if not PDF_AVAILABLE:
        raise RuntimeError("PyPDF2 is not available. Please install it.")
    reader = PdfReader(input_file)
    writer = PdfWriter()
    total = len(reader.pages) if reader.pages else 1
    
    for idx, page in enumerate(reader.pages, start=1):
        try:
            # Attempt to compress content streams
            page.compress_content_streams()
        except Exception:
            pass
        writer.add_page(page)
        
        # Reducing quality of images is hard with just PyPDF2, 
        # so this mainly compresses text streams and structure.
        if progress_callback:
            progress_callback(idx/total, f"Compressing page {idx}/{total}")
        time.sleep(0.01)
        
    with open(output_file, "wb") as f:
        writer.write(f)
    if progress_callback:
        progress_callback(1.0, "Compression complete")

def module_add_watermark(pdf_file, watermark_file, output_file, progress_callback=None):
    if not PDF_AVAILABLE:
        raise RuntimeError("PyPDF2 is not available. Please install it.")
    
    pdf_reader = PdfReader(pdf_file)
    watermark_reader = PdfReader(watermark_file)
    
    if not watermark_reader.pages:
        raise ValueError("Watermark PDF has no pages")
        
    watermark_page = watermark_reader.pages[0]
    writer = PdfWriter()
    total = len(pdf_reader.pages)
    
    for idx, page in enumerate(pdf_reader.pages, start=1):
        try:
            page.merge_page(watermark_page)
        except AttributeError:
            # Fallback for older PyPDF2 versions
            page.mergePage(watermark_page)
            
        writer.add_page(page)
        if progress_callback:
            progress_callback(idx/total, f"Watermarking page {idx}/{total}")
        time.sleep(0.01)
        
    with open(output_file, "wb") as f:
        writer.write(f)
    if progress_callback:
        progress_callback(1.0, "Watermark applied")

def module_word_to_pdf(input_file, output_file, progress_callback=None):
    if not DOCX2PDF_AVAILABLE:
        raise RuntimeError("docx2pdf is not installed or Word is missing.")
    
    if progress_callback:
        progress_callback(0.1, "Starting Word->PDF conversion...")
        
    # docx2pdf conversion is synchronous and blocking
    docx2pdf_convert(input_file, output_file)
    
    if progress_callback:
        progress_callback(1.0, "Conversion complete")


# ----------------------------
# ENHANCED THEME DEFINITIONS
# ----------------------------
FONT_HEADER = ("Segoe UI", 24, "bold")
FONT_SUBHEADER = ("Segoe UI", 11)
FONT_BODY = ("Segoe UI", 10)
FONT_BUTTON = ("Segoe UI", 10, "bold")
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO = ("Courier New", 10)

# "Aura Light"
LIGHT_THEME = {
    "app_bg": "#F8FAFC",
    "sidebar_bg": "#FFFFFF",
    "sidebar_fg": "#475569",
    "sidebar_active_bg": "#F1F5F9",
    "sidebar_active_fg": "#3B82F6",
    "content_bg": "#F8FAFC",
    "card_bg": "#FFFFFF",
    "text_main": "#0F172A",
    "text_muted": "#64748B",
    "primary_bg": "#3B82F6",
    "primary_hover": "#2563EB",
    "primary_fg": "#FFFFFF",
    "secondary_bg": "#F1F5F9",
    "secondary_hover": "#E2E8F0",
    "secondary_fg": "#334155",
    "accent_bg": "#8B5CF6",
    "accent_hover": "#7C3AED",
    "entry_bg": "#FFFFFF",
    "entry_border": "#CBD5E1",
    "entry_focus": "#3B82F6",
    "listbox_bg": "#FFFFFF",
    "listbox_select": "#E0E7FF",
    "danger_bg": "#EF4444",
    "danger_hover": "#DC2626",
    "success_bg": "#10B981",
    "success_hover": "#059669",
    "warning_bg": "#F59E0B",
    "warning_hover": "#D97706",
    "border_color": "#E2E8F0",
    "shadow_color": "#00000010",
}

# "Nebula Dark"
DARK_THEME = {
    "app_bg": "#0F172A",
    "sidebar_bg": "#1E293B",
    "sidebar_fg": "#94A3B8",
    "sidebar_active_bg": "#334155",
    "sidebar_active_fg": "#60A5FA",
    "content_bg": "#0F172A",
    "card_bg": "#1E293B",
    "text_main": "#F1F5F9",
    "text_muted": "#94A3B8",
    "primary_bg": "#3B82F6",
    "primary_hover": "#2563EB",
    "primary_fg": "#FFFFFF",
    "secondary_bg": "#334155",
    "secondary_hover": "#475569",
    "secondary_fg": "#F1F5F9",
    "accent_bg": "#8B5CF6",
    "accent_hover": "#7C3AED",
    "entry_bg": "#0F172A",
    "entry_border": "#475569",
    "entry_focus": "#3B82F6",
    "listbox_bg": "#0F172A",
    "listbox_select": "#334155",
    "danger_bg": "#EF4444",
    "danger_hover": "#DC2626",
    "success_bg": "#10B981",
    "success_hover": "#059669",
    "warning_bg": "#F59E0B",
    "warning_hover": "#D97706",
    "border_color": "#334155",
    "shadow_color": "#00000030",
}


# ----------------------------
# ENHANCED UI COMPONENTS
# ----------------------------
class ModernButton(tk.Button):
    """Enhanced button with gradient effect and smooth animations"""
    def __init__(self, master, text, command, variant="primary", theme=None, icon=None, **kwargs):
        super().__init__(master, text=text, command=command, **kwargs)
        self.variant = variant
        self.current_theme = theme if theme else LIGHT_THEME
        self.icon = icon
        
        # Base styling
        self.config(
            relief="flat",
            font=FONT_BUTTON,
            padx=kwargs.get("padx", 24),
            pady=kwargs.get("pady", 12),
            cursor="hand2",
            bd=0,
            highlightthickness=0
        )
        
        self.apply_colors()
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        
        if icon:
            self.config(text=f"  {icon}  {text}")

    def apply_colors(self):
        t = self.current_theme
        if self.variant == "primary":
            self.bg = t["primary_bg"]
            self.fg = t["primary_fg"]
            self.hover_bg = t["primary_hover"]
        elif self.variant == "accent":
            self.bg = t["accent_bg"]
            self.fg = "#ffffff"
            self.hover_bg = t["accent_hover"]
        elif self.variant == "danger":
            self.bg = t["danger_bg"]
            self.fg = "#ffffff"
            self.hover_bg = t["danger_hover"]
        elif self.variant == "success":
            self.bg = t["success_bg"]
            self.fg = "#ffffff"
            self.hover_bg = t["success_hover"]
        elif self.variant == "warning":
            self.bg = t["warning_bg"]
            self.fg = "#ffffff"
            self.hover_bg = t["warning_hover"]
        else: # secondary
            self.bg = t["secondary_bg"]
            self.fg = t["secondary_fg"]
            self.hover_bg = t["secondary_hover"]

        self.config(bg=self.bg, fg=self.fg, 
                    activebackground=self.hover_bg, 
                    activeforeground=self.fg)

    def on_enter(self, e):
        self.config(bg=self.hover_bg)
        
    def on_leave(self, e):
        self.config(bg=self.bg)
    
    def update_theme(self, new_theme):
        self.current_theme = new_theme
        self.apply_colors()


class ModernCard(tk.Frame):
    """Enhanced card with shadow and rounded corners"""
    def __init__(self, master, theme=None, **kwargs):
        super().__init__(master, **kwargs)
        self.current_theme = theme if theme else LIGHT_THEME
        
        # Configure main frame
        self.config(
            bg=self.current_theme["card_bg"],
            highlightbackground=self.current_theme["border_color"],
            highlightthickness=1,
            padx=30,
            pady=30
        )


class ModernEntry(tk.Frame):
    """Enhanced entry with modern styling and validation support"""
    def __init__(self, master, textvariable=None, theme=None, show=None, placeholder="", **kwargs):
        super().__init__(master)
        self.current_theme = theme if theme else LIGHT_THEME
        self.entry_var = textvariable if textvariable else tk.StringVar()
        self.placeholder = placeholder
        self.show_placeholder = bool(placeholder)
        self.show_char = show  # Store the password masking char
        
        # Container with border effect
        self.config(
            bg=self.current_theme["entry_border"],
            padx=1,
            pady=1
        )
        
        # Inner frame for background
        self.inner = tk.Frame(self, bg=self.current_theme["entry_bg"])
        self.inner.pack(fill="both", expand=True)
        
        # The Entry widget
        self.entry = tk.Entry(
            self.inner,
            textvariable=self.entry_var,
            # If we have a placeholder, don't mask it immediately
            show="" if self.show_placeholder else self.show_char,
            font=FONT_BODY,
            relief="flat",
            bd=0,
            bg=self.current_theme["entry_bg"],
            fg=self.current_theme["text_main"],
            insertbackground=self.current_theme["text_main"],
            highlightthickness=0
        )
        self.entry.pack(fill="x", expand=True, padx=12, pady=10)
        
        # Placeholder handling
        if placeholder:
            self.entry.insert(0, placeholder)
            self.entry.config(fg=self.current_theme["text_muted"])
            self.entry.bind("<FocusIn>", self._clear_placeholder)
            self.entry.bind("<FocusOut>", self._add_placeholder)
        else:
            self.entry.bind("<FocusIn>", self._on_focus)
            self.entry.bind("<FocusOut>", self._on_blur)
        
    def _clear_placeholder(self, e):
        if self.show_placeholder and self.entry.get() == self.placeholder:
            self.entry.delete(0, tk.END)
            self.entry.config(fg=self.current_theme["text_main"])
            # Restore mask char if needed
            if self.show_char:
                self.entry.config(show=self.show_char)
        self._on_focus(e)
        
    def _add_placeholder(self, e):
        if self.show_placeholder and not self.entry.get():
            self.entry.config(show="") # Remove mask to show placeholder text
            self.entry.insert(0, self.placeholder)
            self.entry.config(fg=self.current_theme["text_muted"])
        self._on_blur(e)
        
    def _on_focus(self, e):
        self.config(bg=self.current_theme["entry_focus"])
        
    def _on_blur(self, e):
        self.config(bg=self.current_theme["entry_border"])
        
    def get(self):
        text = self.entry.get()
        if self.show_placeholder and text == self.placeholder:
            return ""
        return text
        
    def insert(self, idx, s):
        if self.show_placeholder and self.entry.get() == self.placeholder:
            self.entry.delete(0, tk.END)
            self.entry.config(fg=self.current_theme["text_main"], show=self.show_char or "")
        self.entry.insert(idx, s)
        
    def delete(self, first, last=None):
        self.entry.delete(first, last)
        
    def update_theme(self, new_theme):
        self.current_theme = new_theme
        self.config(bg=self.current_theme["entry_border"])
        self.inner.config(bg=self.current_theme["entry_bg"])
        self.entry.config(
            bg=self.current_theme["entry_bg"], 
            fg=self.current_theme["text_main"],
            insertbackground=self.current_theme["text_main"]
        )


# ----------------------------
# AUTHENTICATION DIALOGS
# ----------------------------
class AuthDialog:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Login | PDF Toolkit Pro")
        self.root.geometry("480x620")
        self.root.resizable(False, False)
        self._center_window(self.root, 480, 620)
        
        self.t = LIGHT_THEME
        self.root.configure(bg="#EEF2FF")
        
        # Main Card
        card_frame = tk.Frame(self.root, bg=self.t["border_color"], padx=1, pady=1)
        card_frame.place(relx=0.5, rely=0.5, anchor="center", width=400)
        
        card = tk.Frame(card_frame, bg="white", padx=40, pady=40)
        card.pack(fill="both", expand=True)
        
        tk.Label(card, text="🔐", font=("Segoe UI", 42), bg="white").pack(pady=(0, 10))
        
        tk.Label(card, text="Welcome Back", 
                font=("Segoe UI", 22, "bold"), 
                bg="white", fg="#1F2937").pack()
        
        tk.Label(card, text="Sign in to your PDF Toolkit Pro", 
                font=FONT_SUBHEADER, 
                bg="white", fg="#6B7280").pack(pady=(0, 30))

        input_frame = tk.Frame(card, bg="white")
        input_frame.pack(fill="x", pady=(0, 20))
        
        tk.Label(input_frame, text="Username", 
                font=("Segoe UI", 10, "bold"), 
                bg="white", fg="#374151", anchor="w").pack(fill="x")
        
        self.username_var = tk.StringVar()
        self.u_entry = ModernEntry(input_frame, textvariable=self.username_var, 
                                 theme=self.t, placeholder="Enter your username")
        self.u_entry.pack(fill="x", pady=(5, 15))

        tk.Label(input_frame, text="Password", 
                font=("Segoe UI", 10, "bold"), 
                bg="white", fg="#374151", anchor="w").pack(fill="x")
        
        self.password_var = tk.StringVar()
        self.p_entry = ModernEntry(input_frame, textvariable=self.password_var, 
                                 show="•", theme=self.t, placeholder="Enter your password")
        self.p_entry.pack(fill="x", pady=(5, 15))

        self.status_label = tk.Label(card, text="", fg="#DC2626", 
                                    bg="white", font=FONT_SMALL)
        self.status_label.pack(pady=(0, 10))

        ModernButton(card, text="Log In", 
                    command=self._on_login, 
                    variant="primary", 
                    theme=self.t,
                    padx=30,
                    pady=14).pack(fill="x", pady=(0, 15))
        
        reg_frame = tk.Frame(card, bg="white")
        reg_frame.pack(fill="x")
        
        tk.Label(reg_frame, text="No account?", 
                bg="white", fg="#6B7280", 
                font=FONT_SMALL).pack(side="left")
        
        reg_btn = tk.Button(reg_frame, text="Create one", 
                           command=self._on_register,
                           bd=0, bg="white", fg="#3B82F6", 
                           activebackground="white", 
                           activeforeground="#2563EB", 
                           font=("Segoe UI", 9, "bold"), 
                           cursor="hand2")
        reg_btn.pack(side="left", padx=5)

        self.root.bind("<Return>", lambda e: self._on_login())
        self.result_username = None

    def _center_window(self, win, w, h):
        ws = win.winfo_screenwidth()
        hs = win.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 2) - (h // 2)
        win.geometry(f"{w}x{h}+{x}+{y}")

    def _on_login(self):
        uname = self.username_var.get().strip()
        pwd = self.password_var.get()
        
        if not uname or not pwd:
            self.status_label.config(text="Please enter credentials")
            return
            
        self.status_label.config(text="Authenticating...", fg="#3B82F6")
        self.root.update()
        
        if verify_user(uname, pwd):
            self.result_username = uname
            self.root.destroy()
        else:
            self.status_label.config(text="Invalid username or password", fg="#DC2626")

    def _on_register(self):
        RegistrationDialog(self.root)

    def run(self):
        self.root.mainloop()
        return self.result_username


class RegistrationDialog:
    def __init__(self, parent):
        self.win = tk.Toplevel(parent)
        self.win.title("Create Account")
        self.win.geometry("480x620")
        self.win.resizable(False, False)
        self._center_window(self.win, 480, 620)
        
        self.t = LIGHT_THEME
        self.win.configure(bg="#EEF2FF")
        self.win.transient(parent)
        self.win.grab_set()

        card_frame = tk.Frame(self.win, bg=self.t["border_color"], padx=1, pady=1)
        card_frame.place(relx=0.5, rely=0.5, anchor="center", width=400)
        
        card = tk.Frame(card_frame, bg="white", padx=40, pady=40)
        card.pack(fill="both", expand=True)

        tk.Label(card, text="Create Account", 
                font=("Segoe UI", 20, "bold"), 
                bg="white", fg="#1F2937").pack(pady=(0, 25))

        tk.Label(card, text="Username", 
                font=("Segoe UI", 10, "bold"), 
                bg="white", fg="#374151", 
                anchor="w").pack(fill="x")
        
        self.u_var = tk.StringVar()
        ModernEntry(card, textvariable=self.u_var, 
                   theme=self.t, 
                   placeholder="Choose a username").pack(fill="x", pady=(5, 12))

        tk.Label(card, text="Password", 
                font=("Segoe UI", 10, "bold"), 
                bg="white", fg="#374151", 
                anchor="w").pack(fill="x")
        
        self.p_var = tk.StringVar()
        ModernEntry(card, textvariable=self.p_var, 
                   show="•", theme=self.t, 
                   placeholder="Enter a password").pack(fill="x", pady=(5, 12))

        tk.Label(card, text="Confirm Password", 
                font=("Segoe UI", 10, "bold"), 
                bg="white", fg="#374151", 
                anchor="w").pack(fill="x")
        
        self.c_var = tk.StringVar()
        ModernEntry(card, textvariable=self.c_var, 
                   show="•", theme=self.t, 
                   placeholder="Confirm your password").pack(fill="x", pady=(5, 12))

        self.status = tk.Label(card, text="", fg="#DC2626", 
                              bg="white", font=FONT_SMALL)
        self.status.pack(pady=(5, 10))

        ModernButton(card, text="Sign Up", 
                    command=self._do_reg, 
                    variant="primary", 
                    theme=self.t).pack(fill="x", pady=(0, 10))
        
        ModernButton(card, text="Cancel", 
                    command=self.win.destroy, 
                    variant="secondary", 
                    theme=self.t).pack(fill="x")

    def _center_window(self, win, w, h):
        ws = win.winfo_screenwidth()
        hs = win.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 2) - (h // 2)
        win.geometry(f"{w}x{h}+{x}+{y}")

    def _do_reg(self):
        u = self.u_var.get().strip()
        p = self.p_var.get()
        c = self.c_var.get()

        if not u or not p:
            self.status.config(text="All fields required")
            return
            
        if p != c:
            self.status.config(text="Passwords do not match")
            return
            
        if len(p) < 4:
            self.status.config(text="Password must be at least 4 characters")
            return

        try:
            create_user(u, p)
            self.status.config(text="Account created! You can now login.", fg="#10B981")
            self.win.after(1500, self.win.destroy)
        except sqlite3.IntegrityError:
            self.status.config(text="Username already taken")
        except Exception as e:
            self.status.config(text=f"System error: {str(e)}")


# ----------------------------
# MAIN APP
# ----------------------------
BASE_TK = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk

class PDFToolKitApp(BASE_TK):
    def __init__(self, current_user=None):
        super().__init__()
        self.current_user = current_user
        self.title("PDF Toolkit Pro")
        self.logged_out = False  # Flag to track logout state
        
        w, h = config.get("window_size", (1200, 800))
        self.geometry(f"{w}x{h}")
        self.minsize(1050, 700)
        
        self.theme_mode = config.get("theme", "light")
        self.t = LIGHT_THEME if self.theme_mode == "light" else DARK_THEME
        
        self._progress_queue = queue.Queue()
        self.last_folder = config.get("last_folder", os.path.expanduser("~"))

        # Configure ttk style
        self.style = ttk.Style(self)
        try: 
            self.style.theme_use('clam')
        except: 
            pass
        
        self._init_ui()
        self.show_home()
        self.after(100, self._process_progress_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _init_ui(self):
        # Master Container
        self.container = tk.Frame(self, bg=self.t["app_bg"])
        self.container.pack(fill="both", expand=True)

        # --- Sidebar ---
        self.sidebar = tk.Frame(self.container, width=260, bg=self.t["sidebar_bg"])
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Brand Section
        brand_frame = tk.Frame(self.sidebar, bg=self.t["sidebar_bg"], pady=35, padx=20)
        brand_frame.pack(fill="x")
        
        # App Title
        title_frame = tk.Frame(brand_frame, bg=self.t["sidebar_bg"])
        title_frame.pack(anchor="w")
        
        tk.Label(title_frame, text="PDF", 
                font=("Segoe UI", 20, "bold"), 
                bg=self.t["sidebar_bg"], 
                fg=self.t["text_main"]).pack(side="left")
        
        tk.Label(title_frame, text="Toolkit", 
                font=("Segoe UI", 20, "bold"), 
                bg=self.t["sidebar_bg"], 
                fg=self.t["primary_bg"]).pack(side="left")
        
        tk.Label(brand_frame, text="PRO EDITION", 
                font=("Segoe UI", 8, "bold"), 
                bg=self.t["sidebar_bg"], 
                fg=self.t["accent_bg"]).pack(anchor="w", pady=(0, 5))
        
        # User Info
        user_frame = tk.Frame(brand_frame, bg=self.t["sidebar_bg"], pady=15)
        user_frame.pack(fill="x")
        
        tk.Label(user_frame, text="👤", 
                font=("Segoe UI", 12), 
                bg=self.t["sidebar_bg"]).pack(side="left", padx=(0, 8))
        
        user_info = tk.Frame(user_frame, bg=self.t["sidebar_bg"])
        user_info.pack(side="left", fill="x", expand=True)
        
        tk.Label(user_info, text=self.current_user, 
                font=("Segoe UI", 11, "bold"), 
                bg=self.t["sidebar_bg"], 
                fg=self.t["text_main"],
                anchor="w").pack(fill="x")
        
        tk.Label(user_info, text="", 
                font=("Segoe UI", 8), 
                bg=self.t["sidebar_bg"], 
                fg=self.t["success_bg"],
                anchor="w").pack(fill="x")

        sep = tk.Frame(self.sidebar, height=1, bg=self.t["border_color"])
        sep.pack(fill="x", padx=20, pady=10)

        # Navigation Menu
        self.menu_frame = tk.Frame(self.sidebar, bg=self.t["sidebar_bg"])
        self.menu_frame.pack(fill="both", expand=True, pady=10)
        
        self.nav_buttons = []
        menu_items = [
            ("🏠", "Dashboard", self.show_home),
            ("📑", "Merge PDF", self.show_merge),
            ("✂️", "Split PDF", self.show_split),
            ("🗜️", "Compress", self.show_compress),
            ("💧", "Watermark", self.show_watermark),
            ("📄", "Word to PDF", self.show_word),
        ]

        for icon, text, cmd in menu_items:
            btn_frame = tk.Frame(self.menu_frame, bg=self.t["sidebar_bg"])
            btn_frame.pack(fill="x", pady=2, padx=10)
            
            btn = tk.Button(btn_frame, 
                          text=f"  {icon}  {text}", 
                          font=("Segoe UI", 11), 
                          anchor="w",
                          padx=20, 
                          pady=12, 
                          relief="flat", 
                          bd=0, 
                          cursor="hand2",
                          bg=self.t["sidebar_bg"], 
                          fg=self.t["sidebar_fg"],
                          activebackground=self.t["sidebar_active_bg"],
                          activeforeground=self.t["sidebar_active_fg"],
                          command=lambda c=cmd, t=text: self._nav_click(c, t))
            btn.pack(fill="x")
            self.nav_buttons.append(btn)

        # Bottom Sidebar
        btm_frame = tk.Frame(self.sidebar, bg=self.t["sidebar_bg"], pady=20)
        btm_frame.pack(side="bottom", fill="x")

        # LOGOUT BUTTON ADDED HERE
        self.logout_btn = ModernButton(btm_frame, 
                                     text="  🚪  Logout", 
                                     command=self.logout, 
                                     variant="danger", 
                                     theme=self.t, 
                                     padx=20, 
                                     pady=10)
        self.logout_btn.pack(fill="x", padx=20, pady=(0, 10))
        
        # Theme Toggle
        theme_icon = "🌙" if self.theme_mode == "light" else "☀️"
        self.theme_btn = ModernButton(btm_frame, 
                                     text=f"{theme_icon} Toggle Theme", 
                                     command=self.toggle_theme, 
                                     variant="secondary", 
                                     theme=self.t, 
                                     padx=20, 
                                     pady=10)
        self.theme_btn.pack(fill="x", padx=20, pady=(0, 15))
        
        version_frame = tk.Frame(btm_frame, bg=self.t["sidebar_bg"])
        version_frame.pack(fill="x", padx=20)
        
        tk.Label(version_frame, text="Developed By Harsh Vadoliya", 
                font=("Segoe UI", 8), 
                bg=self.t["sidebar_bg"], 
                fg=self.t["text_muted"]).pack(side="left")
        
        tk.Label(version_frame, text="© 2025", 
                font=("Segoe UI", 8), 
                bg=self.t["sidebar_bg"], 
                fg=self.t["text_muted"]).pack(side="right")

        # --- Content Area ---
        self.main_area = tk.Frame(self.container, bg=self.t["app_bg"])
        self.main_area.pack(side="left", fill="both", expand=True)

        # Header
        self.header_frame = tk.Frame(self.main_area, bg=self.t["app_bg"], height=60)
        self.header_frame.pack(fill="x", pady=(0, 1))
        self.header_frame.pack_propagate(False)
        
        header_content = tk.Frame(self.header_frame, bg=self.t["app_bg"])
        header_content.pack(fill="both", expand=True, padx=40, pady=10)
        
        self.header_title = tk.Label(header_content, 
                                    text="Overview", 
                                    font=FONT_HEADER,
                                    bg=self.t["app_bg"], 
                                    fg=self.t["text_main"])
        self.header_title.pack(side="left")
        
        header_sep = tk.Frame(self.main_area, height=1, bg=self.t["border_color"])
        header_sep.pack(fill="x")

        # Dynamic Page Content
        self.page_frame = tk.Frame(self.main_area, bg=self.t["app_bg"], padx=40, pady=30)
        self.page_frame.pack(fill="both", expand=True)

    def _nav_click(self, cmd, text):
        for btn in self.nav_buttons:
            # Reset all buttons
            btn.config(bg=self.t["sidebar_bg"], 
                      fg=self.t["sidebar_fg"], 
                      font=("Segoe UI", 11))
            
            # Highlight active button
            if text in btn['text']:
                btn.config(bg=self.t["sidebar_active_bg"], 
                          fg=self.t["sidebar_active_fg"], 
                          font=("Segoe UI", 11, "bold"))
        cmd()

    def logout(self):
        """Handle logout functionality"""
        self.logged_out = True
        self._on_close()

    def toggle_theme(self):
        self.theme_mode = "dark" if self.theme_mode == "light" else "light"
        self.t = LIGHT_THEME if self.theme_mode == "light" else DARK_THEME
        config["theme"] = self.theme_mode
        save_config(config)
        
        theme_icon = "🌙" if self.theme_mode == "light" else "☀️"
        self.theme_btn.config(text=f"{theme_icon} Toggle Theme")
        
        self._apply_theme_to_ui()

    def _apply_theme_to_ui(self):
        # Main containers
        self.container.config(bg=self.t["app_bg"])
        self.sidebar.config(bg=self.t["sidebar_bg"])
        self.main_area.config(bg=self.t["app_bg"])
        self.header_frame.config(bg=self.t["app_bg"])
        self.page_frame.config(bg=self.t["app_bg"])
        
        self._update_widget_theme(self.sidebar)
        self._update_widget_theme(self.main_area)
        
        if hasattr(self, 'theme_btn'):
            self.theme_btn.update_theme(self.t)
        
        if hasattr(self, 'logout_btn'):
            self.logout_btn.update_theme(self.t)
        
        self.header_title.config(bg=self.t["app_bg"], fg=self.t["text_main"])

    def _update_widget_theme(self, widget):
        try:
            # Skip ttk widgets that don't support bg config
            if isinstance(widget, (ttk.Progressbar, ttk.Scrollbar, ttk.Separator)):
                return

            if isinstance(widget, tk.Frame):
                if widget == self.sidebar or widget.master == self.sidebar:
                    widget.config(bg=self.t["sidebar_bg"])
                elif widget == self.page_frame or widget.master == self.page_frame:
                    widget.config(bg=self.t["app_bg"])
                else:
                    # Generic card logic
                    widget.config(bg=self.t["card_bg"])
                    
            elif isinstance(widget, tk.Label):
                current_fg = widget.cget("fg")
                if current_fg in ["#4F46E5", "#4338CA", "#3B82F6", "#2563EB"]:
                    widget.config(fg=self.t["primary_bg"])
                elif current_fg in ["#6B7280", "#64748B", "#94A3B8"]:
                    widget.config(fg=self.t["text_muted"])
                else:
                    widget.config(fg=self.t["text_main"])
                
                # Check parent bg to match
                if hasattr(widget.master, "cget"):
                    try:
                        widget.config(bg=widget.master.cget("bg"))
                    except:
                        pass
                    
            elif isinstance(widget, tk.Button) and widget not in self.nav_buttons:
                if hasattr(widget, 'update_theme'):
                    widget.update_theme(self.t)
                    
            elif isinstance(widget, ModernEntry):
                widget.update_theme(self.t)
                
            elif isinstance(widget, tk.Listbox):
                widget.config(bg=self.t["listbox_bg"], 
                            fg=self.t["text_main"],
                            selectbackground=self.t["listbox_select"])
                
        except Exception:
            pass
        
        for child in widget.winfo_children():
            self._update_widget_theme(child)

    def _clear_page(self):
        for w in self.page_frame.winfo_children():
            w.destroy()

    def _create_card(self, parent):
        card = ModernCard(parent, theme=self.t)
        return card

    # -------------------------
    # Logic Helpers
    # -------------------------
    def _enqueue_progress(self, p, t):
        self._progress_queue.put((p, t))

    def _process_progress_queue(self):
        try:
            while True:
                p, t = self._progress_queue.get_nowait()
                for w in self.page_frame.winfo_children():
                    if hasattr(w, 'pbar'): 
                        # Handle determinate vs indeterminate
                        if 'determinate' in str(w.pbar.cget('mode')):
                            w.pbar['value'] = p * 100
                    if hasattr(w, 'stat'): 
                        w.stat.config(text=t)
        except queue.Empty: 
            pass
        finally: 
            self.after(100, self._process_progress_queue)

    def _pick_file(self, title, filetypes=(("PDF Files", "*.pdf"),)):
        p = filedialog.askopenfilename(
            initialdir=self.last_folder, 
            title=title, 
            filetypes=filetypes
        )
        if p: 
            self.last_folder = os.path.dirname(p)
            save_config({"last_folder": self.last_folder, "theme": self.theme_mode})
        return p

    def _pick_files(self, title):
        ps = filedialog.askopenfilenames(
            initialdir=self.last_folder, 
            title=title, 
            filetypes=(("PDF Files", "*.pdf"),)
        )
        if ps: 
            self.last_folder = os.path.dirname(ps[0])
            save_config({"last_folder": self.last_folder, "theme": self.theme_mode})
        return list(ps)

    def _save_file(self, title, ext=".pdf"):
        p = filedialog.asksaveasfilename(
            initialdir=self.last_folder, 
            title=title, 
            defaultextension=ext, 
            filetypes=(("PDF", "*.pdf"), ("All", "*.*"))
        )
        if p: 
            self.last_folder = os.path.dirname(p)
            save_config({"last_folder": self.last_folder, "theme": self.theme_mode})
        return p
    
    def _pick_folder(self, title):
        p = filedialog.askdirectory(initialdir=self.last_folder, title=title)
        if p: 
            self.last_folder = p
            save_config({"last_folder": self.last_folder, "theme": self.theme_mode})
        return p

    def _wrap_task(self, func, *args):
        try:
            func(*args, progress_callback=self._enqueue_progress)
            self._enqueue_progress(1.0, "Success!")
            messagebox.showinfo("Done", "Task completed successfully")
        except Exception as e:
            traceback.print_exc()
            self._enqueue_progress(0, f"Error: {str(e)}")
            messagebox.showerror("Error", str(e))
        finally:
            time.sleep(1)
            self._enqueue_progress(0, "")

    # -------------------------
    # ENHANCED PAGES
    # -------------------------
    def show_home(self):
        self._clear_page()
        self.header_title.config(text="Dashboard")
        
        home_container = tk.Frame(self.page_frame, bg=self.t["app_bg"])
        home_container.pack(fill="both", expand=True)
        
        welcome_card = self._create_card(home_container)
        welcome_card.pack(fill="x", pady=(0, 30))
        
        tk.Label(welcome_card, 
                text=f"Welcome back, {self.current_user}! 👋", 
                font=("Segoe UI", 24, "bold"), 
                bg=self.t["card_bg"], 
                fg=self.t["text_main"]).pack(anchor="w", pady=(0, 10))
        
        tk.Label(welcome_card, 
                text="Manage your PDF documents with our professional toolkit.",
                font=FONT_SUBHEADER, 
                bg=self.t["card_bg"], 
                fg=self.t["text_muted"]).pack(anchor="w", pady=(0, 30))
        
        stats_frame = tk.Frame(home_container, bg=self.t["app_bg"])
        stats_frame.pack(fill="x", pady=(0, 30))
        
        stats = [
            ("📊", "Tools Available", "5", self.t["primary_bg"]),
            ("⚡", "Fast Processing", "100%", self.t["success_bg"]),
            ("🔒", "Secure", "Encrypted", self.t["accent_bg"]),
            ("💼", "Professional", "Grade A", self.t["warning_bg"])
        ]
        
        for icon, title, value, color in stats:
            stat_card = tk.Frame(stats_frame, bg=self.t["card_bg"], padx=20, pady=20)
            stat_card.pack(side="left", fill="both", expand=True, padx=(0, 20))
            
            tk.Label(stat_card, text=icon, font=("Segoe UI", 24), 
                    bg=self.t["card_bg"]).pack(anchor="w", pady=(0, 10))
            
            tk.Label(stat_card, text=title, font=FONT_SMALL,
                    bg=self.t["card_bg"], fg=self.t["text_muted"]).pack(anchor="w")
            
            tk.Label(stat_card, text=value, font=("Segoe UI", 18, "bold"),
                    bg=self.t["card_bg"], fg=color).pack(anchor="w", pady=(5, 0))
        
        tk.Label(home_container, text="Quick Actions",
                font=("Segoe UI", 16, "bold"),
                bg=self.t["app_bg"],
                fg=self.t["text_main"]).pack(anchor="w", pady=(0, 20))
        
        grid_frame = tk.Frame(home_container, bg=self.t["app_bg"])
        grid_frame.pack(fill="both", expand=True)
        
        features = [
            ("📑", "Merge PDF", "Combine multiple files", self.show_merge),
            ("✂️", "Split PDF", "Divide into sections", self.show_split),
            ("🗜️", "Compress", "Reduce file size", self.show_compress),
            ("💧", "Watermark", "Add branding", self.show_watermark),
            ("📄", "Convert", "Word to PDF", self.show_word),
        ]
        
        for i, (icon, title, desc, cmd) in enumerate(features):
            row = i // 3
            col = i % 3
            
            if col == 0:
                row_frame = tk.Frame(grid_frame, bg=self.t["app_bg"])
                row_frame.pack(fill="x", pady=(0, 15))
            
            feature_card = tk.Frame(row_frame, bg=self.t["card_bg"], padx=20, pady=25)
            feature_card.pack(side="left", fill="both", expand=True, padx=(0, 20))
            feature_card.config(cursor="hand2")
            
            def make_lambda(c):
                return lambda e: c()
            
            feature_card.bind("<Enter>", lambda e, c=feature_card: c.config(bg=self.t["secondary_bg"]))
            feature_card.bind("<Leave>", lambda e, c=feature_card: c.config(bg=self.t["card_bg"]))
            feature_card.bind("<Button-1>", make_lambda(cmd))
            
            tk.Label(feature_card, text=icon, font=("Segoe UI", 28),
                    bg=feature_card.cget("bg")).pack(anchor="w", pady=(0, 15))
            
            tk.Label(feature_card, text=title, font=("Segoe UI", 12, "bold"),
                    bg=feature_card.cget("bg"), fg=self.t["text_main"]).pack(anchor="w")
            
            tk.Label(feature_card, text=desc, font=FONT_SMALL,
                    bg=feature_card.cget("bg"), fg=self.t["text_muted"]).pack(anchor="w", pady=(5, 0))

    def show_merge(self):
        self._clear_page()
        self.header_title.config(text="Merge PDFs")
        
        wrapper = tk.Frame(self.page_frame, bg=self.t["app_bg"])
        wrapper.pack(fill="both", expand=True)
        
        left_card = self._create_card(wrapper)
        left_card.pack(side="left", fill="both", expand=True, padx=(0, 20))
        
        tk.Label(left_card, text="Files to Merge", 
                font=("Segoe UI", 16, "bold"), 
                bg=self.t["card_bg"], 
                fg=self.t["text_main"]).pack(anchor="w", pady=(0, 20))
        
        list_frame = tk.Frame(left_card, bg=self.t["entry_border"])
        list_frame.pack(fill="both", expand=True, pady=(0, 20))
        
        self.merge_list = tk.Listbox(list_frame, 
                                    font=FONT_BODY, 
                                    activestyle="none",
                                    bg=self.t["listbox_bg"], 
                                    fg=self.t["text_main"],
                                    selectbackground=self.t["listbox_select"],
                                    selectforeground=self.t["text_main"],
                                    relief="flat", 
                                    bd=0,
                                    highlightthickness=0)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.merge_list.pack(side="left", fill="both", expand=True)
        self.merge_list.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.merge_list.yview)
        
        if DND_AVAILABLE:
            def drop(e):
                for p in self.tk.splitlist(e.data):
                    if p.lower().endswith(".pdf"): 
                        self.merge_list.insert(tk.END, p)
            try:
                self.merge_list.drop_target_register(DND_FILES)
                self.merge_list.dnd_bind("<<Drop>>", drop)
                tk.Label(left_card, text="📂 Drag & drop files here", 
                        font=FONT_SMALL, 
                        bg=self.t["card_bg"], 
                        fg=self.t["text_muted"]).pack(anchor="w")
            except Exception:
                pass

        right_card = self._create_card(wrapper)
        right_card.pack(side="right", fill="y")
        right_card.config(width=320)
        
        tk.Label(right_card, text="Actions", 
                font=("Segoe UI", 16, "bold"), 
                bg=self.t["card_bg"], 
                fg=self.t["text_main"]).pack(anchor="w", pady=(0, 20))
        
        def add_files():
            files = self._pick_files("Select PDFs")
            for f in files:
                self.merge_list.insert(tk.END, f)
        
        def remove_selected():
            selected = list(self.merge_list.curselection())
            for i in reversed(selected):
                self.merge_list.delete(i)
        
        ModernButton(right_card, "➕ Add Files", add_files, variant="secondary", theme=self.t).pack(fill="x", pady=5)
        ModernButton(right_card, "🗑️ Remove Selected", remove_selected, variant="secondary", theme=self.t).pack(fill="x", pady=5)
        ModernButton(right_card, "🧹 Clear All", lambda: self.merge_list.delete(0, tk.END), variant="secondary", theme=self.t).pack(fill="x", pady=5)
        
        tk.Label(right_card, text="Reorder", 
                font=("Segoe UI", 12, "bold"), 
                bg=self.t["card_bg"], 
                fg=self.t["text_muted"]).pack(anchor="w", pady=(20, 10))
        
        order_frame = tk.Frame(right_card, bg=self.t["card_bg"])
        order_frame.pack(fill="x")
        
        ModernButton(order_frame, "▲ Move Up", 
                    lambda: self._move_item(-1), 
                    variant="secondary", 
                    theme=self.t).pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ModernButton(order_frame, "▼ Move Down", 
                    lambda: self._move_item(1), 
                    variant="secondary", 
                    theme=self.t).pack(side="left", fill="x", expand=True)
        
        tk.Frame(right_card, height=30, bg=self.t["card_bg"]).pack()
        
        tk.Label(right_card, text="Progress", 
                font=("Segoe UI", 12, "bold"), 
                bg=self.t["card_bg"], 
                fg=self.t["text_muted"]).pack(anchor="w", pady=(0, 10))
        
        self.pbar = ttk.Progressbar(right_card, mode="determinate")
        self.pbar.pack(fill="x", pady=(0, 10))
        
        self.stat = tk.Label(right_card, text="Ready to merge", 
                            font=FONT_SMALL, 
                            bg=self.t["card_bg"], 
                            fg=self.t["text_muted"])
        self.stat.pack(anchor="w")
        
        tk.Frame(right_card, height=20, bg=self.t["card_bg"]).pack()
        
        def run_merge():
            files = list(self.merge_list.get(0, tk.END))
            if not files: 
                messagebox.showerror("Error", "No files selected")
                return
                
            out = self._save_file("Save Merged PDF")
            if out: 
                threading.Thread(
                    target=lambda: self._wrap_task(module_merge_pdf, files, out), 
                    daemon=True
                ).start()
        
        ModernButton(right_card, "🚀 Merge PDFs", 
                    run_merge, 
                    variant="primary", 
                    theme=self.t,
                    pady=14).pack(side="bottom", fill="x")

    def _move_item(self, offset):
        lb = self.merge_list
        sel = lb.curselection()
        if not sel: 
            return
        i = sel[0]
        new_i = i + offset
        if 0 <= new_i < lb.size():
            text = lb.get(i)
            lb.delete(i)
            lb.insert(new_i, text)
            lb.selection_set(new_i)

    def show_split(self):
        self._setup_form_page(
            "Split PDF", 
            "Split a PDF into two parts at a specific page number.",
            [("Input PDF", "file"), ("Split Page Index (0-based)", "text"), ("Output Folder", "folder")],
            self._run_split
        )
    
    def _run_split(self, entries):
        inp, idx, fld = [e.get() for e in entries]
        if not inp or not fld: 
            return
        try:
            threading.Thread(
                target=lambda: self._wrap_task(module_split_pdf, inp, int(idx), fld), 
                daemon=True
            ).start()
        except ValueError: 
            messagebox.showerror("Error", "Invalid page index. Please enter a number.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def show_compress(self):
        self._setup_form_page(
            "Compress PDF",
            "Optimize file size by compressing streams and reducing quality.",
            [("Input PDF", "file"), ("Output PDF", "save")],
            lambda e: threading.Thread(
                target=lambda: self._wrap_task(module_compress_pdf, e[0].get(), e[1].get()), 
                daemon=True
            ).start()
        )

    def show_watermark(self):
        self._setup_form_page(
            "Watermark PDF",
            "Overlay the first page of the Watermark PDF onto Source PDF.",
            [("Source PDF", "file"), ("Watermark PDF", "file"), ("Output PDF", "save")],
            lambda e: threading.Thread(
                target=lambda: self._wrap_task(module_add_watermark, e[0].get(), e[1].get(), e[2].get()), 
                daemon=True
            ).start()
        )

    def show_word(self):
        self._setup_form_page(
            "Word to PDF",
            "Convert a .docx file to PDF format.",
            [("Word File", "file_word"), ("Output PDF", "save")],
            lambda e: threading.Thread(
                target=lambda: self._wrap_task(module_word_to_pdf, e[0].get(), e[1].get()), 
                daemon=True
            ).start()
        )

    def _setup_form_page(self, title, subtitle, fields, run_cb):
        self._clear_page()
        self.header_title.config(text=title)
        
        card = self._create_card(self.page_frame)
        card.pack(fill="both", expand=True)
        
        tk.Label(card, text=subtitle, 
                font=FONT_SUBHEADER, 
                bg=self.t["card_bg"], 
                fg=self.t["text_muted"]).pack(anchor="w", pady=(0, 30))
        
        entries = []
        for label, typ in fields:
            row = tk.Frame(card, bg=self.t["card_bg"])
            row.pack(fill="x", pady=(0, 20))
            
            tk.Label(row, text=label, 
                    font=("Segoe UI", 11, "bold"), 
                    width=25, 
                    anchor="w",
                    bg=self.t["card_bg"], 
                    fg=self.t["text_main"]).grid(row=0, column=0, sticky="w", padx=(0, 15))
            
            if typ == "text":
                e = ModernEntry(row, theme=self.t)
                e.insert(0, "0")
            else:
                e = ModernEntry(row, theme=self.t)
            
            e.grid(row=0, column=1, sticky="ew", padx=(0, 10))
            row.grid_columnconfigure(1, weight=1)
            entries.append(e)

            if typ in ["file", "file_word", "folder", "save"]:
                cmd = None
                if typ == "file": 
                    cmd = lambda x=e: [x.delete(0, tk.END), x.insert(0, self._pick_file("Select File"))]
                elif typ == "file_word": 
                    cmd = lambda x=e: [x.delete(0, tk.END), x.insert(0, self._pick_file("Select Word", (("Word", "*.docx"),)))]
                elif typ == "folder": 
                    cmd = lambda x=e: [x.delete(0, tk.END), x.insert(0, self._pick_folder("Select Folder"))]
                elif typ == "save": 
                    cmd = lambda x=e: [x.delete(0, tk.END), x.insert(0, self._save_file("Save As"))]
                
                ModernButton(row, "Browse", cmd, 
                            variant="secondary", 
                            theme=self.t, 
                            padx=15, 
                            pady=8).grid(row=0, column=2)

        tk.Frame(card, height=30, bg=self.t["card_bg"]).pack()
        
        tk.Label(card, text="Progress", 
                font=("Segoe UI", 12, "bold"), 
                bg=self.t["card_bg"], 
                fg=self.t["text_muted"]).pack(anchor="w", pady=(0, 10))
        
        self.pbar = ttk.Progressbar(card, 
                                   mode="indeterminate" if "word" in title.lower() else "determinate")
        self.pbar.pack(fill="x", pady=(10, 5))
        
        self.stat = tk.Label(card, text="Ready", 
                            font=FONT_SMALL, 
                            bg=self.t["card_bg"], 
                            fg=self.t["text_muted"])
        self.stat.pack(anchor="w")
        
        action_frame = tk.Frame(card, bg=self.t["card_bg"])
        action_frame.pack(fill="x", pady=(20, 0))
        
        def validate_and_run():
            if any(not e.get().strip() for e in entries):
                messagebox.showerror("Error", "Please fill all fields")
                return
            if "word" in title.lower():
                self.pbar.start(10)
            run_cb(entries)
        
        ModernButton(action_frame, f"🚀 Start {title.split()[0]}", 
                    validate_and_run, 
                    variant="primary", 
                    theme=self.t,
                    pady=14).pack(fill="x")

    def _on_close(self):
        try:
            w, h = self.winfo_width(), self.winfo_height()
            config["window_size"] = (w, h)
            save_config(config)
        except: 
            pass
        self.destroy()

def main():
    try:
        if not PDF_AVAILABLE:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror("Dependency Error", "PyPDF2 is required.\nPlease run: pip install PyPDF2")
            r.destroy()
            return
        
        # Application Loop
        while True:
            auth = AuthDialog()
            user = auth.run()
            
            if user:
                app = PDFToolKitApp(current_user=user)
                app.mainloop()
                
                # Check if user explicitly logged out
                if not app.logged_out:
                    break
            else:
                # User closed the login window without logging in
                break

    except Exception as e:
        traceback.print_exc()
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror("Fatal Error", f"An unexpected error occurred:\n{str(e)}")
            r.destroy()
        except:
            print(f"Fatal Error: {e}")

if __name__ == "__main__":
    main()