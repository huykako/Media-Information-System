import os
import sys
import sqlite3
import threading
import time
import webbrowser
from datetime import datetime
from pathlib import Path

import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox

import requests

# =========================
# Application Metadata
# =========================
APP_NAME = "Media Information System | ©Thorsten Bylicki - ©BYLICKILABS"
APP_TITLE = "Media Information System (MIS) | ©Thorsten Bylicki - ©BYLICKILABS"
APP_VERSION = "1.0.0"
APP_COMPANY = "©BYLICKILABS"
GITHUB_URL = "https://github.com/bylickilabs"
DB_FILE = "media_library.db"
POSTER_DIR = "posters"

OMDB_API_KEY = "89e61ff2"  # [REGISTER](https://www.omdbapi.com)


# =========================
# Internationalization
# =========================
TRANSLATIONS = {
    "de": {
        "language_label": "Sprache",
        "tab_movies": "Filme",
        "tab_music": "Musik",
        "tab_programs": "Programme",
        "tab_documents": "Dokumente",
        "btn_add_folder": "Verzeichnis hinzufügen",
        "btn_refresh": "Aktualisieren",
        "btn_print": "Drucken / Exportieren",
        "btn_github": "GitHub",
        "btn_info": "Info",
        "status_ready": "Bereit.",
        "status_indexing": "Indexiere Verzeichnis...",
        "status_done": "Indexierung abgeschlossen.",
        "dialog_select_folder": "Verzeichnis auswählen",
        "info_title": "Über diese Anwendung",
        "info_text": (
            f"{APP_TITLE}\n\n"
            f"Version: {APP_VERSION}\n"
            f"Unternehmen: {APP_COMPANY}\n\n"
            "Media-Bibliotheksverwaltung für Filme, Musik, Programme und Dokumente.\n"
            "Entwickelt für schnelle Indizierung, komfortable Suche und professionelle Auswertung."
        ),
        "msg_no_folder": "Kein Verzeichnis ausgewählt.",
        "msg_index_finished": "Indexierung erfolgreich abgeschlossen.",
        "msg_print_export_success": "Export erfolgreich erstellt.",
        "msg_print_export_error": "Fehler beim Export.",
        "print_dialog_title": "Aktuelle Liste exportieren",
        "print_dialog_text": "Die aktuelle Tabellensicht wird als HTML-Datei exportiert, die Sie über den Browser ausdrucken können.",
        "column_filename": "Dateiname",
        "column_title": "Titel",
        "column_year": "Jahr",
        "column_genre": "Genre",
        "column_rating": "Bewertung",
        "column_artist": "Künstler",
        "column_album": "Album",
        "column_type": "Typ",
        "column_size": "Größe (MB)",
        "column_added": "Hinzugefügt am",
        "msg_omdb_not_configured": "OMDb API-Schlüssel nicht konfiguriert. Filminformationen können nicht automatisch geladen werden.",
        "msg_omdb_error": "Fehler beim Abrufen der Filmdaten.",
    },
    "en": {
        "language_label": "Language",
        "tab_movies": "Movies",
        "tab_music": "Music",
        "tab_programs": "Programs",
        "tab_documents": "Documents",
        "btn_add_folder": "Add Folder",
        "btn_refresh": "Refresh",
        "btn_print": "Print / Export",
        "btn_github": "GitHub",
        "btn_info": "Info",
        "status_ready": "Ready.",
        "status_indexing": "Indexing directory...",
        "status_done": "Indexing finished.",
        "dialog_select_folder": "Select folder",
        "info_title": "About this application",
        "info_text": (
            f"{APP_TITLE}\n\n"
            f"Version: {APP_VERSION}\n"
            f"Company: {APP_COMPANY}\n\n"
            "Media library management for movies, music, programs and documents.\n"
            "Designed for fast indexing, convenient search and professional reporting."
        ),
        "msg_no_folder": "No folder selected.",
        "msg_index_finished": "Indexing completed successfully.",
        "msg_print_export_success": "Export created successfully.",
        "msg_print_export_error": "Error while creating export.",
        "print_dialog_title": "Export current list",
        "print_dialog_text": "The current table view will be exported as an HTML file that you can print from your browser.",
        "column_filename": "Filename",
        "column_title": "Title",
        "column_year": "Year",
        "column_genre": "Genre",
        "column_rating": "Rating",
        "column_artist": "Artist",
        "column_album": "Album",
        "column_type": "Type",
        "column_size": "Size (MB)",
        "column_added": "Added on",
        "msg_omdb_not_configured": "OMDb API key not configured. Movie information cannot be fetched automatically.",
        "msg_omdb_error": "Error while fetching movie data.",
    },
}


# =========================
# Database Layer
# =========================
class MediaDatabase:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._ensure_db()

    def _connect(self):
        return sqlite3.connect(self.db_path)

    def _ensure_db(self):
        conn = self._connect()
        try:
            cur = conn.cursor()

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS movies (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT UNIQUE,
                    filename TEXT,
                    title TEXT,
                    year TEXT,
                    genre TEXT,
                    imdb_rating TEXT,
                    poster_path TEXT,
                    size_bytes INTEGER,
                    added_at TEXT
                );
                """
            )

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS music (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT UNIQUE,
                    filename TEXT,
                    artist TEXT,
                    album TEXT,
                    title TEXT,
                    genre TEXT,
                    size_bytes INTEGER,
                    added_at TEXT
                );
                """
            )

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS programs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT UNIQUE,
                    filename TEXT,
                    app_name TEXT,
                    version TEXT,
                    vendor TEXT,
                    size_bytes INTEGER,
                    added_at TEXT
                );
                """
            )

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS documents (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT UNIQUE,
                    filename TEXT,
                    title TEXT,
                    doc_type TEXT,
                    size_bytes INTEGER,
                    added_at TEXT
                );
                """
            )

            conn.commit()
        finally:
            conn.close()

    # Insert or ignore pattern
    def insert_movie(self, **data):
        conn = self._connect()
        try:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT OR IGNORE INTO movies
                    (file_path, filename, title, year, genre, imdb_rating, poster_path, size_bytes, added_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
                """,
                (
                    data.get("file_path"),
                    data.get("filename"),
                    data.get("title"),
                    data.get("year"),
                    data.get("genre"),
                    data.get("imdb_rating"),
                    data.get("poster_path"),
                    data.get("size_bytes"),
                    data.get("added_at"),
                ),
            )
            conn.commit()
        finally:
            conn.close()

    def insert_music(self, **data):
        conn = self._connect()
        try:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT OR IGNORE INTO music
                    (file_path, filename, artist, album, title, genre, size_bytes, added_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?);
                """,
                (
                    data.get("file_path"),
                    data.get("filename"),
                    data.get("artist"),
                    data.get("album"),
                    data.get("title"),
                    data.get("genre"),
                    data.get("size_bytes"),
                    data.get("added_at"),
                ),
            )
            conn.commit()
        finally:
            conn.close()

    def insert_program(self, **data):
        conn = self._connect()
        try:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT OR IGNORE INTO programs
                    (file_path, filename, app_name, version, vendor, size_bytes, added_at)
                VALUES (?, ?, ?, ?, ?, ?, ?);
                """,
                (
                    data.get("file_path"),
                    data.get("filename"),
                    data.get("app_name"),
                    data.get("version"),
                    data.get("vendor"),
                    data.get("size_bytes"),
                    data.get("added_at"),
                ),
            )
            conn.commit()
        finally:
            conn.close()

    def insert_document(self, **data):
        conn = self._connect()
        try:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT OR IGNORE INTO documents
                    (file_path, filename, title, doc_type, size_bytes, added_at)
                VALUES (?, ?, ?, ?, ?, ?);
                """,
                (
                    data.get("file_path"),
                    data.get("filename"),
                    data.get("title"),
                    data.get("doc_type"),
                    data.get("size_bytes"),
                    data.get("added_at"),
                ),
            )
            conn.commit()
        finally:
            conn.close()

    def fetch_movies(self):
        conn = self._connect()
        try:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT filename, title, year, genre, imdb_rating, size_bytes, added_at
                FROM movies
                ORDER BY title COLLATE NOCASE ASC;
                """
            )
            return cur.fetchall()
        finally:
            conn.close()

    def fetch_music(self):
        conn = self._connect()
        try:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT filename, artist, album, title, genre, size_bytes, added_at
                FROM music
                ORDER BY title COLLATE NOCASE ASC;
                """
            )
            return cur.fetchall()
        finally:
            conn.close()

    def fetch_programs(self):
        conn = self._connect()
        try:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT filename, app_name, version, vendor, size_bytes, added_at
                FROM programs
                ORDER BY app_name COLLATE NOCASE ASC;
                """
            )
            return cur.fetchall()
        finally:
            conn.close()

    def fetch_documents(self):
        conn = self._connect()
        try:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT filename, title, doc_type, size_bytes, added_at
                FROM documents
                ORDER BY title COLLATE NOCASE ASC;
                """
            )
            return cur.fetchall()
        finally:
            conn.close()


# =========================
# Utility Functions
# =========================
MOVIE_EXTENSIONS = {".mp4", ".mkv", ".avi", ".mov"}
MUSIC_EXTENSIONS = {".mp3", ".flac", ".wav", ".aac", ".ogg"}
PROGRAM_EXTENSIONS = {".exe", ".msi", ".bat", ".cmd", ".sh", ".py"}
DOCUMENT_EXTENSIONS = {".pdf", ".doc", ".docx", ".txt", ".xlsx", ".pptx"}


def readable_size(size_bytes: int) -> float:
    if size_bytes is None:
        return 0.0
    return round(size_bytes / (1024 * 1024), 2)


def ensure_poster_dir():
    Path(POSTER_DIR).mkdir(parents=True, exist_ok=True)


def fetch_movie_metadata_from_omdb(filename: str) -> dict:
    """
    Very simple OMDb lookup based on guessed movie title from filename.
    """
    if not OMDB_API_KEY:
        return {}
    # Heuristic: remove extension and common separators
    name = Path(filename).stem
    query = name.replace(".", " ").replace("_", " ").strip()

    try:
        resp = requests.get(
            "http://www.omdbapi.com/",
            params={"apikey": OMDB_API_KEY, "t": query},
            timeout=10,
        )
        if resp.status_code != 200:
            return {}
        data = resp.json()
        if data.get("Response") != "True":
            return {}
        result = {
            "title": data.get("Title") or query,
            "year": data.get("Year") or "",
            "genre": data.get("Genre") or "",
            "imdb_rating": data.get("imdbRating") or "",
        }

        poster_url = data.get("Poster")
        if poster_url and poster_url != "N/A":
            ensure_poster_dir()
            poster_path = Path(POSTER_DIR) / f"{query}.jpg"
            try:
                presp = requests.get(poster_url, timeout=10)
                if presp.status_code == 200:
                    with open(poster_path, "wb") as f:
                        f.write(presp.content)
                    result["poster_path"] = str(poster_path)
            except Exception:
                pass

        return result
    except Exception:
        return {}


# =========================
# GUI Application
# =========================
class MediaApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Core
        self.title(APP_TITLE)
        self.geometry("1200x700")
        self.minsize(1000, 600)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.db = MediaDatabase(DB_FILE)

        # Language
        self.current_lang = "de"
        self.t = TRANSLATIONS[self.current_lang]

        # State
        self.current_tab = "movies"

        # UI
        self._build_ui()
        self._apply_translations()
        self._load_all_tabs()

    # ---------- UI Construction ----------

    def _build_ui(self):
        # Top bar: language, actions, GitHub, Info
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(side="top", fill="x", padx=10, pady=5)

        # Language selector
        self.lang_label = ctk.CTkLabel(top_frame, text="")
        self.lang_label.pack(side="left", padx=(10, 5))

        self.lang_var = ctk.StringVar(value=self.current_lang)
        self.lang_option = ctk.CTkComboBox(
            top_frame,
            variable=self.lang_var,
            values=["de", "en"],
            width=80,
            command=self._on_language_change,
        )
        self.lang_option.pack(side="left", padx=(0, 20))

        # Action buttons
        self.btn_add_folder = ctk.CTkButton(
            top_frame, text="", command=self._on_add_folder_clicked
        )
        self.btn_add_folder.pack(side="left", padx=5)

        self.btn_refresh = ctk.CTkButton(
            top_frame, text="", command=self._on_refresh_clicked
        )
        self.btn_refresh.pack(side="left", padx=5)

        self.btn_print = ctk.CTkButton(
            top_frame, text="", command=self._on_print_clicked
        )
        self.btn_print.pack(side="left", padx=5)

        # Right side buttons
        self.btn_github = ctk.CTkButton(
            top_frame, text="", width=80, command=self._open_github
        )
        self.btn_github.pack(side="right", padx=5)

        self.btn_info = ctk.CTkButton(
            top_frame, text="", width=80, command=self._show_info
        )
        self.btn_info.pack(side="right", padx=5)

        # Notebook / Tabs
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(side="top", fill="both", expand=True, padx=10, pady=5)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # Create tab frames
        self.tab_frames = {}
        self.trees = {}

        self._create_tab("movies")
        self._create_tab("music")
        self._create_tab("programs")
        self._create_tab("documents")

        # Status bar
        status_frame = ctk.CTkFrame(self, height=24)
        status_frame.pack(side="bottom", fill="x")
        self.status_label = ctk.CTkLabel(status_frame, text="")
        self.status_label.pack(side="left", padx=10)
        self._set_status(self.t["status_ready"])

    def _create_tab(self, tab_key: str):
        frame = ctk.CTkFrame(self.notebook)
        self.tab_frames[tab_key] = frame

        # Wrap Treeview in its own frame with scrollbar
        table_frame = ctk.CTkFrame(frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        tree = ttk.Treeview(table_frame, show="headings")
        self.trees[tab_key] = tree

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.notebook.add(frame, text=tab_key)

    # ---------- Translations ----------
    def _apply_translations(self):
        self.t = TRANSLATIONS[self.current_lang]
        self.lang_label.configure(text=self.t["language_label"])
        self.btn_add_folder.configure(text=self.t["btn_add_folder"])
        self.btn_refresh.configure(text=self.t["btn_refresh"])
        self.btn_print.configure(text=self.t["btn_print"])
        self.btn_github.configure(text=self.t["btn_github"])
        self.btn_info.configure(text=self.t["btn_info"])
        self._set_status(self.t["status_ready"])

        self.notebook.tab(self.tab_frames["movies"], text=self.t["tab_movies"])
        self.notebook.tab(self.tab_frames["music"], text=self.t["tab_music"])
        self.notebook.tab(self.tab_frames["programs"], text=self.t["tab_programs"])
        self.notebook.tab(self.tab_frames["documents"], text=self.t["tab_documents"])

        # Rebuild tree columns with translated headers
        self._configure_tree_columns()

    def _configure_tree_columns(self):
        # Movies
        tree = self.trees["movies"]
        tree["columns"] = ("filename", "title", "year", "genre", "rating", "size", "added")
        tree.heading("filename", text=self.t["column_filename"])
        tree.heading("title", text=self.t["column_title"])
        tree.heading("year", text=self.t["column_year"])
        tree.heading("genre", text=self.t["column_genre"])
        tree.heading("rating", text=self.t["column_rating"])
        tree.heading("size", text=self.t["column_size"])
        tree.heading("added", text=self.t["column_added"])
        for col in tree["columns"]:
            tree.column(col, width=150, anchor="w")

        # Music
        tree = self.trees["music"]
        tree["columns"] = ("filename", "artist", "album", "title", "genre", "size", "added")
        tree.heading("filename", text=self.t["column_filename"])
        tree.heading("artist", text=self.t["column_artist"])
        tree.heading("album", text=self.t["column_album"])
        tree.heading("title", text=self.t["column_title"])
        tree.heading("genre", text=self.t["column_genre"])
        tree.heading("size", text=self.t["column_size"])
        tree.heading("added", text=self.t["column_added"])
        for col in tree["columns"]:
            tree.column(col, width=150, anchor="w")

        # Programs
        tree = self.trees["programs"]
        tree["columns"] = ("filename", "name", "version", "vendor", "size", "added")
        tree.heading("filename", text=self.t["column_filename"])
        tree.heading("name", text=self.t["column_title"])
        tree.heading("version", text="Version")
        tree.heading("vendor", text="Vendor")
        tree.heading("size", text=self.t["column_size"])
        tree.heading("added", text=self.t["column_added"])
        for col in tree["columns"]:
            tree.column(col, width=150, anchor="w")

        # Documents
        tree = self.trees["documents"]
        tree["columns"] = ("filename", "title", "type", "size", "added")
        tree.heading("filename", text=self.t["column_filename"])
        tree.heading("title", text=self.t["column_title"])
        tree.heading("type", text=self.t["column_type"])
        tree.heading("size", text=self.t["column_size"])
        tree.heading("added", text=self.t["column_added"])
        for col in tree["columns"]:
            tree.column(col, width=150, anchor="w")

    # ---------- Status ----------

    def _set_status(self, text: str):
        self.status_label.configure(text=text)
        self.status_label.update_idletasks()

    # ---------- Event Handlers ----------

    def _on_language_change(self, _event=None):
        self.current_lang = self.lang_var.get()
        self._apply_translations()
        # Reload current tab in case labels changed formatting
        self._load_current_tab()

    def _on_tab_changed(self, _event=None):
        tab_id = self.notebook.select()
        for key, frame in self.tab_frames.items():
            if str(frame) == str(tab_id):
                self.current_tab = key
                break
        self._load_current_tab()

    def _on_add_folder_clicked(self):
        folder = filedialog.askdirectory(title=self.t["dialog_select_folder"])
        if not folder:
            messagebox.showwarning(APP_TITLE, self.t["msg_no_folder"])
            return

        self._set_status(self.t["status_indexing"])

        # Run indexing in background
        threading.Thread(
            target=self._index_folder_thread, args=(folder, self.current_tab), daemon=True
        ).start()

    def _on_refresh_clicked(self):
        self._load_current_tab()
        self._set_status(self.t["status_done"])

    def _on_print_clicked(self):
        # Export current tree content to simple HTML for printing
        self._export_current_view_to_html()

    # ---------- Core Logic ----------

    def _index_folder_thread(self, folder: str, tab_key: str):
        try:
            if tab_key == "movies":
                self._index_movies(folder)
            elif tab_key == "music":
                self._index_music(folder)
            elif tab_key == "programs":
                self._index_programs(folder)
            elif tab_key == "documents":
                self._index_documents(folder)

            self._load_tab(tab_key)
            self._set_status(self.t["status_done"])
            messagebox.showinfo(APP_TITLE, self.t["msg_index_finished"])
        except Exception as e:
            self._set_status("Error")
            messagebox.showerror(APP_TITLE, str(e))

    def _index_movies(self, folder: str):
        for root, dirs, files in os.walk(folder):
            for fname in files:
                ext = Path(fname).suffix.lower()
                if ext not in MOVIE_EXTENSIONS:
                    continue
                full_path = str(Path(root) / fname)
                stat = os.stat(full_path)
                size = stat.st_size
                added_at = datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d %H:%M:%S")

                meta = {
                    "file_path": full_path,
                    "filename": fname,
                    "title": Path(fname).stem,
                    "year": "",
                    "genre": "",
                    "imdb_rating": "",
                    "poster_path": "",
                    "size_bytes": size,
                    "added_at": added_at,
                }

                # Try to fetch metadata from OMDb
                if OMDB_API_KEY:
                    omdb_data = fetch_movie_metadata_from_omdb(fname)
                    meta.update({k: v for k, v in omdb_data.items() if v})

                self.db.insert_movie(**meta)

    def _index_music(self, folder: str):
        for root, dirs, files in os.walk(folder):
            for fname in files:
                ext = Path(fname).suffix.lower()
                if ext not in MUSIC_EXTENSIONS:
                    continue
                full_path = str(Path(root) / fname)
                stat = os.stat(full_path)
                size = stat.st_size
                added_at = datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d %H:%M:%S")

                meta = {
                    "file_path": full_path,
                    "filename": fname,
                    "artist": "",
                    "album": "",
                    "title": Path(fname).stem,
                    "genre": "",
                    "size_bytes": size,
                    "added_at": added_at,
                }
                self.db.insert_music(**meta)

    def _index_programs(self, folder: str):
        for root, dirs, files in os.walk(folder):
            for fname in files:
                ext = Path(fname).suffix.lower()
                if ext not in PROGRAM_EXTENSIONS:
                    continue
                full_path = str(Path(root) / fname)
                stat = os.stat(full_path)
                size = stat.st_size
                added_at = datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d %H:%M:%S")

                meta = {
                    "file_path": full_path,
                    "filename": fname,
                    "app_name": Path(fname).stem,
                    "version": "",
                    "vendor": "",
                    "size_bytes": size,
                    "added_at": added_at,
                }
                self.db.insert_program(**meta)

    def _index_documents(self, folder: str):
        for root, dirs, files in os.walk(folder):
            for fname in files:
                ext = Path(fname).suffix.lower()
                if ext not in DOCUMENT_EXTENSIONS:
                    continue
                full_path = str(Path(root) / fname)
                stat = os.stat(full_path)
                size = stat.st_size
                added_at = datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d %H:%M:%S")

                meta = {
                    "file_path": full_path,
                    "filename": fname,
                    "title": Path(fname).stem,
                    "doc_type": ext.lstrip("."),
                    "size_bytes": size,
                    "added_at": added_at,
                }
                self.db.insert_document(**meta)

    def _load_all_tabs(self):
        self._load_tab("movies")
        self._load_tab("music")
        self._load_tab("programs")
        self._load_tab("documents")

    def _load_current_tab(self):
        self._load_tab(self.current_tab)

    def _clear_tree(self, tab_key: str):
        tree = self.trees[tab_key]
        for row in tree.get_children():
            tree.delete(row)

    def _load_tab(self, tab_key: str):
        self._clear_tree(tab_key)
        tree = self.trees[tab_key]

        if tab_key == "movies":
            rows = self.db.fetch_movies()
            for filename, title, year, genre, rating, size_bytes, added_at in rows:
                tree.insert(
                    "", "end",
                    values=(
                        filename,
                        title,
                        year,
                        genre,
                        rating,
                        readable_size(size_bytes),
                        added_at,
                    ),
                )
        elif tab_key == "music":
            rows = self.db.fetch_music()
            for filename, artist, album, title, genre, size_bytes, added_at in rows:
                tree.insert(
                    "", "end",
                    values=(
                        filename,
                        artist,
                        album,
                        title,
                        genre,
                        readable_size(size_bytes),
                        added_at,
                    ),
                )
        elif tab_key == "programs":
            rows = self.db.fetch_programs()
            for filename, app_name, version, vendor, size_bytes, added_at in rows:
                tree.insert(
                    "", "end",
                    values=(
                        filename,
                        app_name,
                        version,
                        vendor,
                        readable_size(size_bytes),
                        added_at,
                    ),
                )
        elif tab_key == "documents":
            rows = self.db.fetch_documents()
            for filename, title, doc_type, size_bytes, added_at in rows:
                tree.insert(
                    "", "end",
                    values=(
                        filename,
                        title,
                        doc_type,
                        readable_size(size_bytes),
                        added_at,
                    ),
                )

    # ---------- Export / Print ----------
    def _export_current_view_to_html(self):
        tree = self.trees[self.current_tab]
        columns = tree["columns"]
        rows = [tree.item(item_id, "values") for item_id in tree.get_children()]

        if not rows:
            messagebox.showinfo(APP_TITLE, "Keine Daten zum Exportieren.")
            return

        # Simple HTML table
        html_lines = [
            "<html><head><meta charset='utf-8'><title>Media Export</title></head><body>",
            f"<h2>{APP_TITLE} - {self.t.get('tab_' + self.current_tab, self.current_tab)}</h2>",
            "<table border='1' cellspacing='0' cellpadding='4'>",
            "<tr>",
        ]

        # Column headers
        for col in columns:
            header = ""
            if self.current_tab == "movies":
                mapping = {
                    "filename": self.t["column_filename"],
                    "title": self.t["column_title"],
                    "year": self.t["column_year"],
                    "genre": self.t["column_genre"],
                    "rating": self.t["column_rating"],
                    "size": self.t["column_size"],
                    "added": self.t["column_added"],
                }
                header = mapping.get(col, col)
            elif self.current_tab == "music":
                mapping = {
                    "filename": self.t["column_filename"],
                    "artist": self.t["column_artist"],
                    "album": self.t["column_album"],
                    "title": self.t["column_title"],
                    "genre": self.t["column_genre"],
                    "size": self.t["column_size"],
                    "added": self.t["column_added"],
                }
                header = mapping.get(col, col)
            elif self.current_tab == "programs":
                mapping = {
                    "filename": self.t["column_filename"],
                    "name": self.t["column_title"],
                    "version": "Version",
                    "vendor": "Vendor",
                    "size": self.t["column_size"],
                    "added": self.t["column_added"],
                }
                header = mapping.get(col, col)
            elif self.current_tab == "documents":
                mapping = {
                    "filename": self.t["column_filename"],
                    "title": self.t["column_title"],
                    "type": self.t["column_type"],
                    "size": self.t["column_size"],
                    "added": self.t["column_added"],
                }
                header = mapping.get(col, col)

            html_lines.append(f"<th>{header}</th>")

        html_lines.append("</tr>")

        for row in rows:
            html_lines.append("<tr>")
            for value in row:
                html_lines.append(f"<td>{value}</td>")
            html_lines.append("</tr>")

        html_lines.append("</table></body></html>")

        export_name = f"media_export_{self.current_tab}.html"
        try:
            with open(export_name, "w", encoding="utf-8") as f:
                f.write("\n".join(html_lines))

            webbrowser.open_new_tab(f"file://{Path(export_name).absolute()}")
            messagebox.showinfo(APP_TITLE, self.t["msg_print_export_success"])
        except Exception:
            messagebox.showerror(APP_TITLE, self.t["msg_print_export_error"])

    # ---------- Info & GitHub ----------
    def _open_github(self):
        webbrowser.open_new_tab(GITHUB_URL)

    def _show_info(self):
        messagebox.showinfo(self.t["info_title"], self.t["info_text"])


def main():
    app = MediaApp()
    app.mainloop()


if __name__ == "__main__":
    main()
