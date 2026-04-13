from __future__ import annotations

import csv
import ctypes
import difflib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import zipfile
import hashlib
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, StringVar, IntVar, Tk
from tkinter import ttk

import pystray
from PIL import Image, ImageTk


APP_NAME = "Citizen StarString Helper"
TASK_NAME = "Citizen StarString Helper"
LEGACY_TASK_NAMES = (
    "StarStrings Updater",
    "Citizen StarString Helper",
)
DEFAULT_LIVE_PATH = r"C:\Program Files\Roberts Space Industries\StarCitizen\LIVE"
DEFAULT_REPO = "https://github.com/MrKraken/StarStrings"
APP_UPDATE_REPO = "aj3/CitizenStarStringHelper-v2"
APP_VERSION = "2.2.2"
APP_UPDATE_CHECK_INTERVAL_MS = 6 * 60 * 60 * 1000
BLUEPRINT_SCAN_INTERVAL_MS = 15 * 60 * 1000
LANGUAGE_LINE = "g_language = english."
REFERRAL_URL = "https://www.robertsspaceindustries.com/enlist?referral=STAR-J66D-SPVW"
MAX_LOG_LINES = 300
UPDATER_HELPER_NAME = "Citizen StarString Updater Helper.exe"
MAX_DOWNLOAD_BYTES = 500 * 1024 * 1024  # 500 MB hard cap on any download
BLUEPRINT_LOG_PATTERN = re.compile(r"Received Blueprint:\s*(.*?):")
BLUEPRINT_TAG_PATTERN = re.compile(r"<[^>]+>")
BLUEPRINT_WIKI_BASE_URL = "https://starcitizen.tools/"
BLUEPRINT_WIKI_SEARCH_API_URL = "https://starcitizen.tools/api.php?action=query&list=search&srnamespace=0&srlimit=8&format=json&srsearch="
BLUEPRINT_WIKI_SEARCH_URL = "https://starcitizen.tools/index.php?search="
BLUEPRINT_WIKI_USER_AGENT = f"CitizenStarStringHelper/{APP_VERSION}"
BLUEPRINT_WIKI_TIMEOUT_SECONDS = 5
_blueprint_wiki_cache: dict[str, str] = {}
BLUEPRINT_CATEGORY_OPTIONS = ("Auto", "Armor", "Weapon", "Ammo", "Clothing", "Med", "Tool", "Attachment", "Component", "Unknown")
CRAFTING_DB_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vQQFvtTlpzUucwLkfWSvJ_qdDaqAIsfXP7Y6uH2OIlFi_zWrPHgq_R021aw3Ym6wND4APMIIQJOkp23"
    "/pub?gid=1537513559&single=true&output=csv"
)
CRAFTING_DB_TIMEOUT_SECONDS = 15
CRAFTING_DB_CACHE_MAX_AGE_HOURS = 24
_crafting_db: dict[str, list["CraftingMaterialEntry"]] | None = None
_crafting_lookup_cache: dict[str, list["CraftingMaterialEntry"] | None] = {}


def runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS")).resolve()
    return Path(__file__).resolve().parent


def data_dir() -> Path:
    if sys.platform == "win32":
        base = Path.home() / "AppData" / "Local"
    else:
        base = Path.home() / ".local" / "share"
    return base / "CitizenStarStringHelper"


APP_DIR = runtime_dir()
RESOURCE_DIR = resource_dir()
DATA_DIR = data_dir()
SETTINGS_PATH = DATA_DIR / "starstrings_settings.json"
STATE_PATH = DATA_DIR / "starstrings_state.json"
LOG_PATH = DATA_DIR / "starstrings_updater.log"
BACKUP_ROOT = DATA_DIR / "Backups"
PENDING_UPDATE_DIR = DATA_DIR / "PendingAppUpdate"
UPDATE_TRACE_PATH = PENDING_UPDATE_DIR / "update_trace.log"
CREATE_NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0


def migrate_legacy_data() -> None:
    # Migrate files that sat next to the old single-file exe
    legacy_files = {
        APP_DIR / "starstrings_settings.json": SETTINGS_PATH,
        APP_DIR / "starstrings_state.json": STATE_PATH,
        APP_DIR / "starstrings_updater.log": LOG_PATH,
    }
    for source, target in legacy_files.items():
        if source.exists() and not target.exists():
            shutil.move(str(source), str(target))

    legacy_backup_root = APP_DIR / "Backups"
    if legacy_backup_root.exists() and not BACKUP_ROOT.exists():
        shutil.move(str(legacy_backup_root), str(BACKUP_ROOT))

    # Migrate from old StarStringsUpdater data folder (v4 → v5)
    if sys.platform == "win32":
        old_data = Path.home() / "AppData" / "Local" / "StarStringsUpdater"
    else:
        old_data = Path.home() / ".local" / "share" / "StarStringsUpdater"
    if old_data.exists():
        old_legacy = {
            old_data / "starstrings_settings.json": SETTINGS_PATH,
            old_data / "starstrings_state.json": STATE_PATH,
            old_data / "starstrings_updater.log": LOG_PATH,
        }
        for source, target in old_legacy.items():
            if source.exists() and not target.exists():
                try:
                    shutil.copy2(str(source), str(target))
                except Exception:
                    pass
        old_backups = old_data / "Backups"
        if old_backups.exists() and not BACKUP_ROOT.exists():
            try:
                shutil.copytree(str(old_backups), str(BACKUP_ROOT))
            except Exception:
                pass


def app_command(script_path: Path | None = None) -> list[str]:
    if script_path is None:
        script_path = Path(__file__).resolve()
    if getattr(sys, "frozen", False):
        return [str(Path(sys.executable).resolve())]
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    python_exe = pythonw if pythonw.exists() else Path(sys.executable)
    return [str(python_exe), str(script_path.resolve())]


def run_process(command: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        capture_output=True,
        text=True,
        shell=False,
        creationflags=CREATE_NO_WINDOW,
    )


def add_windows_app_id() -> None:
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("citizen.starstring.helper.desktop")
    except Exception:
        pass


def ensure_taskbar_window(window: Tk) -> None:
    if sys.platform != "win32":
        return
    try:
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        if not hwnd:
            hwnd = window.winfo_id()
        style = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
        style = style & ~0x00000080  # WS_EX_TOOLWINDOW
        style = style | 0x00040000   # WS_EX_APPWINDOW
        ctypes.windll.user32.SetWindowLongW(hwnd, -20, style)
        ctypes.windll.user32.ShowWindow(hwnd, 5)
    except Exception:
        pass


def apply_dark_titlebar(window: Tk) -> None:
    if sys.platform != "win32":
        return
    try:
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        if not hwnd:
            hwnd = window.winfo_id()
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(ctypes.c_int(1)),
            ctypes.sizeof(ctypes.c_int),
        )
    except Exception:
        pass


def remove_appwindow_style(window: Tk) -> None:
    if sys.platform != "win32":
        return
    try:
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        if not hwnd:
            hwnd = window.winfo_id()
        style = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
        style = style & ~0x00040000   # WS_EX_APPWINDOW
        style = style | 0x00000080    # WS_EX_TOOLWINDOW
        ctypes.windll.user32.SetWindowLongW(hwnd, -20, style)
    except Exception:
        pass


@dataclass
class Settings:
    live_path: str = DEFAULT_LIVE_PATH
    interval_hours: int = 6
    github_repo: str = DEFAULT_REPO


@dataclass
class State:
    tracked_release_id: str = ""
    tracked_release_name: str = ""
    last_run_at: str = ""
    last_checked_at: str = ""
    last_update_at: str = ""
    blueprints_last_scanned_at: str = ""
    blueprints_last_scanned_release_id: str = ""
    blueprints_last_scanned_release_name: str = ""
    blueprint_category_overrides: dict[str, str] | None = None


@dataclass
class ReleaseInfo:
    release_id: str
    name: str
    tag: str
    published_at: str
    download_url: str
    asset_name: str


@dataclass
class AppReleaseInfo:
    version: str
    name: str
    download_url: str
    asset_name: str
    published_at: str
    digest: str = ""  # SHA256 hex from GitHub asset digest field, if available


@dataclass
class BlueprintRecord:
    name: str
    normalized_name: str
    inferred_category: str
    category_override: str
    contracts: list[str]
    learned: bool
    learned_count: int
    learned_sources: list[str]

    @property
    def category(self) -> str:
        return self.category_override or self.inferred_category

    @property
    def status(self) -> str:
        if self.learned:
            return "Learned"
        if self.contracts:
            return "Missing"
        return "Unknown"


@dataclass
class CraftingMaterialEntry:
    slot: str
    resource: str
    quantity: float


class UpdaterError(Exception):
    pass


class NoPublishedAppReleaseError(UpdaterError):
    pass


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def strip_markup(text: str) -> str:
    cleaned = BLUEPRINT_TAG_PATTERN.sub("", text or "")
    cleaned = cleaned.replace("\\n", "\n")
    return " ".join(cleaned.split()).strip()


def normalize_search_text(text: str) -> str:
    cleaned = strip_markup(text).lower()
    cleaned = re.sub(r"[^a-z0-9]+", " ", cleaned)
    return " ".join(cleaned.split())


def blueprint_wiki_url(name: str) -> str:
    page_name = name.strip().replace(" ", "_")
    safe_chars = "()!,-._~'"
    return f"{BLUEPRINT_WIKI_BASE_URL}{urllib.parse.quote(page_name, safe=safe_chars)}"


def blueprint_wiki_search_url(name: str) -> str:
    return f"{BLUEPRINT_WIKI_SEARCH_URL}{urllib.parse.quote(name.strip())}"


def resolve_blueprint_wiki_url(name: str) -> str:
    """Resolve the best Star Citizen Wiki URL for a blueprint name.

    Uses the MediaWiki full-text search API (action=query&list=search) which searches
    article content and titles — far more accurate than OpenSearch (title-only) for
    blueprints whose wiki article title differs from the in-game name (e.g. "Antium Arms
    Jet" → "Antium Armor Arms Jet"). Candidates are ranked by word coverage (all query
    words present in the title) weighted above character sequence similarity.
    """
    normalized_name = normalize_search_text(name)
    if not normalized_name:
        return BLUEPRINT_WIKI_BASE_URL
    cached = _blueprint_wiki_cache.get(normalized_name)
    if cached:
        return cached

    fallback = blueprint_wiki_search_url(name)
    request = urllib.request.Request(
        f"{BLUEPRINT_WIKI_SEARCH_API_URL}{urllib.parse.quote(name.strip())}",
        headers={"User-Agent": BLUEPRINT_WIKI_USER_AGENT},
    )

    try:
        with urllib.request.urlopen(request, timeout=BLUEPRINT_WIKI_TIMEOUT_SECONDS) as response:
            payload = json.loads(response.read().decode("utf-8", errors="ignore"))
        results = payload.get("query", {}).get("search", [])
        query_words = set(normalized_name.split())
        candidates: list[tuple[float, str]] = []
        for result in results:
            title = str(result.get("title", "")).strip()
            if not title:
                continue
            normalized_title = normalize_search_text(title)
            if not normalized_title:
                continue
            title_words = set(normalized_title.split())
            # Weight word coverage heavily — catches "Antium Armor Arms Jet" for "Antium Arms Jet"
            word_coverage = len(query_words & title_words) / len(query_words) if query_words else 0
            seq_ratio = difflib.SequenceMatcher(None, normalized_name, normalized_title).ratio()
            score = word_coverage * 0.65 + seq_ratio * 0.35
            url = blueprint_wiki_url(title)
            candidates.append((score, url))
        if candidates:
            candidates.sort(key=lambda item: item[0], reverse=True)
            best_url = candidates[0][1]
            _blueprint_wiki_cache[normalized_name] = best_url
            return best_url
    except Exception:
        pass

    _blueprint_wiki_cache[normalized_name] = fallback
    return fallback


def format_scu(value: float) -> str:
    """Format an SCU quantity from the crafting DB as SCU or µSCU.

    1 µSCU = 0.000001 SCU (1 SCU = 1,000,000 µSCU).
    - Values < 1 SCU are expressed in µSCU (always whole numbers for DB values).
    - Values ≥ 1 SCU are expressed in SCU using :g format (no trailing zeros).
    Float noise is removed before conversion (0.02999999933 → 30,000 µSCU).
    """
    clean = round(value, 6)
    if clean <= 0:
        return "0 µSCU"
    if clean < 1.0:
        muscu = round(clean * 1_000_000)
        return f"{muscu:,} µSCU"
    return f"{clean:g} SCU"


def _crafting_db_cache_path() -> Path:
    return DATA_DIR / "crafting_db.json"


def _parse_crafting_csv(text: str) -> dict[str, list[CraftingMaterialEntry]]:
    reader = csv.DictReader(io.StringIO(text))
    db: dict[str, list[CraftingMaterialEntry]] = {}
    for row in reader:
        name = str(row.get("Blueprint Name") or "").strip()
        slot = str(row.get("Material Slot") or "").strip()
        resource = str(row.get("Resource Name") or "").strip()
        raw_qty = str(row.get("Quantity (SCU)") or "0").strip()
        if not name or not resource:
            continue
        try:
            qty = float(raw_qty)
        except ValueError:
            qty = 0.0
        key = normalize_search_text(name)
        db.setdefault(key, []).append(CraftingMaterialEntry(slot=slot, resource=resource, quantity=qty))
    return db


def _fetch_crafting_db_from_url() -> dict[str, list[CraftingMaterialEntry]]:
    request = urllib.request.Request(CRAFTING_DB_URL, headers={"User-Agent": BLUEPRINT_WIKI_USER_AGENT})
    with urllib.request.urlopen(request, timeout=CRAFTING_DB_TIMEOUT_SECONDS) as response:
        text = response.read().decode("utf-8", errors="ignore")
    return _parse_crafting_csv(text)


def _load_crafting_db_from_cache() -> dict[str, list[CraftingMaterialEntry]] | None:
    cache_path = _crafting_db_cache_path()
    if not cache_path.exists():
        return None
    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
        fetched_at = datetime.fromisoformat(str(payload.get("fetched_at", "")))
        age_hours = (datetime.now() - fetched_at).total_seconds() / 3600
        if age_hours > CRAFTING_DB_CACHE_MAX_AGE_HOURS:
            return None  # Stale — will re-fetch
        db: dict[str, list[CraftingMaterialEntry]] = {}
        for key, items in (payload.get("entries") or {}).items():
            db[key] = [CraftingMaterialEntry(**item) for item in items]
        return db
    except Exception:
        return None


def _save_crafting_db_to_cache(db: dict[str, list[CraftingMaterialEntry]]) -> None:
    try:
        entries_raw = {
            key: [{"slot": e.slot, "resource": e.resource, "quantity": e.quantity} for e in items]
            for key, items in db.items()
        }
        payload = {"fetched_at": datetime.now().isoformat(), "entries": entries_raw}
        _crafting_db_cache_path().write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def ensure_crafting_db_loaded() -> dict[str, list[CraftingMaterialEntry]]:
    """Return the crafting DB, loading from disk cache or fetching live if needed."""
    global _crafting_db
    if _crafting_db is not None:
        return _crafting_db
    cached = _load_crafting_db_from_cache()
    if cached is not None:
        _crafting_db = cached
        return _crafting_db
    db = _fetch_crafting_db_from_url()
    _save_crafting_db_to_cache(db)
    _crafting_db = db
    return _crafting_db


def lookup_crafting_materials(name: str) -> list[CraftingMaterialEntry] | None:
    """Find crafting materials for a blueprint name using exact then fuzzy matching.

    Uses the same word-coverage scoring as wiki resolution so variant names like
    'A03 "Canuto" Sniper Rifle' correctly map to the base 'A03 Sniper Rifle' entry
    when no exact match exists.
    """
    db = _crafting_db
    if db is None:
        return None
    normalized = normalize_search_text(name)
    if normalized in _crafting_lookup_cache:
        return _crafting_lookup_cache[normalized]
    # Exact match
    if normalized in db:
        result = db[normalized]
        _crafting_lookup_cache[normalized] = result
        return result
    # Fuzzy match
    query_words = set(normalized.split())
    best_score, best_key = 0.0, None
    for key in db:
        key_words = set(key.split())
        word_coverage = len(query_words & key_words) / len(query_words) if query_words else 0
        score = word_coverage * 0.65 + difflib.SequenceMatcher(None, normalized, key).ratio() * 0.35
        if score > best_score:
            best_score, best_key = score, key
    result = db[best_key] if best_key and best_score >= 0.72 else None
    _crafting_lookup_cache[normalized] = result
    return result


def fuzzy_query_match(query: str, haystack: str) -> bool:
    if not query:
        return True
    if query in haystack:
        return True
    query_tokens = query.split()
    haystack_tokens = haystack.split()
    for token in query_tokens:
        if token in haystack:
            continue
        if not any(
            token in candidate
            or candidate in token
            or difflib.SequenceMatcher(None, token, candidate).ratio() >= 0.66
            for candidate in haystack_tokens
        ):
            return False
    return True


def infer_blueprint_category(name: str) -> str:
    normalized = normalize_search_text(name)
    if not normalized:
        return "Unknown"

    checks: list[tuple[str, tuple[str, ...]]] = [
        ("Ammo", ("magazine", "battery", "drum", "rocket", "missile", "ammo", "cap")),
        ("Weapon", ("rifle", "pistol", "smg", "shotgun", "sniper", "launcher", "cannon", "gun", "laser", "knife")),
        ("Armor", ("helmet", "arms", "arm", "legs", "leg", "core", "torso", "armor", "vest")),
        ("Clothing", ("jacket", "shirt", "pants", "boots", "gloves", "hat", "mask", "beanie")),
        ("Med", ("medgun", "med gun", "medical", "medpen", "pen", "injector")),
        ("Tool", ("tool", "tractor", "multitool", "multi tool", "salvage", "mining", "cutter")),
        ("Attachment", ("scope", "sight", "barrel", "muzzle", "compensator", "suppressor", "grip", "rail")),
        ("Component", ("cooler", "power plant", "shield", "generator", "drive", "radar", "computer", "turret")),
    ]
    for category, tokens in checks:
        if any(token in normalized for token in tokens):
            return category
    return "Unknown"


def format_timestamp(value: str, fallback: str) -> str:
    text = (value or "").strip()
    if not text:
        return fallback
    try:
        parsed = datetime.fromisoformat(text)
        return parsed.strftime("%b %d, %Y %I:%M:%S %p")
    except ValueError:
        return text


def format_scheduler_timestamp(value: str, fallback: str) -> str:
    text = (value or "").strip()
    if not text:
        return fallback
    lowered = text.lower()
    if lowered in {"n/a", "never", "disabled", "unknown"}:
        return fallback
    if text.startswith("11/30/1999") or text.startswith("11/29/1999"):
        return fallback
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%b %d, %Y %I:%M:%S %p")
        except ValueError:
            pass
    return text


_log_line_count: int = 0  # approximate in-memory count; avoids re-reading the file on every write


def log(message: str) -> None:
    global _log_line_count
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"[{now_text()}] {message}\n")
    _log_line_count += 1
    if _log_line_count > MAX_LOG_LINES * 2:
        try:
            lines = LOG_PATH.read_text(encoding="utf-8").splitlines(keepends=True)
            LOG_PATH.write_text("".join(lines[-MAX_LOG_LINES:]), encoding="utf-8")
            _log_line_count = MAX_LOG_LINES
        except Exception:
            pass


def read_log_tail(limit: int = MAX_LOG_LINES) -> str:
    if not LOG_PATH.exists():
        return "Ready.\n"
    try:
        with LOG_PATH.open("r", encoding="utf-8") as handle:
            lines = handle.readlines()
        return "".join(lines[-limit:]) or "Ready.\n"
    except Exception:
        return "Ready.\n"


def load_settings() -> Settings:
    if not SETTINGS_PATH.exists():
        return Settings()
    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        return Settings(
            live_path=str(data.get("live_path") or DEFAULT_LIVE_PATH),
            interval_hours=max(1, min(24, int(data.get("interval_hours") or 6))),
            github_repo=str(data.get("github_repo") or DEFAULT_REPO),
        )
    except Exception as exc:
        log(f"Failed to load settings, using defaults. {exc}")
        return Settings()


def save_settings(settings: Settings) -> None:
    SETTINGS_PATH.write_text(
        json.dumps(
            {
                "live_path": settings.live_path,
                "interval_hours": settings.interval_hours,
                "github_repo": settings.github_repo,
            },
            indent=2,
        ),
        encoding="utf-8",
    )


def normalize_repo(repo_value: str) -> str:
    value = repo_value.strip().rstrip("/")
    if value.startswith("https://github.com/"):
        value = value.removeprefix("https://github.com/")
    value = value.strip("/")
    parts = [part for part in value.split("/") if part]
    if len(parts) < 2:
        raise UpdaterError("GitHub repository must look like owner/repo or a GitHub URL.")
    return f"{parts[0]}/{parts[1]}"


def canonical_repo_url(repo_value: str) -> str:
    return f"https://github.com/{normalize_repo(repo_value)}"


def compact_repo_name(repo_value: str) -> str:
    return normalize_repo(repo_value)


def repo_api_url(repo_value: str) -> str:
    return f"https://api.github.com/repos/{normalize_repo(repo_value)}/releases/latest"


def parse_version(version_text: str) -> tuple[int, ...]:
    cleaned = version_text.strip().lower().lstrip("v")
    parts: list[int] = []
    for piece in cleaned.split("."):
        digits = "".join(ch for ch in piece if ch.isdigit())
        parts.append(int(digits or "0"))
    return tuple(parts)


def is_newer_version(candidate: str, current: str) -> bool:
    return parse_version(candidate) > parse_version(current)


def load_state() -> State:
    if not STATE_PATH.exists():
        return State()
    try:
        data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        return State(
            tracked_release_id=str(data.get("tracked_release_id") or data.get("last_release_id") or ""),
            tracked_release_name=str(data.get("tracked_release_name") or data.get("last_release_name") or ""),
            last_run_at=str(data.get("last_run_at") or ""),
            last_checked_at=str(data.get("last_checked_at") or ""),
            last_update_at=str(data.get("last_update_at") or ""),
            blueprints_last_scanned_at=str(data.get("blueprints_last_scanned_at") or ""),
            blueprints_last_scanned_release_id=str(data.get("blueprints_last_scanned_release_id") or ""),
            blueprints_last_scanned_release_name=str(data.get("blueprints_last_scanned_release_name") or ""),
            blueprint_category_overrides={
                str(key): str(value)
                for key, value in (data.get("blueprint_category_overrides") or {}).items()
                if isinstance(key, str) and isinstance(value, str)
            },
        )
    except Exception as exc:
        log(f"Failed to load state, using empty state. {exc}")
        return State()


def save_state(state: State) -> None:
    STATE_PATH.write_text(
        json.dumps(
            {
                "tracked_release_id": state.tracked_release_id,
                "tracked_release_name": state.tracked_release_name,
                "last_run_at": state.last_run_at,
                "last_checked_at": state.last_checked_at,
                "last_update_at": state.last_update_at,
                "blueprints_last_scanned_at": state.blueprints_last_scanned_at,
                "blueprints_last_scanned_release_id": state.blueprints_last_scanned_release_id,
                "blueprints_last_scanned_release_name": state.blueprints_last_scanned_release_name,
                "blueprint_category_overrides": state.blueprint_category_overrides or {},
            },
            indent=2,
        ),
        encoding="utf-8",
    )


def github_headers() -> dict[str, str]:
    return {
        "User-Agent": APP_NAME,
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def read_localization_entries(global_ini_path: Path) -> dict[str, str]:
    entries: dict[str, str] = {}
    with global_ini_path.open("r", encoding="utf-8", errors="ignore") as handle:
        for raw_line in handle:
            line = raw_line.rstrip("\r\n")
            if not line or "=" not in line:
                continue
            key, value = line.split("=", 1)
            entries[key.strip()] = value
    return entries


def title_key_candidates(description_key: str) -> list[str]:
    candidates: list[str] = []
    if "_desc" in description_key:
        candidates.append(re.sub(r"_desc(,P)?$", r"_title\1", description_key))
    if "_Repeat_desc" in description_key:
        candidates.append(re.sub(r"_Repeat_desc(,P)?$", r"_Repeat_title\1", description_key))
    if "_Desc_" in description_key:
        candidates.append(description_key.replace("_Desc_", "_Title_"))
    if "_desc_" in description_key:
        candidates.append(description_key.replace("_desc_", "_title_"))
    if description_key.endswith(",P"):
        candidates.append(description_key[:-2])
    seen: set[str] = set()
    ordered: list[str] = []
    for candidate in candidates:
        for variant in (candidate, candidate.replace(",P", ""), f"{candidate},P" if not candidate.endswith(",P") else candidate):
            if variant and variant not in seen:
                seen.add(variant)
                ordered.append(variant)
    return ordered


def extract_blueprint_names_from_description(description: str) -> list[str]:
    lines = [strip_markup(line).strip() for line in description.split("\\n")]
    collecting = False
    blueprints: list[str] = []
    for line in lines:
        if not line:
            continue
        if "Potential Blueprints" in line:
            collecting = True
            continue
        if not collecting:
            continue
        if line.startswith("[") and line.endswith("]"):
            continue
        if "Regional Variants" in line:
            continue
        if line.startswith("- "):
            name = line[2:].strip()
            if name:
                blueprints.append(name)
            continue
        if blueprints:
            break
    return blueprints


def parse_starstrings_blueprints(live_path: str) -> tuple[dict[str, dict[str, object]], Path]:
    global_ini_path = Path(live_path) / "Data" / "Localization" / "english" / "global.ini"
    if not global_ini_path.exists():
        raise UpdaterError(f"Could not find StarStrings localization data at {global_ini_path}")

    entries = read_localization_entries(global_ini_path)
    blueprints: dict[str, dict[str, object]] = {}
    for key, value in entries.items():
        if "Potential Blueprints" not in value:
            continue
        contract_title = ""
        for candidate in title_key_candidates(key):
            if candidate in entries:
                possible_title = strip_markup(entries[candidate]).replace("[BP]*", "").replace("[BP]", "").strip()
                if possible_title and "Potential Blueprints" not in possible_title and len(possible_title) < 140:
                    contract_title = possible_title
                    break
        if not contract_title:
            contract_title = strip_markup(key)
        for blueprint_name in extract_blueprint_names_from_description(value):
            normalized = normalize_search_text(blueprint_name)
            if not normalized:
                continue
            record = blueprints.setdefault(
                normalized,
                {"name": blueprint_name, "contracts": set()},
            )
            record["contracts"].add(contract_title)
    return blueprints, global_ini_path


def parse_learned_blueprints(live_path: str) -> tuple[dict[str, dict[str, object]], list[Path]]:
    live_dir = Path(live_path)
    log_paths: list[Path] = []
    game_log = live_dir / "game.log"
    if game_log.exists():
        log_paths.append(game_log)
    backup_dir = live_dir / "logbackups"
    if backup_dir.exists():
        log_paths.extend(sorted(backup_dir.glob("*.log")))

    learned: dict[str, dict[str, object]] = {}
    for log_path in log_paths:
        try:
            with log_path.open("r", encoding="utf-8", errors="ignore") as handle:
                for line in handle:
                    match = BLUEPRINT_LOG_PATTERN.search(line)
                    if not match:
                        continue
                    blueprint_name = strip_markup(match.group(1))
                    normalized = normalize_search_text(blueprint_name)
                    if not normalized:
                        continue
                    record = learned.setdefault(
                        normalized,
                        {"name": blueprint_name, "count": 0, "sources": set()},
                    )
                    record["count"] += 1
                    record["sources"].add(log_path.name)
        except Exception:
            continue
    return learned, log_paths


def collect_blueprint_records(live_path: str) -> tuple[list[BlueprintRecord], dict[str, object]]:
    starstrings_records, global_ini_path = parse_starstrings_blueprints(live_path)
    learned_records, log_paths = parse_learned_blueprints(live_path)
    state = load_state()
    overrides = state.blueprint_category_overrides or {}

    all_keys = sorted(set(starstrings_records) | set(learned_records))
    results: list[BlueprintRecord] = []
    for key in all_keys:
        starstrings_record = starstrings_records.get(key) or {}
        learned_record = learned_records.get(key) or {}
        contracts = sorted(starstrings_record.get("contracts", set()))
        learned_sources = sorted(learned_record.get("sources", set()))
        display_name = (
            str(learned_record.get("name") or "")
            or str(starstrings_record.get("name") or "")
            or key
        )
        results.append(
            BlueprintRecord(
                name=display_name,
                normalized_name=key,
                inferred_category=infer_blueprint_category(display_name),
                category_override=overrides.get(key, ""),
                contracts=contracts,
                learned=bool(learned_record),
                learned_count=int(learned_record.get("count") or 0),
                learned_sources=learned_sources,
            )
        )

    metadata = {
        "global_ini_path": global_ini_path,
        "log_paths": log_paths,
        "learned_count": sum(1 for record in results if record.learned),
        "available_count": sum(1 for record in results if record.contracts),
        "missing_count": sum(1 for record in results if record.contracts and not record.learned),
        "total_count": len(results),
        "scanned_at": datetime.now().isoformat(),
        "tracked_release_id": state.tracked_release_id,
        "tracked_release_name": state.tracked_release_name,
    }
    return results, metadata


def fetch_latest_release(repo_value: str) -> ReleaseInfo:
    request = urllib.request.Request(repo_api_url(repo_value), headers=github_headers())
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.URLError as exc:
        raise UpdaterError(f"Could not reach GitHub: {exc}") from exc

    assets = payload.get("assets") or []
    zip_asset = next((asset for asset in assets if str(asset.get("name", "")).lower().endswith(".zip")), None)
    download_url = zip_asset["browser_download_url"] if zip_asset else payload.get("zipball_url")
    if not download_url:
        raise UpdaterError("GitHub release did not include a downloadable ZIP.")

    asset_name = zip_asset["name"] if zip_asset else "StarStrings-source.zip"
    return ReleaseInfo(
        release_id=str(payload.get("id") or ""),
        name=str(payload.get("name") or payload.get("tag_name") or "Unnamed release"),
        tag=str(payload.get("tag_name") or ""),
        published_at=str(payload.get("published_at") or ""),
        download_url=str(download_url),
        asset_name=str(asset_name),
    )


def fetch_latest_app_release() -> AppReleaseInfo:
    request = urllib.request.Request(repo_api_url(APP_UPDATE_REPO), headers=github_headers())
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise NoPublishedAppReleaseError("No app release has been published on GitHub yet.") from exc
        raise UpdaterError(f"Could not reach GitHub for app updates: {exc}") from exc
    except urllib.error.URLError as exc:
        raise UpdaterError(f"Could not reach GitHub for app updates: {exc}") from exc

    assets = payload.get("assets") or []
    exe_asset = next((a for a in assets if str(a.get("name", "")).lower().endswith(".exe")), None)
    if not exe_asset:
        raise UpdaterError("Latest app release does not include a downloadable .exe asset.")

    version = str(payload.get("tag_name") or payload.get("name") or "").strip()
    if not version:
        raise UpdaterError("Latest app release did not include a version tag.")

    # GitHub may provide a digest in the format "sha256:<hex>" for the asset.
    raw_digest = str(exe_asset.get("digest") or "")
    sha256_digest = raw_digest.removeprefix("sha256:") if raw_digest.startswith("sha256:") else ""

    return AppReleaseInfo(
        version=version,
        name=str(payload.get("name") or version),
        download_url=str(exe_asset.get("browser_download_url") or ""),
        asset_name=str(exe_asset.get("name") or "Citizen StarString Helper.exe"),
        published_at=str(payload.get("published_at") or ""),
        digest=sha256_digest,
    )


def _copy_response_bounded(response, output, max_bytes: int = MAX_DOWNLOAD_BYTES) -> None:
    """Copy an HTTP response to a file, raising UpdaterError if it exceeds max_bytes."""
    received = 0
    chunk = 1024 * 64
    while True:
        block = response.read(chunk)
        if not block:
            break
        received += len(block)
        if received > max_bytes:
            raise UpdaterError(
                f"Download exceeded the {max_bytes // (1024 * 1024)} MB safety limit and was aborted."
            )
        output.write(block)


def download_file(url: str, destination: Path) -> Path:
    request = urllib.request.Request(url, headers=github_headers())
    try:
        with urllib.request.urlopen(request, timeout=60) as response, destination.open("wb") as output:
            _copy_response_bounded(response, output)
    except urllib.error.URLError as exc:
        raise UpdaterError(f"Could not download update file: {exc}") from exc
    return destination




def append_update_trace(message: str) -> None:
    try:
        PENDING_UPDATE_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with UPDATE_TRACE_PATH.open("a", encoding="utf-8") as handle:
            handle.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass


def cleanup_pending_update_artifacts() -> None:
    try:
        PENDING_UPDATE_DIR.mkdir(parents=True, exist_ok=True)
        cutoff = time.time() - (12 * 60 * 60)
        for path in PENDING_UPDATE_DIR.iterdir():
            if path.name in {"last_update.json", "update_trace.log"}:
                continue
            if path.is_dir():
                continue
            if path.stat().st_mtime >= cutoff:
                continue
            if path.suffix.lower() in {".cmd", ".ps1", ".exe"}:
                try:
                    path.unlink()
                except Exception:
                    pass
    except Exception:
        pass


def updater_helper_resource() -> Path:
    helper_path = RESOURCE_DIR / UPDATER_HELPER_NAME
    if not helper_path.exists():
        raise UpdaterError(f"Bundled updater helper is missing: {helper_path}")
    return helper_path


def install_app_update(downloaded_exe: Path, release_version: str) -> None:
    """Self-update via a bundled helper executable extracted from the main app."""
    if not getattr(sys, "frozen", False):
        raise UpdaterError("App self-update is only available from the packaged executable.")

    current_exe = Path(sys.executable).resolve()
    current_pid = os.getpid()
    PENDING_UPDATE_DIR.mkdir(parents=True, exist_ok=True)
    cleanup_pending_update_artifacts()
    append_update_trace(f"Staging update to v{release_version}. exe={current_exe}")

    staged_exe = PENDING_UPDATE_DIR / f"Citizen StarString Helper-{release_version}.exe"
    shutil.copy2(downloaded_exe, staged_exe)
    helper_exe = PENDING_UPDATE_DIR / f"updater-helper-{uuid.uuid4().hex[:8]}.exe"
    shutil.copy2(updater_helper_resource(), helper_exe)

    backup_exe  = PENDING_UPDATE_DIR / f"CSH_old_{uuid.uuid4().hex[:8]}.exe"
    result_file = PENDING_UPDATE_DIR / "last_update.json"
    append_update_trace(f"Launching updater helper {helper_exe.name} for v{release_version}.")

    subprocess.Popen(
        [
            str(helper_exe),
            "--current-exe", str(current_exe),
            "--new-exe", str(staged_exe),
            "--backup-exe", str(backup_exe),
            "--result-file", str(result_file),
            "--trace-file", str(UPDATE_TRACE_PATH),
            "--version", release_version,
            "--pid", str(current_pid),
        ],
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        close_fds=True,
    )


def ensure_live_path(live_path: Path, allow_prompt: bool, parent: Tk | None = None) -> Path:
    if live_path.is_dir():
        return live_path
    if not allow_prompt:
        raise UpdaterError(f"Configured LIVE path was not found: {live_path}")
    chosen = filedialog.askdirectory(
        parent=parent,
        title="Select your Star Citizen LIVE folder",
        mustexist=True,
        initialdir=live_path.drive + "\\" if live_path.drive else str(Path.home()),
    )
    if not chosen:
        raise UpdaterError("No LIVE folder was selected.")
    return Path(chosen)


def find_release_content(extract_root: Path) -> tuple[Path, Path]:
    data_dir: Path | None = None
    user_cfg: Path | None = None
    for path in extract_root.rglob("*"):
        if data_dir is None and path.name == "Data" and path.is_dir():
            data_dir = path
        if user_cfg is None and path.is_file() and path.name.lower() == "user.cfg":
            user_cfg = path
        if data_dir is not None and user_cfg is not None:
            break
    if data_dir is None:
        raise UpdaterError("Downloaded release did not include a Data directory.")
    if user_cfg is None:
        raise UpdaterError("Downloaded release did not include USER.cfg.")
    return data_dir, user_cfg


def create_backup_snapshot(live_path: Path, source_data_path: Path, tracked_release_name: str) -> Path:
    BACKUP_ROOT.mkdir(parents=True, exist_ok=True)
    backup_path = BACKUP_ROOT / datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path.mkdir(parents=True, exist_ok=True)
    (backup_path / "backup_meta.json").write_text(
        json.dumps(
            {
                "tracked_release_name": tracked_release_name,
                "created_at": datetime.now().isoformat(),
            },
            indent=2,
        ),
        encoding="utf-8",
    )

    target_data_root = live_path / "Data"
    if target_data_root.exists():
        for source_file in source_data_path.rglob("*"):
            if not source_file.is_file():
                continue
            relative = source_file.relative_to(source_data_path)
            target_file = target_data_root / relative
            if target_file.is_file():
                destination = backup_path / "Data" / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target_file, destination)

    user_cfg = live_path / "USER.cfg"
    if user_cfg.is_file():
        shutil.copy2(user_cfg, backup_path / "USER.cfg")

    return backup_path


def merge_user_cfg(target_user_cfg: Path, source_user_cfg: Path) -> str:
    if not target_user_cfg.exists():
        shutil.copy2(source_user_cfg, target_user_cfg)
        return "Copied USER.cfg"

    existing_lines = target_user_cfg.read_text(encoding="utf-8", errors="ignore").splitlines()
    cleaned = [line for line in existing_lines if line.strip().lower() != LANGUAGE_LINE.lower()]
    if cleaned and cleaned[-1] != "":
        cleaned.append("")
    cleaned.append(LANGUAGE_LINE)
    target_user_cfg.write_text("\n".join(cleaned) + "\n", encoding="utf-8")
    return "Merged USER.cfg"


def overlay_directory(source_dir: Path, target_dir: Path) -> None:
    for source_path in source_dir.rglob("*"):
        relative = source_path.relative_to(source_dir)
        destination = target_dir / relative
        if source_path.is_dir():
            destination.mkdir(parents=True, exist_ok=True)
        else:
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_path, destination)


def list_backups() -> list[Path]:
    """Return backup snapshot folders sorted newest-first."""
    if not BACKUP_ROOT.exists():
        return []
    return sorted(
        (p for p in BACKUP_ROOT.iterdir() if p.is_dir()),
        reverse=True,
    )


def read_backup_release_name(backup_path: Path) -> str:
    meta_path = backup_path / "backup_meta.json"
    if not meta_path.exists():
        return "Unknown release"
    try:
        data = json.loads(meta_path.read_text(encoding="utf-8"))
        return str(data.get("tracked_release_name") or "Unknown release")
    except Exception:
        return "Unknown release"


def restore_backup(backup_path: Path, live_path: Path) -> str:
    """Copy a backup snapshot back into the Star Citizen LIVE folder."""
    data_src = backup_path / "Data"
    cfg_src  = backup_path / "USER.cfg"
    restored: list[str] = []

    if data_src.exists():
        target_data = live_path / "Data"
        target_data.mkdir(parents=True, exist_ok=True)
        overlay_directory(data_src, target_data)
        restored.append("Data/")

    if cfg_src.exists():
        shutil.copy2(cfg_src, live_path / "USER.cfg")
        restored.append("USER.cfg")

    if not restored:
        raise UpdaterError("Backup snapshot contains no restorable files.")

    return f"Restored {', '.join(restored)} from backup '{backup_path.name}'."


def install_release(release: ReleaseInfo, live_path: Path) -> tuple[str, Path]:
    with tempfile.TemporaryDirectory(prefix="starstrings-updater-") as temp_dir:
        temp_root = Path(temp_dir)
        zip_path = temp_root / release.asset_name
        extract_path = temp_root / "extract"
        extract_path.mkdir(parents=True, exist_ok=True)

        request = urllib.request.Request(release.download_url, headers=github_headers())
        try:
            with urllib.request.urlopen(request, timeout=60) as response, zip_path.open("wb") as output:
                _copy_response_bounded(response, output)
        except urllib.error.URLError as exc:
            raise UpdaterError(f"Could not download release ZIP: {exc}") from exc

        extract_resolved = extract_path.resolve()
        with zipfile.ZipFile(zip_path) as archive:
            for member in archive.namelist():
                member_path = (extract_path / member).resolve()
                if not member_path.is_relative_to(extract_resolved):
                    raise UpdaterError(f"ZIP archive contains unsafe path: {member}")
            archive.extractall(extract_path)

        source_data_path, source_user_cfg = find_release_content(extract_path)
        backup_path = create_backup_snapshot(live_path, source_data_path, release.name)

        target_data = live_path / "Data"
        target_data.mkdir(parents=True, exist_ok=True)
        overlay_directory(source_data_path, target_data)
        user_cfg_result = merge_user_cfg(live_path / "USER.cfg", source_user_cfg)
        return user_cfg_result, backup_path


def run_update(settings: Settings, allow_prompt: bool, parent: Tk | None = None, force_update: bool = False) -> str:
    live_path = ensure_live_path(Path(settings.live_path), allow_prompt=allow_prompt, parent=parent)
    if str(live_path) != settings.live_path:
        settings.live_path = str(live_path)
        save_settings(settings)

    state = load_state()
    release = fetch_latest_release(settings.github_repo)
    now_iso = datetime.now().isoformat()

    if not force_update and not state.tracked_release_id:
        state.tracked_release_id = release.release_id
        state.tracked_release_name = release.name
        state.last_run_at = now_iso
        state.last_checked_at = now_iso
        save_state(state)
        message = f"Tracking initialized for '{settings.github_repo}' at release '{release.name}'. No files changed."
        log(message)
        return message

    if not force_update and state.tracked_release_id == release.release_id:
        state.last_run_at = now_iso
        state.last_checked_at = now_iso
        save_state(state)
        message = f"No GitHub changes detected for '{settings.github_repo}'. Latest release is still '{release.name}'."
        log(message)
        return message

    previous_name = state.tracked_release_name or "untracked"
    if force_update:
        log(f"Manual update requested for '{settings.github_repo}'. Reinstalling release '{release.name}' regardless of tracked state.")
    else:
        log(f"GitHub change detected for '{settings.github_repo}': '{previous_name}' -> '{release.name}'.")

    user_cfg_result, backup_path = install_release(release, live_path)
    state.tracked_release_id = release.release_id
    state.tracked_release_name = release.name
    state.last_run_at = now_iso
    state.last_checked_at = now_iso
    state.last_update_at = now_iso
    save_state(state)
    log(f"USER.cfg action: {user_cfg_result}.")
    if force_update:
        message = f"Manual update completed with '{release.name}' in '{live_path}'. {user_cfg_result}. Backup: '{backup_path}'."
    else:
        message = f"Updated to '{release.name}' in '{live_path}'. {user_cfg_result}. Backup: '{backup_path}'."
    log(message)
    return message


def scheduled_task_exists() -> bool:
    return any(_query_task_lines(task_name) is not None for task_name in LEGACY_TASK_NAMES)


def _query_task_lines(task_name: str) -> dict[str, str] | None:
    result = run_process(["schtasks", "/Query", "/TN", task_name, "/FO", "LIST", "/V"])
    if result.returncode != 0:
        return None
    lines: dict[str, str] = {}
    for raw_line in result.stdout.splitlines():
        if ":" not in raw_line:
            continue
        key, value = raw_line.split(":", 1)
        lines[key.strip()] = value.strip()
    return lines


def query_scheduled_task() -> str:
    lines = None
    active_task_name = None
    for task_name in LEGACY_TASK_NAMES:
        lines = _query_task_lines(task_name)
        if lines is not None:
            active_task_name = task_name
            break
    if lines is None:
        return "Scheduled task is not registered."

    status = lines.get("Status", "Unknown")
    last_run = lines.get("Last Run Time", "Unknown")
    next_run = lines.get("Next Run Time", "Unknown")
    if active_task_name and active_task_name != TASK_NAME:
        return f"Scheduled ({active_task_name}): {status} | Last Run: {last_run} | Next Run: {next_run}"
    return f"Scheduled: {status} | Last Run: {last_run} | Next Run: {next_run}"


def register_scheduled_task(interval_hours: int) -> None:
    if interval_hours < 1:
        raise UpdaterError("Interval must be at least 1 hour.")
    command = app_command()
    taskrun = " ".join(f'"{part}"' if " " in part else part for part in (command + ["--scheduled"]))
    if interval_hours >= 24:
        start_time = (datetime.now() + timedelta(minutes=2)).strftime("%H:%M")
        args = [
            "schtasks",
            "/Create",
            "/F",
            "/SC",
            "DAILY",
            "/MO",
            "1",
            "/ST",
            start_time,
            "/TN",
            TASK_NAME,
            "/TR",
            taskrun,
        ]
    else:
        args = [
            "schtasks",
            "/Create",
            "/F",
            "/SC",
            "HOURLY",
            "/MO",
            str(interval_hours),
            "/TN",
            TASK_NAME,
            "/TR",
            taskrun,
        ]

    # /F already overwrites an existing task with the same name; no pre-delete needed.
    # Legacy task names are cleaned up once at startup via migrate_legacy_data.
    result = run_process(args)
    if result.returncode != 0:
        raise UpdaterError(result.stderr.strip() or result.stdout.strip() or "Could not create scheduled task.")


def unregister_scheduled_task() -> None:
    errors: list[str] = []
    for task_name in LEGACY_TASK_NAMES:
        result = run_process(["schtasks", "/Delete", "/F", "/TN", task_name])
        if result.returncode == 0:
            continue
        output = result.stderr.strip() or result.stdout.strip()
        lowered = output.lower()
        if "cannot find the file specified" in lowered or "the system cannot find the file specified" in lowered:
            continue
        errors.append(f"{task_name}: {output or 'Could not delete scheduled task.'}")
    if errors:
        raise UpdaterError(" | ".join(errors))


class StarStringsApp:
    def __init__(self) -> None:
        add_windows_app_id()
        self.settings = load_settings()
        self.state = load_state()
        self.tray_icon: pystray.Icon | None = None
        self.tray_thread: threading.Thread | None = None
        self.is_quitting = False
        self.is_restoring = False
        self.icon_image = self._load_icon_image()
        cleanup_pending_update_artifacts()

        self.root = Tk()
        self.root.title(APP_NAME)
        self.root.geometry("900x940")
        self.root.minsize(840, 880)
        self.root.configure(bg="#080c10")
        self.root.after(50, lambda: ensure_taskbar_window(self.root))
        self.root.after(100, lambda: apply_dark_titlebar(self.root))
        self.root.bind("<Configure>", self._handle_configure)
        self.root.bind("<Unmap>", self._handle_unmap)
        self.root.protocol("WM_DELETE_WINDOW", self._handle_close)
        self.root.bind("<Map>", self._handle_map)
        self._apply_window_icon()

        self.live_path_var = StringVar(value=self.settings.live_path)
        self.interval_var = IntVar(value=self.settings.interval_hours)
        self.github_repo_var = StringVar(value=canonical_repo_url(self.settings.github_repo))
        self.repo_display_var = StringVar(value=compact_repo_name(self.settings.github_repo))
        self.schedule_var = StringVar(value="")
        self.operations_meta_var = StringVar(value=f"App version: v{APP_VERSION}")
        self.current_app_path_var = StringVar(value=str(Path(sys.executable).resolve() if getattr(sys, "frozen", False) else Path(__file__).resolve()))
        self.release_var = StringVar(value=self.state.tracked_release_name or "No release tracked yet")
        self.backup_var = StringVar(value=f"Backups stored in:\n{BACKUP_ROOT}")
        self.last_checked_var = StringVar(value=format_timestamp(self.state.last_checked_at, "Not checked yet"))
        self.last_updated_var = StringVar(value=format_timestamp(self.state.last_update_at, "No update applied yet"))
        self.auto_state_var = StringVar(value="")
        self.blueprint_search_var = StringVar(value="")
        self.blueprint_filter_var = StringVar(value="All")
        self.blueprint_type_filter_var = StringVar(value="All Types")
        self.blueprint_search_mode_var = StringVar(value="Strict")
        self.blueprint_status_var = StringVar(value="Scan your StarStrings data to see available and learned blueprints.")
        self.blueprint_summary_var = StringVar(value="No blueprint scan has been run yet.")
        self.blueprint_detail_title_var = StringVar(value="Select a blueprint")
        self.blueprint_detail_status_var = StringVar(value="")
        self.app_update_button_var = StringVar(value="Check for Updates")
        self.app_update_available = False
        self.app_update_release: AppReleaseInfo | None = None
        self.app_update_pulse_on = False
        self.app_update_pulse_job: str | None = None
        self.app_update_check_job: str | None = None
        self.blueprint_auto_scan_job: str | None = None
        self.settings_save_job: str | None = None
        self._app_update_checking = False  # guard against concurrent update checks
        self.settings_loaded = False
        self.current_view = "setup"
        self.inline_info_labels: list[ttk.Label] = []
        self.meta_labels: list[ttk.Label] = []
        self.header_art_photo = None
        self.blueprint_records: list[BlueprintRecord] = []
        self.filtered_blueprint_records: list[BlueprintRecord] = []
        self.blueprint_scan_metadata: dict[str, object] = {}
        self.blueprint_scan_in_progress = False
        self.blueprint_sort_column = "blueprint"
        self.blueprint_sort_desc = False
        self.selected_blueprint_record: BlueprintRecord | None = None
        self._search_debounce_job: str | None = None
        self._crafting_db_loading = False

        self._build_style()
        self._build_ui()
        self._refresh_toggle()
        self._refresh_status_vars()
        self._load_log()
        self._report_completed_staged_update()
        self.append_log(f"Running from: {self.current_app_path_var.get()}")
        self._bind_auto_save()
        self.blueprint_search_var.trace_add("write", lambda *_: self._schedule_blueprint_search())
        self.settings_loaded = True
        self.root.after(1500, self.check_for_app_update_silent)
        self.root.after(2500, self._refresh_blueprint_freshness)
        self.root.after(4000, self._schedule_blueprint_auto_scan)

    def _build_style(self) -> None:
        # ── RSI-inspired palette ─────────────────────────────────────────────
        BG       = "#080c10"   # near-black root
        CARD     = "#0d1219"   # dark navy cards
        PANEL    = "#111922"   # info-row panels
        GOLD     = "#c09040"   # RSI gold accent
        GOLD_H   = "#d4a84e"   # gold hover / highlight
        FG       = "#e8edf2"   # primary text
        MUTED    = "#6e8096"   # secondary / muted text
        BTN      = "#162030"   # button resting
        BTN_H    = "#1e2d3d"   # button hover
        SUCCESS  = "#1a4a2e"   # enabled / on green
        SUCC_H   = "#215e38"   # green hover
        WARN_A   = "#7a2a08"   # alert pulse A
        WARN_B   = "#9a3c10"   # alert pulse B
        INPUT_BG = "#0a0e14"   # text-input fields

        style = ttk.Style(self.root)
        style.theme_use("clam")

        # Frames
        style.configure("Root.TFrame",    background=BG)
        style.configure("Hero.TFrame",    background=BG)
        style.configure("Card.TFrame",    background=CARD)
        style.configure("SideCard.TFrame",background=CARD)
        style.configure("Panel.TFrame",   background=PANEL)
        style.configure("Stat.TFrame",    background=PANEL)
        style.configure("Accent.TFrame",  background=GOLD)
        style.configure("TSeparator",     background="#1e2d3d")

        # Labels — backgrounds must match their parent frame
        style.configure("HeroTitle.TLabel",      background=BG,    foreground=FG,     font=("Segoe UI", 17, "bold"))
        style.configure("HeroText.TLabel",        background=BG,    foreground=MUTED,  font=("Segoe UI", 9))
        style.configure("SectionTitle.TLabel",    background=CARD,  foreground=FG,     font=("Segoe UI Semibold", 12))
        style.configure("CardTitle.TLabel",       background=CARD,  foreground=FG,     font=("Segoe UI Semibold", 12))
        style.configure("Muted.TLabel",           background=CARD,  foreground=MUTED,  font=("Segoe UI", 8))
        style.configure("MutedSide.TLabel",       background=CARD,  foreground=MUTED,  font=("Segoe UI", 9))
        style.configure("SmallAccent.TLabel",     background=CARD,  foreground=GOLD,   font=("Segoe UI Semibold", 8))
        style.configure("SmallAccentSide.TLabel", background=PANEL, foreground=GOLD,   font=("Segoe UI Semibold", 8))
        style.configure("InfoValue.TLabel",       background=PANEL, foreground=FG,     font=("Segoe UI", 9))
        style.configure("InlineInfo.TLabel",      background=PANEL, foreground=FG,     font=("Segoe UI", 8))
        style.configure("StatTitle.TLabel",       background=PANEL, foreground=GOLD,   font=("Segoe UI Semibold", 9))
        style.configure("StatValue.TLabel",       background=PANEL, foreground=FG,     font=("Segoe UI Semibold", 12))

        # Buttons
        _btn_opts = dict(borderwidth=0, relief="flat", padding=(14, 9))
        # Primary — gold CTA (RSI style)
        style.configure("Primary.TButton",        background=GOLD,    foreground="#080c10", font=("Segoe UI Semibold", 9), **_btn_opts)
        style.map(       "Primary.TButton",        background=[("active", GOLD_H), ("disabled", BTN)])

        style.configure("Secondary.TButton",      background=BTN,     foreground=FG,        font=("Segoe UI Semibold", 9), **_btn_opts)
        style.map(       "Secondary.TButton",      background=[("active", BTN_H)])

        style.configure("AppUpdateIdle.TButton",  background=BTN,     foreground=FG,        font=("Segoe UI Semibold", 9), **_btn_opts)
        style.map(       "AppUpdateIdle.TButton",  background=[("active", BTN_H)])

        style.configure("AppUpdateAlertA.TButton",background=WARN_A,  foreground="#fef3c7", font=("Segoe UI Semibold", 9), **_btn_opts)
        style.map(       "AppUpdateAlertA.TButton",background=[("active", WARN_B)])

        style.configure("AppUpdateAlertB.TButton",background=WARN_B,  foreground="#fff7ed", font=("Segoe UI Semibold", 9), **_btn_opts)
        style.map(       "AppUpdateAlertB.TButton",background=[("active", WARN_A)])

        style.configure("ToggleOff.TButton",      background=BTN,     foreground=MUTED,     font=("Segoe UI Semibold", 9), **_btn_opts)
        style.map(       "ToggleOff.TButton",      background=[("active", BTN_H)])

        style.configure("ToggleOn.TButton",       background=SUCCESS,  foreground="#d1fae5", font=("Segoe UI Semibold", 9), **_btn_opts)
        style.map(       "ToggleOn.TButton",       background=[("active", SUCC_H)])

        style.configure("ViewActive.TButton",     background=PANEL,   foreground=GOLD,      font=("Segoe UI Semibold", 9), borderwidth=0, padding=(18, 8))
        style.map(       "ViewActive.TButton",     background=[("active", BTN_H)])

        style.configure("ViewIdle.TButton",       background=BG,      foreground=MUTED,     font=("Segoe UI Semibold", 9), borderwidth=0, padding=(18, 8))
        style.map(       "ViewIdle.TButton",       background=[("active", CARD)])

        # Activity tab with update badge — gold dot visible even when tab is idle
        style.configure("ViewIdleBadge.TButton",  background=BG,      foreground=GOLD,      font=("Segoe UI Semibold", 9), borderwidth=0, padding=(18, 8))
        style.map(       "ViewIdleBadge.TButton",  background=[("active", CARD)])

        # Inputs
        style.configure("TEntry",   fieldbackground=INPUT_BG, foreground=FG, insertcolor=FG,
                         bordercolor="#1e2d3d", lightcolor="#1e2d3d", darkcolor="#1e2d3d", padding=(8, 6))
        style.configure("TSpinbox", fieldbackground=INPUT_BG, foreground=FG,
                         arrowsize=14, padding=(6, 4))

    def _load_icon_image(self) -> Image.Image | None:
        png_path = RESOURCE_DIR / "app_icon.png"
        icon_path = RESOURCE_DIR / "app_icon.ico"
        try:
            if png_path.exists():
                return Image.open(png_path).convert("RGBA")
            if icon_path.exists():
                return Image.open(icon_path).convert("RGBA")
        except Exception:
            return None
        return None

    def _apply_window_icon(self) -> None:
        if self.icon_image is None:
            return
        try:
            window_icon = ImageTk.PhotoImage(self.icon_image)
            self.root.iconphoto(True, window_icon)
            self.root._icon_photo = window_icon  # type: ignore[attr-defined]
        except Exception:
            try:
                self.root.iconbitmap(default=str(RESOURCE_DIR / "app_icon.ico"))
            except Exception:
                pass

    def _show_window(self, *_args) -> None:
        self.root.after(0, self._restore_window)

    def _restore_window(self) -> None:
        self.is_restoring = True
        self._stop_tray_icon()
        ensure_taskbar_window(self.root)
        self.root.deiconify()
        self.root.state("normal")
        self.root.lift()
        self.root.focus_force()
        self.root.after(50, self._force_foreground)
        self.root.after(300, self._finish_restore)

    def _force_foreground(self) -> None:
        try:
            self.root.attributes("-topmost", True)
            self.root.after(50, lambda: self.root.attributes("-topmost", False))
        except Exception:
            pass

    def _finish_restore(self) -> None:
        self.is_restoring = False

    def _quit_from_tray(self, *_args) -> None:
        self.root.after(0, self._quit_app)

    def _open_current_app_folder(self) -> None:
        target = Path(self.current_app_path_var.get()).resolve()
        folder = target.parent if target.suffix else target
        try:
            if sys.platform == "win32":
                os.startfile(str(folder))  # type: ignore[attr-defined]
            else:
                webbrowser.open(folder.as_uri())
        except Exception as exc:
            self.append_log(f"Could not open app folder: {exc}")

    def _open_log_folder(self) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(str(DATA_DIR))  # type: ignore[attr-defined]
            else:
                webbrowser.open(DATA_DIR.as_uri())
        except Exception as exc:
            self.append_log(f"Could not open log folder: {exc}")

    def _open_path(self, path: Path) -> None:
        try:
            target = path if path.is_dir() else path.parent
            if sys.platform == "win32":
                os.startfile(str(target))  # type: ignore[attr-defined]
            else:
                webbrowser.open(target.as_uri())
        except Exception as exc:
            self.append_log(f"Could not open path: {exc}")

    def _parse_update_result(self, message: str) -> dict[str, str]:
        summary = "StarStrings check completed."
        release_name = self.state.tracked_release_name or "Current release"
        user_cfg_action = "Verified USER.cfg settings."
        backup_path = ""

        release_match = re.search(r"'([^']+)'", message)
        if release_match:
            release_name = release_match.group(1)

        if "Manual update completed" in message or "Updated to" in message:
            summary = "StarStrings was updated successfully."
        elif "No GitHub changes detected" in message:
            summary = "StarStrings is already up to date."
        elif "Tracking initialized" in message:
            summary = "Tracking was initialized successfully."

        if "Merged USER.cfg" in message:
            user_cfg_action = "Merged your existing USER.cfg settings."
        elif "Copied USER.cfg" in message:
            user_cfg_action = "Copied the packaged USER.cfg into LIVE."

        backup_match = re.search(r"Backup:\s*'([^']+)'", message)
        if backup_match:
            backup_path = backup_match.group(1)

        return {
            "summary": summary,
            "release_name": release_name,
            "user_cfg_action": user_cfg_action,
            "backup_path": backup_path,
        }

    def _show_run_result_dialog(self, message: str, is_error: bool) -> None:
        if is_error:
            dialog = tk.Toplevel(self.root)
            dialog.title(APP_NAME)
            dialog.configure(bg="#0d1219")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.minsize(460, 200)
            tk.Frame(dialog, bg="#7a2a08", height=2).pack(fill="x", side="top")

            content = tk.Frame(dialog, bg="#0d1219")
            content.pack(fill="both", expand=True, padx=24, pady=20)
            tk.Label(content, text="StarStrings Update Failed",
                     fg="#e8edf2", bg="#0d1219", font=("Segoe UI Semibold", 12)).pack(anchor="w")
            tk.Label(content, text="The update did not finish. You can review the details below and try again.",
                     fg="#6e8096", bg="#0d1219", font=("Segoe UI", 9), wraplength=400, justify="left").pack(anchor="w", pady=(6, 12))
            tk.Label(content, text=message,
                     fg="#e8edf2", bg="#080c10", font=("Segoe UI", 9), wraplength=400, justify="left",
                     padx=12, pady=10).pack(fill="x")
            btn_row = tk.Frame(dialog, bg="#0d1219")
            btn_row.pack(fill="x", padx=24, pady=(0, 20))
            tk.Button(btn_row, text="Close", command=dialog.destroy,
                      bg="#162030", fg="#e8edf2", activebackground="#1e2d3d", activeforeground="#e8edf2",
                      relief="flat", padx=16, pady=8, font=("Segoe UI Semibold", 9),
                      cursor="hand2", bd=0).pack(side="right")
            dialog.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
            dialog.geometry(f"+{x}+{y}")
            apply_dark_titlebar(dialog)
            dialog.grab_set()
            return

        info = self._parse_update_result(message)
        dialog = tk.Toplevel(self.root)
        dialog.title(APP_NAME)
        dialog.configure(bg="#0d1219")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.minsize(520, 265)
        tk.Frame(dialog, bg="#c09040", height=2).pack(fill="x", side="top")

        content = tk.Frame(dialog, bg="#0d1219")
        content.pack(fill="both", expand=True, padx=24, pady=20)
        tk.Label(content, text=info["summary"],
                 fg="#e8edf2", bg="#0d1219", font=("Segoe UI Semibold", 12)).pack(anchor="w")
        tk.Label(content, text="Your StarStrings files were checked and the result is ready below.",
                 fg="#6e8096", bg="#0d1219", font=("Segoe UI", 9), wraplength=460, justify="left").pack(anchor="w", pady=(6, 14))

        details = tk.Frame(content, bg="#111922")
        details.pack(fill="x")
        rows = [
            ("Release", info["release_name"]),
            ("LIVE Folder", Path(self.settings.live_path).name or self.settings.live_path),
            ("USER.cfg", info["user_cfg_action"]),
        ]
        for label_text, value_text in rows:
            row = tk.Frame(details, bg="#111922")
            row.pack(fill="x", padx=14, pady=8)
            tk.Label(row, text=f"{label_text}:", fg="#c09040", bg="#111922",
                     font=("Segoe UI Semibold", 9)).pack(side="left")
            tk.Label(row, text=f" {value_text}", fg="#e8edf2", bg="#111922",
                     font=("Segoe UI", 9), wraplength=320, justify="left").pack(side="left", fill="x", expand=True)

        footer = tk.Frame(content, bg="#0d1219")
        footer.pack(fill="x", pady=(14, 0))

        countdown_var = tk.StringVar(value="This message closes automatically in 30 seconds.")
        tk.Label(footer, textvariable=countdown_var,
                 fg="#6e8096", bg="#0d1219", font=("Segoe UI", 8)).pack(anchor="w")

        link_row = tk.Frame(footer, bg="#0d1219")
        link_row.pack(fill="x", pady=(8, 0))
        if info["backup_path"]:
            tk.Label(link_row, text="Backup:", fg="#6e8096", bg="#0d1219", font=("Segoe UI", 9)).pack(side="left")
            backup_link = tk.Label(link_row, text=" Backup Location", fg="#c09040", bg="#0d1219",
                                   cursor="hand2", font=("Segoe UI Semibold", 9, "underline"))
            backup_link.pack(side="left")
            backup_link.bind("<Button-1>", lambda _event, p=Path(info["backup_path"]): self._open_path(p))

        action_row = tk.Frame(content, bg="#0d1219")
        action_row.pack(fill="x", pady=(14, 0))
        tk.Button(action_row, text="Close", command=dialog.destroy,
                  bg="#162030", fg="#e8edf2", activebackground="#1e2d3d", activeforeground="#e8edf2",
                  relief="flat", padx=16, pady=8, font=("Segoe UI Semibold", 9),
                  cursor="hand2", bd=0).pack(side="right")

        def tick(seconds_left: int = 30) -> None:
            if not dialog.winfo_exists():
                return
            if seconds_left <= 0:
                dialog.destroy()
                return
            countdown_var.set(f"This message closes automatically in {seconds_left} second{'s' if seconds_left != 1 else ''}.")
            dialog.after(1000, lambda: tick(seconds_left - 1))

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        apply_dark_titlebar(dialog)
        dialog.grab_set()
        tick()

    def _open_restore_backup_dialog(self) -> None:
        backups = list_backups()
        if not backups:
            messagebox.showinfo(APP_NAME, "No backups found.\n\nBackups are created automatically each time StarStrings is updated.", parent=self.root)
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Restore Backup")
        dialog.configure(bg="#0d1219")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.minsize(620, 330)

        tk.Frame(dialog, bg="#c09040", height=2).pack(fill="x", side="top")

        content = tk.Frame(dialog, bg="#0d1219")
        content.pack(fill="both", padx=24, pady=20)

        tk.Label(content, text="Restore a StarStrings Backup",
                 fg="#e8edf2", bg="#0d1219", font=("Segoe UI Semibold", 11)).pack(anchor="w")
        tk.Label(content, text="Select a backup to restore to your Star Citizen LIVE folder.",
                 fg="#6e8096", bg="#0d1219", font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 12))

        # Listbox with scrollbar
        list_frame = tk.Frame(content, bg="#0d1219")
        list_frame.pack(fill="both")
        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        listbox = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set, selectmode="single",
            bg="#080c10", fg="#e8edf2", selectbackground="#c09040", selectforeground="#080c10",
            relief="flat", font=("Consolas", 10), height=8, width=60, activestyle="none",
        )
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)

        def _fmt(p: Path) -> str:
            n = p.name  # YYYYMMDD-HHMMSS
            release_name = read_backup_release_name(p)
            try:
                dt = datetime.strptime(n, "%Y%m%d-%H%M%S")
                return f"{dt.strftime('%b %d, %Y  %I:%M:%S %p')}  —  {release_name}"
            except ValueError:
                return f"{n}  —  {release_name}"

        for b in backups:
            listbox.insert("end", _fmt(b))
        listbox.selection_set(0)

        btn_row = tk.Frame(dialog, bg="#0d1219")
        btn_row.pack(fill="x", padx=24, pady=(0, 20))

        def do_restore() -> None:
            sel = listbox.curselection()
            if not sel:
                return
            chosen = backups[sel[0]]
            live = Path(self.settings.live_path)
            if not live.is_dir():
                messagebox.showerror(APP_NAME, f"LIVE folder not found:\n{live}", parent=dialog)
                return
            confirm = messagebox.askyesno(
                "Confirm Restore",
                f"Restore backup from:\n{_fmt(chosen)}\n\nThis will overwrite files in your LIVE folder. Continue?",
                parent=dialog,
            )
            if not confirm:
                return
            dialog.destroy()
            def worker() -> None:
                try:
                    msg = restore_backup(chosen, live)
                    log(msg)
                    self.root.after(0, lambda: self.append_log(msg))
                except Exception as exc:
                    err = str(exc)
                    log(f"Restore failed. {err}")
                    self.root.after(0, lambda: self.append_log(f"Restore failed: {err}"))
            threading.Thread(target=worker, daemon=True).start()

        tk.Button(btn_row, text="Restore Selected", command=do_restore,
                  bg="#c09040", fg="#080c10", activebackground="#d4a84e", activeforeground="#080c10",
                  relief="flat", padx=14, pady=8, font=("Segoe UI Semibold", 9),
                  cursor="hand2", bd=0).pack(side="left", padx=(0, 10))
        tk.Button(btn_row, text="Cancel", command=dialog.destroy,
                  bg="#162030", fg="#e8edf2", activebackground="#1e2d3d", activeforeground="#e8edf2",
                  relief="flat", padx=14, pady=8, font=("Segoe UI Semibold", 9),
                  cursor="hand2", bd=0).pack(side="left")

        dialog.update_idletasks()
        width = max(dialog.winfo_width(), 620)
        height = max(dialog.winfo_height(), 330)
        x = self.root.winfo_x() + (self.root.winfo_width() - width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - height) // 2
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        apply_dark_titlebar(dialog)
        dialog.grab_set()

    def _quit_app(self) -> None:
        self.is_quitting = True
        self._remove_traces()
        self._cancel_scheduled_jobs()
        self._stop_tray_icon()
        self.root.destroy()

    def _quit_for_update(self) -> None:
        append_update_trace("App is shutting down for staged update.")
        self.is_quitting = True
        self._remove_traces()
        self._cancel_scheduled_jobs()
        self._stop_tray_icon()
        try:
            self.root.destroy()
        except Exception:
            pass
        os._exit(0)

    def _cancel_scheduled_jobs(self) -> None:
        for job_name in ("app_update_pulse_job", "app_update_check_job", "settings_save_job", "blueprint_auto_scan_job", "_search_debounce_job"):
            job = getattr(self, job_name, None)
            if job is not None:
                try:
                    self.root.after_cancel(job)
                except Exception:
                    pass
                setattr(self, job_name, None)

    def _start_tray_icon(self) -> None:
        if self.tray_icon is not None:
            return
        tray_image = self.icon_image.copy() if self.icon_image else Image.new("RGBA", (64, 64), "#143044")
        tray_image = tray_image.resize((64, 64), Image.LANCZOS)
        menu = pystray.Menu(
            pystray.MenuItem("Open", self._show_window, default=True),
            pystray.MenuItem("Exit", self._quit_from_tray),
        )
        self.tray_icon = pystray.Icon("CitizenStarStringHelper", tray_image, APP_NAME, menu)
        self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        self.tray_thread.start()

    def _stop_tray_icon(self) -> None:
        if self.tray_icon is not None:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
            self.tray_icon = None
        if self.tray_thread is not None:
            self.tray_thread.join(timeout=2)
            self.tray_thread = None

    def _minimize_to_tray(self) -> None:
        if self.is_quitting:
            return
        self._start_tray_icon()
        remove_appwindow_style(self.root)
        self.root.withdraw()

    def _handle_map(self, _event) -> None:
        if self.is_restoring:
            return

    def _handle_unmap(self, _event) -> None:
        if self.is_quitting or self.is_restoring:
            return
        state = self.root.state()
        if state == "iconic":
            self.root.after(10, self._handle_minimize)

    def _handle_configure(self, _event) -> None:
        if self.is_quitting or self.is_restoring:
            return
        if self.root.state() == "iconic":
            self.root.after(10, self._handle_minimize)

    def _handle_minimize(self) -> None:
        if self.is_quitting or self.is_restoring:
            return
        self._minimize_to_tray()

    def _handle_close(self) -> None:
        self._ask_minimize_or_exit()

    def _ask_minimize_or_exit(self) -> None:
        """Show a dialog asking whether to minimize to tray or exit completely."""
        result: dict[str, str] = {}

        dialog = tk.Toplevel(self.root)
        dialog.title(APP_NAME)
        dialog.configure(bg="#0d1219")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.minsize(430, 170)
        dialog.columnconfigure(0, weight=1)

        tk.Frame(dialog, bg="#c09040", height=2).grid(row=0, column=0, sticky="ew")

        content = tk.Frame(dialog, bg="#0d1219")
        content.grid(row=1, column=0, sticky="nsew", padx=28, pady=(22, 14))
        content.columnconfigure(0, weight=1)
        tk.Label(content, text="Close Citizen StarString Helper",
                 fg="#e8edf2", bg="#0d1219", font=("Segoe UI Semibold", 11)).grid(row=0, column=0, sticky="w")
        tk.Label(content, text="Continue running in the system tray, or exit completely?",
                 fg="#6e8096", bg="#0d1219", font=("Segoe UI", 9), wraplength=360, justify="left").grid(row=1, column=0, sticky="w", pady=(5, 0))

        btn_row = tk.Frame(dialog, bg="#0d1219")
        btn_row.grid(row=2, column=0, sticky="ew", padx=28, pady=(0, 22))
        btn_row.columnconfigure(0, weight=1)
        btn_row.columnconfigure(1, weight=1)

        def do_minimize() -> None:
            result["action"] = "minimize"
            dialog.destroy()

        def do_exit() -> None:
            result["action"] = "exit"
            dialog.destroy()

        tk.Button(btn_row, text="Minimize to Tray", command=do_minimize,
                  bg="#162030", fg="#e8edf2", activebackground="#1e2d3d", activeforeground="#e8edf2",
                  relief="flat", padx=14, pady=8, font=("Segoe UI Semibold", 9),
                  cursor="hand2", bd=0).grid(row=0, column=0, sticky="ew", padx=(0, 10))
        tk.Button(btn_row, text="Exit", command=do_exit,
                  bg="#c09040", fg="#080c10", activebackground="#d4a84e", activeforeground="#080c10",
                  relief="flat", padx=14, pady=8, font=("Segoe UI Semibold", 9),
                  cursor="hand2", bd=0).grid(row=0, column=1, sticky="ew")

        dialog.update_idletasks()
        width = max(dialog.winfo_width(), 430)
        height = max(dialog.winfo_height(), 170)
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        apply_dark_titlebar(dialog)
        dialog.grab_set()
        dialog.focus_force()

        dialog.wait_window()

        action = result.get("action")
        if action == "minimize":
            self._minimize_to_tray()
        elif action == "exit":
            self._quit_app()

    def _build_ui(self) -> None:
        # Thin accent stripe at the very top of the window
        accent_bar = ttk.Frame(self.root, style="Accent.TFrame", height=3)
        accent_bar.pack(fill="x", side="top")

        root_frame = ttk.Frame(self.root, style="Root.TFrame", padding=18)
        root_frame.pack(fill="both", expand=True)
        root_frame.columnconfigure(0, weight=1)
        root_frame.rowconfigure(2, weight=1)

        topbar = ttk.Frame(root_frame, style="Hero.TFrame", padding=(6, 6, 6, 10))
        topbar.grid(row=0, column=0, sticky="ew")
        topbar.columnconfigure(0, weight=1)
        topbar.columnconfigure(1, weight=0)
        topbar.rowconfigure(1, weight=0)

        title_col = ttk.Frame(topbar, style="Hero.TFrame")
        title_col.grid(row=0, column=0, sticky="ew")
        title_col.columnconfigure(0, weight=1)
        ttk.Label(title_col, text="CITIZEN STARSTRING HELPER", style="HeroTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(title_col, text=f"v{APP_VERSION}  ·  Star Citizen String Manager", style="HeroText.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 0))

        if self.icon_image is not None:
            try:
                header_image = self.icon_image.copy()
                header_image.thumbnail((132, 132), Image.LANCZOS)
                self.header_art_photo = ImageTk.PhotoImage(header_image)
                art_wrap = ttk.Frame(topbar, style="Hero.TFrame")
                art_wrap.grid(row=0, column=1, rowspan=2, sticky="ne", padx=(18, 0))
                tk.Label(
                    art_wrap,
                    image=self.header_art_photo,
                    bg="#080c10",
                    bd=0,
                    highlightthickness=0,
                ).pack(anchor="ne")
            except Exception:
                self.header_art_photo = None

        view_row = ttk.Frame(topbar, style="Hero.TFrame")
        view_row.grid(row=1, column=0, sticky="w", pady=(12, 0))
        view_row.columnconfigure(0, weight=0)
        view_row.columnconfigure(1, weight=0)
        view_row.columnconfigure(2, weight=0)
        self.setup_view_button = ttk.Button(view_row, text="SETUP", style="ViewActive.TButton", command=lambda: self._show_view("setup"))
        self.setup_view_button.grid(row=0, column=0, padx=(0, 8))
        self.activity_view_button = ttk.Button(view_row, text="ACTIVITY", style="ViewIdle.TButton", command=lambda: self._show_view("activity"))
        self.activity_view_button.grid(row=0, column=1, padx=(0, 8))
        self.blueprints_view_button = ttk.Button(view_row, text="BLUEPRINTS", style="ViewIdle.TButton", command=lambda: self._show_view("blueprints"))
        self.blueprints_view_button.grid(row=0, column=2)

        # Gold accent line between header and content (RSI nav-bar style)
        tk.Frame(root_frame, bg="#c09040", height=1).grid(row=1, column=0, sticky="ew", pady=(0, 14))

        body = ttk.Frame(root_frame, style="Root.TFrame")
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        self.content_host = ttk.Frame(body, style="Root.TFrame")
        self.content_host.grid(row=0, column=0, sticky="nsew")
        self.content_host.columnconfigure(0, weight=1)
        self.content_host.rowconfigure(0, weight=1)

        footer = ttk.Frame(root_frame, style="Card.TFrame", padding=(0, 10, 0, 0))
        footer.grid(row=3, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        setup_view = ttk.Frame(self.content_host, style="Root.TFrame")
        setup_view.grid(row=0, column=0, sticky="nsew")
        setup_view.columnconfigure(0, weight=1)
        setup_view.rowconfigure(0, weight=3)
        setup_view.rowconfigure(1, weight=2)

        main = ttk.Frame(setup_view, style="Card.TFrame", padding=16)
        main.grid(row=0, column=0, sticky="nsew", pady=(0, 12))
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=0)

        status = ttk.Frame(setup_view, style="SideCard.TFrame", padding=12)
        status.grid(row=1, column=0, sticky="nsew")
        status.columnconfigure(0, weight=1)

        ttk.Label(main, text="Deployment Settings", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(main, text="Update the install path, repository, and StarStrings schedule.", style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 10))

        form = ttk.Frame(main, style="Card.TFrame")
        form.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 14))
        form.columnconfigure(0, weight=0)
        form.columnconfigure(1, weight=1)
        form.columnconfigure(2, weight=0)

        ttk.Label(form, text="LIVE FOLDER", style="SmallAccent.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        path_row = ttk.Frame(form, style="Card.TFrame")
        path_row.grid(row=0, column=1, columnspan=2, sticky="ew", pady=(0, 10))
        path_row.columnconfigure(0, weight=1)
        self.path_entry = ttk.Entry(path_row, textvariable=self.live_path_var, font=("Segoe UI", 10))
        self.path_entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(path_row, text="Browse", style="Secondary.TButton", command=self.choose_folder).grid(row=0, column=1, padx=(8, 0))

        ttk.Label(form, text="REPOSITORY", style="SmallAccent.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        repo_row = ttk.Frame(form, style="Card.TFrame")
        repo_row.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(0, 10))
        repo_row.columnconfigure(0, weight=1)
        self.repo_entry = ttk.Entry(repo_row, textvariable=self.github_repo_var, font=("Segoe UI", 10))
        self.repo_entry.grid(row=0, column=0, sticky="ew")

        ttk.Label(form, text="UPDATE INTERVAL", style="SmallAccent.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 10))
        interval_row = ttk.Frame(form, style="Card.TFrame")
        interval_row.grid(row=2, column=1, sticky="w")
        self.interval_spin = ttk.Spinbox(interval_row, from_=1, to=24, textvariable=self.interval_var, width=5, font=("Segoe UI", 10))
        self.interval_spin.grid(row=0, column=0, sticky="w")
        ttk.Label(interval_row, text="hours", style="Muted.TLabel").grid(row=0, column=1, sticky="w", padx=(8, 0))

        button_row = ttk.Frame(main, style="Card.TFrame")
        button_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        button_row.columnconfigure(0, weight=1)
        button_row.columnconfigure(1, weight=1)
        button_row.rowconfigure(1, weight=0)
        self.run_button = ttk.Button(button_row, text="Run StarStrings Now", style="Primary.TButton", command=self.run_manual_update)
        self.run_button.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self.toggle_button = ttk.Button(button_row, text="StarStrings Auto Update: Off", command=self.toggle_schedule)
        self.toggle_button.grid(row=0, column=1, sticky="ew")
        restore_button = ttk.Button(button_row, text="Restore Previous Backup", style="Secondary.TButton", command=self._open_restore_backup_dialog)
        restore_button.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        footer_bar = tk.Frame(footer, bg="#080c10")
        footer_bar.grid(row=0, column=0, sticky="ew")
        footer_bar.grid_columnconfigure(0, weight=1)
        footer_bar.grid_columnconfigure(1, weight=0)
        footer_bar.grid_columnconfigure(2, weight=1)

        referral_wrap = tk.Frame(footer_bar, bg="#080c10")
        referral_wrap.grid(row=0, column=1, pady=(0, 8))
        tk.Label(
            referral_wrap,
            text="New to Star Citizen? Enlist with referral code",
            fg="#6e8096", bg="#080c10", font=("Segoe UI", 9),
        ).pack(side="left")
        referral_code = tk.Label(
            referral_wrap,
            text="  STAR-J66D-SPVW",
            fg="#c09040", bg="#080c10", cursor="hand2",
            font=("Segoe UI", 9, "bold"),
        )
        referral_code.pack(side="left")
        referral_code.bind("<Button-1>", lambda _event: webbrowser.open(REFERRAL_URL))

        exit_wrap = tk.Frame(footer_bar, bg="#080c10")
        exit_wrap.grid(row=1, column=2, sticky="e")
        tk.Button(
            exit_wrap,
            text="Exit",
            command=self._ask_minimize_or_exit,
            bg="#162030", fg="#6e8096",
            activebackground="#1e2d3d", activeforeground="#e8edf2",
            relief="flat", padx=12, pady=4,
            font=("Segoe UI", 9), cursor="hand2", bd=0,
        ).pack(side="right")

        ttk.Label(status, text="Status", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))

        self._make_info_card(status, 1, "REPOSITORY", self.repo_display_var)
        self._make_info_card(status, 2, "TRACKED RELEASE", self.release_var)
        self._make_info_card(status, 3, "STARSTRINGS AUTO", self.auto_state_var)
        self._make_info_card(status, 4, "LAST CHECKED", self.last_checked_var)
        self._make_info_card(status, 5, "LAST UPDATED", self.last_updated_var)
        self._make_info_card(status, 6, "TARGET INSTALL", self.live_path_var)

        activity = ttk.Frame(self.content_host, style="Root.TFrame")
        activity.grid(row=0, column=0, sticky="nsew")
        activity.columnconfigure(0, weight=1)
        activity.rowconfigure(0, weight=1)

        feed_card = ttk.Frame(activity, style="Card.TFrame", padding=20)
        feed_card.grid(row=0, column=0, sticky="nsew")
        feed_card.columnconfigure(0, weight=1)
        feed_card.rowconfigure(5, weight=1)

        ttk.Label(feed_card, text="Operations Feed", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        activity_button_row = ttk.Frame(feed_card, style="Card.TFrame")
        activity_button_row.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        activity_button_row.columnconfigure(0, weight=1)
        activity_button_row.columnconfigure(1, weight=1)
        activity_button_row.grid_columnconfigure(0, uniform="activity")
        activity_button_row.grid_columnconfigure(1, uniform="activity")
        self.app_update_button = ttk.Button(
            activity_button_row,
            textvariable=self.app_update_button_var,
            style="AppUpdateIdle.TButton",
            command=self._on_app_update_button_click,
        )
        self.app_update_button.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=(0, 8))
        ttk.Button(activity_button_row, text="Open App Folder", style="Secondary.TButton", command=self._open_current_app_folder).grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=(0, 8))
        ttk.Button(activity_button_row, text="Open Log Folder", style="Secondary.TButton", command=self._open_log_folder).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(activity_button_row, text="Clear Log", style="Secondary.TButton", command=self._clear_log).grid(row=1, column=1, sticky="ew", padx=(8, 0))

        ops_meta = ttk.Label(feed_card, textvariable=self.operations_meta_var, style="Muted.TLabel", wraplength=860, justify="left")
        ops_meta.grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.meta_labels.append(ops_meta)
        app_path_label = ttk.Label(feed_card, textvariable=self.current_app_path_var, style="Muted.TLabel", wraplength=860, justify="left")
        app_path_label.grid(row=3, column=0, sticky="w", pady=(0, 2))
        self.meta_labels.append(app_path_label)
        schedule_label = ttk.Label(feed_card, textvariable=self.schedule_var, style="Muted.TLabel", wraplength=860, justify="left")
        schedule_label.grid(row=4, column=0, sticky="w", pady=(0, 10))
        self.meta_labels.append(schedule_label)

        self.log_text = ttk.Treeview(feed_card, show="", height=8)
        self.log_text.grid_remove()

        self.log_widget = tk.Text(feed_card, height=10, wrap="word", bg="#080c10", fg="#e8edf2", insertbackground="#e8edf2", relief="flat", font=("Consolas", 9), padx=10, pady=8)
        self.log_widget.grid(row=5, column=0, sticky="nsew")
        self.log_widget.tag_configure("ts",      foreground="#5a4418")  # dim gold timestamp
        self.log_widget.tag_configure("info",    foreground="#e8edf2")  # normal
        self.log_widget.tag_configure("success", foreground="#4ade80")  # green
        self.log_widget.tag_configure("error",   foreground="#f87171")  # red
        self.log_widget.tag_configure("muted",   foreground="#6e8096")  # grey
        self.log_widget.configure(state="disabled")

        blueprints = ttk.Frame(self.content_host, style="Root.TFrame")
        blueprints.grid(row=0, column=0, sticky="nsew")
        blueprints.columnconfigure(0, weight=1)
        blueprints.rowconfigure(2, weight=1)
        blueprints.rowconfigure(3, weight=1)

        bp_toolbar = ttk.Frame(blueprints, style="Card.TFrame", padding=16)
        bp_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        bp_toolbar.columnconfigure(1, weight=1)
        bp_toolbar.columnconfigure(3, weight=0)
        bp_toolbar.columnconfigure(5, weight=0)
        bp_toolbar.columnconfigure(7, weight=0)
        bp_toolbar.columnconfigure(9, weight=0)

        ttk.Label(bp_toolbar, text="Blueprints", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=10, sticky="w")
        ttk.Label(bp_toolbar, text="Search learned and mission-linked blueprints from your installed StarStrings data.", style="Muted.TLabel").grid(row=1, column=0, columnspan=10, sticky="w", pady=(4, 12))
        ttk.Label(bp_toolbar, text="SEARCH", style="SmallAccent.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 10))
        search_entry = ttk.Entry(bp_toolbar, textvariable=self.blueprint_search_var, font=("Segoe UI", 10))
        search_entry.grid(row=2, column=1, sticky="ew", padx=(0, 12))
        ttk.Label(bp_toolbar, text="FILTER", style="SmallAccent.TLabel").grid(row=2, column=2, sticky="w", padx=(0, 10))
        self.blueprint_filter_combo = ttk.Combobox(
            bp_toolbar,
            state="readonly",
            values=("All", "Learned", "Missing"),
            textvariable=self.blueprint_filter_var,
            width=18,
            font=("Segoe UI", 9),
        )
        self.blueprint_filter_combo.grid(row=2, column=3, sticky="ew", padx=(0, 12))
        self.blueprint_filter_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_blueprint_list())
        ttk.Label(bp_toolbar, text="TYPE", style="SmallAccent.TLabel").grid(row=2, column=4, sticky="w", padx=(0, 10))
        self.blueprint_type_filter_combo = ttk.Combobox(
            bp_toolbar,
            state="readonly",
            values=("All Types", "Armor", "Weapon", "Ammo", "Clothing", "Med", "Tool", "Attachment", "Component", "Unknown"),
            textvariable=self.blueprint_type_filter_var,
            width=14,
            font=("Segoe UI", 9),
        )
        self.blueprint_type_filter_combo.grid(row=2, column=5, sticky="ew", padx=(0, 12))
        self.blueprint_type_filter_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_blueprint_list())
        ttk.Label(bp_toolbar, text="MODE", style="SmallAccent.TLabel").grid(row=2, column=6, sticky="w", padx=(0, 10))
        self.blueprint_search_mode_combo = ttk.Combobox(
            bp_toolbar,
            state="readonly",
            values=("Strict", "Fuzzy"),
            textvariable=self.blueprint_search_mode_var,
            width=10,
            font=("Segoe UI", 9),
        )
        self.blueprint_search_mode_combo.grid(row=2, column=7, sticky="ew", padx=(0, 12))
        self.blueprint_search_mode_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_blueprint_list())
        self.blueprint_scan_button = ttk.Button(bp_toolbar, text="Scan Blueprints", style="Primary.TButton", command=self.scan_blueprints)
        self.blueprint_scan_button.grid(row=2, column=9, sticky="ew")

        ttk.Label(bp_toolbar, textvariable=self.blueprint_status_var, style="Muted.TLabel").grid(row=3, column=0, columnspan=10, sticky="w", pady=(12, 0))

        bp_summary = ttk.Frame(blueprints, style="Card.TFrame", padding=12)
        bp_summary.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        bp_summary.columnconfigure(0, weight=1)
        ttk.Label(bp_summary, textvariable=self.blueprint_summary_var, style="MutedSide.TLabel").grid(row=0, column=0, sticky="w")

        bp_results = ttk.Frame(blueprints, style="Card.TFrame", padding=14)
        bp_results.grid(row=2, column=0, sticky="nsew", pady=(0, 12))
        bp_results.columnconfigure(0, weight=1)
        bp_results.rowconfigure(1, weight=1)
        ttk.Label(bp_results, text="Blueprint Matches", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))
        self.blueprint_tree = ttk.Treeview(bp_results, columns=("wiki", "blueprint", "type", "status", "materials"), show="headings", height=12)
        self.blueprint_tree.heading("wiki", text="SC Wiki")
        self.blueprint_tree.heading("blueprint", text="Blueprint", command=lambda: self._sort_blueprints("blueprint"))
        self.blueprint_tree.heading("type", text="Type  ▾", command=lambda: self._sort_blueprints("type"))
        self.blueprint_tree.heading("status", text="Status", command=lambda: self._sort_blueprints("status"))
        self.blueprint_tree.heading("materials", text="Materials")
        self.blueprint_tree.grid(row=1, column=0, sticky="nsew")
        self.blueprint_tree.column("wiki", width=64, anchor="center", stretch=False)
        self.blueprint_tree.column("blueprint", width=340, anchor="w")
        self.blueprint_tree.column("type", width=120, anchor="w")
        self.blueprint_tree.column("status", width=100, anchor="w")
        self.blueprint_tree.column("materials", width=72, anchor="center", stretch=False)
        self.blueprint_tree.bind("<<TreeviewSelect>>", self._on_blueprint_selected)
        self.blueprint_tree.bind("<Button-1>", self._on_blueprint_tree_click)
        self.blueprint_tree.bind("<Return>", self._open_selected_blueprint_wiki)
        self.blueprint_tree.tag_configure("learned", background="#153225", foreground="#dff8e8")
        bp_scroll = ttk.Scrollbar(bp_results, orient="vertical", command=self.blueprint_tree.yview)
        self.blueprint_tree.configure(yscrollcommand=bp_scroll.set)
        bp_scroll.grid(row=1, column=1, sticky="ns")

        bp_detail = ttk.Frame(blueprints, style="SideCard.TFrame", padding=16)
        bp_detail.grid(row=3, column=0, sticky="nsew")
        bp_detail.columnconfigure(0, weight=1)
        ttk.Label(bp_detail, textvariable=self.blueprint_detail_title_var, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(bp_detail, text="Possible Contract Sources", style="SmallAccent.TLabel").grid(row=1, column=0, sticky="w", pady=(8, 6))
        self.blueprint_contracts_text = tk.Text(bp_detail, height=12, wrap="word", bg="#080c10", fg="#e8edf2", insertbackground="#e8edf2", relief="flat", font=("Segoe UI", 9), padx=10, pady=8)
        self.blueprint_contracts_text.grid(row=2, column=0, sticky="nsew")
        self.blueprint_contracts_text.configure(state="disabled")
        bp_detail.rowconfigure(2, weight=1)

        self.setup_view = setup_view
        self.activity_view = activity
        self.blueprints_view = blueprints
        self._show_view("setup")
        self.root.bind("<Configure>", self._refresh_compact_wraps, add="+")

    def _make_info_card(self, parent: ttk.Frame, row: int, title: str, variable: StringVar) -> None:
        card = ttk.Frame(parent, style="Panel.TFrame", padding=(10, 8))
        card.grid(row=row, column=0, sticky="ew", pady=(0, 6))
        card.columnconfigure(0, weight=0)
        card.columnconfigure(1, weight=1)
        ttk.Label(card, text=f"{title}:", style="SmallAccentSide.TLabel").grid(row=0, column=0, sticky="w")
        value_label = ttk.Label(card, textvariable=variable, style="InlineInfo.TLabel", wraplength=500, justify="left")
        value_label.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.inline_info_labels.append(value_label)

    def _refresh_compact_wraps(self, _event=None) -> None:
        width = max(self.root.winfo_width(), 720)
        info_wrap = max(220, width - 290)
        meta_wrap = max(420, width - 110)
        for label in getattr(self, "inline_info_labels", []):
            try:
                label.configure(wraplength=info_wrap)
            except Exception:
                pass
        for label in getattr(self, "meta_labels", []):
            try:
                label.configure(wraplength=meta_wrap)
            except Exception:
                pass

    def _show_view(self, view_name: str) -> None:
        self.current_view = view_name
        if view_name == "setup":
            self.activity_view.grid_remove()
            self.blueprints_view.grid_remove()
            self.setup_view.grid()
            self.setup_view_button.configure(style="ViewActive.TButton")
            self.activity_view_button.configure(style="ViewIdle.TButton")
            self.blueprints_view_button.configure(style="ViewIdle.TButton")
        elif view_name == "activity":
            self.setup_view.grid_remove()
            self.blueprints_view.grid_remove()
            self.activity_view.grid()
            self.setup_view_button.configure(style="ViewIdle.TButton")
            self.activity_view_button.configure(style="ViewActive.TButton")
            self.blueprints_view_button.configure(style="ViewIdle.TButton")
        else:
            self.setup_view.grid_remove()
            self.activity_view.grid_remove()
            self.blueprints_view.grid()
            self.setup_view_button.configure(style="ViewIdle.TButton")
            self.activity_view_button.configure(style="ViewIdle.TButton")
            self.blueprints_view_button.configure(style="ViewActive.TButton")
            if not self.blueprint_records and not self.blueprint_scan_in_progress:
                self.scan_blueprints()

    def choose_folder(self) -> None:
        chosen = filedialog.askdirectory(title="Select your Star Citizen LIVE folder", mustexist=True)
        if chosen:
            self.live_path_var.set(chosen)

    def _bind_auto_save(self) -> None:
        self._trace_ids = [
            self.live_path_var.trace_add("write", self._schedule_auto_save),
            self.github_repo_var.trace_add("write", self._schedule_auto_save),
            self.interval_var.trace_add("write", self._schedule_auto_save),
        ]

    def _remove_traces(self) -> None:
        trace_ids = getattr(self, "_trace_ids", [])
        pairs = [
            (self.live_path_var,    trace_ids[0] if len(trace_ids) > 0 else None),
            (self.github_repo_var,  trace_ids[1] if len(trace_ids) > 1 else None),
            (self.interval_var,     trace_ids[2] if len(trace_ids) > 2 else None),
        ]
        for var, tid in pairs:
            if tid:
                try:
                    var.trace_remove("write", tid)
                except Exception:
                    pass
        self._trace_ids = []

    def _schedule_auto_save(self, *_args) -> None:
        if not self.settings_loaded:
            return
        if self.settings_save_job is not None:
            try:
                self.root.after_cancel(self.settings_save_job)
            except Exception:
                pass
        self.settings_save_job = self.root.after(500, self._auto_save_settings)

    def _auto_save_settings(self) -> None:
        self.settings_save_job = None
        self.save_settings(reapply_schedule=True, log_message=False)

    def save_settings(self, reapply_schedule: bool = True, log_message: bool = True) -> None:
        self.settings.live_path = self.live_path_var.get().strip()
        try:
            self.settings.interval_hours = max(1, min(24, int(self.interval_var.get())))
        except (ValueError, tk.TclError):
            self.settings.interval_hours = 6
            self.interval_var.set(6)
        try:
            self.settings.github_repo = normalize_repo(self.github_repo_var.get())
        except UpdaterError:
            if log_message:
                raise
            return
        full_repo_url = canonical_repo_url(self.settings.github_repo)
        compact_repo = compact_repo_name(self.settings.github_repo)
        self.github_repo_var.set(full_repo_url)
        self.repo_display_var.set(compact_repo)
        save_settings(self.settings)
        if reapply_schedule and self.auto_state_var.get() == "Enabled":
            register_scheduled_task(self.settings.interval_hours)
            if log_message:
                self.append_log(f"Settings saved. Automatic updates rescheduled to every {self.settings.interval_hours} hour(s).")
            self._refresh_toggle()
        else:
            if log_message:
                self.append_log("Settings saved.")

    def _log_level(self, message: str) -> str:
        m = message.lower()
        if any(k in m for k in ("fail", "error", "could not", "skipped", "denied")):
            return "error"
        if any(k in m for k in ("updated", "completed", "installed", "enabled", "up to date",
                                 "saved", "downloaded", "tracking initialized", "merged", "copied")):
            return "success"
        if any(k in m for k in ("checking", "no changes", "no github", "no release", "declined")):
            return "muted"
        return "info"

    def _clear_log(self) -> None:
        self.log_widget.configure(state="normal")
        self.log_widget.delete("1.0", "end")
        self.log_widget.configure(state="disabled")

    def _schedule_blueprint_auto_scan(self) -> None:
        if self.blueprint_auto_scan_job is not None:
            try:
                self.root.after_cancel(self.blueprint_auto_scan_job)
            except Exception:
                pass
        self.blueprint_auto_scan_job = self.root.after(BLUEPRINT_SCAN_INTERVAL_MS, self._run_blueprint_auto_scan)

    def _run_blueprint_auto_scan(self) -> None:
        self.blueprint_auto_scan_job = None
        if not self.blueprint_scan_in_progress:
            self.scan_blueprints(silent=True)
        self._schedule_blueprint_auto_scan()

    def _refresh_blueprint_freshness(self) -> None:
        self.state = load_state()
        scanned_at = format_timestamp(self.state.blueprints_last_scanned_at, "Not scanned yet")
        scanned_release = self.state.blueprints_last_scanned_release_name or "unknown StarStrings release"
        tracked_release = self.state.tracked_release_name or "No release tracked yet"
        if not self.state.blueprints_last_scanned_at:
            self.blueprint_status_var.set("Blueprints have not been scanned yet. Run a scan to map learned rewards and possible contract sources.")
            self.blueprint_summary_var.set(f"No blueprint scan has been run yet. Current tracked StarStrings release: {tracked_release}")
            return
        if self.state.tracked_release_id and self.state.blueprints_last_scanned_release_id != self.state.tracked_release_id:
            self.blueprint_status_var.set("Blueprint data is older than the currently tracked StarStrings release. A rescan will refresh contract sources.")
        else:
            self.blueprint_status_var.set("Blueprint data reflects the currently installed StarStrings files and your learned blueprint logs.")
        self.blueprint_summary_var.set(
            f"Last blueprint scan: {scanned_at}   |   Scanned release: {scanned_release}   |   Current tracked release: {tracked_release}"
        )

    def scan_blueprints(self, silent: bool = False) -> None:
        if self.blueprint_scan_in_progress:
            return
        self.blueprint_scan_in_progress = True
        self.blueprint_scan_button.configure(state="disabled")
        if not silent:
            self.blueprint_status_var.set("Scanning StarStrings data and local logs for blueprint information...")

        def worker() -> None:
            try:
                records, metadata = collect_blueprint_records(self.live_path_var.get().strip())
                self.root.after(0, lambda records=records, metadata=metadata, silent=silent: self._complete_blueprint_scan(records, metadata, silent))
            except Exception as exc:
                self.root.after(0, lambda exc=exc, silent=silent: self._fail_blueprint_scan(exc, silent))

        threading.Thread(target=worker, daemon=True).start()

    def _complete_blueprint_scan(self, records: list[BlueprintRecord], metadata: dict[str, object], silent: bool) -> None:
        self.blueprint_scan_in_progress = False
        self.blueprint_scan_button.configure(state="normal")
        self.blueprint_records = records
        self.blueprint_scan_metadata = metadata
        self.state = load_state()
        self.state.blueprints_last_scanned_at = str(metadata.get("scanned_at") or datetime.now().isoformat())
        self.state.blueprints_last_scanned_release_id = str(metadata.get("tracked_release_id") or "")
        self.state.blueprints_last_scanned_release_name = str(metadata.get("tracked_release_name") or "")
        save_state(self.state)
        total = int(metadata.get("total_count") or 0)
        learned = int(metadata.get("learned_count") or 0)
        missing = int(metadata.get("missing_count") or 0)
        available = int(metadata.get("available_count") or 0)
        scanned_at = format_timestamp(str(metadata.get("scanned_at") or ""), "just now")
        release_name = str(metadata.get("tracked_release_name") or "unknown release")
        self.blueprint_status_var.set(f"Blueprint scan complete. Parsed {total} blueprints from StarStrings data and local logs.")
        self.blueprint_summary_var.set(
            f"Learned: {learned}   |   Contract-linked: {available}   |   Missing: {missing}   |   Last scan: {scanned_at}   |   Release: {release_name}"
        )
        if silent:
            self.append_log(f"Background blueprint refresh complete. Learned {learned}, contract-linked {available}, missing {missing}.")
        else:
            self.append_log(f"Blueprint scan complete. Learned {learned}, contract-linked {available}, missing {missing}.")
        self._refresh_blueprint_list()
        self._refresh_blueprint_freshness()

    def _fail_blueprint_scan(self, exc: Exception, silent: bool) -> None:
        self.blueprint_scan_in_progress = False
        self.blueprint_scan_button.configure(state="normal")
        self.blueprint_status_var.set("Blueprint scan could not be completed.")
        self.append_log(f"Blueprint scan failed. {exc}")
        if not silent:
            messagebox.showerror(APP_NAME, str(exc), parent=self.root)

    def _refresh_blueprint_list(self) -> None:
        if not hasattr(self, "blueprint_tree"):
            return
        query = normalize_search_text(self.blueprint_search_var.get())
        filter_value = self.blueprint_filter_var.get().strip() or "All"
        type_filter = self.blueprint_type_filter_var.get().strip() or "All Types"
        search_mode = (self.blueprint_search_mode_var.get().strip() or "Strict").lower()

        self.blueprint_tree.delete(*self.blueprint_tree.get_children())
        self.filtered_blueprint_records = []
        for record in self.blueprint_records:
            if filter_value == "Learned" and not record.learned:
                continue
            if filter_value == "Missing" and (record.learned or not record.contracts):
                continue
            if type_filter != "All Types" and record.category != type_filter:
                continue
            if query:
                haystack = normalize_search_text(" ".join([record.name, record.category, *record.contracts]))
                if search_mode == "strict":
                    if query not in haystack:
                        continue
                else:
                    if not fuzzy_query_match(query, haystack):
                        continue
            self.filtered_blueprint_records.append(record)

        self._sort_filtered_blueprint_records()

        for index, record in enumerate(self.filtered_blueprint_records):
            self.blueprint_tree.insert(
                "",
                "end",
                iid=f"bp-{index}",
                values=("🔗", record.name, record.category, record.status, "⚒"),
                tags=("learned",) if record.learned else (),
            )

        if self.filtered_blueprint_records:
            first_id = self.blueprint_tree.get_children()[0]
            self.blueprint_tree.selection_set(first_id)
            self.blueprint_tree.focus(first_id)
            self._show_blueprint_details(self.filtered_blueprint_records[0])
        else:
            self._show_blueprint_details(None)

    def _sort_filtered_blueprint_records(self) -> None:
        sort_key = self.blueprint_sort_column

        def key_func(record: BlueprintRecord):
            if sort_key == "blueprint":
                return record.name.lower()
            if sort_key == "type":
                return (record.category == "Unknown", record.category.lower(), record.name.lower())
            if sort_key == "status":
                status_order = {"Learned": 0, "Missing": 1, "Unknown": 2}
                return (status_order.get(record.status, 9), record.name.lower())
            return record.name.lower()

        self.filtered_blueprint_records.sort(key=key_func, reverse=self.blueprint_sort_desc)

    def _sort_blueprints(self, column: str) -> None:
        if self.blueprint_sort_column == column:
            self.blueprint_sort_desc = not self.blueprint_sort_desc
        else:
            self.blueprint_sort_column = column
            self.blueprint_sort_desc = False
        self._refresh_blueprint_list()

    def _on_blueprint_selected(self, _event=None) -> None:
        selection = self.blueprint_tree.selection()
        if not selection:
            self._show_blueprint_details(None)
            return
        item_id = selection[0]
        try:
            index = int(item_id.split("-", 1)[1])
            record = self.filtered_blueprint_records[index]
        except Exception:
            self._show_blueprint_details(None)
            return
        self._show_blueprint_details(record)

    def _get_selected_blueprint_record(self) -> BlueprintRecord | None:
        selection = self.blueprint_tree.selection()
        if not selection:
            return None
        item_id = selection[0]
        try:
            index = int(item_id.split("-", 1)[1])
            return self.filtered_blueprint_records[index]
        except Exception:
            return None

    def _open_selected_blueprint_wiki(self, _event=None) -> None:
        record = self._get_selected_blueprint_record()
        if record is None:
            return
        def worker():
            url = resolve_blueprint_wiki_url(record.name)
            self.root.after(0, lambda: webbrowser.open(url))
        threading.Thread(target=worker, daemon=True).start()

    def _on_blueprint_tree_click(self, event) -> None:
        """Handle left-clicks in the blueprint treeview.

        - Column #1 (SC Wiki): open the wiki page for that row in a background thread.
        - Column #3 (Type): show an inline combobox overlay for changing the type override.
        Clicks on headings or outside rows are ignored so sorting commands still work.
        """
        region = self.blueprint_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.blueprint_tree.identify_column(event.x)
        item_id = self.blueprint_tree.identify_row(event.y)
        if not item_id:
            return
        try:
            index = int(item_id.split("-", 1)[1])
            record = self.filtered_blueprint_records[index]
        except Exception:
            return
        if column == "#1":  # SC Wiki column
            def worker():
                url = resolve_blueprint_wiki_url(record.name)
                self.root.after(0, lambda: webbrowser.open(url))
            threading.Thread(target=worker, daemon=True).start()
        elif column == "#3":  # Type column
            self._show_inline_type_combobox(item_id, column, record)
        elif column == "#5":  # Materials column
            self._on_materials_click(record)

    def _show_inline_type_combobox(self, item_id: str, column: str, record: BlueprintRecord) -> None:
        """Show a dark-themed popup listbox anchored below the Type cell."""
        bbox = self.blueprint_tree.bbox(item_id, column)
        if not bbox:
            return
        x, y, cell_w, cell_h = bbox

        abs_x = self.blueprint_tree.winfo_rootx() + x
        abs_y = self.blueprint_tree.winfo_rooty() + y + cell_h

        ITEM_H = 22
        popup_w = max(cell_w, 150)
        popup_h = len(BLUEPRINT_CATEGORY_OPTIONS) * ITEM_H + 2

        # Flip above the cell if popup would clip below the screen
        screen_h = self.root.winfo_screenheight()
        if abs_y + popup_h > screen_h - 40:
            abs_y = self.blueprint_tree.winfo_rooty() + y - popup_h

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.configure(bg="#c09040")          # 1 px gold border via padx/pady below
        popup.geometry(f"{popup_w}x{popup_h}+{abs_x}+{abs_y}")
        popup.lift()

        lb = tk.Listbox(
            popup,
            bg="#0d1219",
            fg="#e8edf2",
            selectbackground="#c09040",
            selectforeground="#080c10",
            relief="flat",
            font=("Segoe UI", 9),
            bd=0,
            highlightthickness=0,
            activestyle="none",
            exportselection=False,
        )
        lb.pack(fill="both", expand=True, padx=1, pady=1)

        current = record.category_override or "Auto"
        for i, option in enumerate(BLUEPRINT_CATEGORY_OPTIONS):
            lb.insert("end", f"  {option}")
            if option == current:
                lb.selection_set(i)
                lb.activate(i)
                lb.see(i)

        def _close(_event=None) -> None:
            try:
                popup.destroy()
            except Exception:
                pass

        def on_pick(event=None) -> None:
            # Use nearest() so we don't depend on curselection() being set —
            # more reliable than <ButtonRelease-1> on Windows overrideredirect windows.
            if event is not None:
                idx = lb.nearest(event.y)
            else:
                sel = lb.curselection()
                idx = sel[0] if sel else -1
            if idx < 0 or idx >= lb.size():
                return
            choice = lb.get(idx).strip()
            _close()
            self._apply_type_override(record, choice)

        def on_focus_out(_event=None) -> None:
            # Small delay prevents premature dismiss when focus shifts between
            # the Toplevel and its child Listbox during initial show.
            popup.after(80, _close)

        lb.bind("<Button-1>", on_pick)
        lb.bind("<Return>", on_pick)
        lb.bind("<Escape>", _close)
        lb.bind("<FocusOut>", on_focus_out)

        popup.focus_force()
        lb.focus_set()

    def _apply_type_override(self, record: BlueprintRecord, choice: str) -> None:
        """Persist a user-chosen type override for a blueprint record."""
        override = "" if choice == "Auto" else choice
        if override == record.category_override:
            return
        record.category_override = override
        if self.state.blueprint_category_overrides is None:
            self.state.blueprint_category_overrides = {}
        if override:
            self.state.blueprint_category_overrides[record.normalized_name] = override
        else:
            self.state.blueprint_category_overrides.pop(record.normalized_name, None)
        save_state(self.state)
        self._refresh_blueprint_list()

    def _on_materials_click(self, record: BlueprintRecord) -> None:
        """Show crafting materials for the clicked blueprint, loading the DB if needed."""
        if _crafting_db is None:
            if self._crafting_db_loading:
                return
            self._crafting_db_loading = True
            self.blueprint_status_var.set("Fetching crafting database...")

            def worker():
                try:
                    ensure_crafting_db_loaded()
                    self.root.after(0, lambda r=record: self._on_crafting_db_ready(r))
                except Exception as exc:
                    self.root.after(0, lambda e=exc: self._on_crafting_db_error(e))

            threading.Thread(target=worker, daemon=True).start()
            return
        self._show_crafting_popup(record)

    def _on_crafting_db_ready(self, record: BlueprintRecord) -> None:
        self._crafting_db_loading = False
        self.blueprint_status_var.set("Crafting database loaded.")
        self._show_crafting_popup(record)

    def _on_crafting_db_error(self, exc: Exception) -> None:
        self._crafting_db_loading = False
        self.blueprint_status_var.set("Could not load crafting database.")
        self.append_log(f"Crafting database fetch failed: {exc}")

    def _show_crafting_popup(self, record: BlueprintRecord) -> None:
        materials = lookup_crafting_materials(record.name)

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Crafting Materials")
        dialog.configure(bg="#0d1219")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # Gold accent stripe
        tk.Frame(dialog, bg="#c09040", height=2).pack(fill="x", side="top")

        content = tk.Frame(dialog, bg="#0d1219")
        content.pack(fill="both", expand=True, padx=24, pady=(18, 10))

        tk.Label(
            content, text=record.name,
            fg="#e8edf2", bg="#0d1219",
            font=("Segoe UI Semibold", 11), anchor="w",
        ).pack(fill="x")
        tk.Label(
            content, text="Crafting Materials",
            fg="#6e8096", bg="#0d1219",
            font=("Segoe UI", 9), anchor="w",
        ).pack(fill="x", pady=(2, 14))

        if not materials:
            tk.Label(
                content,
                text="No crafting data found for this blueprint.",
                fg="#6e8096", bg="#0d1219",
                font=("Segoe UI", 9),
            ).pack(anchor="w")
        else:
            # Column header row
            hdr = tk.Frame(content, bg="#162030")
            hdr.pack(fill="x", pady=(0, 2))
            for label, w, anchor in (("SLOT", 160, "w"), ("MATERIAL", 160, "w"), ("QUANTITY", 90, "e")):
                tk.Label(
                    hdr, text=label,
                    fg="#c09040", bg="#162030",
                    font=("Segoe UI Semibold", 8),
                    width=0, anchor=anchor, padx=10, pady=5,
                ).pack(side="left", ipadx=0)
                # Spacer to reach fixed width
                tk.Frame(hdr, bg="#162030", width=w).pack(side="left")

            # Clear the spacer approach — use a grid Frame instead
            hdr.destroy()
            hdr = tk.Frame(content, bg="#162030")
            hdr.pack(fill="x", pady=(0, 2))
            hdr.columnconfigure(0, minsize=170)
            hdr.columnconfigure(1, minsize=170)
            hdr.columnconfigure(2, minsize=90)
            tk.Label(hdr, text="SLOT",     fg="#c09040", bg="#162030", font=("Segoe UI Semibold", 8), anchor="w", padx=10, pady=5).grid(row=0, column=0, sticky="ew")
            tk.Label(hdr, text="MATERIAL", fg="#c09040", bg="#162030", font=("Segoe UI Semibold", 8), anchor="w", padx=10, pady=5).grid(row=0, column=1, sticky="ew")
            tk.Label(hdr, text="QUANTITY", fg="#c09040", bg="#162030", font=("Segoe UI Semibold", 8), anchor="e", padx=10, pady=5).grid(row=0, column=2, sticky="ew")

            # Material rows
            for i, mat in enumerate(materials):
                row_bg = "#0d1219" if i % 2 == 0 else "#111922"
                row = tk.Frame(content, bg=row_bg)
                row.pack(fill="x")
                row.columnconfigure(0, minsize=170)
                row.columnconfigure(1, minsize=170)
                row.columnconfigure(2, minsize=90)
                tk.Label(row, text=mat.slot,              fg="#e8edf2", bg=row_bg, font=("Segoe UI", 9), anchor="w", padx=10, pady=4).grid(row=0, column=0, sticky="ew")
                tk.Label(row, text=mat.resource,          fg="#e8edf2", bg=row_bg, font=("Segoe UI", 9), anchor="w", padx=10, pady=4).grid(row=0, column=1, sticky="ew")
                tk.Label(row, text=format_scu(mat.quantity), fg="#c09040", bg=row_bg, font=("Segoe UI Semibold", 9), anchor="e", padx=10, pady=4).grid(row=0, column=2, sticky="ew")

        tk.Label(
            content,
            text="Data sourced from community crafting database.",
            fg="#3a4a5a", bg="#0d1219",
            font=("Segoe UI", 8),
        ).pack(anchor="w", pady=(12, 0))

        btn_row = tk.Frame(dialog, bg="#0d1219")
        btn_row.pack(fill="x", padx=24, pady=(4, 18))
        tk.Button(
            btn_row, text="Close", command=dialog.destroy,
            bg="#162030", fg="#e8edf2",
            activebackground="#1e2d3d", activeforeground="#e8edf2",
            relief="flat", padx=14, pady=8,
            font=("Segoe UI Semibold", 9), cursor="hand2", bd=0,
        ).pack(side="right")

        dialog.update_idletasks()
        w = max(dialog.winfo_reqwidth(), 460)
        h = dialog.winfo_reqheight()
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        dialog.geometry(f"{w}x{h}+{x}+{y}")
        apply_dark_titlebar(dialog)
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.grab_set()
        dialog.focus_force()

    def _schedule_blueprint_search(self) -> None:
        """Debounce blueprint search box keystrokes (150 ms) to avoid O(n) work per key."""
        if self._search_debounce_job is not None:
            try:
                self.root.after_cancel(self._search_debounce_job)
            except Exception:
                pass
        self._search_debounce_job = self.root.after(150, self._refresh_blueprint_list)

    def _show_blueprint_details(self, record: BlueprintRecord | None) -> None:
        self.selected_blueprint_record = record
        self.blueprint_contracts_text.configure(state="normal")
        self.blueprint_contracts_text.delete("1.0", "end")
        if record is None:
            self.blueprint_detail_title_var.set("No blueprint selected")
            self.blueprint_contracts_text.insert("1.0", "Possible contract sources will appear here once a blueprint is selected.")
            self.blueprint_contracts_text.configure(state="disabled")
            return

        self.blueprint_detail_title_var.set(record.name)
        if record.contracts:
            for contract in record.contracts:
                self.blueprint_contracts_text.insert("end", f"• {contract}\n")
        else:
            self.blueprint_contracts_text.insert("1.0", "No contract source was found in the installed StarStrings data for this blueprint.")
        self.blueprint_contracts_text.configure(state="disabled")

    def _on_app_update_button_click(self) -> None:
        if self.app_update_available and self.app_update_release is not None:
            self._download_and_install_app_update(self.app_update_release)
        else:
            self.check_for_app_update()

    def _set_activity_badge(self, active: bool) -> None:
        if not hasattr(self, "activity_view_button"):
            return
        text = "ACTIVITY  ●" if active else "ACTIVITY"
        idle_style = "ViewIdleBadge.TButton" if active else "ViewIdle.TButton"
        self.activity_view_button.configure(text=text)
        if self.current_view != "activity":
            self.activity_view_button.configure(style=idle_style)

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        level = self._log_level(message)
        self.log_widget.configure(state="normal")
        self.log_widget.insert("end", f"[{timestamp}] ", "ts")
        self.log_widget.insert("end", f"{message}\n", level)
        line_count = int(self.log_widget.index("end-1c").split(".")[0])
        if line_count > MAX_LOG_LINES:
            trim_to = line_count - MAX_LOG_LINES
            self.log_widget.delete("1.0", f"{trim_to + 1}.0")
        self.log_widget.see("end")
        self.log_widget.configure(state="disabled")

    def _refresh_status_vars(self) -> None:
        """Refresh the status panel StringVars from self.state — call after mutating state."""
        self.release_var.set(self.state.tracked_release_name or "No release tracked yet")
        self.last_checked_var.set(format_timestamp(self.state.last_checked_at, "Not checked yet"))
        self.last_updated_var.set(format_timestamp(self.state.last_update_at, "No update applied yet"))
        self.backup_var.set(f"Backups stored in:\n{BACKUP_ROOT}")
        self._refresh_blueprint_freshness()

    def _report_completed_staged_update(self) -> None:
        result_path = PENDING_UPDATE_DIR / "last_update.json"
        if not result_path.exists():
            return
        try:
            payload = json.loads(result_path.read_text(encoding="utf-8"))
            version = str(payload.get("version", "")).strip() or "latest release"
            applied_at = format_timestamp(str(payload.get("applied_at", "")).strip(), "recently")
            self.append_log(f"Application updated to v{version}. Applied on {applied_at}.")
            append_update_trace(f"App confirmed staged update to v{version} on startup.")
        except Exception:
            self.append_log("Application update was applied.")
        finally:
            try:
                result_path.unlink()
            except Exception:
                pass

    def _load_log(self) -> None:
        self.log_widget.configure(state="normal")
        self.log_widget.delete("1.0", "end")
        for line in read_log_tail().splitlines():
            if not line:
                self.log_widget.insert("end", "\n", "info")
                continue
            if line.startswith("[") and "] " in line:
                split = line.index("] ") + 2
                ts_part, msg_part = line[:split], line[split:]
                self.log_widget.insert("end", ts_part, "ts")
                self.log_widget.insert("end", f"{msg_part}\n", self._log_level(msg_part))
            else:
                self.log_widget.insert("end", f"{line}\n", "info")
        self.log_widget.see("end")
        self.log_widget.configure(state="disabled")

    def _refresh_toggle(self) -> None:
        """Query the scheduled task on a worker thread to avoid blocking the UI."""
        def worker() -> None:
            result = run_process(["schtasks", "/Query", "/TN", TASK_NAME, "/FO", "LIST", "/V"])
            self.root.after(0, lambda: self._apply_toggle_result(result))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_toggle_result(self, result) -> None:
        if result.returncode == 0:
            lines: dict[str, str] = {}
            for raw_line in result.stdout.splitlines():
                if ":" not in raw_line:
                    continue
                key, value = raw_line.split(":", 1)
                lines[key.strip()] = value.strip()
            status = lines.get("Status", "Unknown")
            last_run = format_scheduler_timestamp(lines.get("Last Run Time", ""), "Not run yet")
            next_run = format_scheduler_timestamp(lines.get("Next Run Time", ""), "Not scheduled yet")
            self.toggle_button.configure(text="StarStrings Auto Update: On", style="ToggleOn.TButton")
            self.auto_state_var.set("Enabled")
            self.schedule_var.set(f"Scheduled: {status} | Last Run: {last_run} | Next Run: {next_run}")
        else:
            self.toggle_button.configure(text="StarStrings Auto Update: Off", style="ToggleOff.TButton")
            self.auto_state_var.set("Disabled")
            self.schedule_var.set("Scheduled task is not registered.")

    def run_manual_update(self) -> None:
        self.save_settings()
        self.run_button.configure(state="disabled")
        self.append_log("Checking GitHub for updates...")

        def worker() -> None:
            try:
                result = run_update(self.settings, allow_prompt=True, parent=self.root, force_update=True)
                self.root.after(0, lambda: self._complete_run(result, is_error=False))
            except Exception as exc:
                log(f"Manual run failed. {exc}")
                self.root.after(0, lambda: self._complete_run(str(exc), is_error=True))

        threading.Thread(target=worker, daemon=True).start()

    def _complete_run(self, message: str, is_error: bool) -> None:
        self.run_button.configure(state="normal")
        self.state = load_state()
        self._refresh_status_vars()
        self._refresh_toggle()
        self.append_log(message)
        if not is_error:
            self.scan_blueprints(silent=True)
        self._show_run_result_dialog(message, is_error=is_error)

    def check_for_app_update(self) -> None:
        if self._app_update_checking:
            return
        self._app_update_checking = True
        self._set_app_update_checking_state()
        self.append_log(f"Checking {APP_UPDATE_REPO} for a new app release...")

        def worker() -> None:
            try:
                release = fetch_latest_app_release()
                self.root.after(0, lambda release=release: self._handle_app_release_check(release, notify_if_current=True))
            except NoPublishedAppReleaseError as exc:
                self.root.after(0, lambda exc=exc: self._handle_app_update_error(exc))
            except Exception as exc:
                self.root.after(0, lambda exc=exc: self._handle_app_update_error(exc))
            finally:
                self.root.after(0, self._clear_update_checking_flag)

        threading.Thread(target=worker, daemon=True).start()

    def check_for_app_update_silent(self) -> None:
        if self._app_update_checking:
            self.root.after(0, self._schedule_next_app_update_check)
            return
        self._app_update_checking = True
        self._set_app_update_checking_state()

        def worker() -> None:
            try:
                release = fetch_latest_app_release()
                self.root.after(0, lambda release=release: self._handle_app_release_check(release, notify_if_current=False))
            except NoPublishedAppReleaseError:
                self.root.after(0, self._set_app_update_no_release_state)
            except Exception:
                self.root.after(0, self._set_app_update_unknown_state)
            finally:
                self.root.after(0, self._clear_update_checking_flag)
                self.root.after(0, self._schedule_next_app_update_check)

        threading.Thread(target=worker, daemon=True).start()

    def _clear_update_checking_flag(self) -> None:
        self._app_update_checking = False

    def _schedule_next_app_update_check(self) -> None:
        if self.app_update_check_job is not None:
            try:
                self.root.after_cancel(self.app_update_check_job)
            except Exception:
                pass
        self.app_update_check_job = self.root.after(APP_UPDATE_CHECK_INTERVAL_MS, self.check_for_app_update_silent)

    def _handle_app_release_check(self, release: AppReleaseInfo, notify_if_current: bool) -> None:
        if not is_newer_version(release.version, APP_VERSION):
            self._set_app_update_idle_state()
            self.append_log(f"Latest app release is {release.version}. Current version is v{APP_VERSION}.")
            if notify_if_current:
                self.append_log(f"Application is up to date. Current version is v{APP_VERSION}.")
            return

        self._set_app_update_available_state(release)
        self.append_log(f"New version available: v{APP_VERSION} → {release.version}. Click 'Update Application' to install.")

    def _download_and_install_app_update(self, release: AppReleaseInfo) -> None:
        self.append_log(f"Downloading app update {release.version}...")
        self.app_update_button_var.set(f"Downloading {release.version}...")
        self.app_update_button.configure(style="AppUpdateIdle.TButton")
        self._stop_app_update_pulse()

        def worker() -> None:
            temp_dir = Path(tempfile.mkdtemp(prefix="starstrings-app-update-"))
            try:
                download_path = temp_dir / release.asset_name
                download_file(release.download_url, download_path)
                # Pass temp_dir so _finalize_app_update can clean it up after the
                # PowerShell installer has copied the exe out of it.
                self.root.after(0, lambda dp=download_path, rel=release, td=temp_dir: self._finalize_app_update(dp, rel, td))
            except Exception as exc:
                shutil.rmtree(temp_dir, ignore_errors=True)
                self.root.after(0, lambda exc=exc: self._handle_app_update_error(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _finalize_app_update(self, download_path: Path, release: AppReleaseInfo, temp_dir: Path) -> None:
        try:
            append_update_trace(f"Finalize requested for app update v{release.version}. downloaded_exe={download_path}")

            # Verify integrity against the SHA256 digest published with the release asset.
            if release.digest:
                actual = sha256_file(download_path)
                if actual != release.digest:
                    raise UpdaterError(
                        f"Downloaded update failed integrity check (expected {release.digest[:16]}…, got {actual[:16]}…). "
                        "The file may be corrupted or tampered with."
                    )
                append_update_trace(f"Integrity check passed for v{release.version}.")
            else:
                append_update_trace(f"No digest published for v{release.version}; skipping integrity check.")

            install_app_update(download_path, release.version)
        except Exception as exc:
            shutil.rmtree(temp_dir, ignore_errors=True)
            append_update_trace(f"Finalize failed for v{release.version}: {exc}")
            self._handle_app_update_error(exc)
            return
        # temp_dir is cleaned up on the next OS temp sweep once the helper copies the exe out.

        self.append_log(f"App update {release.version} downloaded. Applying update — the app will close and relaunch automatically.")
        self._quit_for_update()

    def _handle_app_update_error(self, exc: Exception) -> None:
        error_text = str(exc).strip() or "Application update status could not be refreshed."
        if isinstance(exc, NoPublishedAppReleaseError):
            self._set_app_update_idle_state()
            self.append_log("Application is up to date.")
            return
        self._set_app_update_idle_state()
        self.append_log(f"Application update check skipped. {error_text}")

    def _set_app_update_checking_state(self) -> None:
        self.app_update_available = False
        self.app_update_release = None
        self.app_update_button_var.set("Checking for Updates...")
        self.app_update_button.configure(style="AppUpdateIdle.TButton")
        self._stop_app_update_pulse()

    def _set_app_update_idle_state(self) -> None:
        self.app_update_available = False
        self.app_update_release = None
        self.app_update_button_var.set("Check for Updates")
        self.app_update_button.configure(style="AppUpdateIdle.TButton")
        self._stop_app_update_pulse()

    def _set_app_update_available_state(self, release: AppReleaseInfo) -> None:
        self.app_update_available = True
        self.app_update_release = release
        self.app_update_button_var.set(f"Update Application  →  {release.version}")
        self._show_view("activity")
        self._set_activity_badge(True)
        self._start_app_update_pulse()

    def _set_app_update_unknown_state(self) -> None:
        self.app_update_available = False
        self.app_update_release = None
        self.app_update_button_var.set("Check for Updates")
        self.app_update_button.configure(style="AppUpdateIdle.TButton")
        self._stop_app_update_pulse()

    def _set_app_update_no_release_state(self) -> None:
        self.app_update_available = False
        self.app_update_release = None
        self.app_update_button_var.set("Check for Updates")
        self.app_update_button.configure(style="AppUpdateIdle.TButton")
        self._stop_app_update_pulse()

    def _start_app_update_pulse(self) -> None:
        self._stop_app_update_pulse()
        self.app_update_pulse_on = False
        self._pulse_app_update_button()

    def _stop_app_update_pulse(self) -> None:
        if self.app_update_pulse_job is not None:
            try:
                self.root.after_cancel(self.app_update_pulse_job)
            except Exception:
                pass
        self.app_update_pulse_job = None
        self._set_activity_badge(False)

    def _pulse_app_update_button(self) -> None:
        if not self.app_update_available:
            self.app_update_button.configure(style="AppUpdateIdle.TButton")
            self.app_update_pulse_job = None
            return
        self.app_update_pulse_on = not self.app_update_pulse_on
        self.app_update_button.configure(style="AppUpdateAlertA.TButton" if self.app_update_pulse_on else "AppUpdateAlertB.TButton")
        self.app_update_pulse_job = self.root.after(600, self._pulse_app_update_button)

    def toggle_schedule(self) -> None:
        try:
            self.save_settings(reapply_schedule=False)
            if self.auto_state_var.get() == "Enabled":
                unregister_scheduled_task()
                self.append_log("Automatic updates disabled. Scheduled task removed.")
            else:
                register_scheduled_task(self.settings.interval_hours)
                self.append_log(f"Automatic updates enabled every {self.settings.interval_hours} hour(s).")
            self._refresh_toggle()
        except Exception as exc:
            log(f"Failed to change automatic updates. {exc}")
            self.append_log(f"Failed to change automatic updates: {exc}")
            messagebox.showerror(APP_NAME, str(exc), parent=self.root)

    def run(self) -> None:
        self.root.mainloop()


def run_scheduled() -> int:
    if not _acquire_scheduled_instance():
        log("Scheduled run skipped because another scheduled instance is already active.")
        return 0
    settings = load_settings()
    try:
        log("Scheduled StarStrings update check started.")
        message = run_update(settings, allow_prompt=False, force_update=False)
        log(message)
        log("Scheduled StarStrings update check finished; exiting.")
        return 0
    except Exception as exc:
        log(f"Scheduled run failed. {exc}")
        return 1


_INSTANCE_MUTEX: object = None   # keeps the handle alive for the process lifetime
_SCHEDULED_MUTEX: object = None  # prevents overlapping headless scheduled runs


def _acquire_single_instance() -> bool:
    """Return True if this is the first instance, False if another is already running."""
    global _INSTANCE_MUTEX
    if sys.platform != "win32":
        return True
    import ctypes
    handle = ctypes.windll.kernel32.CreateMutexW(None, True, "CitizenStarStringHelper_SingleInstance")
    err = ctypes.windll.kernel32.GetLastError()
    _INSTANCE_MUTEX = handle  # prevent GC / premature release
    return err != 183  # ERROR_ALREADY_EXISTS


def _acquire_scheduled_instance() -> bool:
    """Return True if this is the only scheduled run, False if another is active."""
    global _SCHEDULED_MUTEX
    if sys.platform != "win32":
        return True
    handle = ctypes.windll.kernel32.CreateMutexW(None, True, "CitizenStarStringHelper_ScheduledRun")
    err = ctypes.windll.kernel32.GetLastError()
    _SCHEDULED_MUTEX = handle  # prevent GC / premature release
    return err != 183  # ERROR_ALREADY_EXISTS


def main(argv: list[str]) -> int:
    # Bootstrap: create data directory and migrate legacy files before anything else runs.
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    migrate_legacy_data()

    if "--scheduled" in argv:
        return run_scheduled()
    if "--warmup" in argv:
        # All module-level imports have already executed, so every DLL this exe
        # needs has been extracted to %TEMP%\_MEI* and loaded by the time we
        # reach here.  Windows Defender will scan them during this window.
        # The batch update helper waits for this process to exit before
        # launching the real instance, so Defender's cache is warm by then.
        return 0
    if not _acquire_single_instance():
        # Another instance is already running — exit silently.
        return 0
    app = StarStringsApp()
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
