"""
Number File Submission Bot
Python 3.10+
Install: pip install -r requirements.txt
Run: python main.py

Set BOT_TOKEN and ADMIN_IDS in hosting environment variables before running.
Handles number-list files (.xlsx, .csv, .txt) with manual admin review.
Categories and payment methods are fully managed from the Admin Panel.
It intentionally does not collect passwords, cookies, 2FA codes, or session credentials.
"""

import asyncio
import logging
import os
import re
import sqlite3
import uuid
import html
import json
import inspect
import ast
import hashlib
import unicodedata
import functools
import contextvars
from datetime import datetime, timezone
from pathlib import Path

from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup,
    InlineKeyboardButton as _TGInlineKeyboardButton
)
from telegram.constants import ChatAction
from telegram.error import TelegramError, InvalidToken, BadRequest, Forbidden, RetryAfter, NetworkError, TimedOut
from telegram.constants import ParseMode
from telegram.request import HTTPXRequest
from telegram import Bot, Message, CallbackQuery
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ContextTypes, ApplicationHandlerStop, TypeHandler, filters
)

# =====================================================================
#                 এখানে আপনার BOT TOKEN ও ADMIN ID বসান
# =====================================================================
#
# ১) BOT TOKEN:
# নিচের কোটেশনের (" ") ভেতরে BotFather থেকে পাওয়া সম্পূর্ণ Token বসান।
# উদাহরণ: BOT_TOKEN = "123456789:আপনার-আসল-token"
# নিজের আসল Token বসাবেন; উদাহরণের Token ব্যবহার করবেন না।
#
BOT_TOKEN = "9152728729:AAEXrxBrqC4Qw30N5bE-DPnmm_6DlgIr2H4"  # নিরাপত্তার জন্য token Hosting Environment Variable-এ দিন।

# ২) ADMIN TELEGRAM ID:
# নিচে নিজের Telegram Numeric User ID বসান।
# উদাহরণ: ADMIN_IDS = {123456789}
# একাধিক Admin হলে: ADMIN_IDS = {123456789, 987654321}
#
ADMIN_IDS = {6800119501}
# ---------------------------------------------------------------------
# বিকল্প: Hosting-এর Environment Variables-এ BOT_TOKEN / TELEGRAM_BOT_TOKEN
# এবং ADMIN_IDS সেট করলে সেগুলো উপরের কোডের মানকে Override করবে।
# ---------------------------------------------------------------------
_env_bot_token = (
    os.getenv("BOT_TOKEN")
    or os.getenv("TELEGRAM_BOT_TOKEN")
    or ""
).strip().strip('"').strip("'")
if _env_bot_token:
    BOT_TOKEN = _env_bot_token
else:
    BOT_TOKEN = BOT_TOKEN.strip().strip('"').strip("'")

_env_admin_ids = os.getenv("ADMIN_IDS", "").strip()
if _env_admin_ids:
    ADMIN_IDS = {
        int(value.strip())
        for value in _env_admin_ids.split(",")
        if value.strip().lstrip("-").isdigit()
    }

BOT_NAME = "Number File Submission Bot"
DEFAULT_SUPPORT_LINK = "@YourSupport"  # Admin Panel থেকে পরে পরিবর্তন করা যাবে
CURRENCY = "BDT"
DB_PATH = "bot_data.sqlite3"
UPLOAD_DIR = Path("uploads")
MAX_FILE_MB = 20
# =====================================================================

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO
)
log = logging.getLogger("number-file-bot")

# =====================================================================
#   🎨 PREMIUM BUTTON COLORS (Telegram Bot API 9.4+ button "style")
#   primary = নীল | success = সবুজ | danger = লাল
#   পুরনো Telegram ক্লায়েন্টে রঙ না দেখালেও বাটন স্বাভাবিক কাজ করে।
# =====================================================================
_DANGER_MARKERS = ("❌", "🗑", "🚫", "🔴", "⛔", "➖", "reject", "cancel", "delete", "remove")
_SUCCESS_MARKERS = ("✅", "🟢", "➕", "📤", "💾", "approve", "confirm", "save", "add ", "send")


def _auto_button_style(text):
    """বাটনের লেখা দেখে রঙ ঠিক করে: বাতিল/মোছা = লাল, নিশ্চিত/যোগ/পাঠানো = সবুজ, বাকি সব = নীল।"""
    lowered = str(text).lower()
    if any(marker in lowered for marker in _DANGER_MARKERS):
        return "danger"
    if any(marker in lowered for marker in _SUCCESS_MARKERS):
        return "success"
    return "primary"


def InlineKeyboardButton(text, *args, **kwargs):
    """telegram.InlineKeyboardButton-এর মতোই কাজ করে, শুধু স্বয়ংক্রিয়ভাবে রঙ যোগ করে।
    আলাদা রঙ চাইলে: InlineKeyboardButton("..", callback_data="..", style="danger")"""
    api_kwargs = dict(kwargs.pop("api_kwargs", None) or {})
    explicit = kwargs.pop("style", None)
    api_kwargs["style"] = explicit or api_kwargs.get("style") or _auto_button_style(text)
    try:
        text, _icon = apply_button_override(text)  # Add Single Emoji ওভাররাইড
        if _icon:
            api_kwargs["icon_custom_emoji_id"] = _icon
    except Exception:
        pass
    return _TGInlineKeyboardButton(fancy_text(text, "bold"), *args, api_kwargs=api_kwargs, **kwargs)


def ColorButton(text, style):
    """রঙিন Reply Keyboard বাটন (style: primary / success / danger)।"""
    api = {"style": style}
    try:
        text, _icon = apply_button_override(text)  # Add Single Emoji ওভাররাইড
        if _icon:
            api["icon_custom_emoji_id"] = _icon
    except Exception:
        pass
    return KeyboardButton(text, api_kwargs=api)


# =====================================================================
#   🔤 STYLISH FONTS (Unicode Math Alphanumeric)
#   bold        → 𝐄𝐗𝐀𝐌𝐏𝐋𝐄 𝐅𝐎𝐍𝐓      (ইনলাইন বাটন ও মেসেজের হেডিং)
#   bold_italic_sans → 𝙀𝙓𝘼𝙈𝙋𝙇𝙀 𝙁𝙊𝙉𝙏  (নিচের মেইন মেনু বাটন)
#   ইউজারনেম, ফাইলের নাম, সংখ্যা, লিংক ও বাংলা লেখা অপরিবর্তিত থাকে।
# =====================================================================
_FANCY_BASES = {"bold": (0x1D400, 0x1D41A), "bold_italic_sans": (0x1D63C, 0x1D656)}
_FANCY_SKIP = re.compile(r"[@._0-9]|^https?:|^/")


def fancy_text(text, style="bold", upper=True):
    up_base, low_base = _FANCY_BASES[style]

    def convert(token):
        if not token or token.isspace() or _FANCY_SKIP.search(token):
            return token
        out = []
        for ch in token:
            if "A" <= ch <= "Z":
                out.append(chr(up_base + ord(ch) - 65))
            elif "a" <= ch <= "z":
                out.append(chr((up_base if upper else low_base) + ord(ch) - 97))
            else:
                out.append(ch)
        return "".join(out)

    return "".join(convert(part) for part in re.split(r"(\s+)", str(text)))


# মেইন মেনুর বাটন লেবেল (text_router এই ধ্রুবকগুলো দিয়েই মেলায়)
BTN_SUBMIT = fancy_text("📩 Submit File", "bold_italic_sans")
BTN_PROFILE = fancy_text("👤 My Profile", "bold_italic_sans")
BTN_PAYMENTS = fancy_text("💳 My Payments", "bold_italic_sans")
BTN_PRICES = fancy_text("🎯 Price & Rules", "bold_italic_sans")
BTN_HOWTO = fancy_text("📖 How To Use", "bold_italic_sans")
BTN_SUPPORT = fancy_text("💬 Support", "bold_italic_sans")
BTN_PAYISSUE = "⚠️ পেমেন্ট ইস্যু"
BTN_ADMIN = fancy_text("🛠 Admin Panel", "bold_italic_sans")

# পুরনো কিবোর্ড (আগের লেবেল) যাদের ফোনে এখনো আছে, তাদের চাপও যেন কাজ করে
LEGACY_MENU_ALIASES = {
    old: new for old, new in {
        "📩 Submit File": BTN_SUBMIT, "👤 My Profile": BTN_PROFILE, "💳 My Payments": BTN_PAYMENTS,
        "🎯 Price & Rules": BTN_PRICES, "📖 How To Use": BTN_HOWTO, "💬 Support": BTN_SUPPORT,
        "⚠️ পেমেন্ট ইস্যু": BTN_PAYISSUE, "🛠 Admin Panel": BTN_ADMIN,
    }.items() if old != new
}

MAIN_MENU = [
    [ColorButton(BTN_SUBMIT, "success")],
    [ColorButton(BTN_PROFILE, "primary"), ColorButton(BTN_PAYMENTS, "success")],
    [ColorButton(BTN_PRICES, "success"), ColorButton(BTN_HOWTO, "primary")],
    [ColorButton(BTN_SUPPORT, "primary"), ColorButton(BTN_PAYISSUE, "danger")],
]

DB_SCHEMA = """
CREATE TABLE IF NOT EXISTS users(
  user_id INTEGER PRIMARY KEY, username TEXT, first_name TEXT,
  balance REAL NOT NULL DEFAULT 0, joined_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS categories(
  id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL,
  rate REAL NOT NULL DEFAULT 0, is_open INTEGER NOT NULL DEFAULT 1,
  group_id TEXT, notice TEXT, submit_note TEXT
);
CREATE TABLE IF NOT EXISTS submissions(
  id TEXT PRIMARY KEY, user_id INTEGER NOT NULL, category_id INTEGER NOT NULL,
  file_id TEXT NOT NULL, file_name TEXT NOT NULL, total_items INTEGER NOT NULL,
  status TEXT NOT NULL DEFAULT 'PENDING', valid_items INTEGER DEFAULT 0,
  invalid_items INTEGER DEFAULT 0, rate REAL NOT NULL, amount REAL DEFAULT 0,
  submitted_at TEXT NOT NULL, reviewed_at TEXT, admin_note TEXT,
  payment_method_id INTEGER, payment_method_name TEXT, payment_details TEXT,
  paid_at TEXT, group_message_id INTEGER
);
CREATE TABLE IF NOT EXISTS payments(
  id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL,
  submission_id TEXT NOT NULL, amount REAL NOT NULL, note TEXT,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS payment_methods(
  id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL,
  account_info TEXT, is_open INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS payment_reports(
  id INTEGER PRIMARY KEY AUTOINCREMENT, submission_id TEXT NOT NULL,
  user_id INTEGER NOT NULL, amount REAL NOT NULL, message TEXT,
  file_id TEXT, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS bot_settings(
  setting_key TEXT PRIMARY KEY, setting_value TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS additional_admins(
  user_id INTEGER PRIMARY KEY, added_by INTEGER NOT NULL, added_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS force_join_channels(
  id INTEGER PRIMARY KEY AUTOINCREMENT, chat_id INTEGER NOT NULL UNIQUE,
  join_link TEXT NOT NULL, title TEXT NOT NULL DEFAULT '', added_at TEXT NOT NULL
);
"""

# Existing databases-এ নতুন column যোগ করার জন্য (পুরনো ডেটা নষ্ট হবে না)
SUBMISSION_NEW_COLUMNS = [
    ("payment_method_id", "INTEGER"),
    ("payment_method_name", "TEXT"),
    ("payment_details", "TEXT"),
    ("paid_at", "TEXT"),
    ("group_message_id", "INTEGER"),
    ("user_message", "TEXT"),
    ("local_path", "TEXT"),
]

CATEGORY_NEW_COLUMNS = [
    ("group_id", "TEXT"),
    ("notice", "TEXT"),
    ("submit_note", "TEXT"),
    ("is_deleted", "INTEGER NOT NULL DEFAULT 0"),
]

# User submission flow-এর temporary state
SUBMISSION_STATE_KEYS = (
    "selected_category",
    "selected_payment_method",
    "selected_payment_details",
    "awaiting_payment_details",
)

# Admin input flow-এর temporary state
ADMIN_STATE_KEYS = (
    "admin_action",
    "admin_submission_id",
    "admin_category_id",
    "admin_pm_id",
    "settle_valid_count",
    "balance_user_id",
    "balance_direction",
    "report_sid",
    "report_amount",
    "report_message",
    "admin_group_cat_id",
    "admin_notice_cat_id",
    "admin_submit_note_cat_id",
    "broadcast_pending",
    "emoji_edit_ref",
    "emoji_edit_scope",
    "emoji_edit_page",
)


def db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA busy_timeout=30000")
    except sqlite3.Error:
        pass
    return conn


def init_db():
    with db() as con:
        con.executescript(DB_SCHEMA)

        # নোট: আগে ক্যাটাগরি টেবিল খালি হলে প্রতিবার বট চালুর সময় Facebook/Instagram Numbers
        # নিজে থেকে আবার তৈরি হতো — এটাই "ডিলিট করার পরেও ক্যাটাগরি দেখানো"র কারণ ছিল।
        # এখন আর কোনো ক্যাটাগরি অটো-তৈরি হয় না; ক্যাটাগরি শুধু Admin Panel থেকে যোগ হবে।

        # পুরনো submissions টেবিলে নতুন column যোগ (safe migration)
        existing_sub = {row[1] for row in con.execute("PRAGMA table_info(submissions)").fetchall()}
        for col_name, col_type in SUBMISSION_NEW_COLUMNS:
            if col_name not in existing_sub:
                con.execute(f"ALTER TABLE submissions ADD COLUMN {col_name} {col_type}")
                log.info("Migration: submissions.%s column added", col_name)

        # পুরনো categories টেবিলে নতুন column যোগ
        existing_cat = {row[1] for row in con.execute("PRAGMA table_info(categories)").fetchall()}
        for col_name, col_type in CATEGORY_NEW_COLUMNS:
            if col_name not in existing_cat:
                con.execute(f"ALTER TABLE categories ADD COLUMN {col_name} {col_type}")
                log.info("Migration: categories.%s column added", col_name)


def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def is_main_admin(user_id):
    return user_id in ADMIN_IDS


def is_admin(user_id):
    if is_main_admin(user_id):
        return True
    try:
        with db() as con:
            return con.execute("SELECT 1 FROM additional_admins WHERE user_id=?", (user_id,)).fetchone() is not None
    except sqlite3.Error:
        return False


def get_admin_ids():
    with db() as con:
        rows = con.execute("SELECT user_id FROM additional_admins").fetchall()
    return set(ADMIN_IDS) | {int(row[0]) for row in rows}


def get_welcome_message():
    value = get_setting("welcome_message")
    if value:
        return RawHTML(rich_html(value))
    return RawHTML(
        f"👋 <b>{html.escape(fancy_text('WELCOME', 'bold'))}</b>\n"
        f"<b>{html.escape(BOT_NAME)}</b>\n\n"
        "<blockquote>📩 আপনার নাম্বার-লিস্ট ফাইল সাবমিট করুন — অ্যাডমিন যাচাই করে পেমেন্ট যোগ করবে।</blockquote>\n\n"
        "📎 <b>Accepted formats:</b> XLSX • CSV • TXT\n\n"
        "👇 নিচের মেনু থেকে শুরু করুন।"
    )


def get_payment_issue_message():
    """Admin-configured message shown from the user's Payment Issue button."""
    value = get_setting("payment_issue_message")
    if value:
        return RawHTML(rich_html(value))
    return RawHTML("⚠️ পেমেন্ট সংক্রান্ত কোনো নির্দেশনা এখনো সেট করা হয়নি। অনুগ্রহ করে পরে আবার চেষ্টা করুন।")


def get_support_link():
    """Return the support username/link configured by the admin."""
    with db() as con:
        row = con.execute(
            "SELECT setting_value FROM bot_settings WHERE setting_key='support_link'"
        ).fetchone()
    return row[0].strip() if row and row[0].strip() else DEFAULT_SUPPORT_LINK


def normalize_support_url(value):
    value = value.strip()
    if value.startswith("@"):
        return "https://t.me/" + value[1:]
    if re.fullmatch(r"[A-Za-z0-9_]{5,}", value):
        return "https://t.me/" + value
    if re.match(r"^https?://t\.me/[A-Za-z0-9_+/?=-]+$", value):
        return value
    return None


def get_premium_emoji_map():
    """Load the admin-managed ordinary-emoji -> Telegram custom emoji ID map."""
    try:
        with db() as con:
            row = con.execute(
                "SELECT setting_value FROM bot_settings WHERE setting_key='premium_emojis'"
            ).fetchone()
        data = json.loads(row[0]) if row and row[0] else {}
        if not isinstance(data, dict):
            return {}
        return {str(k): str(v) for k, v in data.items()
                if isinstance(k, str) and k and str(v).isdigit()}
    except (sqlite3.Error, json.JSONDecodeError, TypeError, ValueError):
        log.exception("Could not load premium emoji map")
        return {}


# =====================================================================
#   ✨ AUTO PREMIUM EMOJI
#   ডিফল্ট ID গুলো পাবলিক GitHub ক্যাটালগ থেকে নেওয়া। বট চালুর সময় Telegram-কে দিয়েই
#   প্রতিটি ID যাচাই করা হয় (ID বৈধ কি না + ইমোজিটা মিলছে কি না) — ভুল ID কখনো সেভ হয় না।
# =====================================================================
DEFAULT_PREMIUM_EMOJI_IDS = {
    "✅": "5920052658743283381",
    "❌": "5258226313285607065",
    "⬅️": "5877536313623711363",
    "🏠": "5807868868886009920",
    "⬇️": "5258336354642697821",
    "⚠️": "5258474669769497337",
    "⚙️": "5258420634785947640",
    "🚫": "5260293700088511294",
    "📄": "5258477770735885832",
    "💎": "5807465992363710697",
    "⭐️": "5874948844935974490",
    "⚡️": "5843553939672274145",
    "📁": "5875206779196935950",
    "⬆️": "5260652420052032852",
    "ℹ️": "5258503720928288433",
    "❗": "5274099962655816924",
    "❓": "5436113877181941026",
    "🚨": "5456140674028019486",
    "📈": "5449683594425410231",
}

_EMOJI_REGEX = re.compile("[\U0001F300-\U0001FAFF\u2600-\u27BF\u2B00-\u2BFF\u2300-\u23FF\u2B50](?:\ufe0f)?")


def _norm_emoji(value):
    return str(value or "").replace("\ufe0f", "")


def get_bot_emojis():
    """এই বটের কোডে ব্যবহৃত সব সাধারণ ইমোজির তালিকা (ক্রম বজায় রেখে, ডুপ্লিকেট ছাড়া)।"""
    try:
        source = Path(__file__).read_text(encoding="utf-8")
    except Exception:
        return []
    found = []
    for item in _EMOJI_REGEX.findall(source):
        if item not in found:
            found.append(item)
    return found


async def verify_emoji_ids(bot, mapping):
    """Telegram-কে জিজ্ঞেস করে শুধু সেই ID গুলো রাখে যেগুলো সত্যিই আছে এবং ইমোজি মিলছে।"""
    ids = sorted(set(mapping.values()))
    valid = {}
    for i in range(0, len(ids), 100):
        stickers = await bot.get_custom_emoji_stickers(ids[i:i + 100])
        for sticker in stickers:
            valid[str(sticker.custom_emoji_id)] = _norm_emoji(sticker.emoji)
    return {e: cid for e, cid in mapping.items() if valid.get(cid) == _norm_emoji(e)}


async def import_emoji_pack(bot, pack_name):
    """Custom Emoji Pack স্ক্যান করে বটের ইমোজিগুলোর সাথে মিলিয়ে ID বের করে (Telegram থেকেই সরাসরি)।"""
    sticker_set = await bot.get_sticker_set(pack_name)
    by_norm = {}
    for sticker in sticker_set.stickers:
        cid = getattr(sticker, "custom_emoji_id", None)
        if cid and sticker.emoji:
            by_norm.setdefault(_norm_emoji(sticker.emoji), str(cid))
    result = {e: by_norm[_norm_emoji(e)] for e in get_bot_emojis() if _norm_emoji(e) in by_norm}
    return result, len(sticker_set.stickers)


def save_premium_emoji_map(mapping):
    set_setting("premium_emojis", json.dumps(mapping, ensure_ascii=False))


async def seed_default_premium_emojis(application):
    """প্রথমবার চালুর সময় ডিফল্ট প্রিমিয়াম ইমোজি যাচাই করে সেভ করে। Admin 'Reset/Off' করলে আর ফিরে আসে না।"""
    try:
        if get_setting("premium_defaults_seeded"):
            return
        verified = await verify_emoji_ids(application.bot, DEFAULT_PREMIUM_EMOJI_IDS)
        if verified:
            existing = get_premium_emoji_map()
            save_premium_emoji_map({**verified, **existing})  # Admin-এর নিজের ম্যাপিং অগ্রাধিকার পায়
        set_setting("premium_defaults_seeded", "1")
        log.info("Default premium emojis verified & saved: %d/%d", len(verified), len(DEFAULT_PREMIUM_EMOJI_IDS))
    except Exception:
        log.exception("Could not seed default premium emojis (পরের বার চালুর সময় আবার চেষ্টা হবে)")


class RawHTML(str):
    """আগে থেকে তৈরি Telegram-HTML টেক্সট (Admin-এর ফরম্যাটিংসহ) — escape করা হবে না।"""


def get_setting(key, default=""):
    with db() as con:
        row = con.execute("SELECT setting_value FROM bot_settings WHERE setting_key=?", (key,)).fetchone()
    return row[0] if row and row[0] else default


def set_setting(key, value):
    with db() as con:
        con.execute(
            "INSERT INTO bot_settings(setting_key,setting_value) VALUES(?,?) "
            "ON CONFLICT(setting_key) DO UPDATE SET setting_value=excluded.setting_value",
            (key, value)
        )


def delete_settings(*keys):
    with db() as con:
        for key in keys:
            con.execute("DELETE FROM bot_settings WHERE setting_key=?", (key,))


def html_to_plain(text):
    """HTML parse ব্যর্থ হলে fallback: ট্যাগ সরিয়ে লিংকটি লেখার পাশে রাখে।"""
    text = re.sub(r'<a\s+href="([^"]*)"[^>]*>(.*?)</a>', r'\2 (\1)', str(text), flags=re.S)
    text = re.sub(r"<[^>]+>", "", text)
    return html.unescape(text)


RICH_MARK = "<!--rich-->"


def rich_store(message):
    """Admin-এর পাঠানো মেসেজ ফরম্যাটিংসহ (HTML) সেভ করার মান তৈরি করে।"""
    body = getattr(message, "text_html", None) or html.escape(message.text or "", quote=False)
    return RICH_MARK + body


def rich_html(value):
    """সেভ করা মান → পাঠানোর উপযোগী HTML (পুরনো প্লেইন মানও নিরাপদে চলে)।"""
    value = str(value or "")
    if value.startswith(RICH_MARK):
        return value[len(RICH_MARK):]
    return html.escape(value, quote=False)


def rich_plain(value):
    """Admin-কে দেখানোর জন্য ট্যাগ ছাড়া লেখা।"""
    return html_to_plain(rich_html(value))


def _emoji_pattern(mapping):
    """সব ইমোজি একসাথে এক regex-এ (বড় আগে)। FE0F (variation selector) থাকুক বা না থাকুক মিলবে।"""
    keys = sorted({k.replace("\ufe0f", "") for k in mapping if k.replace("\ufe0f", "")},
                  key=len, reverse=True)
    return re.compile("|".join(re.escape(k) + "\ufe0f?" for k in keys))


def _emoji_lookup(mapping):
    return {k.replace("\ufe0f", ""): v for k, v in mapping.items()}


def format_premium_html(text):
    """Admin-এর তৈরি HTML-এ configured Premium Emoji বসায় (ইতিমধ্যে tg-emoji থাকলে ছোঁয় না)।"""
    mapping = get_premium_emoji_map()
    if not mapping:
        return str(text)
    pattern = _emoji_pattern(mapping)
    lookup = _emoji_lookup(mapping)
    # tg-emoji ট্যাগ এবং যেকোনো HTML ট্যাগের ভেতরের অংশ (attribute) ছোঁয়া হবে না
    parts = re.split(r"(<tg-emoji[^>]*>.*?</tg-emoji>|<[^>]+>)", str(text), flags=re.S)
    for i in range(0, len(parts), 2):
        parts[i] = pattern.sub(
            lambda m: f'<tg-emoji emoji-id="{lookup[m.group(0).replace(chr(0xfe0f), "")]}">{m.group(0)}</tg-emoji>', parts[i])
    return "".join(parts)


_LABEL_RE = re.compile(r"^(?P<pre>[^\w<&;]{0,8})(?P<label>[^\W\d_][\w ]{0,28}?):(?=\s|$)")


def _auto_style_line(line):
    if not line.strip():
        return line
    # হেডিং: শুধু বড় হাতের ইংরেজি অক্ষর (যেমন "📥 ALL SUBMISSIONS")
    if ":" not in line and not re.search(r"[a-z0-9]", line) and re.search(r"[A-Z]{3,}", line):
        return fancy_text(line, "bold")
    match = _LABEL_RE.match(line)
    if match:
        return f"{match.group('pre')}<b>{match.group('label')}:</b>{line[match.end():]}"
    return line


def auto_style_plain(escaped):
    """বটের সাধারণ (escape করা) মেসেজে স্বয়ংক্রিয় হেডিং/লেবেল স্টাইল।"""
    return "\n".join(_auto_style_line(line) for line in str(escaped).split("\n"))


def format_premium_text(text):
    """Escape all source text, then replace configured emoji with Telegram HTML tags."""
    if text is None:
        return ""
    escaped = auto_style_plain(html.escape(str(text), quote=False))
    if not get_premium_emoji_map():
        return escaped
    return format_premium_html(escaped)


def parse_premium_emoji_file_ex(content):
    """'emoji = id' লাইন বা শুধু ID-র তালিকা পড়ে। ফেরত: (mapping, ids, invalid_lines, duplicate_emojis)।"""
    result, ids, invalid, dups = {}, [], [], []
    try:
        text = content.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = content.decode("utf-8", errors="ignore")
    for number, raw in enumerate(text.splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith(("#", "//")):
            continue
        if "=" in line or ":" in line:
            left, right = re.split(r"[=:]", line, maxsplit=1)
            found_ids = re.findall(r"(?<!\d)\d{15,25}(?!\d)", right)
            left = left.strip()
            if left and found_ids:
                key = left.replace("\ufe0f", "")
                if any(k.replace("\ufe0f", "") == key for k in result):
                    dups.append(left)  # একই ইমোজি দুইবার: শেষেরটি ধরা হবে
                    for k in [k for k in result if k.replace("\ufe0f", "") == key]:
                        del result[k]
                result[left] = found_ids[0]
                continue
            invalid.append((number, line[:40]))
            continue
        found = re.findall(r"(?<!\d)\d{15,25}(?!\d)", line)
        if found:
            ids.extend(found)
        else:
            invalid.append((number, line[:40]))
    return result, list(dict.fromkeys(ids)), invalid, dups


def parse_premium_emoji_file(content):
    """আগের সিগনেচার (compat)।"""
    result, ids, _invalid, _dups = parse_premium_emoji_file_ex(content)
    return result, ids


async def resolve_premium_emoji_file(bot, content, report=None):
    """Telegram-এর মাধ্যমে ID যাচাই করে mapping বানায়। report dict-এ invalid/dups/rejected ভরে।"""
    explicit, ids, invalid, dups = parse_premium_emoji_file_ex(content)
    resolved = {}
    rejected = []
    if explicit:
        verified = await verify_emoji_ids(bot, explicit)
        resolved.update(verified)
        rejected = [e for e in explicit if e not in verified]
    for start in range(0, len(ids), 100):
        stickers = await bot.get_custom_emoji_stickers(ids[start:start + 100])
        found_ids = set()
        for sticker in stickers:
            custom_id = str(getattr(sticker, "custom_emoji_id", "") or "")
            emoji = getattr(sticker, "emoji", None)
            if custom_id and emoji:
                resolved[emoji] = custom_id
                found_ids.add(custom_id)
        rejected += [i for i in ids[start:start + 100] if i not in found_ids]
    if report is not None:
        report.update(invalid=invalid, dups=dups, rejected=rejected)
    return resolved, len(ids), len(explicit)


# =====================================================================
#   ➕ ADD SINGLE EMOJI — প্রতিটি ইমোজি-রেফারেন্স আলাদাভাবে বদলানো
#   অগ্রাধিকার: Single Emoji Override  >  Premium TXT Mapping  >  মূল ইমোজি
#   সংরক্ষণ: bot_settings টেবিলের 'emoji_overrides' কী (TXT ম্যাপিং 'premium_emojis' অপরিবর্তিত)
# =====================================================================
_EM_BASE = ("[\u00a9\u00ae\u203c\u2049\u2122\u2139\u2194-\u21aa\u231a-\u23ff\u24c2\u25aa-\u25fe"
            "\u2600-\u27bf\u2934\u2935\u2b05-\u2b55\u3030\u303d\u3297\u3299\U0001F000-\U0001FAFF]")
_EM_ELEM = "(?:[\U0001F1E6-\U0001F1FF]{2}|" + _EM_BASE + "\ufe0f?[\U0001F3FB-\U0001F3FF]?)"
_EMOJI_SEQ = re.compile(_EM_ELEM + "(?:\u200d" + _EM_ELEM + ")*")
_ONE_EMOJI = re.compile("(?:" + _EM_ELEM + "(?:\u200d" + _EM_ELEM + ")*)")

_OVR_KEY = "emoji_overrides"
_OVR_CACHE = {"data": None}
_REF_CACHE = {"refs": None}
_EMOJI_READY = [False]  # main()-এ DB তৈরির পরে True হয়; তার আগে (import-সময়ে) কিছুই বদলায় না
_PH_OPEN, _PH_CLOSE = "\ue000", "\ue001"


class NoOverrideHTML(RawHTML):
    """Admin-এর ইমোজি তালিকার মেসেজ: Single-Emoji ওভাররাইড প্রয়োগ হবে না (মূল ইমোজি দেখাবে)।"""


def _nkey(value):
    """তুলনার জন্য স্বাভাবিক রূপ: বোল্ড/ফ্যান্সি ফন্ট → সাধারণ অক্ষর, ছোট হাত, FE0F বাদ।"""
    out = []
    for ch in str(value):
        if ch == "\ufe0f":
            continue
        out.append(unicodedata.normalize("NFKC", ch).casefold())
    return "".join(out)


def _noemoji(value):
    value = _EMOJI_SEQ.sub("", str(value)).replace("\ufe0f", "")
    return re.sub(r"\s+", " ", value).strip()


def get_emoji_overrides():
    if not _EMOJI_READY[0]:
        return {}
    if _OVR_CACHE["data"] is None:
        try:
            raw = get_setting(_OVR_KEY)
        except Exception:
            return {}
        data = {}
        try:
            parsed = json.loads(raw) if raw else {}
        except ValueError:
            parsed = {}
        if isinstance(parsed, dict):
            for key, val in parsed.items():
                if not isinstance(val, dict):
                    continue
                if val.get("kind") == "unicode" and val.get("new"):
                    data[str(key)] = val
                elif val.get("kind") == "custom" and str(val.get("id", "")).isdigit():
                    data[str(key)] = val
        _OVR_CACHE["data"] = data
    return _OVR_CACHE["data"]


def save_emoji_overrides(data):
    set_setting(_OVR_KEY, json.dumps(data, ensure_ascii=False))
    _OVR_CACHE["data"] = data


def _scan_source_emoji_refs():
    """সোর্স কোড (AST) স্ক্যান করে প্রতিটি ইমোজি-রেফারেন্স ও তার লোকেশন বের করে।"""
    src = Path(__file__).read_text(encoding="utf-8")
    tree = ast.parse(src)
    skip_names = {"_DANGER_MARKERS", "_SUCCESS_MARKERS", "DEFAULT_PREMIUM_EMOJI_IDS", "_EMOJI_REGEX",
                  "LEGACY_MENU_ALIASES", "_EM_BASE", "_EM_ELEM", "_EMOJI_SEQ", "_ONE_EMOJI"}
    skip = set()

    def mark(node):
        for child in ast.walk(node):
            skip.add(id(child))

    reply_literals = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            body = node.body
            if (body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant)
                    and isinstance(body[0].value.value, str)):
                mark(body[0])
        elif isinstance(node, ast.Assign):
            if any(isinstance(t, ast.Name) and t.id in skip_names for t in node.targets):
                mark(node)
        elif isinstance(node, ast.Compare):
            mark(node)  # রাউটিং তুলনা (text == "...") — দেখানোর লেখা নয়
        elif isinstance(node, ast.Call):
            fn = node.func
            if isinstance(fn, ast.Attribute) and isinstance(fn.value, ast.Name) and fn.value.id in ("log", "logging"):
                mark(node)
            if isinstance(fn, ast.Name) and fn.id in ("ColorButton", "KeyboardButton") and node.args:
                first = node.args[0]
                if isinstance(first, ast.Constant) and isinstance(first.value, str):
                    reply_literals.add(first.value)

    refs = {}
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Constant) and isinstance(node.value, str)) or id(node) in skip:
            continue
        text = node.value
        for m in _EMOJI_SEQ.finditer(text):
            before = _nkey(text[max(0, m.start() - 12):m.start()])
            after = _nkey(text[m.end():m.end() + 24])
            nemoji = _nkey(m.group(0))
            rid = hashlib.sha1(f"{nemoji}|{before}|{after}".encode("utf-8")).hexdigest()[:8]
            label = text[m.start():m.end() + 26].replace("\n", " ").strip()
            ref = refs.setdefault(rid, {"id": rid, "emoji": m.group(0), "nemoji": nemoji,
                                        "before": before, "after": after, "label": label,
                                        "lines": [], "first": node.lineno})
            if node.lineno not in ref["lines"]:
                ref["lines"].append(node.lineno)
            ref["first"] = min(ref["first"], node.lineno)
    ordered = sorted(refs.values(), key=lambda r: (r["first"], r["id"]))
    return ordered, reply_literals


def get_emoji_refs():
    if _REF_CACHE["refs"] is None:
        refs, literals = _scan_source_emoji_refs()
        by_emoji = {}
        for ref in refs:
            by_emoji.setdefault(ref["nemoji"], []).append(ref)
        _REF_CACHE.update(refs=refs, by_emoji=by_emoji, by_id={r["id"]: r for r in refs},
                          reply_literals=literals)
    return _REF_CACHE["refs"]


def _match_overrides(text):
    """text-এর যেসব ইমোজি ওভাররাইডেড রেফারেন্সের সাথে মেলে: [(start, end, override)]।"""
    overrides = get_emoji_overrides()
    if not overrides or not text:
        return []
    get_emoji_refs()
    by_emoji = _REF_CACHE["by_emoji"]
    found = []
    for m in _EMOJI_SEQ.finditer(text):
        cands = by_emoji.get(_nkey(m.group(0)))
        if not cands or not any(c["id"] in overrides for c in cands):
            continue
        nb, na = _nkey(text[:m.start()]), _nkey(text[m.end():])
        best = None
        for cand in cands:
            if nb.endswith(cand["before"]) and na.startswith(cand["after"]):
                score = len(cand["before"]) + len(cand["after"])
                if best is None or score > best[0]:
                    best = (score, cand)
        if best and best[1]["id"] in overrides:
            found.append((m.start(), m.end(), overrides[best[1]["id"]]))
    return found


def _override_text(plain):
    """মেসেজ টেক্সটে ওভাররাইডেড ইমোজির জায়গায় প্লেসহোল্ডার বসায়। ফেরত: (text, subs)।"""
    matches = _match_overrides(plain)
    if not matches:
        return plain, []
    out, subs, last = [], [], 0
    for start, end, ovr in matches:
        out.append(plain[last:start])
        out.append(_PH_OPEN + chr(0xE100 + len(subs)) + _PH_CLOSE)
        subs.append(ovr)
        last = end
    out.append(plain[last:])
    return "".join(out), subs


def _resolve_subs(text, subs, html_mode):
    if not subs:
        return text

    def repl(m):
        idx = ord(m.group(1)) - 0xE100
        if not 0 <= idx < len(subs):
            return ""
        ovr = subs[idx]
        if ovr["kind"] == "unicode":
            return ovr["new"]
        fallback = ovr.get("fallback") or "⭐"
        if html_mode:
            return f'<tg-emoji emoji-id="{ovr["id"]}">{fallback}</tg-emoji>'
        return fallback

    return re.sub(_PH_OPEN + "([\ue100-\ue1ff])" + _PH_CLOSE, repl, text)


def apply_button_override(text):
    """বাটনের লেখায় ওভাররাইড বসায়। ফেরত: (নতুন লেখা, icon_custom_emoji_id বা None)।"""
    try:
        matches = _match_overrides(text)
    except Exception:
        log.exception("Button emoji override failed")
        return text, None
    if not matches:
        return text, None
    parts, last, icon = [], 0, None
    for start, end, ovr in matches:
        parts.append(text[last:start])
        if ovr["kind"] == "unicode":
            parts.append(ovr["new"])
        elif ovr.get("btn_icon") and icon is None:
            icon = ovr["id"]  # বাটনের আইকন স্লটে বসবে; লেখা থেকে ইমোজি সরে যায়
        else:
            parts.append(ovr.get("fallback") or text[start:end])
        last = end
    parts.append(text[last:])
    new_text = re.sub(r"\s{2,}", " ", "".join(parts)).strip() if icon else "".join(parts)
    if icon and not new_text:
        # শুধু ইমোজিই বাটনের লেখা হলে আইকন-ফাঁকা বাটন বৈধ নয় → সাধারণ ইমোজি fallback
        new_text = next((o.get("fallback") for _s, _e, o in matches if o["kind"] == "custom"), None) or text
        icon = None
    return new_text, icon


def _reply_candidates():
    get_emoji_refs()
    legacy = set(LEGACY_MENU_ALIASES)
    ordered = [t for t in MENU_BUTTON_TEXTS if t not in legacy] + [t for t in MENU_BUTTON_TEXTS if t in legacy]
    return ordered + sorted(_REF_CACHE.get("reply_literals", set()))


def canonical_reply_text(text):
    """ইমোজি বদলানো Reply-Keyboard বাটনের চাপ → মূল লেবেলে ফিরিয়ে আনে (রাউটিং অক্ষুণ্ণ রাখতে)।"""
    try:
        if not get_emoji_overrides() or not text:
            return text
        if text in MENU_BUTTON_TEXTS or text == "❌ Cancel":
            return text
        key = _noemoji(text)
        if not key:
            return text
        for canon in _reply_candidates():
            if _noemoji(canon) == key and _match_overrides(canon):
                return canon
    except Exception:
        log.exception("Reply button canonicalisation failed")
    return text


# ---------------- Admin UI ----------------
_EMOJI_PAGE_SIZE = 20


def _eb(text, callback_data):
    """ইমোজি-তালিকার বাটন (ওভাররাইড/ফ্যান্সি ফন্ট প্রয়োগ ছাড়া)।"""
    return _TGInlineKeyboardButton(text, callback_data=callback_data)


def _ref_status(ref):
    ovr = get_emoji_overrides().get(ref["id"])
    if not ovr:
        return ""
    if ovr["kind"] == "unicode":
        return f" → {ovr['new']}"
    return f" → ✨{ovr.get('fallback') or ''}(Premium)"


async def _emoji_show_list(q, context, page):
    refs = get_emoji_refs()
    total_pages = max(1, (len(refs) + _EMOJI_PAGE_SIZE - 1) // _EMOJI_PAGE_SIZE)
    page = max(0, min(page, total_pages - 1))
    chunk = refs[page * _EMOJI_PAGE_SIZE:(page + 1) * _EMOJI_PAGE_SIZE]
    lines = [f"➕ <b>ADD SINGLE EMOJI</b>  ({page + 1}/{total_pages})",
             f"মোট ইমোজি-রেফারেন্স: {len(refs)}টি | পরিবর্তিত: {len(get_emoji_overrides())}টি",
             "যেটি বদলাতে চান সেটির বাটন চাপুন:\n"]
    for n, ref in enumerate(chunk, 1):
        lines.append(f"{n}. {ref['emoji']} {html.escape(ref['label'][:34])} · L{ref['first']}"
                     f"{'×' + str(len(ref['lines'])) if len(ref['lines']) > 1 else ''}"
                     f"{html.escape(_ref_status(ref))}")
    rows, row = [], []
    for n, ref in enumerate(chunk, 1):
        row.append(_eb(f"{n} {ref['emoji']}", f"adm:emojiref:{ref['id']}:{page}"))
        if len(row) == 4:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    nav = []
    if page > 0:
        nav.append(_eb("⬅️ Previous", f"adm:emojiadd:{page - 1}"))
    if page + 1 < total_pages:
        nav.append(_eb("Next ➡️", f"adm:emojiadd:{page + 1}"))
    if nav:
        rows.append(nav)
    rows.append([_eb("⬅️ System", "adm:system")])
    await q.edit_message_text(NoOverrideHTML("\n".join(lines)), reply_markup=InlineKeyboardMarkup(rows))


async def emoji_editor_callback(update, context, q, data):
    """adm:emojiadd / emojiref / emojiset / emojiglob / emojiundo — শুধু Admin (admin_callback আগেই চেক করে)।"""
    action = data[1]
    refs = get_emoji_refs()
    by_id = _REF_CACHE["by_id"]
    page = 0
    if action == "emojiadd":
        if context.user_data.get("admin_action") == "emoji_single_wait":
            clear_admin_state(context)
        page = int(data[2]) if len(data) > 2 and data[2].isdigit() else 0
        await _emoji_show_list(q, context, page)
        return
    ref_id = data[2] if len(data) > 2 else ""
    ref = by_id.get(ref_id)
    if not ref:
        await q.edit_message_text("এই ইমোজি-রেফারেন্সটি আর পাওয়া যায়নি।",
                                  reply_markup=InlineKeyboardMarkup([[_eb("⬅️ Emoji List", "adm:emojiadd:0")]]))
        return
    if action == "emojiref":
        page = int(data[3]) if len(data) > 3 and data[3].isdigit() else 0
        if context.user_data.get("admin_action") == "emoji_single_wait":
            clear_admin_state(context)
        ovr = get_emoji_overrides().get(ref_id)
        same = [r for r in refs if r["nemoji"] == ref["nemoji"]]
        in_map = any(k.replace("\ufe0f", "") == ref["nemoji"] for k in get_premium_emoji_map())
        cur = "মূল ইমোজি"
        if ovr:
            cur = (f"সাধারণ ইমোজি {ovr['new']}" if ovr["kind"] == "unicode" else
                   f"Premium ইমোজি (ID {ovr['id']}), বাটনে আইকন: {'হ্যাঁ' if ovr.get('btn_icon') else 'না (সাধারণ ইমোজি দেখাবে)'}")
        text = (f"{ref['emoji']} <b>EMOJI DETAILS</b>\n\n"
                f"🆔 Ref: <code>{ref_id}</code>\n"
                f"📍 লেখা: {html.escape(ref['label'])}\n"
                f"📄 সোর্স লাইন: {', '.join('L' + str(n) for n in ref['lines'][:10])}\n"
                f"🔁 একই ইমোজি মোট {len(same)}টি জায়গায় আছে\n"
                f"💎 TXT Premium ম্যাপিং: {'আছে' if in_map else 'নেই'}\n"
                f"🎯 বর্তমান: {html.escape(cur)}")
        rows = [[_eb("✏️ Change Only This One", f"adm:emojiset:{ref_id}:{page}")],
                [_eb("🌐 Replace Everywhere (All Same Emoji)", f"adm:emojiglob:{ref_id}:{page}")]]
        if ovr:
            rows.append([_eb("♻️ Restore Original", f"adm:emojiundo:{ref_id}:{page}")])
        rows.append([_eb("⬅️ Emoji List", f"adm:emojiadd:{page}")])
        await q.edit_message_text(NoOverrideHTML(text), reply_markup=InlineKeyboardMarkup(rows))
        return
    page = int(data[3]) if len(data) > 3 and data[3].isdigit() else 0
    if action in ("emojiset", "emojiglob"):
        clear_admin_state(context)
        context.user_data["admin_action"] = "emoji_single_wait"
        context.user_data["emoji_edit_ref"] = ref_id
        context.user_data["emoji_edit_scope"] = "all" if action == "emojiglob" else "one"
        context.user_data["emoji_edit_page"] = page
        scope_note = ("⚠️ এটি বটের সব জায়গায় এই ইমোজির সব রেফারেন্স বদলাবে।\n\n"
                      if action == "emojiglob" else "")
        await q.edit_message_text(
            NoOverrideHTML(f"{scope_note}{ref['emoji']} Please send the new emoji you want to use for this selection.\n\n"
                           "একটি সাধারণ ইমোজি অথবা একটি Premium Custom Emoji পাঠান (শুধু ইমোজিটি, অন্য লেখা ছাড়া)।"),
            reply_markup=InlineKeyboardMarkup([
                [_eb("⬅️ Back", f"adm:emojiref:{ref_id}:{page}")],
                [_eb("❌ Cancel", "adm:cancelinput")]]))
        return
    if action == "emojiundo":
        data_now = dict(get_emoji_overrides())
        if data_now.pop(ref_id, None) is not None:
            save_emoji_overrides(data_now)
        await emoji_editor_callback(update, context, q, ["adm", "emojiref", ref_id, str(page)])
        return


async def emoji_single_receive(update, context):
    """Admin-এর পাঠানো নতুন ইমোজি যাচাই → সেভ → প্রিভিউ।"""
    msg = update.message
    refs = get_emoji_refs()
    ref = _REF_CACHE["by_id"].get(context.user_data.get("emoji_edit_ref"))
    scope = context.user_data.get("emoji_edit_scope", "one")
    page = context.user_data.get("emoji_edit_page", 0)
    cancel_kb = InlineKeyboardMarkup([[_eb("❌ Cancel", "adm:cancelinput")]])
    if not ref:
        clear_admin_state(context)
        await msg.reply_text("নির্বাচিত ইমোজি পাওয়া যায়নি। আবার চেষ্টা করুন।")
        return True
    text = (msg.text or "").strip()
    entities = [e for e in (msg.entities or []) if getattr(e, "type", "") == "custom_emoji"]
    new = None
    if entities:
        if len(entities) != 1 or msg.parse_entity(entities[0]).strip() != text:
            await msg.reply_text("❌ শুধু একটি Premium ইমোজি পাঠান (অন্য লেখা/একাধিক ইমোজি নয়)।", reply_markup=cancel_kb)
            return True
        cid = str(entities[0].custom_emoji_id)
        try:
            stickers = await context.bot.get_custom_emoji_stickers([cid])
        except TelegramError:
            await msg.reply_text("❌ Telegram থেকে ইমোজিটি যাচাই করা যায়নি। একটু পরে আবার পাঠান।", reply_markup=cancel_kb)
            return True
        if not stickers:
            await msg.reply_text("❌ এই Premium ইমোজি ID Telegram-এ পাওয়া যায়নি।", reply_markup=cancel_kb)
            return True
        fallback = getattr(stickers[0], "emoji", None) or text
        new = {"kind": "custom", "id": cid, "fallback": fallback, "btn_icon": False}
    elif text and _ONE_EMOJI.fullmatch(text):
        new = {"kind": "unicode", "new": text}
    else:
        await msg.reply_text("❌ বৈধ ইমোজি নয়। শুধু একটি সাধারণ ইমোজি বা একটি Premium Custom Emoji পাঠান।",
                             reply_markup=cancel_kb)
        return True

    targets = [r for r in refs if r["nemoji"] == ref["nemoji"]] if scope == "all" else [ref]
    # বাটনে Premium আইকন সমর্থিত কিনা — সত্যিকারের Telegram পরীক্ষায় যাচাই (ব্যর্থ হলে সাধারণ ইমোজি fallback)
    btn_ok = None
    if new["kind"] == "custom":
        try:
            await context.bot.send_message(
                msg.chat_id, NoOverrideHTML("🧪 Button icon test"),
                reply_markup=InlineKeyboardMarkup([[_TGInlineKeyboardButton(
                    "Preview", callback_data="adm:emojiadd:0",
                    api_kwargs={"icon_custom_emoji_id": new["id"]})]]))
            btn_ok = True
        except TelegramError as exc:
            btn_ok = False
            log.warning("Button custom-emoji icon not accepted by Telegram: %s", exc)
        new["btn_icon"] = bool(btn_ok)
    current = dict(get_emoji_overrides())
    for target in targets:
        current[target["id"]] = dict(new)
    save_emoji_overrides(current)
    clear_admin_state(context)

    if new["kind"] == "custom":
        shown = f'<tg-emoji emoji-id="{new["id"]}">{html.escape(new["fallback"])}</tg-emoji>'
        note = ("✅ বাটনেও Premium আইকন কাজ করছে।" if btn_ok else
                "ℹ️ Telegram এই বটের বাটনে Premium আইকন নেয়নি, তাই বাটনে সাধারণ ইমোজি "
                f"{html.escape(new['fallback'])} দেখাবে; মেসেজ টেক্সটে Premium চেষ্টা হবে।")
    else:
        shown, note = html.escape(new["new"]), "✅ সাধারণ ইমোজি সেট হয়েছে।"
    await msg.reply_text(
        NoOverrideHTML(f"✅ <b>Emoji updated successfully!</b>\n\nনতুন ইমোজি: {shown}\n"
                       f"পরিবর্তিত রেফারেন্স: {len(targets)}টি\n{note}"),
        reply_markup=InlineKeyboardMarkup([[_eb("⬅️ Emoji List", f"adm:emojiadd:{page}")],
                                            [_eb("⚙️ System", "adm:system")]]))
    return True


_PREMIUM_BUSY = contextvars.ContextVar("premium_busy", default=False)
_TAG_RE = re.compile(r"</?(?:b|strong|i|em|u|s|strike|del|code|pre|a|tg-emoji|blockquote|tg-spoiler|span)(?:\s|>|/)", re.I)


def install_premium_message_wrappers():
    """বটের মেসেজ/ক্যাপশনে ফরম্যাটিং কেন্দ্রীয়ভাবে প্রয়োগ করে।

    - ভেতরের (nested) কলে আর দ্বিতীয়বার ফরম্যাট হয় না।
    - Telegram HTML/ইমোজি রিজেক্ট করলে ট্যাগ ছাড়া পরিষ্কার plain text পাঠায় (আগে `<b>` লেখা দেখা যেত)।
    - 502 Bad Gateway / Timeout হলে নিজে থেকে কয়েকবার retry করে; তারপর ইমোজি ছাড়া পাঠায়।
    """
    targets = [
        (Bot, "send_message", "text"),
        (Bot, "send_photo", "caption"),
        (Bot, "send_document", "caption"),
        (Bot, "send_video", "caption"),
        (Bot, "edit_message_text", "text"),
        (Bot, "edit_message_caption", "caption"),
    ]
    for cls, method_name, field in targets:
        original = getattr(cls, method_name, None)
        if original is None or getattr(original, "_premium_wrapped", False):
            continue
        signature = inspect.signature(original)

        async def wrapped(self, *args, __original=original, __signature=signature,
                          __field=field, **kwargs):
            bound = __signature.bind_partial(self, *args, **kwargs)
            source = bound.arguments.get(__field)
            if source is None:
                return await __original(self, *args, **kwargs)
            plain = str(source)
            is_raw = isinstance(source, RawHTML)
            subs = []
            if not isinstance(source, NoOverrideHTML):
                try:
                    plain, subs = _override_text(plain)  # Add Single Emoji (অগ্রাধিকার: TXT ম্যাপের উপরে)
                except Exception:
                    log.exception("Single emoji override failed; skipping")
                    plain, subs = str(source), []
            has_html = bool(_TAG_RE.search(plain))
            has_parse = "parse_mode" in __signature.parameters
            try:
                bound.arguments[__field] = _resolve_subs(
                    format_premium_html(plain) if (is_raw or has_html)
                    else format_premium_text(plain), subs, True)
                if has_parse:
                    bound.arguments["parse_mode"] = ParseMode.HTML
            except Exception:
                log.exception("Premium emoji formatting failed; sending safe plain text")
                bound.arguments[__field] = _resolve_subs(
                    html_to_plain(plain) if (is_raw or has_html) else plain, subs, False)
                if has_parse:
                    bound.arguments["parse_mode"] = None

            safe_plain = _resolve_subs(html_to_plain(plain) if (is_raw or has_html) else plain, subs, False)
            fallback_kwargs = None

            for attempt in range(4):
                try:
                    return await __original(*bound.args, **bound.kwargs)
                except BadRequest as exc:
                    error_text = str(exc).lower()
                    if "message is not modified" in error_text:
                        raise
                    if not any(k in error_text for k in (
                            "custom_emoji", "parse entities", "end tag", "start tag", "entity", "entities")):
                        raise
                    retry = __signature.bind_partial(self, *args, **kwargs)
                    retry.arguments[__field] = safe_plain
                    if has_parse:
                        retry.arguments["parse_mode"] = None
                    log.warning("Premium emoji/HTML rejected by Telegram; retrying plain text: %s", exc)
                    return await __original(*retry.args, **retry.kwargs)
                except (NetworkError, TimedOut) as exc:
                    # 502 Bad Gateway ইত্যাদি সাময়িক সমস্যা: অপেক্ষা করে আবার চেষ্টা
                    if attempt >= 3:
                        raise
                    delay = 1.5 * (attempt + 1)
                    log.warning("Telegram network error (%s); retry %d/3 in %.1fs", exc, attempt + 1, delay)
                    await asyncio.sleep(delay)
                    if attempt == 1:
                        # দুইবার ব্যর্থ হলে সবচেয়ে হালকা (ইমোজি/HTML ছাড়া) মেসেজে নামি
                        bound = __signature.bind_partial(self, *args, **kwargs)
                        bound.arguments[__field] = safe_plain
                        if has_parse:
                            bound.arguments["parse_mode"] = None

        wrapped._premium_wrapped = True
        functools.update_wrapper(wrapped, original)
        wrapped._premium_wrapped = True
        setattr(cls, method_name, wrapped)


def money(value):
    return f"{value:,.2f} {CURRENCY}"


def main_keyboard(user_id=None):
    if _EMOJI_READY[0] and get_emoji_overrides():
        # Add Single Emoji: মূল লেবেল থেকে নতুন করে বাটন বানানো (রং বজায় থাকে)
        menu = [[ColorButton(b.text, (getattr(b, "api_kwargs", None) or {}).get("style", "primary"))
                 for b in row] for row in MAIN_MENU]
    else:
        menu = [list(row) for row in MAIN_MENU]
    if user_id is not None and is_admin(user_id):
        menu.append([ColorButton(BTN_ADMIN, "danger")])
    return ReplyKeyboardMarkup(menu, resize_keyboard=True)


def register_user(user):
    with db() as con:
        con.execute(
            """INSERT INTO users(user_id,username,first_name,joined_at)
               VALUES(?,?,?,?)
               ON CONFLICT(user_id) DO UPDATE SET username=excluded.username,
               first_name=excluded.first_name""",
            (user.id, user.username or "", user.first_name or "", now())
        )


def get_categories(open_only=False):
    with db() as con:
        q = "SELECT * FROM categories WHERE is_deleted=0"
        if open_only:
            q += " AND is_open=1"
        return con.execute(q + " ORDER BY id").fetchall()


def get_payment_methods(open_only=False):
    with db() as con:
        q = "SELECT * FROM payment_methods"
        if open_only:
            q += " WHERE is_open=1"
        return con.execute(q + " ORDER BY id").fetchall()


def clear_submission_state(context):
    for key in SUBMISSION_STATE_KEYS:
        context.user_data.pop(key, None)


def clear_admin_state(context):
    for key in ADMIN_STATE_KEYS:
        context.user_data.pop(key, None)


# Reply Keyboard-এর সব বাটন — এগুলো কখনোই "ইনপুট" হিসেবে গণ্য হবে না
MENU_BUTTON_TEXTS = ({button.text for row in MAIN_MENU for button in row}
                     | {BTN_ADMIN} | set(LEGACY_MENU_ALIASES))


def reset_all_input_state(context):
    """ইউজার ও অ্যাডমিন — সব অপেক্ষমাণ ইনপুট স্টেট একসাথে পরিষ্কার করে।"""
    clear_admin_state(context)
    clear_submission_state(context)


def count_file_items(path: Path) -> int:
    """Count non-empty data rows/lines; excludes a header only when admin later confirms."""
    suffix = path.suffix.lower()
    try:
        if suffix == ".txt":
            return sum(1 for line in path.read_text(encoding="utf-8-sig", errors="ignore").splitlines()
                       if line.strip())
        if suffix == ".csv":
            import csv
            with path.open("r", encoding="utf-8-sig", errors="ignore", newline="") as f:
                return sum(1 for row in csv.reader(f) if any(str(x).strip() for x in row))
        if suffix == ".xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            n = sum(1 for row in ws.iter_rows(values_only=True)
                    if any(v is not None and str(v).strip() for v in row))
            wb.close()
            return n
    except Exception:
        log.exception("Could not count rows in uploaded file")
    return 0


# ========================================================
# ================  REUSABLE ADMIN KEYBOARDS  ============
# ========================================================

def submission_action_keyboard(sid):
    """Legacy /pending ও /review command-এর keyboard অপরিবর্তিত রাখে।"""
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📄 Review", callback_data=f"adm:review:{sid}")],
        [InlineKeyboardButton("✅ Approve", callback_data=f"adm:approve:{sid}")],
        [InlineKeyboardButton("❌ Reject", callback_data=f"adm:reject:{sid}")],
    ])


def category_submission_keyboard(sid):
    """Category File Management-এর জন্য শুধু Approve / Reject।"""
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Approve", callback_data=f"adm:approve:{sid}"),
        InlineKeyboardButton("❌ Reject", callback_data=f"adm:reject:{sid}"),
    ]])


def approved_submission_keyboard(sid):
    """Approved Submission-এর Report ও Reject action।"""
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("📤 Send Report", callback_data=f"adm:report:{sid}"),
        InlineKeyboardButton("❌ Reject", callback_data=f"adm:reject:{sid}"),
    ]])


def admin_home_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📂 Manage Categories", callback_data="adm:cats:0"),
         InlineKeyboardButton("📥 Submissions", callback_data="adm:subs")],
        [InlineKeyboardButton("👥 Users & Balances", callback_data="adm:users:0"),
         InlineKeyboardButton("📊 Statistics", callback_data="adm:stats")],
        [InlineKeyboardButton("➕ Add Category", callback_data="adm:addcat"),
         InlineKeyboardButton("❌ Close", callback_data="adm:close")],
        [InlineKeyboardButton("💳 Payment Methods", callback_data="adm:pms:0")],
        [InlineKeyboardButton("📢 Broadcast", callback_data="adm:broadcast")],
        [InlineKeyboardButton("⚙️ System", callback_data="adm:system")],
    ])


def get_force_join_channels():
    with db() as con:
        return con.execute("SELECT * FROM force_join_channels ORDER BY id").fetchall()


def force_join_keyboard(channels):
    rows = []
    for channel in channels:
        title = channel["title"] or "চ্যানেল/গ্রুপ"
        rows.append([InlineKeyboardButton(f"📢 Join {title}", url=channel["join_link"])])
    rows.append([InlineKeyboardButton("✅ Joined / Verify", callback_data="fj:verify")])
    return InlineKeyboardMarkup(rows)


async def check_force_join_membership(bot, user_id):
    """Return (all_joined, failed_channels). API errors fail closed for safety."""
    channels = get_force_join_channels()
    failed = []
    for channel in channels:
        try:
            member = await bot.get_chat_member(
                chat_id=int(channel["chat_id"]), user_id=int(user_id)
            )
            status = str(member.status).lower()
            is_member = status in ("creator", "administrator", "member")
            if status == "restricted":
                is_member = bool(getattr(member, "is_member", False))
            if not is_member:
                failed.append(channel)
        except TelegramError:
            log.exception("Force Join membership check failed for chat_id=%s", channel["chat_id"])
            failed.append(channel)
    return not failed, failed


async def send_force_join_prompt(message, channels):
    await message.reply_text(
        "🔒 বট ব্যবহার করতে নিচের সব চ্যানেল/গ্রুপে Join করুন।\n"
        "সবগুলোতে Join করার পর নিচের Joined / Verify বাটনে চাপ দিন।",
        reply_markup=force_join_keyboard(channels)
    )


async def force_join_guard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Gate private user interactions until all configured channels are joined."""
    user = update.effective_user
    chat = update.effective_chat
    if not user or not chat or chat.type != "private" or is_admin(user.id):
        return

    channels = get_force_join_channels()
    if not channels:
        return

    query = update.callback_query
    if query and query.data == "fj:verify":
        await query.answer()
        joined, failed = await check_force_join_membership(context.bot, user.id)
        if joined:
            try:
                await query.edit_message_text("✅ Membership verified! এখন বট ব্যবহার করতে পারবেন।")
            except TelegramError:
                pass
            register_user(user)
            clear_submission_state(context)
            await context.bot.send_message(
                chat_id=chat.id, text=get_welcome_message(),
                reply_markup=main_keyboard(user.id)
            )
        else:
            try:
                await query.edit_message_text(
                    "❌ এখনো সব চ্যানেল/গ্রুপে Join করা হয়নি। বাকি লিংকগুলোতে Join করে আবার Verify করুন।",
                    reply_markup=force_join_keyboard(failed)
                )
            except TelegramError:
                await send_force_join_prompt(query.message, failed)
        raise ApplicationHandlerStop

    joined, failed = await check_force_join_membership(context.bot, user.id)
    if joined:
        return

    if update.message:
        await send_force_join_prompt(update.message, failed)
    elif query:
        await query.answer("বট ব্যবহার করতে আগে সব চ্যানেল/গ্রুপে Join করুন।", show_alert=True)
        await context.bot.send_message(
            chat_id=chat.id,
            text="🔒 বট ব্যবহার করতে আগে সব চ্যানেল/গ্রুপে Join করুন।",
            reply_markup=force_join_keyboard(failed)
        )
    raise ApplicationHandlerStop


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    register_user(user)
    reset_all_input_state(context)
    await update.message.reply_text(
        get_welcome_message(),
        reply_markup=main_keyboard(user.id)
    )


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Use the menu to submit a number-list file, view your balance and payment history, "
        "or contact support.\n\n"
        "Admin হলে মেনুতে 🛠 Admin Panel বাটন দেখা যাবে।",
        reply_markup=main_keyboard(update.effective_user.id)
    )


async def submit_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_submission_state(context)
    cats = get_categories(open_only=True)
    if not cats:
        await update.message.reply_text("দুঃখিত, বর্তমানে কোনো ক্যাটাগরি খোলা নেই।")
        return
    rows = [[InlineKeyboardButton(f"🟢 {c['name']} — {money(c['rate'])}/valid", callback_data=f"cat:{c['id']}")]
            for c in cats]
    rows.append([InlineKeyboardButton("❌ Cancel", callback_data="cancel")])
    await update.message.reply_text(
        RawHTML("📂 <b>𝐒𝐞𝐥𝐞𝐜𝐭 𝐜𝐚𝐭𝐞𝐠𝐨𝐫𝐲</b>\n\n"
                "<blockquote> 𝐈𝐧 𝐰𝐡𝐢𝐜𝐡 𝐜𝐚𝐭𝐞𝐠𝐨𝐫𝐲 𝐰𝐢𝐥𝐥 𝐲𝐨𝐮 𝐬𝐮𝐛𝐦𝐢𝐭 𝐭𝐡𝐞 𝐟𝐢𝐥𝐞? 𝐏𝐥𝐞𝐚𝐬𝐞 𝐜𝐡𝐨𝐨𝐬𝐞 𝐟𝐫𝐨𝐦 𝐭𝐡𝐞 𝐥𝐢𝐬𝐭 𝐛𝐞𝐥𝐨𝐰.।</blockquote>"),
        reply_markup=InlineKeyboardMarkup(rows))


async def category_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    # ❌ Cancel
    if q.data == "cancel":
        clear_submission_state(context)
        await q.edit_message_text("Submission cancelled.")
        return

    # 💳 Payment Method নির্বাচন
    if q.data.startswith("pm:"):
        pm_id = int(q.data.split(":")[1])
        with db() as con:
            pm = con.execute(
                "SELECT * FROM payment_methods WHERE id=? AND is_open=1", (pm_id,)
            ).fetchone()
        if not pm:
            await q.edit_message_text("এই পেমেন্ট মেথডটি বর্তমানে বন্ধ। আবার Submit File চাপুন।")
            return
        context.user_data["selected_payment_method"] = pm["id"]
        context.user_data["awaiting_payment_details"] = True
        hint = f"\nℹ️ Admin-এর তথ্য: {pm['account_info']}" if pm["account_info"] else ""
        await q.edit_message_text(
            f"💳 Payment Method: {pm['name']}{hint}\n\n"
            "এই মেথডে আপনার নিজের অ্যাকাউন্ট নম্বর/ID লিখুন\n"
            "(যেমন: 01XXXXXXXXX অথবা Email/ID)।\n\n"
            "বাতিল করতে নিচের ❌ Cancel বাটন চাপুন।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("❌ Cancel", callback_data="cancel")]
            ])
        )
        return

    # 📂 Category নির্বাচন
    cat_id = int(q.data.split(":")[1])
    with db() as con:
        cat = con.execute("SELECT * FROM categories WHERE id=? AND is_open=1 AND is_deleted=0", (cat_id,)).fetchone()
    if not cat:
        await q.edit_message_text("এই ক্যাটাগরিটি বর্তমানে বন্ধ।")
        return

    context.user_data["selected_category"] = cat_id
    context.user_data.pop("selected_payment_method", None)
    context.user_data.pop("selected_payment_details", None)
    context.user_data.pop("awaiting_payment_details", None)

    methods = get_payment_methods(open_only=True)
    if not methods:
        context.user_data.pop("selected_category", None)
        await q.edit_message_text(
            f"📂 Category: {cat['name']}\n\n"
            "দুঃখিত, আপনার Admin এখনো কোনো Payment Method খোলেননি।\n"
            "অনুগ্রহ করে Support-এ যোগাযোগ করুন।"
        )
        return

    rows = [[InlineKeyboardButton(f"💳 {m['name']}", callback_data=f"pm:{m['id']}")] for m in methods]
    rows.append([InlineKeyboardButton("❌ Cancel", callback_data="cancel")])
    notice_html = (f"\n\n📋 <b>Notice</b>\n<blockquote>{rich_html(cat['notice'])}</blockquote>"
                   if cat['notice'] else "")
    await q.edit_message_text(
        RawHTML(f"📂 <b>Category:</b> {html.escape(cat['name'])}{notice_html}\n\n"
                "💳 <b>পেমেন্ট মেথড নির্বাচন করুন:</b>"),
        reply_markup=InlineKeyboardMarkup(rows)
    )


async def handle_payment_details(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = (update.message.text or "").strip()
    if not raw:
        await update.message.reply_text("পেমেন্ট ডিটেইলস খালি রাখা যাবে না। আবার লিখুন।")
        return
    if len(raw) > 100:
        await update.message.reply_text("পেমেন্ট ডিটেইলস অনেক বড় (সর্বোচ্চ ১০০ অক্ষর)। আবার লিখুন।")
        return

    context.user_data["selected_payment_details"] = raw
    context.user_data["awaiting_payment_details"] = False

    default_submit_note = (
        "✅ <b>পেমেন্ট ডিটেইলস সংরক্ষিত হয়েছে!</b>\n\n"
        "📎 এখন আপনার নাম্বার লিস্টের <b>XLSX, CSV অথবা TXT</b> ফাইল পাঠান।\n"
        "💬 চাইলে ফাইলের Caption-এ অ্যাডমিনের জন্য মেসেজ লিখতে পারেন <i>(ঐচ্ছিক)</i>।\n\n"
        "<blockquote>🔒 শুধু নাম্বার-সংক্রান্ত ডেটা পাঠাবেন; পাসওয়ার্ড, কুকি বা 2FA তথ্য পাঠাবেন না।</blockquote>"
    )
    cat_id = context.user_data.get("selected_category")
    with db() as con:
        cat_note = con.execute(
            "SELECT submit_note FROM categories WHERE id=?", (cat_id,)
        ).fetchone() if cat_id else None
    submit_note = (rich_html(cat_note["submit_note"])
                   if cat_note and cat_note["submit_note"] else default_submit_note)
    await update.message.reply_text(
        RawHTML(submit_note),
        reply_markup=main_keyboard(update.effective_user.id)
    )


def get_broadcast_user_ids():
    with db() as con:
        return [r[0] for r in con.execute("SELECT user_id FROM users ORDER BY user_id").fetchall()]


async def run_broadcast(application, admin_id, from_chat_id, message_id, user_ids):
    """copy_message দিয়ে হুবহু কপি পাঠায় (লেখা, ফরম্যাটিং, লিংক, ইমোজি, ছবি, ভিডিও, ফাইল, ভয়েস,
    স্টিকার সব কিছু) — 'Forwarded from' ট্যাগ ছাড়া। ব্যাকগ্রাউন্ডে চলে, বট আটকায় না।"""
    bot = application.bot
    sent = blocked = failed = 0
    for uid in user_ids:
        for attempt in range(2):
            try:
                await bot.copy_message(chat_id=uid, from_chat_id=from_chat_id, message_id=message_id)
                sent += 1
                break
            except RetryAfter as exc:
                delay = exc.retry_after
                delay = delay.total_seconds() if hasattr(delay, "total_seconds") else float(delay)
                if attempt == 0:
                    await asyncio.sleep(delay + 1)
                    continue
                failed += 1
                break
            except Forbidden:
                blocked += 1  # ইউজার বট ব্লক করেছে / ডিঅ্যাক্টিভ
                break
            except TelegramError as exc:
                failed += 1
                log.warning("Broadcast failed for %s: %s", uid, exc)
                break
            except Exception:
                failed += 1
                log.exception("Broadcast unexpected error for %s", uid)
                break
        await asyncio.sleep(0.05)  # Telegram-এর limit মেনে ~২০ মেসেজ/সেকেন্ড

    try:
        await bot.send_message(
            admin_id,
            "📢 ব্রডকাস্ট সম্পন্ন!\n\n"
            f"✅ সফল: {sent}\n"
            f"🚫 ব্লক/নিষ্ক্রিয়: {blocked}\n"
            f"⚠️ ব্যর্থ: {failed}\n"
            f"👥 মোট: {len(user_ids)}"
        )
    except Exception:
        log.exception("Could not send broadcast summary")


async def broadcast_capture(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Broadcast মোডে Admin যা-ই পাঠাবে (যেকোনো ধরনের মেসেজ) সেটি ধরে Confirm চায়।"""
    msg = update.message
    if not msg or not update.effective_user or not is_admin(update.effective_user.id):
        return
    if context.user_data.get("admin_action") != "broadcast":
        return
    context.user_data.pop("admin_action", None)
    context.user_data["broadcast_pending"] = (msg.chat_id, msg.message_id)
    total = len(get_broadcast_user_ids())
    try:  # প্রিভিউ: ইউজাররা ঠিক যেমন দেখবে
        await context.bot.copy_message(chat_id=msg.chat_id, from_chat_id=msg.chat_id, message_id=msg.message_id)
    except Exception:
        log.exception("Broadcast preview failed")
    await context.bot.send_message(
        msg.chat_id,
        f"👆 উপরের মেসেজটি সব ইউজারকে হুবহু পাঠানো হবে।\n👥 মোট ইউজার: {total}\n\nনিশ্চিত করুন:",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton(f"✅ Send to {total} users", callback_data="adm:bcsend")],
            [InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")],
        ])
    )


async def handle_video(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin How to Use-এর ভিডিও (ক্যাপশনসহ বা ছাড়া) সেভ করে।"""
    if not (context.user_data.get("admin_action") == "howto_content"
            and is_admin(update.effective_user.id)):
        return
    video = update.message.video
    set_setting("howto_video_id", video.file_id)
    if update.message.caption:
        set_setting("howto_text_html", update.message.caption_html)
    context.user_data.pop("admin_action", None)
    await update.message.reply_text("✅ How to Use ভিডিও সেভ হয়েছে। ইউজার এভাবে দেখবে:")
    await send_howto(context, update.effective_chat.id, main_keyboard(update.effective_user.id))


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Broadcast চলাকালীন Admin-এর ফাইল সাবমিশন নয়, ব্রডকাস্ট কনটেন্ট (broadcast_capture ধরবে)
    if (context.user_data.get("admin_action") == "broadcast"
            and is_admin(update.effective_user.id)):
        return
    # Admin How to Use চলাকালীন ভিডিও File হিসেবে এলে সাবমিশন হিসেবে ধরা হবে না
    if (context.user_data.get("admin_action") == "howto_content"
            and is_admin(update.effective_user.id)):
        await update.message.reply_text(
            "❌ ভিডিওটি File হিসেবে এসেছে। অনুগ্রহ করে Video হিসেবে (Send as video) পাঠান।")
        return
    # Admin: Premium Emoji TXT upload (mapping rows or ID-only list).
    if (context.user_data.get("admin_action") == "premium_emoji_upload"
            and is_admin(update.effective_user.id)):
        doc = update.message.document
        filename = (doc.file_name or "").lower()
        if not filename.endswith(".txt"):
            await update.message.reply_text("❌ শুধু .txt ফাইল Upload করুন।")
            return
        if doc.file_size and doc.file_size > 1024 * 1024:
            await update.message.reply_text("❌ Premium Emoji ফাইল সর্বোচ্চ ১ MB হতে পারবে।")
            return
        try:
            telegram_file = await context.bot.get_file(doc.file_id)
            content = bytes(await telegram_file.download_as_bytearray())
            if len(content) > 1024 * 1024:
                await update.message.reply_text("❌ Premium Emoji ফাইল সর্বোচ্চ ১ MB হতে পারবে।")
                return
            report = {}
            emoji_map, id_count, explicit_count = await resolve_premium_emoji_file(
                context.bot, content, report
            )
            if not emoji_map:
                await update.message.reply_text(
                    "❌ কোনো বৈধ Premium Emoji ID পাওয়া যায়নি। TXT ফাইলে প্রতি লাইনে শুধু ID দিন, "
                    "অথবা ✅ = 5412345678901234567 লিখুন।"
                )
                return
            merged = dict(get_premium_emoji_map())
            merged.update(emoji_map)
            save_premium_emoji_map(merged)
            context.user_data.pop("admin_action", None)
            await update.message.reply_text(
                f"✅ {len(emoji_map)}টি Premium Emoji শনাক্ত ও সেভ হয়েছে.\\n"
                f"📄 ID পাওয়া গেছে: {id_count}টি | Mapping format: {explicit_count}টি\\n"
                f"✨ মোট সক্রিয় Mapping: {len(merged)}টি"
            )
        except (TelegramError, OSError, ValueError, UnicodeError):
            log.exception("Premium emoji TXT upload failed")
            await update.message.reply_text(
                "❌ ফাইলটি প্রসেস করা যায়নি। UTF-8 TXT এবং বৈধ Emoji ID দিয়ে আবার চেষ্টা করুন।"
            )
        except Exception:
            log.exception("Unexpected error while importing premium emoji TXT")
            await update.message.reply_text(
                "❌ Premium Emoji আপলোডে সমস্যা হয়েছে। বট বন্ধ হয়নি; আবার চেষ্টা করুন।"
            )
        return
    # Admin: Payment Report-এর Screenshot
    if context.user_data.get("admin_action") == "report_screenshot" and is_admin(update.effective_user.id):
        await handle_report_screenshot(update, context)
        return

    user = update.effective_user
    register_user(user)

    cat_id = context.user_data.get("selected_category")
    pm_id = context.user_data.get("selected_payment_method")
    payment_details = context.user_data.get("selected_payment_details")

    if not cat_id or not pm_id or not payment_details:
        clear_submission_state(context)
        await update.message.reply_text(
            "প্রথমে 📩 Submit File চাপুন, ক্যাটাগরি ও পেমেন্ট মেথড নির্বাচন করুন।"
        )
        return

    doc = update.message.document
    filename = doc.file_name or "upload"
    suffix = Path(filename).suffix.lower()
    if suffix not in {".xlsx", ".csv", ".txt"}:
        await update.message.reply_text("❌ শুধু XLSX, CSV অথবা TXT ফাইল গ্রহণ করা হচ্ছে।")
        return
    if doc.file_size and doc.file_size > MAX_FILE_MB * 1024 * 1024:
        await update.message.reply_text(f"❌ ফাইলের সাইজ সর্বোচ্চ {MAX_FILE_MB} MB হতে পারবে।")
        return

    with db() as con:
        cat = con.execute("SELECT * FROM categories WHERE id=? AND is_open=1 AND is_deleted=0", (cat_id,)).fetchone()
        pm = con.execute("SELECT * FROM payment_methods WHERE id=? AND is_open=1", (pm_id,)).fetchone()

    if not cat or not pm:
        clear_submission_state(context)
        await update.message.reply_text("ক্যাটাগরি বা পেমেন্ট মেথড বন্ধ হয়ে গেছে। আবার Submit File চাপুন।")
        return

    UPLOAD_DIR.mkdir(exist_ok=True)
    tg_file = await context.bot.get_file(doc.file_id)
    local_path = UPLOAD_DIR / f"{uuid.uuid4().hex}_{Path(filename).name}"
    await tg_file.download_to_drive(custom_path=str(local_path))
    total = count_file_items(local_path)
    if total <= 0:
        local_path.unlink(missing_ok=True)
        await update.message.reply_text("❌ ফাইল থেকে কোনো ডেটা/লাইন পড়া যায়নি। ফাইলটি পরীক্ষা করে আবার পাঠান।")
        return

    submission_id = uuid.uuid4().hex[:10].upper()
    user_message = (update.message.caption or "").strip()[:300]
    with db() as con:
        con.execute(
            """INSERT INTO submissions
            (id,user_id,category_id,file_id,file_name,total_items,rate,submitted_at,
             payment_method_id,payment_method_name,payment_details,user_message,local_path)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (submission_id, user.id, cat_id, doc.file_id, filename, total, cat["rate"], now(),
             pm["id"], pm["name"], payment_details, user_message, str(local_path))
        )

    clear_submission_state(context)

    await update.message.reply_text(
        RawHTML(
            "✅ <b>File submitted successfully!</b>\n\n"
            f"🆔 <b>Submission ID:</b> <code>{html.escape(submission_id)}</code>\n"
            f"📂 <b>Category:</b> {html.escape(cat['name'])}\n"
            f"📄 <b>File:</b> {html.escape(filename)}\n"
            f"🔢 <b>Detected rows/lines:</b> {total}\n"
            f"💳 <b>Payment:</b> {html.escape(pm['name'])} — {html.escape(payment_details)}\n"
            "📌 <b>Status:</b> ⏳ Pending Review\n\n"
            "<blockquote>অ্যাডমিন যাচাই করার পর Valid সংখ্যার ভিত্তিতে পেমেন্ট যোগ হবে।</blockquote>"
        ),
        reply_markup=main_keyboard(user.id)
    )

    # ---- Admin Inbox: File + নির্দিষ্ট User Details + শুধু Send Report button ----
    for admin_id in get_admin_ids():
        try:
            admin_username = f"@{user.username}" if user.username else user.full_name
            await context.bot.send_document(
                chat_id=admin_id,
                document=doc.file_id,
                caption=(f"👤 User: {admin_username}\n"
                         f"📂 Category: {cat['name']}\n"
                         f"💵 Rate: {money(cat['rate'])}\n"
                         f"💳 Payment Method: {pm['name']}\n"
                         f"🔑 Payment Details: {payment_details}"),
                reply_markup=InlineKeyboardMarkup([[
                    InlineKeyboardButton("📤 Send Report", callback_data=f"adm:report:{submission_id}")
                ]])
            )
        except Exception:
            log.exception("Could not notify admin %s", admin_id)

    # ---- Category Group-এ notification (Name, Username, Payment Details, Message, File) ----
    if cat['group_id']:
        try:
            # Telegram send_document-এ invite URL নয়, resolved numeric chat ID প্রয়োজন।
            group_chat_id = str(cat['group_id']).strip()
            if not re.fullmatch(r"-?\d+", group_chat_id):
                raise ValueError("Category group_id is not a numeric Telegram chat ID; set it using /setgroup inside that group.")
            with db() as con:
                srow = con.execute("SELECT * FROM submissions WHERE id=?", (submission_id,)).fetchone()
            group_caption = build_group_caption(srow, GROUP_STATUS_LABELS["PENDING"])
            msg = await context.bot.send_document(
                chat_id=int(group_chat_id), document=doc.file_id,
                caption=group_caption,
                reply_markup=group_keyboard(submission_id, "PENDING")
            )
            with db() as con:
                con.execute("UPDATE submissions SET group_message_id=? WHERE id=?", (msg.message_id, submission_id))
            log.info("Submission %s forwarded to category group %s", submission_id, group_chat_id)
        except Exception as exc:
            log.exception("Could not send submission %s to category group %s", submission_id, cat['group_id'])
            # Group delivery ব্যর্থ হলে Admin-কে কারণ জানানো হবে, যাতে ভুল Group ID/permission ধরা যায়।
            for admin_id in get_admin_ids():
                try:
                    await context.bot.send_message(
                        chat_id=admin_id,
                        text=(f"⚠️ Private OTP Group-এ File Forward ব্যর্থ হয়েছে.\n"
                              f"Submission: {submission_id}\nCategory: {cat['name']}\n"
                              f"Group ID: {cat['group_id']}\nError: {str(exc)[:700]}\n\n"
                              "বটকে ওই Group-এ Admin করুন এবং Category Group ID যাচাই করুন।")
                    )
                except Exception:
                    log.exception("Could not notify admin %s about group forwarding failure", admin_id)

async def profile(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    with db() as con:
        u = con.execute("SELECT * FROM users WHERE user_id=?", (user_id,)).fetchone()
        stats = con.execute(
            "SELECT COUNT(*) total, SUM(CASE WHEN status IN ('APPROVED','FULLY_COMPLETED') THEN 1 ELSE 0 END) approved, "
            "SUM(CASE WHEN status='PENDING' THEN 1 ELSE 0 END) pending FROM submissions WHERE user_id=?",
            (user_id,)
        ).fetchone()
    await update.message.reply_text(
        RawHTML(
            f"👤 <b>{fancy_text('MY PROFILE', 'bold')}</b>\n\n"
            f"🙍 <b>Name:</b> {html.escape(u['first_name'] or '-')}\n"
            f"🆔 <b>User ID:</b> <code>{user_id}</code>\n\n"
            f"💰 <b>Balance:</b> {html.escape(money(u['balance']))}\n\n"
            "<blockquote>"
            f"📦 Total submissions: {stats['total'] or 0}\n"
            f"✅ Approved: {stats['approved'] or 0}\n"
            f"⏳ Pending: {stats['pending'] or 0}"
            "</blockquote>"
        ),
        reply_markup=main_keyboard(update.effective_user.id)
    )


async def payments(update: Update, context: ContextTypes.DEFAULT_TYPE):
    with db() as con:
        rows = con.execute(
            "SELECT * FROM payments WHERE user_id=? ORDER BY id DESC LIMIT 10",
            (update.effective_user.id,)
        ).fetchall()
    if not rows:
        await update.message.reply_text("এখনো কোনো পেমেন্ট রেকর্ড নেই।", reply_markup=main_keyboard(update.effective_user.id))
        return
    lines = [f"💳 <b>{fancy_text('PAYMENT HISTORY', 'bold')}</b> <i>(Last 10)</i>"]
    for r in rows:
        lines.append(
            f"\n💰 <b>{html.escape(money(r['amount']))}</b>  <i>#{r['id']}</i>\n"
            f"🆔 <code>{html.escape(str(r['submission_id']))}</code> • 📅 {html.escape(r['created_at'][:10])}")
    await update.message.reply_text(RawHTML("\n".join(lines)), reply_markup=main_keyboard(update.effective_user.id))


async def send_price_rules(context, chat_id, reply_markup=None, custom_html=None):
    """Admin-এর সেট করা Price & Rules দেখায়; না থাকলে ডিফল্ট Price List।"""
    custom = custom_html if custom_html is not None else get_setting("price_rules_html")
    if custom:
        await context.bot.send_message(chat_id=chat_id, text=RawHTML(custom), reply_markup=reply_markup)
        return
    cats = get_categories()
    lines = [f"🎯 <b>{fancy_text('PRICE LIST', 'bold')}</b>"]
    for c in cats:
        status = "🟢 Open" if c["is_open"] else "🔴 Closed"
        lines.append(
            f"\n📂 <b>{html.escape(c['name'])}</b>\n"
            f"<blockquote>💵 Rate: {html.escape(money(c['rate']))} / valid number\n📌 Status: {status}</blockquote>")
    if not cats:
        lines.append("\nএখনো কোনো ক্যাটাগরি যোগ করা হয়নি।")
    await context.bot.send_message(chat_id=chat_id, text=RawHTML("\n".join(lines)), reply_markup=reply_markup)


async def prices(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_price_rules(context, update.effective_chat.id,
                           main_keyboard(update.effective_user.id))


DEFAULT_HOWTO_TEXT = (
    f"📖 <b>{fancy_text('HOW TO USE', 'bold')}</b>\n\n"
    "1️⃣ <b>Submit File</b> চাপুন\n"
    "2️⃣ ক্যাটাগরি নির্বাচন করুন\n"
    "3️⃣ Payment Method ও নিজের Payment Details দিন\n"
    "4️⃣ XLSX / CSV / TXT নাম্বার লিস্ট পাঠান\n"
    "5️⃣ অ্যাডমিন ফাইল যাচাই করবে\n"
    "6️⃣ Valid নাম্বারের সংখ্যা অনুযায়ী ব্যালেন্স যোগ হবে\n\n"
    "<blockquote>⚠️ একই ফাইল বারবার সাবমিট করবেন না।</blockquote>"
)


async def send_howto(context, chat_id, reply_markup=None):
    """Admin-এর সেট করা How to Use (ভিডিও + লেখা) দেখায়; না থাকলে ডিফল্ট লেখা।"""
    text_html = get_setting("howto_text_html")
    video_id = get_setting("howto_video_id")

    if video_id:
        fits = bool(text_html) and len(text_html) <= 1000
        try:
            await context.bot.send_video(
                chat_id=chat_id, video=video_id,
                caption=RawHTML(text_html) if fits else None,
                reply_markup=reply_markup if (fits or not text_html) else None)
            if text_html and not fits:
                await context.bot.send_message(chat_id=chat_id, text=RawHTML(text_html),
                                               reply_markup=reply_markup)
            return
        except Exception:
            log.exception("Could not send How to Use video")
            # ভিডিও পাঠানো না গেলে শুধু লেখা দেখানো হবে

    if text_html:
        await context.bot.send_message(chat_id=chat_id, text=RawHTML(text_html), reply_markup=reply_markup)
    elif not video_id:
        await context.bot.send_message(chat_id=chat_id, text=RawHTML(DEFAULT_HOWTO_TEXT), reply_markup=reply_markup)


async def howto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_howto(context, update.effective_chat.id, main_keyboard(update.effective_user.id))


async def support(update: Update, context: ContextTypes.DEFAULT_TYPE):
    support_link = get_support_link()
    support_url = normalize_support_url(support_link)
    intro = (
        "💬 <b>SUPPORT</b>\n\n"
        "কোনো সমস্যা বা প্রশ্ন আছে? আমাদের Support টিম আপনাকে সাহায্য করতে প্রস্তুত।\n\n"
    )
    outro = "\n\n⏳ উত্তর পেতে কিছুটা সময় লাগতে পারে, অনুগ্রহ করে ধৈর্য ধরুন।"

    if support_url:
        label = support_link.strip() if support_link.strip().startswith("@") else "Support-এ মেসেজ করুন"
        text = (intro
                + f'📩 সরাসরি মেসেজ করুন: <a href="{html.escape(support_url, quote=True)}">{html.escape(label)}</a>'
                + outro)
        # RawHTML: প্রিমিয়াম-ইমোজি র‍্যাপার HTML ট্যাগ escape করে "কোডের মতো" দেখাত — এখন ঠিক থাকবে
        await update.message.reply_text(
            RawHTML(text),
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("💬 Contact Support", url=support_url)]])
        )
    else:
        text = intro + f"📩 যোগাযোগ: {html.escape(support_link)}" + outro
        await update.message.reply_text(RawHTML(text))


async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    with db() as con:
        pending = con.execute("SELECT COUNT(*) FROM submissions WHERE status='PENDING'").fetchone()[0]
        users = con.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        total = con.execute("SELECT COALESCE(SUM(balance),0) FROM users").fetchone()[0]
        pm_open = con.execute("SELECT COUNT(*) FROM payment_methods WHERE is_open=1").fetchone()[0]
    await update.message.reply_text(
        f"🛠 ADMIN CONTROL PANEL\n\n⏳ Pending: {pending}\n👥 Users: {users}\n"
        f"💰 Total user balances: {money(total)}\n💳 Open Payment Methods: {pm_open}\n\n"
        "নিচের বাটন থেকে সবকিছু পরিচালনা করুন।",
        reply_markup=admin_home_keyboard()
    )


async def addcategory(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    raw = update.message.text.partition(" ")[2]
    if "|" not in raw:
        await update.message.reply_text("Usage: /addcategory Category Name | 0.50")
        return
    name, rate = [x.strip() for x in raw.split("|", 1)]
    try: rate = float(rate)
    except ValueError:
        await update.message.reply_text("Rate must be a number."); return
    if not name or rate < 0:
        await update.message.reply_text("Invalid category name or rate."); return
    try:
        with db() as con:
            con.execute("INSERT INTO categories(name,rate,is_open) VALUES(?,?,1)", (name, rate))
        await update.message.reply_text(f"✅ Added: {name} — {money(rate)} per valid number")
    except sqlite3.IntegrityError:
        await update.message.reply_text("এই নামে ক্যাটাগরি আগে থেকেই আছে।")


async def categories_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    cats = get_categories()
    await update.message.reply_text("\n".join(
        f"ID {c['id']} | {c['name']} | {money(c['rate'])} | {'OPEN' if c['is_open'] else 'CLOSED'}"
        for c in cats
    ) or "No categories.")


async def setrate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    try:
        cat_id, rate = int(context.args[0]), float(context.args[1])
        if rate < 0: raise ValueError()
        with db() as con:
            cur = con.execute("UPDATE categories SET rate=? WHERE id=?", (rate, cat_id))
        await update.message.reply_text("✅ Rate updated." if cur.rowcount else "Category not found.")
    except (IndexError, ValueError):
        await update.message.reply_text("Usage: /setrate CATEGORY_ID RATE")


async def toggle_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    try: cat_id = int(context.args[0])
    except (IndexError, ValueError):
        await update.message.reply_text("Usage: /toggle CATEGORY_ID"); return
    with db() as con:
        cur = con.execute("UPDATE categories SET is_open=CASE is_open WHEN 1 THEN 0 ELSE 1 END WHERE id=?", (cat_id,))
    await update.message.reply_text("✅ Category status changed." if cur.rowcount else "Category not found.")


async def pending_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    with db() as con:
        rows = con.execute("SELECT * FROM submissions WHERE status='PENDING' ORDER BY submitted_at").fetchall()
    if not rows:
        await update.message.reply_text("No pending submissions."); return
    for s in rows[:30]:
        await update.message.reply_text(
            f"🆔 ID: {s['id']}\nUser: {s['user_id']}\nFile: {s['file_name']}\n"
            f"Rows/lines: {s['total_items']}\nRate snapshot: {money(s['rate'])}",
            reply_markup=submission_action_keyboard(s["id"])
        )


async def review_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    if not context.args:
        await update.message.reply_text("Usage: /review SUBMISSION_ID"); return
    sid = context.args[0].upper()
    with db() as con:
        s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
    if not s:
        await update.message.reply_text("Submission not found."); return
    cat_name = str(s["category_id"])
    try:
        with db() as con:
            c = con.execute("SELECT name FROM categories WHERE id=?", (s["category_id"],)).fetchone()
            if c:
                cat_name = c["name"]
    except Exception:
        log.exception("Could not load category name")
    await context.bot.send_document(
        chat_id=update.effective_chat.id, document=s["file_id"],
        caption=(f"Submission {sid}\nUser: {s['user_id']}\nFile: {s['file_name']}\n"
                 f"Category: {cat_name}\n"
                 f"Detected rows/lines: {s['total_items']}\nRate: {money(s['rate'])}\n"
                 f"💳 Payment: {s['payment_method_name'] or '-'} — {s['payment_details'] or '-'}\n"
                 f"Status: {s['status']}")
    )
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=f"📄 Submission {sid} — নিচের বাটন থেকে Approve অথবা Reject করুন।",
        reply_markup=submission_action_keyboard(sid)
    )


async def settle_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    try:
        sid = context.args[0].upper()
        valid, invalid = int(context.args[1]), int(context.args[2])
        if valid < 0 or invalid < 0: raise ValueError()
    except (IndexError, ValueError):
        await update.message.reply_text("Usage: /settle SUBMISSION_ID VALID_COUNT INVALID_COUNT"); return
    with db() as con:
        s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
        if not s:
            await update.message.reply_text("Submission not found."); return
        if s["status"] != "PENDING":
            await update.message.reply_text("এই সাবমিশনটি আগে থেকেই রিভিউ করা হয়েছে।"); return
        if valid + invalid > s["total_items"]:
            await update.message.reply_text(
                f"Valid + Invalid ({valid+invalid}) detected rows/lines ({s['total_items']})-এর বেশি হতে পারবে না।"
            ); return
        amount = valid * s["rate"]
        con.execute(
            "UPDATE submissions SET status='APPROVED',valid_items=?,invalid_items=?,amount=?,reviewed_at=? WHERE id=?",
            (valid, invalid, amount, now(), sid)
        )
        con.execute("UPDATE users SET balance=balance+? WHERE user_id=?", (amount, s["user_id"]))
        con.execute(
            "INSERT INTO payments(user_id,submission_id,amount,note,created_at) VALUES(?,?,?,?,?)",
            (s["user_id"], sid, amount, f"{valid} valid × {money(s['rate'])}", now())
        )
    await update.message.reply_text(f"✅ Approved {sid}\nValid: {valid}\nInvalid: {invalid}\nPayment: {money(amount)}")
    try:
        await context.bot.send_message(
            s["user_id"],
            RawHTML(f"✅ <b>ফাইল রিভিউ সম্পন্ন!</b>\n\n🆔 <b>Submission:</b> <code>{sid}</code>\n✅ <b>Valid:</b> {valid}\n❌ <b>Invalid:</b> {invalid}\n💰 <b>Added to balance:</b> {money(amount)}")
        )
    except Exception:
        log.exception("Could not notify user")


async def reject_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    if len(context.args) < 1:
        await update.message.reply_text("Usage: /reject SUBMISSION_ID"); return
    sid = context.args[0].upper()
    with db() as con:
        s = con.execute("SELECT status FROM submissions WHERE id=?", (sid,)).fetchone()
    if not s:
        await update.message.reply_text("Submission not found."); return
    if s["status"] not in ("PENDING", "APPROVED"):
        await update.message.reply_text("এই সাবমিশনটি আগে থেকেই রিভিউ করা হয়েছে।"); return
    ok, err = await delete_submission_completely(context, sid)
    await update.message.reply_text(f"🗑 Submission {sid} rejected ও সম্পূর্ণ মুছে ফেলা হয়েছে।" if ok else err)


async def balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    try:
        uid, amount = int(context.args[0]), float(context.args[1])
    except (IndexError, ValueError):
        await update.message.reply_text("Usage: /balance USER_ID AMOUNT (negative amount deducts)")
        return
    with db() as con:
        u = con.execute("SELECT balance FROM users WHERE user_id=?", (uid,)).fetchone()
        if not u:
            await update.message.reply_text("User not found."); return
        if u["balance"] + amount < 0:
            await update.message.reply_text("Balance cannot become negative."); return
        con.execute("UPDATE users SET balance=balance+? WHERE user_id=?", (amount, uid))
    await update.message.reply_text(f"✅ Balance adjusted by {money(amount)}.")
    try:
        await context.bot.send_message(uid, f"Your balance was adjusted by {money(amount)}.")
    except Exception:
        pass


# ========================================================
# ===========  SUBMISSION STATUS / GROUP HELPERS  ========
# ========================================================

GROUP_STATUS_LABELS = {
    "PENDING": "⏳ PENDING",
    "APPROVED": "✅ APPROVED",
    "FULLY_COMPLETED": "✅ FULLY COMPLETED",
}


def build_group_caption(s, status_line):
    """Group message-এর caption: Name, Username, Payment Details, Message।"""
    with db() as con:
        u = con.execute("SELECT username, first_name FROM users WHERE user_id=?", (s["user_id"],)).fetchone()
        c = con.execute("SELECT name FROM categories WHERE id=?", (s["category_id"],)).fetchone()
        paid = None
        if s["status"] == "FULLY_COMPLETED":
            paid = con.execute(
                "SELECT COALESCE(SUM(amount),0) FROM payment_reports WHERE submission_id=?", (s["id"],)
            ).fetchone()[0]
    name = u["first_name"] if u and u["first_name"] else str(s["user_id"])
    username = f"@{u['username']}" if u and u["username"] else "N/A"
    message = (s["user_message"] or "").strip() or "-"
    lines = [
        f"📥 SUBMISSION — {status_line}",
        f"🆔 ID: {s['id']}",
        f"👤 Name: {name}",
        f"🔗 Username: {username}",
        f"🆔 User ID: {s['user_id']}",
        f"📂 Category: {c['name'] if c else s['category_id']}",
        f"📎 File: {s['file_name']}",
        f"🔢 Rows/lines: {s['total_items']}",
        f"💵 Rate: {money(s['rate'])}",
        f"💳 Payment: {s['payment_method_name'] or '-'} — {s['payment_details'] or '-'}",
        f"💬 Message: {message}",
    ]
    if paid is not None:
        lines.append(f"💰 Paid: {money(paid)}")
    return "\n".join(lines)[:1000]


def group_keyboard(sid, status):
    """Status অনুযায়ী Group message-এর বাটন।"""
    if status == "PENDING":
        return InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Approve", callback_data=f"grp:approve:{sid}"),
            InlineKeyboardButton("❌ Reject", callback_data=f"grp:reject:{sid}"),
        ]])
    if status == "APPROVED":
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("📤 Send Report", callback_data=f"grp:report:{sid}")],
            [InlineKeyboardButton("❌ Reject", callback_data=f"grp:reject:{sid}")],
        ])
    return None  # FULLY_COMPLETED — আর কোনো বাটন নেই


async def refresh_group_message(context, sid):
    """Category Group-এর submission message-এর Status ও বাটন আপডেট করে।"""
    with db() as con:
        s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
        g = con.execute("SELECT group_id FROM categories WHERE id=?", (s["category_id"],)).fetchone() if s else None
    if not (s and g and g["group_id"] and s["group_message_id"]):
        return
    try:
        await context.bot.edit_message_caption(
            chat_id=g["group_id"], message_id=s["group_message_id"],
            caption=build_group_caption(s, GROUP_STATUS_LABELS.get(s["status"], s["status"])),
            reply_markup=group_keyboard(sid, s["status"])
        )
    except Exception:
        log.exception("Could not update group message")


async def remove_group_message(context, sid):
    """Group থেকে submission message সম্পূর্ণ মুছে ফেলে।"""
    with db() as con:
        g = con.execute(
            "SELECT c.group_id, s.group_message_id FROM submissions s "
            "JOIN categories c ON s.category_id=c.id WHERE s.id=?", (sid,)
        ).fetchone()
    if g and g["group_id"] and g["group_message_id"]:
        try:
            await context.bot.delete_message(chat_id=g["group_id"], message_id=g["group_message_id"])
        except Exception:
            log.exception("Could not delete group message")


async def delete_submission_completely(context, sid, notify_user=True):
    """Reject = Submission সম্পূর্ণ মুছে ফেলা: Database, Group message ও ডিস্কের ফাইল।
    (ok, error_text) ফেরত দেয়। Payments (ব্যালেন্স হিস্ট্রি) অপরিবর্তিত থাকে।"""
    with db() as con:
        s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
    if not s:
        return False, "Submission not found."

    # Group থেকে মেসেজ মুছে ফেলা (ডাটা মোছার আগে, কারণ group_message_id লাগবে)
    await remove_group_message(context, sid)

    # Database থেকে সম্পূর্ণ মুছে ফেলা
    with db() as con:
        con.execute("DELETE FROM payment_reports WHERE submission_id=?", (sid,))
        con.execute("DELETE FROM submissions WHERE id=?", (sid,))

    # ডিস্কে সংরক্ষিত ফাইল মুছে ফেলা
    try:
        if s["local_path"]:
            Path(s["local_path"]).unlink(missing_ok=True)
    except Exception:
        log.exception("Could not delete local file for %s", sid)

    if notify_user:
        try:
            await context.bot.send_message(s["user_id"], RawHTML(
                f"❌ <b>আপনার সাবমিশন Reject করা হয়েছে।</b>\n\n🆔 <b>Submission:</b> <code>{sid}</code>"))
        except Exception:
            log.exception("Could not notify user")
    return True, ""


async def approve_submission(context, sid):
    """Approve: শুধু Status APPROVED + ইউজারকে Approval Notification। (ok, error_text) ফেরত দেয়।"""
    with db() as con:
        s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
        if not s:
            return False, "Submission not found."
        if s["status"] != "PENDING":
            return False, "এই সাবমিশনটি আগে থেকেই রিভিউ করা হয়েছে।"
        con.execute("UPDATE submissions SET status='APPROVED', reviewed_at=? WHERE id=?", (now(), sid))
    try:
        await context.bot.send_message(
            s["user_id"], RawHTML(
                f"✅ <b>আপনার সাবমিশন Approve করা হয়েছে!</b>\n\n🆔 <b>Submission:</b> <code>{sid}</code>\n\n"
                "<blockquote>অ্যাডমিন পেমেন্ট জানালে আপনাকে নোটিফিকেশন দেওয়া হবে।</blockquote>"))
    except Exception:
        log.exception("Could not notify user")
    await refresh_group_message(context, sid)
    return True, ""


async def resolve_group(context, raw):
    """Admin-এর দেওয়া Group ID / public @username / public লিংক থেকে chat id বের করে।"""
    ref = raw.strip()
    if re.search(r"t\.me/(\+|joinchat/)", ref):
        return None, (
            "⚠️ Private invite লিংক থেকে বট গ্রুপ চিনতে পারে না (Telegram-এর সীমাবদ্ধতা)।\n\n"
            "যেকোনো একটি উপায়ে যুক্ত করুন:\n"
            "1) বটকে গ্রুপে Admin করে গ্রুপের ভিতরে লিখুন: /setgroup CATEGORY_ID\n"
            "2) অথবা এখানে Group ID (যেমন -100xxxxxxxxxx) পাঠান।"
        )
    m = re.search(r"(?:t\.me/|^@)([A-Za-z][A-Za-z0-9_]{3,})", ref)
    if m:
        ref = "@" + m.group(1)
    elif not re.fullmatch(r"-?\d+", ref):
        return None, "সঠিক Group ID (যেমন -100xxxxxxxxxx) অথবা public @username দিন।"
    try:
        chat = await context.bot.get_chat(ref)
    except Exception:
        return None, ("বট এই গ্রুপটি খুঁজে পায়নি। বটকে গ্রুপে Admin হিসেবে যোগ করে আবার পাঠান, "
                      "অথবা ❌ Cancel চাপুন।")
    return str(chat.id), None


async def setgroup_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """গ্রুপের ভিতরে /setgroup CATEGORY_ID লিখলে সেই গ্রুপ ক্যাটাগরির সাথে যুক্ত হবে।"""
    user, chat, msg = update.effective_user, update.effective_chat, update.effective_message
    if not user or not is_admin(user.id):
        return
    if chat.type == "private":
        await msg.reply_text("এই কমান্ডটি Private Group-এর ভিতরে লিখুন: /setgroup CATEGORY_ID")
        return
    if not context.args or not context.args[0].isdigit():
        await msg.reply_text("Usage: /setgroup CATEGORY_ID")
        return
    with db() as con:
        cur = con.execute("UPDATE categories SET group_id=? WHERE id=?", (str(chat.id), int(context.args[0])))
    await msg.reply_text("✅ এই গ্রুপ ক্যাটাগরির সাথে যুক্ত হয়েছে।" if cur.rowcount else "Category not found.")


# ========================================================
# ====================  ADMIN PANEL  =====================
# ========================================================

async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if not is_admin(q.from_user.id):
        await q.edit_message_text("⛔ Admin access only.")
        return
    data = q.data.split(":")
    action = data[1] if len(data) > 1 else ""

    if action in ("emojiadd", "emojiref", "emojiset", "emojiglob", "emojiundo"):
        await emoji_editor_callback(update, context, q, data)
        return

    # ⚙️ System settings (Main Admin only for privileged controls)
    if action == "system":
        rows = [
            [InlineKeyboardButton("✉️ Welcome Message", callback_data="adm:welcome")],
            [InlineKeyboardButton("⚠️ এডিট ইস্যু", callback_data="adm:editissue")],
            [InlineKeyboardButton("🎯 Price and Rules", callback_data="adm:pricerules")],
            [InlineKeyboardButton("📖 How to Use", callback_data="adm:howtoedit")],
            [InlineKeyboardButton("💬 Edit Support Link", callback_data="adm:editsupport")],
            [InlineKeyboardButton("📢 Force to Join", callback_data="adm:forcejoin")],
            [InlineKeyboardButton("✨ Upload Premium Emojis", callback_data="adm:premiumemojis")],
            [InlineKeyboardButton("➕ Add Single Emoji", callback_data="adm:emojiadd:0")],
            [InlineKeyboardButton("📝 Edit Category Note", callback_data="adm:editcatnote")],
        ]
        if is_main_admin(q.from_user.id):
            rows.append([InlineKeyboardButton("➕ Add Admin", callback_data="adm:addadmin")])
        rows.append([InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")])
        await q.edit_message_text("⚙️ SYSTEM SETTINGS\nSelect an option:", reply_markup=InlineKeyboardMarkup(rows))
        return

    if action == "premiumemojis":
        context.user_data["admin_action"] = "premium_emoji_upload"
        current_count = len(get_premium_emoji_map())
        rows = [
            [InlineKeyboardButton("🔗 Import from Emoji Pack (Auto)", callback_data="adm:premiumpack")],
            [InlineKeyboardButton("🗑 Reset / Off Premium Emojis", callback_data="adm:premiumreset")],
            [InlineKeyboardButton("⬅️ System", callback_data="adm:system")],
        ]
        await q.edit_message_text(
            "✨ PREMIUM EMOJI SETTINGS\n\n"
            "📄 TXT ফাইল Upload করুন — প্রতি লাইনে শুধু Premium Emoji ID দিলেই হবে। "
            "বট Telegram থেকে ID শনাক্ত করে সংশ্লিষ্ট ইমোজি নিজে থেকেই সেট করবে।\\n\\n"
            "আগের ফরম্যাটও চলবে: সাধারণ ইমোজি = Custom Emoji ID\\n\\n"
            "উদাহরণ (শুধু ID):\\n"
            "5412345678901234567\\n"
            "5498765432109876543\\n\\n"
            f"বর্তমানে সেভ করা ইমোজি: {current_count}টি\n"
            "ভুল ফরম্যাট/খালি লাইন উপেক্ষা করা হবে।",
            reply_markup=InlineKeyboardMarkup(rows)
        )
        return

    if action == "premiumpack":
        context.user_data["admin_action"] = "premium_pack_link"
        await q.edit_message_text(
            "🔗 EMOJI PACK IMPORT\n\n"
            "আপনার পছন্দের Custom Emoji Pack-এর লিংক পাঠান (একাধিক হলে আলাদা লাইনে):\n"
            "https://t.me/addemoji/PackName\n\n"
            "বট প্যাকটি Telegram থেকে স্ক্যান করে বটের সব ইমোজির জন্য মিল খুঁজে বের করবে এবং ID সেভ করবে। "
            "প্যাকে যে ইমোজি নেই সেগুলো সাধারণই থাকবে।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ Back", callback_data="adm:premiumemojis")],
                [InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")],
            ])
        )
        return

    if action == "premiumreset":
        with db() as con:
            con.execute(
                "DELETE FROM bot_settings WHERE setting_key='premium_emojis'"
            )
        context.user_data.pop("admin_action", None)
        await q.edit_message_text(
            "🗑 Premium Emoji System বন্ধ করা হয়েছে। এখন থেকে সাধারণ ইমোজি ব্যবহার হবে।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ System", callback_data="adm:system")]
            ])
        )
        return

    if action == "forcejoin":
        channels = get_force_join_channels()
        rows = [
            [InlineKeyboardButton("➕ Add Channel / Group", callback_data="adm:fjadd")],
            [InlineKeyboardButton("📋 View / Manage Links", callback_data="adm:fjlist")],
            [InlineKeyboardButton("⬅️ System", callback_data="adm:system")],
        ]
        await q.edit_message_text(
            f"📢 FORCE TO JOIN SETTINGS\nবর্তমানে যুক্ত আছে: {len(channels)}টি চ্যানেল/গ্রুপ।\n"
            "নতুন লিংক যোগ বা আগের লিংক ম্যানেজ করুন।",
            reply_markup=InlineKeyboardMarkup(rows)
        )
        return

    if action == "fjadd":
        context.user_data["admin_action"] = "force_join_add"
        await q.edit_message_text(
            "চ্যানেল/গ্রুপের Join link পাঠান।\n\n"
            "পাবলিক চ্যানেল: https://t.me/channelname অথবা @channelname\n"
            "প্রাইভেট চ্যানেল/গ্রুপ: invite link | chat_id\n"
            "উদাহরণ: https://t.me/+AbCdEf | -1001234567890\n\n"
            "নোট: বটকে সংশ্লিষ্ট চ্যানেল/গ্রুপে Admin করতে হবে।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]])
        )
        return

    if action == "fjlist":
        channels = get_force_join_channels()
        rows = []
        for channel in channels:
            title = (channel["title"] or "চ্যানেল/গ্রুপ")[:35]
            rows.append([InlineKeyboardButton(f"🗑 Delete: {title}", callback_data=f"adm:fjdel:{channel['id']}")])
        rows.append([InlineKeyboardButton("⬅️ Force to Join", callback_data="adm:forcejoin")])
        if not channels:
            await q.edit_message_text("এখনো কোনো Force to Join লিংক যোগ করা হয়নি।", reply_markup=InlineKeyboardMarkup(rows))
        else:
            details = "\n".join(f"{i+1}. {c['title']} — {c['join_link']}" for i, c in enumerate(channels))
            await q.edit_message_text("📋 যুক্ত করা চ্যানেল/গ্রুপ:\n\n" + details, reply_markup=InlineKeyboardMarkup(rows))
        return

    if action == "fjdel":
        if len(data) < 3 or not data[2].isdigit():
            await q.answer("Invalid channel selection.", show_alert=True)
            return
        with db() as con:
            con.execute("DELETE FROM force_join_channels WHERE id=?", (int(data[2]),))
        await q.answer("লিংকটি মুছে ফেলা হয়েছে।")
        channels = get_force_join_channels()
        rows = []
        for channel in channels:
            rows.append([InlineKeyboardButton(f"🗑 Delete: {(channel['title'] or 'চ্যানেল/গ্রুপ')[:35]}", callback_data=f"adm:fjdel:{channel['id']}")])
        rows.append([InlineKeyboardButton("⬅️ Force to Join", callback_data="adm:forcejoin")])
        details = "\n".join(f"{i+1}. {c['title']} — {c['join_link']}" for i, c in enumerate(channels))
        await q.edit_message_text("📋 যুক্ত করা চ্যানেল/গ্রুপ:\n\n" + (details or "কোনো লিংক নেই।"), reply_markup=InlineKeyboardMarkup(rows))
        return

    # 📢 Broadcast
    if action == "broadcast":
        clear_admin_state(context)
        context.user_data["admin_action"] = "broadcast"
        total = len(get_broadcast_user_ids())
        await q.edit_message_text(
            "📢 BROADCAST\n\n"
            f"👥 মোট ইউজার: {total}\n\n"
            "এখন যা পাঠাতে চান সেটি এখানে পাঠান — লেখা, লিংক, ইমোজি, ছবি, ভিডিও, ফাইল, "
            "ভয়েস, স্টিকার বা অন্য যেকোনো মেসেজ। ফরম্যাটিংসহ হুবহু সবার কাছে যাবে।\n\n"
            "পাঠানোর আগে একবার Confirm করতে হবে।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]])
        )
        return

    if action == "bcsend":
        pending = context.user_data.pop("broadcast_pending", None)
        if not pending:
            await q.edit_message_text(
                "কোনো ব্রডকাস্ট অপেক্ষমাণ নেই।",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]])
            )
            return
        user_ids = get_broadcast_user_ids()
        await q.edit_message_text(
            f"🚀 ব্রডকাস্ট শুরু হয়েছে — {len(user_ids)} জন ইউজারকে পাঠানো হচ্ছে।\n"
            "শেষ হলে রিপোর্ট পাঠানো হবে। এর মধ্যে বট স্বাভাবিকভাবেই কাজ করবে।"
        )
        context.application.create_task(
            run_broadcast(context.application, q.from_user.id, pending[0], pending[1], user_ids))
        return

    # 🎯 Price and Rules — Admin টেক্সট সেট করবে
    if action == "pricerules":
        context.user_data["admin_action"] = "price_rules_text"
        rows = []
        if get_setting("price_rules_html"):
            rows.append([InlineKeyboardButton("♻️ Reset to Default", callback_data="adm:pricerulesreset")])
        rows.append([InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")])
        await q.edit_message_text(
            "🎯 PRICE AND RULES\n\n"
            "ইউজাররা Price & Rules বাটনে যে লেখা দেখবে, সেটি এখন পাঠান।\n"
            "• Telegram-এর Bold / Italic / Link / Monospace ফরম্যাটিং সাপোর্ট করে।\n"
            "• যেকোনো লিংক সরাসরি পেস্ট করতে পারেন।\n\n"
            + ("বর্তমানে আপনার কাস্টম লেখা সেট করা আছে।" if get_setting("price_rules_html")
               else "বর্তমানে ডিফল্ট Price List দেখানো হচ্ছে।"),
            reply_markup=InlineKeyboardMarkup(rows)
        )
        return

    if action == "pricerulesreset":
        delete_settings("price_rules_html")
        context.user_data.pop("admin_action", None)
        await q.edit_message_text(
            "♻️ Price & Rules ডিফল্টে ফিরিয়ে আনা হয়েছে।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ System", callback_data="adm:system")]])
        )
        return

    # 📖 How to Use — Admin লেখা ও/অথবা ভিডিও সেট করবে
    if action == "howtoedit":
        context.user_data["admin_action"] = "howto_content"
        has_video = bool(get_setting("howto_video_id"))
        has_text = bool(get_setting("howto_text_html"))
        rows = []
        if has_video:
            rows.append([InlineKeyboardButton("🗑 Remove Video", callback_data="adm:howtoremovevideo")])
        if has_video or has_text:
            rows.append([InlineKeyboardButton("♻️ Reset to Default", callback_data="adm:howtoreset")])
        rows.append([InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")])
        await q.edit_message_text(
            "📖 HOW TO USE\n\n"
            "নিচের যেকোনো একটি পাঠান:\n"
            "• শুধু লেখা → লেখা সেভ হবে (আগের ভিডিও থাকবে)\n"
            "• ভিডিও + ক্যাপশন → ভিডিও ও লেখা একসাথে সেভ হবে\n"
            "• শুধু ভিডিও → ভিডিও সেভ হবে (আগের লেখা থাকবে)\n\n"
            "Bold / Italic / Link ফরম্যাটিং ও সব ধরনের লিংক সাপোর্ট করে।\n"
            "ভিডিও অবশ্যই Video হিসেবে পাঠান (File/Document হিসেবে নয়)।\n\n"
            f"বর্তমান: লেখা {'✅' if has_text else '—'} | ভিডিও {'✅' if has_video else '—'}",
            reply_markup=InlineKeyboardMarkup(rows)
        )
        return

    if action == "howtoremovevideo":
        delete_settings("howto_video_id")
        context.user_data.pop("admin_action", None)
        await q.edit_message_text(
            "🗑 How to Use-এর ভিডিও মুছে ফেলা হয়েছে।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ System", callback_data="adm:system")]])
        )
        return

    if action == "howtoreset":
        delete_settings("howto_video_id", "howto_text_html")
        context.user_data.pop("admin_action", None)
        await q.edit_message_text(
            "♻️ How to Use ডিফল্টে ফিরিয়ে আনা হয়েছে।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ System", callback_data="adm:system")]])
        )
        return

    if action == "editissue":
        context.user_data["admin_action"] = "payment_issue_message"
        current = html_to_plain(get_payment_issue_message())
        await q.edit_message_text(
            "ইউজারদের Payment Issue বাটনে যে মেসেজ দেখাতে চান, সেটি লিখে পাঠান।\n\n"
            f"বর্তমান মেসেজ:\n{current}\n\nবাতিল করতে Cancel চাপুন।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]])
        )
        return

    if action == "editsupport":
        context.user_data["admin_action"] = "support_link"
        current = get_support_link()
        await q.edit_message_text(
            "💬 Support Link সেট করুন।\n\n"
            "আপনার Support username অথবা link পাঠান।\n"
            "উদাহরণ: @YourSupport অথবা https://t.me/YourSupport\n\n"
            f"বর্তমান Support: {current}\n\nবাতিল করতে Cancel চাপুন।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]])
        )
        return

    if action == "welcome":
        context.user_data["admin_action"] = "welcome_message"
        current = html_to_plain(get_welcome_message())
        await q.edit_message_text(
            "বর্তমান Welcome Message নিচে দেওয়া হলো। নতুন মেসেজ লিখে পাঠান; সেটি Save হবে।\n\n"
            f"{current}\n\nবাতিল করতে Cancel চাপুন।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]])
        )
        return

    if action == "addadmin":
        if not is_main_admin(q.from_user.id):
            await q.answer("শুধু Main Admin এই কাজটি করতে পারবেন।", show_alert=True)
            return
        context.user_data["admin_action"] = "add_admin"
        await q.edit_message_text(
            "যে User-কে Admin করবেন তার numeric Telegram ID পাঠান।\n"
            "শুধু ID গ্রহণ করা হবে।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]])
        )
        return

    # 📝 Edit Category Note: category list and per-category note editor.
    if action == "editcatnote":
        cats = get_categories()
        if not cats:
            await q.edit_message_text("কোনো Category পাওয়া যায়নি।", reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ System", callback_data="adm:system")]
            ]))
            return
        rows = [[InlineKeyboardButton(c["name"], callback_data=f"adm:catnotesel:{c['id']}")] for c in cats]
        rows.append([InlineKeyboardButton("⬅️ System", callback_data="adm:system")])
        await q.edit_message_text("📝 যে Category-এর Submit Note পরিবর্তন করতে চান সেটি Select করুন:",
                                  reply_markup=InlineKeyboardMarkup(rows))
        return

    if action == "catnotesel":
        cat_id = int(data[2])
        with db() as con:
            cat = con.execute("SELECT id,name,submit_note FROM categories WHERE id=?", (cat_id,)).fetchone()
        if not cat:
            await q.edit_message_text("Category not found.", reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ Category List", callback_data="adm:editcatnote")]
            ]))
            return
        context.user_data["admin_action"] = "cat_submit_note"
        context.user_data["admin_submit_note_cat_id"] = cat_id
        current_note = rich_plain(cat["submit_note"]) if cat["submit_note"] else "Default Message (Custom Note সেট করা নেই)"
        await q.edit_message_text(
            f"📂 Category: {cat['name']}\n\nবর্তমান Note:\n{current_note}\n\n"
            "নতুন সম্পূর্ণ Note লিখে পাঠান। একাধিক লাইন ব্যবহার করতে পারবেন।\n"
            "Custom Note সরিয়ে Default Message চালু করতে /resetnote লিখুন।",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]])
        )
        return

    # ❌ Close
    if action == "close":
        await q.edit_message_text("Admin panel closed.")
        return

    # 🏠 Admin Home
    if action == "home":
        with db() as con:
            pending = con.execute("SELECT COUNT(*) FROM submissions WHERE status='PENDING'").fetchone()[0]
            users = con.execute("SELECT COUNT(*) FROM users").fetchone()[0]
            total = con.execute("SELECT COALESCE(SUM(balance),0) FROM users").fetchone()[0]
            pm_open = con.execute("SELECT COUNT(*) FROM payment_methods WHERE is_open=1").fetchone()[0]
        await q.edit_message_text(
            f"🛠 ADMIN CONTROL PANEL\n\n⏳ Pending: {pending}\n👥 Users: {users}\n"
            f"💰 Total balances: {money(total)}\n💳 Open Payment Methods: {pm_open}",
            reply_markup=admin_home_keyboard())
        return

    # 📂 Categories
    if action == "cats":
        page = int(data[2]) if len(data) > 2 else 0
        cats = get_categories()
        rows = [[InlineKeyboardButton(
            f"{'🟢' if c['is_open'] else '🔴'} {c['name']} | {money(c['rate'])}",
            callback_data=f"adm:cat:{c['id']}")] for c in cats[page*8:(page+1)*8]]
        rows.append([InlineKeyboardButton("➕ Add Category", callback_data="adm:addcat")])
        nav = []
        if page > 0: nav.append(InlineKeyboardButton("⬅️ Previous", callback_data=f"adm:cats:{page-1}"))
        if (page+1)*8 < len(cats): nav.append(InlineKeyboardButton("Next ➡️", callback_data=f"adm:cats:{page+1}"))
        if nav: rows.append(nav)
        rows.append([InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")])
        await q.edit_message_text("📂 CATEGORY MANAGER\nSelect a category to edit:", reply_markup=InlineKeyboardMarkup(rows))
        return

    # 📂 Single Category
    if action == "cat":
        cat_id = int(data[2])
        with db() as con:
            c = con.execute("SELECT * FROM categories WHERE id=?", (cat_id,)).fetchone()
        if not c:
            await q.edit_message_text("Category not found."); return
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✏️ Rename", callback_data=f"adm:rename:{cat_id}"),
             InlineKeyboardButton("💰 Edit Rate", callback_data=f"adm:rate:{cat_id}")],
            [InlineKeyboardButton("📝 Notice", callback_data=f"adm:catnotice:{cat_id}"),
             InlineKeyboardButton("👥 Group", callback_data=f"adm:catgroup:{cat_id}")],
            [InlineKeyboardButton("🔄 Open/Close", callback_data=f"adm:toggle:{cat_id}"),
             InlineKeyboardButton("🗑 Delete", callback_data=f"adm:delete:{cat_id}")],
            [InlineKeyboardButton("⬅️ Categories", callback_data="adm:cats:0"),
             InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]
        ])
        await q.edit_message_text(
            f"📂 {c['name']}\nID: {c['id']}\nRate: {money(c['rate'])} per valid number\n"
            f"Status: {'OPEN' if c['is_open'] else 'CLOSED'}\n"
            f"👥 Group: {c['group_id'] or 'Not Set'}\n"
            f"📝 Notice: {rich_plain(c['notice']) if c['notice'] else 'Not Set'}",
            reply_markup=kb)
        return

    # ✏️ Rename / 💰 Rate / 📝 Notice / 👥 Group
    if action in {"rename", "rate", "catnotice", "catgroup"}:
        cat_id = int(data[2])
        context.user_data["admin_action"] = action
        context.user_data["admin_category_id"] = cat_id
        prompt = {
            "rename": "নতুন ক্যাটাগরি নাম লিখুন:",
            "rate": "নতুন Rate লিখুন (যেমন 0.50):",
            "catnotice": "File Submission Notice/Instructions লিখুন:",
            "catgroup": (
                "👥 Private Group যুক্ত করুন:\n"
                "1) বটকে গ্রুপে Admin করুন\n"
                "2) Group ID (যেমন -100xxxxxxxxxx) অথবা public @username/লিংক পাঠান\n"
                f"   অথবা গ্রুপের ভিতরে লিখুন: /setgroup {cat_id}\n"
                "Group সরাতে লিখুন: skip"
            )
        }.get(action, "লিখুন:")
        await q.edit_message_text(
            prompt + "\nবাতিল করতে নিচের ❌ Cancel বাটন চাপুন।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]
            ]))
        return

    # 🔄 Open/Close
    if action == "toggle":
        cat_id = int(data[2])
        with db() as con:
            con.execute("UPDATE categories SET is_open=CASE is_open WHEN 1 THEN 0 ELSE 1 END WHERE id=?", (cat_id,))
        await q.edit_message_text("✅ Category status updated.", reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("⬅️ Back to Categories", callback_data="adm:cats:0")],
            [InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]
        ]))
        return

    # 🗑 Delete — ইউজারদের কাছ থেকে ও Admin-এর ক্যাটাগরি লিস্ট থেকে সম্পূর্ণ সরে যায়
    if action == "delete":
        cat_id = int(data[2])
        with db() as con:
            cat = con.execute("SELECT name FROM categories WHERE id=?", (cat_id,)).fetchone()
            used = con.execute(
                "SELECT COUNT(*) FROM submissions WHERE category_id=? AND status IN ('PENDING','APPROVED')",
                (cat_id,)).fetchone()[0]
            if not cat:
                result = "ক্যাটাগরিটি আগেই মুছে ফেলা হয়েছে।"
            elif used:
                # চলমান সাবমিশনের তথ্য ঠিক রাখতে রো-টি লুকানো থাকে; ইউজার বা ক্যাটাগরি লিস্টে আর দেখাবে না।
                con.execute(
                    "UPDATE categories SET is_deleted=1, is_open=0, name=? WHERE id=?",
                    (f"{cat['name']} (deleted #{cat_id})", cat_id))
                result = (f"🗑 Category deleted। ইউজাররা আর এটি দেখবে না।\n"
                          f"চলমান {used}টি সাবমিশন Submissions মেনু থেকে আগের মতো প্রসেস করা যাবে।")
            else:
                con.execute("DELETE FROM categories WHERE id=?", (cat_id,))
                result = "🗑 Category deleted."
        await q.edit_message_text(result, reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("⬅️ Categories", callback_data="adm:cats:0")],
            [InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]
        ]))
        return

    # ⏭ Skip group link (নতুন ক্যাটাগরি তৈরির সময়)
    if action == "grpskip":
        await q.edit_message_text(
            "✅ Group ছাড়াই ক্যাটাগরি তৈরি আছে। পরে Manage Categories → Group থেকে যুক্ত করতে পারবেন।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("📂 Categories", callback_data="adm:cats:0")],
                [InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]
            ]))
        return

    # ➕ Add Category
    if action == "addcat":
        context.user_data["admin_action"] = "addcat"
        await q.edit_message_text(
            "নতুন ক্যাটাগরির নাম লিখুন, এভাবে: Facebook Numbers | 0.50\n"
            "বাতিল করতে নিচের ❌ Cancel বাটন চাপুন।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]
            ]))
        return

    # =====================================================
    # ============  📥 SUBMISSIONS (Category-wise)  ========
    # =====================================================
    if action == "subs":
        with db() as con:
            cats = con.execute(
                "SELECT c.id, c.name, COUNT(s.id) AS n FROM categories c "
                "LEFT JOIN submissions s ON s.category_id=c.id "
                "AND s.status NOT IN ('REJECTED','FULLY_COMPLETED') GROUP BY c.id "
                "HAVING c.is_deleted=0 OR COUNT(s.id)>0 ORDER BY c.id"
            ).fetchall()
            total = con.execute(
                "SELECT COUNT(*) FROM submissions WHERE status NOT IN ('REJECTED','FULLY_COMPLETED')"
            ).fetchone()[0]
        rows = [[InlineKeyboardButton(f"🌐 All Submissions ({total})", callback_data="adm:subsall:0")]]
        rows += [[InlineKeyboardButton(
            f"📂 {c['name']} ({c['n']})", callback_data=f"adm:subscat:{c['id']}:0"
        )] for c in cats]
        rows.append([InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")])
        await q.edit_message_text(
            "📥 SUBMISSION MANAGER\n\nকোন ক্যাটাগরির সাবমিশন দেখতে চান?",
            reply_markup=InlineKeyboardMarkup(rows))
        return

    # 🌐 All Submissions
    if action == "subsall":
        page = int(data[2]) if len(data) > 2 else 0
        with db() as con:
            rows = con.execute(
                "SELECT s.id, s.status, s.file_name, s.user_id, s.total_items, c.name AS category_name "
                "FROM submissions s LEFT JOIN categories c ON c.id=s.category_id "
                "WHERE s.status NOT IN ('REJECTED','FULLY_COMPLETED') "
                "ORDER BY s.submitted_at DESC LIMIT 200"
            ).fetchall()
        batch = rows[page*8:(page+1)*8]
        buttons = [[InlineKeyboardButton(
            f"{s['status']} | {s['file_name'][:18]} | {(s['category_name'] or '-')[:12]} | {s['total_items']}",
            callback_data=f"adm:sub:{s['id']}")] for s in batch]
        nav = []
        if page > 0: nav.append(InlineKeyboardButton("⬅️ Previous", callback_data=f"adm:subsall:{page-1}"))
        if (page+1)*8 < len(rows): nav.append(InlineKeyboardButton("Next ➡️", callback_data=f"adm:subsall:{page+1}"))
        if nav: buttons.append(nav)
        buttons.append([InlineKeyboardButton("📂 Categories", callback_data="adm:subs"),
                        InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")])
        await q.edit_message_text(
            f"📥 ALL SUBMISSIONS\nTotal: {len(rows)}",
            reply_markup=InlineKeyboardMarkup(buttons))
        return

    # 📂 Category-wise Submissions — ফাইল, নির্দিষ্ট Caption ও Approve/Reject
    if action == "subscat":
        cat_id = int(data[2])
        page = int(data[3]) if len(data) > 3 else 0
        with db() as con:
            cat = con.execute("SELECT * FROM categories WHERE id=?", (cat_id,)).fetchone()
            rows = con.execute(
                "SELECT s.*, u.username, u.first_name FROM submissions s "
                "LEFT JOIN users u ON u.user_id=s.user_id "
                "WHERE s.category_id=? AND s.status NOT IN ('REJECTED','FULLY_COMPLETED') "
                "ORDER BY s.submitted_at DESC",
                (cat_id,)
            ).fetchall()
        if not cat:
            await q.edit_message_text("Category not found."); return

        batch = rows[page*8:(page+1)*8]
        for submission in batch:
            username = (f"@{submission['username']}" if submission['username']
                        else (submission['first_name'] or str(submission['user_id'])))
            caption = (
                f"👤 User: {username}\n"
                f"📂 Category: {cat['name']}\n"
                f"💵 Rate: {money(submission['rate'])}\n"
                f"💳 Payment Method: {submission['payment_method_name'] or '-'}\n"
                f"🔑 Payment Details: {(submission['payment_details'] or '-')[:300]}"
            )
            try:
                await context.bot.send_document(
                    chat_id=q.message.chat_id,
                    document=submission['file_id'],
                    caption=caption,
                    reply_markup=(category_submission_keyboard(submission['id'])
                                  if submission['status'] == 'PENDING'
                                  else approved_submission_keyboard(submission['id']))
                )
            except Exception:
                log.exception("Could not send category submission file %s", submission['id'])
                await context.bot.send_message(
                    chat_id=q.message.chat_id,
                    text=f"⚠️ Submission {submission['id']}-এর ফাইল পাঠানো যায়নি।"
                )

        buttons = []
        nav = []
        if page > 0:
            nav.append(InlineKeyboardButton("⬅️ Previous", callback_data=f"adm:subscat:{cat_id}:{page-1}"))
        if (page+1)*8 < len(rows):
            nav.append(InlineKeyboardButton("Next ➡️", callback_data=f"adm:subscat:{cat_id}:{page+1}"))
        if nav:
            buttons.append(nav)
        buttons.append([InlineKeyboardButton("📂 All Categories", callback_data="adm:subs"),
                        InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")])
        await q.edit_message_text(
            f"📥 SUBMISSIONS — {cat['name']}\nTotal: {len(rows)}\n"
            f"এই পেজে ফাইল: {len(batch)}",
            reply_markup=InlineKeyboardMarkup(buttons))
        return

    # 📄 Single Submission (File + User + Category + Payment Details)
    if action == "sub":
        sid = data[2]
        with db() as con:
            s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
            u = con.execute(
                "SELECT user_id,username,first_name FROM users WHERE user_id=?",
                (s["user_id"],)
            ).fetchone() if s else None
            c = con.execute(
                "SELECT name FROM categories WHERE id=?", (s["category_id"],)
            ).fetchone() if s else None
            r = con.execute(
                "SELECT COUNT(*) AS n, COALESCE(SUM(amount),0) AS total FROM payment_reports WHERE submission_id=?",
                (sid,)
            ).fetchone() if s else None
        if not s:
            await q.edit_message_text("Submission not found."); return

        user_name = u["first_name"] if u else str(s["user_id"])
        user_tag = f" (@{u['username']})" if u and u["username"] else ""
        cat_name = c["name"] if c else str(s["category_id"])
        report_count = r["n"] if r else 0
        report_total = r["total"] if r else 0

        detail_text = (
            f"📄 SUBMISSION: {sid}\n"
            f"👤 User: {user_name}{user_tag}\n"
            f"🆔 User ID: {s['user_id']}\n"
            f"📂 Category: {cat_name}\n"
            f"📎 File: {s['file_name']}\n"
            f"🔢 Detected rows: {s['total_items']}\n"
            f"📌 Status: {s['status']}\n"
            f"✅ Valid: {s['valid_items'] or 0} | ❌ Invalid: {s['invalid_items'] or 0}\n"
            f"💵 Rate snapshot: {money(s['rate'])}\n"
            f"💰 Amount: {money(s['amount'] or 0)}\n"
            f"💳 Payment Method: {s['payment_method_name'] or '-'}\n"
            f"🔑 Payment Details: {s['payment_details'] or '-'}\n"
            f"💬 Message: {s['user_message'] or '-'}\n"
            f"📨 Reports sent: {report_count} (Total {money(report_total)})"
        )

        status = s["status"]
        kb_rows = [
            [InlineKeyboardButton("📄 Review", callback_data=f"adm:review:{sid}")],
            [InlineKeyboardButton("⬇️ Download File", callback_data=f"adm:download:{sid}")],
        ]
        if status == "PENDING":
            kb_rows.append([InlineKeyboardButton("✅ Approve", callback_data=f"adm:approve:{sid}")])
        if status == "APPROVED":
            kb_rows.append([InlineKeyboardButton("📤 Send Report", callback_data=f"adm:report:{sid}")])
        if status in ("PENDING", "APPROVED"):
            kb_rows.append([InlineKeyboardButton("❌ Reject", callback_data=f"adm:reject:{sid}")])
        kb_rows.append([InlineKeyboardButton("⬅️ Submissions", callback_data="adm:subs")])
        kb = InlineKeyboardMarkup(kb_rows)
        await q.edit_message_text(detail_text, reply_markup=kb)
        return

    # 📄 Review (ফাইলসহ বিস্তারিত)
    if action == "review":
        sid = data[2]
        with db() as con:
            s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
            c = con.execute(
                "SELECT name FROM categories WHERE id=?", (s["category_id"],)
            ).fetchone() if s else None
        if not s:
            await q.edit_message_text("Submission not found."); return
        cat_name = c["name"] if c else str(s["category_id"])
        try:
            await context.bot.send_document(
                chat_id=q.message.chat_id, document=s["file_id"],
                caption=(f"📄 Submission {sid}\n👤 User ID: {s['user_id']}\n"
                         f"📂 Category: {cat_name}\n📎 File: {s['file_name']}\n"
                         f"🔢 Rows: {s['total_items']}\n"
                         f"💳 Payment: {s['payment_method_name'] or '-'} — {s['payment_details'] or '-'}\n"
                         f"📌 Status: {s['status']}")
            )
        except Exception:
            log.exception("Could not send review document")
            await context.bot.send_message(
                chat_id=q.message.chat_id,
                text="⚠️ ফাইলটি পাঠানো যায়নি। ⬇️ Download File বাটন চেষ্টা করুন।")
        return

    # ⬇️ Download File
    if action == "download":
        sid = data[2]
        with db() as con:
            s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
        if s:
            await context.bot.send_document(q.message.chat_id, s["file_id"], caption=f"{sid} — {s['file_name']}")
        return

    # ✅ Approve — Status/Notification processing অপরিবর্তিত; media caption-ও নিরাপদে update
    if action in {"settle", "approve"}:
        sid = data[2]
        ok, err = await approve_submission(context, sid)
        if not ok:
            await context.bot.send_message(chat_id=q.from_user.id, text=err)
            return
        approved_text = (
            f"✅ APPROVED\n🆔 Submission: {sid}\n\n"
            "ইউজারকে Approval Notification পাঠানো হয়েছে।"
        )
        try:
            if q.message.document or q.message.photo:
                await q.edit_message_caption(
                    caption=q.message.caption or "",
                    reply_markup=approved_submission_keyboard(sid)
                )
            else:
                await q.edit_message_text(
                    approved_text + "\nপেমেন্ট জানাতে 📤 Send Report চাপুন।",
                    reply_markup=approved_submission_keyboard(sid)
                )
        except Exception:
            log.exception("Could not update admin submission message after approval")
            await context.bot.send_message(
                chat_id=q.from_user.id,
                text=approved_text + "\nপেমেন্ট জানাতে Send Report ব্যবহার করুন।",
                reply_markup=approved_submission_keyboard(sid)
            )
        return

    # ❌ Reject — এক ট্যাপেই Submission সম্পূর্ণ মুছে যায় (Database + Admin Panel + Group)
    if action == "reject":
        sid = data[2]
        with db() as con:
            current = con.execute("SELECT status FROM submissions WHERE id=?", (sid,)).fetchone()
        if not current or current['status'] not in ("PENDING", "APPROVED"):
            await q.answer("এই Submission-এ এখন Reject করা যাবে না।", show_alert=True)
            try:
                await q.message.delete()
            except Exception:
                pass
            return
        ok, err = await delete_submission_completely(context, sid)
        await q.answer("🗑 Rejected — সম্পূর্ণ মুছে ফেলা হয়েছে।" if ok else err, show_alert=not ok)
        if ok:
            try:
                await q.message.delete()  # Admin Panel থেকে ফাইলের মেসেজ সরানো
            except Exception:
                try:
                    await q.message.edit_reply_markup(reply_markup=None)
                except Exception:
                    pass
        return

    # Cancel admin text input
    if action == "cancelinput":
        clear_admin_state(context)
        await q.edit_message_text("Cancelled.", reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]
        ]))
        return

    # =====================================================
    # ============  💳 PAYMENT METHOD MANAGER  ============
    # =====================================================
    if action == "pms":
        page = int(data[2]) if len(data) > 2 else 0
        rows = get_payment_methods()
        batch = rows[page*8:(page+1)*8]
        buttons = [[InlineKeyboardButton(
            f"{'🟢' if m['is_open'] else '🔴'} {m['name']}"
            + (f" | {m['account_info'][:16]}" if m["account_info"] else ""),
            callback_data=f"adm:pm:{m['id']}")] for m in batch]
        nav = []
        if page > 0: nav.append(InlineKeyboardButton("⬅️ Previous", callback_data=f"adm:pms:{page-1}"))
        if (page+1)*8 < len(rows): nav.append(InlineKeyboardButton("Next ➡️", callback_data=f"adm:pms:{page+1}"))
        if nav: buttons.append(nav)
        buttons.append([InlineKeyboardButton("➕ Add Payment Method", callback_data="adm:addpm")])
        buttons.append([InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")])
        await q.edit_message_text(
            "💳 PAYMENT METHOD MANAGER\n\nএকটি মেথড সিলেক্ট করে Edit করুন:",
            reply_markup=InlineKeyboardMarkup(buttons))
        return

    if action == "pm":
        pm_id = int(data[2])
        with db() as con:
            m = con.execute("SELECT * FROM payment_methods WHERE id=?", (pm_id,)).fetchone()
        if not m:
            await q.edit_message_text("Payment method not found."); return
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✏️ Edit Name", callback_data=f"adm:pmname:{pm_id}"),
             InlineKeyboardButton("📝 Edit Account Info", callback_data=f"adm:pmacc:{pm_id}")],
            [InlineKeyboardButton("🔄 Open/Close", callback_data=f"adm:pmtoggle:{pm_id}"),
             InlineKeyboardButton("🗑 Delete", callback_data=f"adm:pmdelete:{pm_id}")],
            [InlineKeyboardButton("⬅️ Payment Methods", callback_data="adm:pms:0"),
             InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]
        ])
        await q.edit_message_text(
            f"💳 {m['name']}\nID: {m['id']}\nAccount: {m['account_info'] or '—'}\n"
            f"Status: {'OPEN' if m['is_open'] else 'CLOSED'}",
            reply_markup=kb)
        return

    if action in {"pmname", "pmacc"}:
        pm_id = int(data[2])
        context.user_data["admin_action"] = action
        context.user_data["admin_pm_id"] = pm_id
        prompt = ("নতুন Payment Method-এর নাম লিখুন:" if action == "pmname"
                  else "নতুন Account Info লিখুন (যেমন 01XXXXXXXXX অথবা ID):")
        await q.edit_message_text(
            prompt + "\nবাতিল করতে নিচের ❌ Cancel বাটন চাপুন।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]
            ]))
        return

    if action == "pmtoggle":
        pm_id = int(data[2])
        with db() as con:
            con.execute("UPDATE payment_methods SET is_open=CASE is_open WHEN 1 THEN 0 ELSE 1 END WHERE id=?", (pm_id,))
        await q.edit_message_text("✅ Payment method status updated.", reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("⬅️ Payment Methods", callback_data="adm:pms:0")],
            [InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]
        ]))
        return

    if action == "pmdelete":
        pm_id = int(data[2])
        with db() as con:
            used = con.execute(
                "SELECT COUNT(*) FROM submissions WHERE payment_method_id=?", (pm_id,)
            ).fetchone()[0]
            if used:
                await q.edit_message_text(
                    "এই Payment Method দিয়ে সাবমিশন আছে, তাই নিরাপত্তার জন্য Delete করা যাবে না। "
                    "চাইলে Open/Close করে বন্ধ রাখুন।",
                    reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Payment Methods", callback_data="adm:pms:0")]]))
                return
            con.execute("DELETE FROM payment_methods WHERE id=?", (pm_id,))
        await q.edit_message_text("🗑 Payment method deleted.", reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("⬅️ Payment Methods", callback_data="adm:pms:0")],
            [InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]
        ]))
        return

    if action == "addpm":
        context.user_data["admin_action"] = "addpm"
        await q.edit_message_text(
            "নতুন Payment Method লিখুন, এভাবে: bKash | 01XXXXXXXXX\n"
            "বাতিল করতে নিচের ❌ Cancel বাটন চাপুন।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]
            ]))
        return

    # =====================================================
    # ============  📤 SEND PAYMENT REPORT  ===============
    # =====================================================
    if action == "report":
        sid = data[2]
        with db() as con:
            s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
        if not s:
            await context.bot.send_message(chat_id=q.from_user.id, text="Submission not found.")
            return
        if s["status"] != "APPROVED":
            await context.bot.send_message(
                chat_id=q.from_user.id,
                text="Send Report পাঠানোর আগে Submission Approve করতে হবে।"
            )
            return
        context.user_data["admin_action"] = "report_amount"
        context.user_data["report_sid"] = sid
        context.user_data.pop("report_amount", None)
        context.user_data.pop("report_message", None)
        report_prompt = (
            f"💳 SEND PAYMENT REPORT\nSubmission: {sid}\n\n"
            "কত টাকা Payment পাঠাবেন তা লিখুন (শুধু সংখ্যা, যেমন 150)।\n"
            "বাতিল করতে নিচের ❌ Cancel বাটন চাপুন।"
        )
        try:
            # Callback টি document/caption message থেকেও আসতে পারে; তাই edit_message_text নয়।
            await context.bot.send_message(
                chat_id=q.from_user.id, text=report_prompt,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("❌ Cancel", callback_data=f"adm:reportcancel:{sid}")]
                ])
            )
        except Exception:
            clear_admin_state(context)
            log.exception("Could not send payment report prompt to admin")
            return
        return

    if action == "reportcancel":
        clear_admin_state(context)
        await q.edit_message_text("Cancelled.", reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]
        ]))
        return

    if action == "reportskip":
        sid = data[2]
        amount = context.user_data.get("report_amount")
        message = context.user_data.get("report_message", "")
        if amount is None:
            await q.edit_message_text("Amount পাওয়া যায়নি। আবার শুরু করুন।")
            return
        with db() as con:
            s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
        if not s:
            await q.edit_message_text("Submission not found."); return
        if s["status"] != "APPROVED":
            clear_admin_state(context)
            await q.edit_message_text("এই Submission আর Approved অবস্থায় নেই; Report পাঠানো হয়নি।")
            return
        ok = await send_payment_report(
            context, q.message.chat_id, sid, s["user_id"], amount, message
        )
        clear_admin_state(context)
        if ok:
            await q.edit_message_text(
                f"✅ Payment Report পাঠানো হয়েছে।\nSubmission: {sid}\nAmount: {money(amount)}\n"
                "🏁 FULLY COMPLETED — ফাইলটি Admin Panel থেকে সরানো হয়েছে।",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]]))
        return

    # =====================================================
    # ====================  👥 USERS  =====================
    # =====================================================
    if action == "users":
        page = int(data[2]) if len(data) > 2 else 0
        with db() as con:
            rows = con.execute("SELECT user_id,first_name,balance FROM users ORDER BY user_id DESC LIMIT 100").fetchall()
        buttons = [[InlineKeyboardButton(
            f"{u['first_name'][:18]} | {u['user_id']} | {money(u['balance'])}",
            callback_data=f"adm:user:{u['user_id']}")] for u in rows[page*8:(page+1)*8]]
        nav = []
        if page > 0: nav.append(InlineKeyboardButton("⬅️ Previous", callback_data=f"adm:users:{page-1}"))
        if (page+1)*8 < len(rows): nav.append(InlineKeyboardButton("Next ➡️", callback_data=f"adm:users:{page+1}"))
        if nav: buttons.append(nav)
        buttons.append([InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")])
        await q.edit_message_text("👥 USER MANAGER\nSelect a user:", reply_markup=InlineKeyboardMarkup(buttons))
        return

    if action == "user":
        uid = int(data[2])
        with db() as con:
            u = con.execute("SELECT * FROM users WHERE user_id=?", (uid,)).fetchone()
            n = con.execute("SELECT COUNT(*) FROM submissions WHERE user_id=?", (uid,)).fetchone()[0]
        if not u:
            await q.edit_message_text("User not found."); return
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("➕ Add Balance", callback_data=f"adm:baladd:{uid}"),
             InlineKeyboardButton("➖ Deduct Balance", callback_data=f"adm:balminus:{uid}")],
            [InlineKeyboardButton("⬅️ Users", callback_data="adm:users:0")]
        ])
        await q.edit_message_text(
            f"👤 {u['first_name']}\nUser ID: {uid}\nUsername: @{u['username']}\n"
            f"Balance: {money(u['balance'])}\nSubmissions: {n}",
            reply_markup=kb)
        return

    if action in {"baladd", "balminus"}:
        uid = int(data[2])
        context.user_data["admin_action"] = "balance"
        context.user_data["balance_user_id"] = uid
        context.user_data["balance_direction"] = 1 if action == "baladd" else -1
        label = "কত টাকা Add করবেন?" if action == "baladd" else "কত টাকা Deduct করবেন?"
        await q.edit_message_text(
            label + "\nবাতিল করতে নিচের ❌ Cancel বাটন চাপুন।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")]
            ]))
        return

    # 📊 Statistics
    if action == "stats":
        with db() as con:
            total_sub = con.execute("SELECT COUNT(*) FROM submissions").fetchone()[0]
            approved = con.execute("SELECT COUNT(*) FROM submissions WHERE status IN ('APPROVED','FULLY_COMPLETED')").fetchone()[0]
            pending = con.execute("SELECT COUNT(*) FROM submissions WHERE status='PENDING'").fetchone()[0]
            paid = con.execute("SELECT COALESCE(SUM(amount),0) FROM payments").fetchone()[0]
            reported = con.execute("SELECT COALESCE(SUM(amount),0) FROM payment_reports").fetchone()[0]
            users = con.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        await q.edit_message_text(
            f"📊 STATISTICS\nUsers: {users}\nSubmissions: {total_sub}\n"
            f"Approved: {approved}\nPending: {pending}\nTotal credited: {money(paid)}\n"
            f"Total reported paid: {money(reported)}",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🏠 Admin Home", callback_data="adm:home")]]))
        return


# ========================================================
# ==============  📤 PAYMENT REPORT HELPERS  =============
# ========================================================

async def send_payment_report(context, admin_chat_id, submission_id, user_id,
                              amount, message, file_id=None, is_photo=False):
    """ইউজারকে Payment Report পাঠায়, record রাখে এবং Status = FULLY_COMPLETED করে।"""
    caption = (
        f"💳 <b>{fancy_text('PAYMENT REPORT', 'bold')}</b>\n\n"
        f"🆔 <b>Submission:</b> <code>{html.escape(str(submission_id))}</code>\n"
        f"💰 <b>Amount:</b> {html.escape(money(amount))}\n"
    )
    if message:
        caption += f"\n<blockquote>{rich_html(message)}</blockquote>\n"
    caption += "\n🙏 ধন্যবাদ!"
    caption = RawHTML(caption)

    try:
        if file_id and is_photo:
            await context.bot.send_photo(chat_id=user_id, photo=file_id, caption=caption)
        elif file_id:
            await context.bot.send_document(chat_id=user_id, document=file_id, caption=caption)
        else:
            await context.bot.send_message(chat_id=user_id, text=caption)
    except Exception:
        log.exception("Could not send payment report to user %s", user_id)
        try:
            await context.bot.send_message(
                chat_id=admin_chat_id,
                text="⚠️ ইউজারকে মেসেজ পাঠানো যায়নি (Bot Blocked অথবা ভুল User ID)।"
            )
        except Exception:
            log.exception("Could not notify admin about failed report")
        return False

    with db() as con:
        con.execute(
            "INSERT INTO payment_reports(submission_id,user_id,amount,message,file_id,created_at) "
            "VALUES(?,?,?,?,?,?)",
            (submission_id, user_id, amount, message, file_id, now())
        )
        # Approve করা submission-এর report সফল হলে FULLY_COMPLETED (Admin Panel থেকে সরে যায়)
        con.execute(
            "UPDATE submissions SET status='FULLY_COMPLETED', paid_at=? WHERE id=? AND status='APPROVED'",
            (now(), submission_id)
        )

    # Group message: "FULLY COMPLETED" দেখাবে, বাটন থাকবে না
    await refresh_group_message(context, submission_id)
    return True


async def handle_report_screenshot(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.file_size and doc.file_size > MAX_FILE_MB * 1024 * 1024:
        await update.message.reply_text(f"❌ ছবির সাইজ সর্বোচ্চ {MAX_FILE_MB} MB হতে পারবে।")
        return

    sid = context.user_data.get("report_sid")
    amount = context.user_data.get("report_amount")
    message = context.user_data.get("report_message", "")

    with db() as con:
        s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
    if not s:
        clear_admin_state(context)
        await update.message.reply_text("Submission not found.")
        return
    if s["status"] != "APPROVED" or amount is None:
        clear_admin_state(context)
        await update.message.reply_text("এই Submission Approved নয় বা Report Amount পাওয়া যায়নি। Report পাঠানো হয়নি।")
        return

    is_photo = (doc.mime_type or "").startswith("image/")
    ok = await send_payment_report(
        context, update.effective_chat.id, sid, s["user_id"],
        amount, message, doc.file_id, is_photo
    )
    clear_admin_state(context)

    if ok:
        await update.message.reply_text(
            f"✅ Payment Report পাঠানো হয়েছে।\n"
            f"Submission: {sid}\nAmount: {money(amount)}\n"
            "🏁 FULLY COMPLETED — ফাইলটি Admin Panel থেকে সরানো হয়েছে।",
            reply_markup=main_keyboard(update.effective_user.id)
        )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin সাধারণ Photo হিসেবে Payment Screenshot পাঠালে।"""
    if not (context.user_data.get("admin_action") == "report_screenshot" and is_admin(update.effective_user.id)):
        return
    photo = update.message.photo[-1]
    sid = context.user_data.get("report_sid")
    amount = context.user_data.get("report_amount")
    message = context.user_data.get("report_message", "")
    with db() as con:
        s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
    if not s:
        clear_admin_state(context)
        await update.message.reply_text("Submission not found.")
        return
    if s["status"] != "APPROVED" or amount is None:
        clear_admin_state(context)
        await update.message.reply_text("এই Submission Approved নয় বা Report Amount পাওয়া যায়নি। Report পাঠানো হয়নি।")
        return
    ok = await send_payment_report(
        context, update.effective_chat.id, sid, s["user_id"], amount, message, photo.file_id, True)
    clear_admin_state(context)
    if ok:
        await update.message.reply_text(
            f"✅ Payment Report পাঠানো হয়েছে।\n"
            f"Submission: {sid}\nAmount: {money(amount)}\n"
            "🏁 FULLY COMPLETED — ফাইলটি Admin Panel থেকে সরানো হয়েছে।",
            reply_markup=main_keyboard(update.effective_user.id)
        )


# ========================================================
# =================  ADMIN TEXT INPUT  ===================
# ========================================================

async def admin_text_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return False

    action = context.user_data.get("admin_action")
    if not action:
        return False

    raw = (update.message.text or "").strip()

    if action == "emoji_single_wait":
        return await emoji_single_receive(update, context)

    # Broadcast মোড: টেক্সটও ব্রডকাস্ট কনটেন্ট — broadcast_capture হ্যান্ডলার ধরবে
    if action == "broadcast":
        return True

    # Screenshot ধাপে টেক্সট এলে সহায়ক মেসেজ
    if action == "report_screenshot":
        await update.message.reply_text(
            "এখন Screenshot পাঠান, অথবা ⏭ Skip বাটন চাপুন।",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⏭ Skip", callback_data=f"adm:reportskip:{context.user_data.get('report_sid')}")],
                [InlineKeyboardButton("❌ Cancel", callback_data="adm:cancelinput")],
            ]))
        return True

    try:
        with db() as con:
            # ---------- System: Add Force Join Channel / Group ----------
            if action == "force_join_add":
                if not raw:
                    await update.message.reply_text("লিংক খালি রাখা যাবে না।")
                    return True
                parts = [part.strip() for part in raw.split("|", 1)]
                link = parts[0]
                chat_id = None
                if len(parts) == 2:
                    if not re.fullmatch(r"-?\d+", parts[1]):
                        await update.message.reply_text("chat_id অবশ্যই সংখ্যায় দিতে হবে।")
                        return True
                    chat_id = int(parts[1])
                else:
                    # Public channel links and @usernames can be resolved by Telegram.
                    candidate = link
                    match = re.search(r"(?:https?://)?(?:www\.)?(?:t\.me|telegram\.me)/([A-Za-z0-9_]{5,})", link)
                    if match:
                        candidate = "@" + match.group(1)
                    elif link.startswith("@"):
                        candidate = link
                    else:
                        await update.message.reply_text(
                            "প্রাইভেট invite link হলে এভাবে পাঠান:\ninvite_link | -100xxxxxxxxxx"
                        )
                        return True
                    try:
                        chat = await context.bot.get_chat(candidate)
                        chat_id = int(chat.id)
                    except TelegramError:
                        await update.message.reply_text(
                            "চ্যানেল শনাক্ত করা যায়নি। বটকে চ্যানেলে Admin করুন অথবা প্রাইভেট লিংকের সঙ্গে chat_id দিন।"
                        )
                        return True
                try:
                    chat = await context.bot.get_chat(chat_id)
                    title = getattr(chat, "title", None) or getattr(chat, "username", None) or str(chat_id)
                    # Verify the bot can inspect membership (requires admin rights in channels).
                    bot_member = await context.bot.get_chat_member(chat_id=chat_id, user_id=context.bot.id)
                    if str(bot_member.status).lower() not in ("administrator", "creator"):
                        await update.message.reply_text("বটকে ওই চ্যানেল/গ্রুপে Admin করে আবার চেষ্টা করুন।")
                        return True
                except TelegramError:
                    await update.message.reply_text(
                        "চ্যানেল/গ্রুপ যাচাই করা যায়নি। সঠিক chat_id দিন এবং বটকে সেখানে Admin করুন।"
                    )
                    return True
                if not (link.startswith("https://") or link.startswith("http://")):
                    if link.startswith("@"):
                        link = "https://t.me/" + link[1:]
                    elif getattr(chat, "username", None):
                        link = "https://t.me/" + chat.username
                    else:
                        await update.message.reply_text("প্রাইভেট চ্যানেল/গ্রুপের সম্পূর্ণ invite link দিতে হবে।")
                        return True
                try:
                    with db() as con:
                        con.execute(
                            "INSERT INTO force_join_channels(chat_id,join_link,title,added_at) VALUES(?,?,?,?)",
                            (chat_id, link, title, now())
                        )
                except sqlite3.IntegrityError:
                    await update.message.reply_text("এই চ্যানেল/গ্রুপটি আগে থেকেই যুক্ত আছে।")
                    return True
                context.user_data.pop("admin_action", None)
                await update.message.reply_text(f"✅ Force to Join-এ যুক্ত হয়েছে: {title}")
                return True

            # ---------- System: Premium Emoji Pack Import ----------
            if action == "premium_pack_link":
                names = []
                for token in re.split(r"[\s,]+", raw):
                    m = re.search(r"(?:addemoji|addstickers)/([A-Za-z0-9_]+)", token) or re.fullmatch(r"([A-Za-z0-9_]{3,})", token)
                    if m and m.group(1) not in names:
                        names.append(m.group(1))
                if not names:
                    await update.message.reply_text("সঠিক Emoji Pack লিংক দিন। উদাহরণ: https://t.me/addemoji/PackName")
                    return True
                merged = dict(get_premium_emoji_map())
                report = []
                added = {}
                for name in names[:5]:
                    try:
                        found, total_in_pack = await import_emoji_pack(context.bot, name)
                        merged.update(found)
                        added.update(found)
                        report.append(f"✅ {name}: প্যাকে {total_in_pack}টি, মিলেছে {len(found)}টি")
                    except Exception as exc:
                        log.warning("Emoji pack import failed for %s: %s", name, exc)
                        report.append(f"❌ {name}: প্যাকটি পাওয়া যায়নি বা Custom Emoji প্যাক নয়")
                context.user_data.pop("admin_action", None)
                if added:
                    save_premium_emoji_map(merged)
                await update.message.reply_text(
                    "✨ Emoji Pack Import\n\n" + "\n".join(report) + "\n\n"
                    f"মোট প্রিমিয়াম ইমোজি এখন: {len(merged)}টি"
                    + (("\n\nপ্রিভিউ: " + " ".join(list(added)[:40])) if added else ""),
                    reply_markup=main_keyboard(update.effective_user.id))
                return True

            # ---------- System: Price and Rules ----------
            if action == "price_rules_text":
                if not raw or len(raw) > 4000:
                    await update.message.reply_text("লেখা খালি রাখা যাবে না এবং সর্বোচ্চ ৪০০০ অক্ষর হতে পারবে।")
                    return True
                formatted = update.message.text_html or html.escape(raw, quote=False)
                set_setting("price_rules_html", formatted)
                context.user_data.pop("admin_action", None)
                await update.message.reply_text("✅ Price and Rules সেভ হয়েছে। ইউজার এভাবে দেখবে:")
                await send_price_rules(context, update.effective_chat.id,
                                       main_keyboard(update.effective_user.id), custom_html=formatted)
                return True

            # ---------- System: How to Use (শুধু লেখা) ----------
            if action == "howto_content":
                if not raw or len(raw) > 4000:
                    await update.message.reply_text("লেখা খালি রাখা যাবে না এবং সর্বোচ্চ ৪০০০ অক্ষর হতে পারবে।")
                    return True
                set_setting("howto_text_html", update.message.text_html or html.escape(raw, quote=False))
                context.user_data.pop("admin_action", None)
                await update.message.reply_text("✅ How to Use সেভ হয়েছে। ইউজার এভাবে দেখবে:")
                await send_howto(context, update.effective_chat.id, main_keyboard(update.effective_user.id))
                return True

            # ---------- System: Payment Issue Message ----------
            if action == "payment_issue_message":
                if not raw or len(raw) > 3500:
                    await update.message.reply_text("মেসেজ খালি রাখা যাবে না এবং সর্বোচ্চ ৩৫০০ অক্ষর হতে পারবে।")
                    return True
                con.execute(
                    "INSERT INTO bot_settings(setting_key,setting_value) VALUES('payment_issue_message',?) "
                    "ON CONFLICT(setting_key) DO UPDATE SET setting_value=excluded.setting_value",
                    (rich_store(update.message),)
                )
                context.user_data.pop("admin_action", None)
                await update.message.reply_text("✅ Payment Issue Message সফলভাবে Save হয়েছে।")
                return True

            # ---------- System: Support Link ----------
            if action == "support_link":
                support_url = normalize_support_url(raw)
                if not support_url:
                    await update.message.reply_text(
                        "সঠিক Support username বা Telegram link দিন।\n"
                        "উদাহরণ: @YourSupport অথবা https://t.me/YourSupport"
                    )
                    return True
                con.execute(
                    "INSERT INTO bot_settings(setting_key,setting_value) VALUES('support_link',?) "
                    "ON CONFLICT(setting_key) DO UPDATE SET setting_value=excluded.setting_value",
                    (raw,)
                )
                context.user_data.pop("admin_action", None)
                await update.message.reply_text("✅ Support Link সফলভাবে Save হয়েছে।")
                return True

            # ---------- System: Welcome Message ----------
            if action == "welcome_message":
                if not raw or len(raw) > 3500:
                    await update.message.reply_text("মেসেজ খালি রাখা যাবে না এবং সর্বোচ্চ ৩৫০০ অক্ষর হতে পারবে।")
                    return True
                con.execute(
                    "INSERT INTO bot_settings(setting_key,setting_value) VALUES('welcome_message',?) "
                    "ON CONFLICT(setting_key) DO UPDATE SET setting_value=excluded.setting_value",
                    (rich_store(update.message),)
                )
                context.user_data.pop("admin_action", None)
                await update.message.reply_text("✅ Welcome Message সফলভাবে Save হয়েছে।")
                return True

            # ---------- System: Category-specific Submit Note ----------
            if action == "cat_submit_note":
                cat_id = context.user_data.get("admin_submit_note_cat_id")
                if raw == "/resetnote":
                    con.execute("UPDATE categories SET submit_note=NULL WHERE id=?", (cat_id,))
                    context.user_data.pop("admin_action", None)
                    context.user_data.pop("admin_submit_note_cat_id", None)
                    await update.message.reply_text("✅ Custom Note সরানো হয়েছে। এখন এই Category-তে Default Message দেখাবে।")
                    return True
                if not raw or len(raw) > 3500:
                    await update.message.reply_text("Note খালি রাখা যাবে না এবং সর্বোচ্চ ৩৫০০ অক্ষর হতে পারবে।")
                    return True
                cur = con.execute("UPDATE categories SET submit_note=? WHERE id=?", (rich_store(update.message), cat_id))
                if not cur.rowcount:
                    await update.message.reply_text("Category পাওয়া যায়নি।")
                    return True
                context.user_data.pop("admin_action", None)
                context.user_data.pop("admin_submit_note_cat_id", None)
                await update.message.reply_text("✅ এই Category-এর Submit Note সফলভাবে Save হয়েছে।")
                return True

            # ---------- System: Add Admin ----------
            if action == "add_admin":
                if not is_main_admin(update.effective_user.id):
                    context.user_data.pop("admin_action", None)
                    await update.message.reply_text("শুধু Main Admin এই কাজটি করতে পারবেন।")
                    return True
                if not raw.isdigit() or int(raw) <= 0:
                    await update.message.reply_text("সঠিক numeric Telegram ID পাঠান।")
                    return True
                new_admin_id = int(raw)
                if new_admin_id in ADMIN_IDS:
                    context.user_data.pop("admin_action", None)
                    await update.message.reply_text("এই User আগে থেকেই Main Admin হিসেবে সেট করা আছে।")
                    return True
                con.execute(
                    "INSERT OR IGNORE INTO additional_admins(user_id,added_by,added_at) VALUES(?,?,?)",
                    (new_admin_id, update.effective_user.id, now())
                )
                context.user_data.pop("admin_action", None)
                await update.message.reply_text(
                    f"✅ Telegram ID {new_admin_id} এখন Admin। তাকে বটের /start দিতে বলুন।"
                )
                try:
                    await context.bot.send_message(new_admin_id, "আপনাকে এই বটের Admin হিসেবে যুক্ত করা হয়েছে। /start দিন।")
                except Exception:
                    pass
                return True

            # ---------- Category: Add ----------
            if action == "addcat":
                if "|" not in raw:
                    await update.message.reply_text("এইভাবে লিখুন: Category Name | Rate")
                    return True
                name, rate = [x.strip() for x in raw.split("|", 1)]
                rate = float(rate)
                if not name or rate < 0: raise ValueError()
                cur = con.execute("INSERT INTO categories(name,rate,is_open) VALUES(?,?,1)", (name, rate))
                new_id = cur.lastrowid
                await update.message.reply_text(
                    f"✅ নতুন ক্যাটাগরি তৈরি হয়েছে: {name} — {money(rate)}\n\n"
                    "এই ক্যাটাগরির জন্য Private Group যুক্ত করতে চান?",
                    reply_markup=InlineKeyboardMarkup([
                        [InlineKeyboardButton("🔗 Send Group Link", callback_data=f"adm:catgroup:{new_id}")],
                        [InlineKeyboardButton("⏭ Skip", callback_data=f"adm:grpskip:{new_id}")],
                    ]))
                context.user_data.pop("admin_action", None)
                return True

            # ---------- Category: Rename ----------
            if action == "rename":
                name = raw
                if not name or len(name) > 60: raise ValueError()
                con.execute("UPDATE categories SET name=? WHERE id=?", (name, context.user_data["admin_category_id"]))
                await update.message.reply_text("✅ ক্যাটাগরির নাম পরিবর্তন হয়েছে।")
                context.user_data.pop("admin_action", None)
                context.user_data.pop("admin_category_id", None)
                return True

            # ---------- Category: Rate ----------
            if action == "rate":
                rate = float(raw)
                if rate < 0: raise ValueError()
                con.execute("UPDATE categories SET rate=? WHERE id=?", (rate, context.user_data["admin_category_id"]))
                await update.message.reply_text("✅ ক্যাটাগরির রেট পরিবর্তন হয়েছে।")
                context.user_data.pop("admin_action", None)
                context.user_data.pop("admin_category_id", None)
                return True

            # ---------- Category: Notice ----------
            if action == "catnotice":
                notice = raw
                if len(notice) > 500: raise ValueError()
                con.execute("UPDATE categories SET notice=? WHERE id=?", (rich_store(update.message), context.user_data["admin_category_id"]))
                await update.message.reply_text("✅ ক্যাটাগরির Notice পরিবর্তন হয়েছে।")
                context.user_data.pop("admin_action", None)
                context.user_data.pop("admin_category_id", None)
                return True

            # ---------- Category: Group ----------
            if action == "catgroup":
                if raw.lower() == "skip":
                    group_id = None
                else:
                    group_id, err = await resolve_group(context, raw)
                    if err:
                        await update.message.reply_text(err)
                        return True
                con.execute("UPDATE categories SET group_id=? WHERE id=?", (group_id, context.user_data["admin_category_id"]))
                await update.message.reply_text(
                    "✅ ক্যাটাগরির Group যুক্ত হয়েছে।" if group_id else "✅ ক্যাটাগরির Group সরানো হয়েছে।")
                context.user_data.pop("admin_action", None)
                context.user_data.pop("admin_category_id", None)
                return True

            # ---------- Payment Method: Add ----------
            if action == "addpm":
                if "|" not in raw:
                    await update.message.reply_text("এইভাবে লিখুন: Payment Method Name | Account Info")
                    return True
                name, acc = [x.strip() for x in raw.split("|", 1)]
                if not name or len(name) > 40: raise ValueError()
                con.execute(
                    "INSERT INTO payment_methods(name,account_info,is_open,created_at) VALUES(?,?,1,?)",
                    (name, acc, now())
                )
                await update.message.reply_text(f"✅ নতুন Payment Method তৈরি হয়েছে: {name}")
                context.user_data.pop("admin_action", None)
                return True

            # ---------- Payment Method: Edit Name ----------
            if action == "pmname":
                name = raw
                if not name or len(name) > 40: raise ValueError()
                con.execute("UPDATE payment_methods SET name=? WHERE id=?", (name, context.user_data["admin_pm_id"]))
                await update.message.reply_text("✅ Payment method-এর নাম পরিবর্তন হয়েছে।")
                context.user_data.pop("admin_action", None)
                context.user_data.pop("admin_pm_id", None)
                return True

            # ---------- Payment Method: Edit Account Info ----------
            if action == "pmacc":
                acc = raw
                if len(acc) > 120: raise ValueError()
                con.execute("UPDATE payment_methods SET account_info=? WHERE id=?", (acc, context.user_data["admin_pm_id"]))
                await update.message.reply_text("✅ Payment method-এর তথ্য আপডেট হয়েছে।")
                context.user_data.pop("admin_action", None)
                context.user_data.pop("admin_pm_id", None)
                return True

            # ---------- Report: Amount ----------
            if action == "report_amount":
                amount = float(raw)
                if amount < 0: raise ValueError()
                context.user_data["report_amount"] = amount
                context.user_data["admin_action"] = "report_message"
                await update.message.reply_text(
                    "এবার ইউজারকে যে মেসেজ যাবে তা লিখুন।\n"
                    "বাতিল করতে নিচের ❌ Cancel বাটন চাপুন।",
                    reply_markup=ReplyKeyboardMarkup(
                        [[ColorButton("❌ Cancel", "danger")]], resize_keyboard=True)
                )
                return True

            # ---------- Report: Message ----------
            if action == "report_message":
                context.user_data["report_message"] = rich_store(update.message)
                context.user_data["admin_action"] = "report_screenshot"
                sid = context.user_data.get("report_sid")
                await update.message.reply_text(
                    "এখন পেমেন্টের Screenshot পাঠান।\n"
                    "স্ক্রিনশট না পাঠিয়ে ⏭ Skip চাপলেও চলবে।",
                    reply_markup=InlineKeyboardMarkup([
                        [InlineKeyboardButton("⏭ Skip", callback_data=f"adm:reportskip:{sid}")],
                        [InlineKeyboardButton("❌ Cancel", callback_data=f"adm:reportcancel:{sid}")],
                    ]))
                return True

            # ---------- Settle: Valid ----------
            if action == "settle_valid":
                valid = int(raw)
                if valid < 0: raise ValueError()
                context.user_data["settle_valid_count"] = valid
                context.user_data["admin_action"] = "settle_invalid"
                await update.message.reply_text(
                    "এবার Invalid নাম্বারের সংখ্যা লিখুন।\n"
                    "বাতিল করতে নিচের ❌ Cancel বাটন চাপুন।",
                    reply_markup=ReplyKeyboardMarkup(
                        [[ColorButton("❌ Cancel", "danger")]], resize_keyboard=True)
                )
                return True

            # ---------- Settle: Invalid ----------
            if action == "settle_invalid":
                invalid = int(raw)
                valid = int(context.user_data["settle_valid_count"])
                sid = context.user_data["admin_submission_id"]
                s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
                if not s:
                    await update.message.reply_text("Submission not found.")
                elif s["status"] != "PENDING":
                    await update.message.reply_text("এই সাবমিশনটি আগে থেকেই রিভিউ করা হয়েছে।")
                elif invalid < 0 or valid + invalid > s["total_items"]:
                    await update.message.reply_text(
                        f"Valid + Invalid ({valid+invalid}) ফাইলের মোট row/line ({s['total_items']})-এর বেশি। আবার Invalid সংখ্যা দিন।")
                    return True
                else:
                    amount = valid * s["rate"]
                    con.execute(
                        "UPDATE submissions SET status='APPROVED',valid_items=?,invalid_items=?,amount=?,reviewed_at=? WHERE id=?",
                        (valid, invalid, amount, now(), sid))
                    con.execute("UPDATE users SET balance=balance+? WHERE user_id=?", (amount, s["user_id"]))
                    con.execute(
                        "INSERT INTO payments(user_id,submission_id,amount,note,created_at) VALUES(?,?,?,?,?)",
                        (s["user_id"], sid, amount, f"{valid} valid × {money(s['rate'])}", now()))
                    await update.message.reply_text(
                        f"✅ Submission approved!\nValid: {valid}\nInvalid: {invalid}\nPayment: {money(amount)}",
                        reply_markup=main_keyboard(update.effective_user.id))
                    try:
                        await context.bot.send_message(s["user_id"],
                            RawHTML(f"✅ <b>ফাইল রিভিউ সম্পন্ন!</b>\n\n🆔 <b>Submission:</b> <code>{sid}</code>\n✅ <b>Valid:</b> {valid}\n❌ <b>Invalid:</b> {invalid}\n💰 <b>Added to balance:</b> {money(amount)}"))
                    except Exception:
                        log.exception("Could not notify user")
                    # Update group message
                    try:
                        with db() as con2:
                            g = con2.execute("SELECT c.group_id, s.group_message_id FROM submissions s JOIN categories c ON s.category_id=c.id WHERE s.id=?", (sid,)).fetchone()
                        if g and g['group_id'] and g['group_message_id']:
                            await context.bot.edit_message_caption(
                                chat_id=g['group_id'], message_id=g['group_message_id'],
                                caption=f"✅ APPROVED\nSubmission: {sid}\nValid: {valid}\nInvalid: {invalid}\nAmount: {money(amount)}")
                    except Exception:
                        log.exception("Could not update group message")
                context.user_data.pop("admin_action", None)
                context.user_data.pop("admin_submission_id", None)
                context.user_data.pop("settle_valid_count", None)
                return True

            # ---------- Reject: Reason ----------
            if action == "reject_reason":
                sid = context.user_data["admin_submission_id"]
                s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
                if not s:
                    await update.message.reply_text("Submission not found.")
                elif s["status"] not in ("PENDING", "APPROVED"):
                    await update.message.reply_text("এই সাবমিশনটি আগে থেকেই রিভিউ করা হয়েছে।")
                else:
                    con.execute("UPDATE submissions SET status='REJECTED',admin_note=?,reviewed_at=? WHERE id=?",
                                (raw, now(), sid))
                    await update.message.reply_text(
                        "❌ Submission rejected — Admin Panel ও Group থেকে সরানো হয়েছে।",
                        reply_markup=main_keyboard(update.effective_user.id))
                    try:
                        await context.bot.send_message(s["user_id"], f"❌ Submission {sid} was rejected.\nReason: {raw}")
                    except Exception:
                        log.exception("Could not notify user")
                    # Group থেকে সম্পূর্ণ Remove
                    await remove_group_message(context, sid)
                context.user_data.pop("admin_action", None)
                context.user_data.pop("admin_submission_id", None)
                return True

            # ---------- Balance Add / Deduct ----------
            if action == "balance":
                uid = int(context.user_data["balance_user_id"])
                amount = float(raw) * int(context.user_data.get("balance_direction", 1))
                if amount == 0: raise ValueError()
                u = con.execute("SELECT balance FROM users WHERE user_id=?", (uid,)).fetchone()
                if not u:
                    await update.message.reply_text("User not found.")
                elif u["balance"] + amount < 0:
                    await update.message.reply_text("Balance শূন্যের নিচে যেতে পারবে না। আবার amount দিন।")
                    return True
                else:
                    con.execute("UPDATE users SET balance=balance+? WHERE user_id=?", (amount, uid))
                    await update.message.reply_text(f"✅ Balance updated by {money(amount)}.")
                    try:
                        await context.bot.send_message(uid, f"Your balance was adjusted by {money(amount)}.")
                    except Exception:
                        pass
                context.user_data.pop("admin_action", None)
                context.user_data.pop("balance_user_id", None)
                context.user_data.pop("balance_direction", None)
                return True

    except (ValueError, sqlite3.IntegrityError):
        await update.message.reply_text("সঠিক তথ্য দিন। আবার চেষ্টা করুন, অথবা ❌ Cancel চাপুন।")
        return True

    return True


# ========================================================
# ====================  TEXT ROUTER  =====================
# ========================================================

async def group_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Category Group-এর বাটন (Approve / Reject / Send Report)।"""
    q = update.callback_query
    if not is_admin(q.from_user.id):
        await q.answer("⛔ Admin access only.", show_alert=True)
        return
    data = q.data.split(":")
    action, sid = data[1], data[2]

    with db() as con:
        s = con.execute("SELECT * FROM submissions WHERE id=?", (sid,)).fetchone()
    if not s:
        await q.answer("Submission not found.", show_alert=True)
        return

    # ✅ Approve — এক ট্যাপেই
    if action == "approve":
        ok, err = await approve_submission(context, sid)
        await q.answer("✅ Approved — ইউজারকে নোটিফিকেশন পাঠানো হয়েছে।" if ok else err, show_alert=not ok)
        return

    # ❌ Reject — এক ট্যাপেই সম্পূর্ণ মুছে যায়
    if action == "reject":
        if s["status"] not in ("PENDING", "APPROVED"):
            await q.answer("এই সাবমিশনে এখন Reject করা যাবে না।", show_alert=True)
            return
        ok, err = await delete_submission_completely(context, sid)
        await q.answer("🗑 Rejected — সম্পূর্ণ মুছে ফেলা হয়েছে।" if ok else err, show_alert=not ok)
        return

    # 📤 Send Report — মেসেজ লিখতে Admin-এর Private Chat-এ নিয়ে যায়
    if action == "report":
        if s["status"] != "APPROVED":
            await q.answer("আগে Approve করুন, তারপর Send Report করা যাবে।", show_alert=True)
            return
        clear_admin_state(context)
        context.user_data["admin_action"] = "report_amount"
        context.user_data["report_sid"] = sid
        text = (f"💳 SEND PAYMENT REPORT\nSubmission: {sid}\n\n"
                "কত টাকা Payment পাঠাবেন তা লিখুন (শুধু সংখ্যা, যেমন 150)।\n"
                "বাতিল করতে নিচের ❌ Cancel বাটন চাপুন।")
        markup = InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancel", callback_data=f"adm:reportcancel:{sid}")]])
    else:
        await q.answer()
        return

    try:
        await context.bot.send_message(q.from_user.id, text, reply_markup=markup)
    except Exception:
        clear_admin_state(context)
        await q.answer("বটকে প্রাইভেটে একবার /start দিয়ে আবার চেষ্টা করুন।", show_alert=True)
        return
    await q.answer("প্রাইভেট চ্যাটে বটের মেসেজ দেখুন।")


async def text_router(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (update.message.text or "").strip()
    text = canonical_reply_text(text)  # Add Single Emoji: বদলানো ইমোজিসহ বাটন → মূল লেবেল
    text = LEGACY_MENU_ALIASES.get(text, text)  # পুরনো কিবোর্ডের লেবেল → নতুন লেবেল

    # 1) ❌ Cancel — Admin input flow বা User submission flow
    if text == "❌ Cancel" and (
        context.user_data.get("admin_action") or context.user_data.get("awaiting_payment_details")
    ):
        clear_admin_state(context)
        clear_submission_state(context)
        await update.message.reply_text(
            "Cancelled.", reply_markup=main_keyboard(update.effective_user.id))
        return

    # 1.5) গ্লোবাল নিয়ম: কোনো Reply Keyboard বাটন চাপলে সেটি কখনো ইনপুট নয় —
    #      আগের ইনপুট-অপেক্ষা (Admin বা User) বাতিল করে সরাসরি ওই বাটনের কাজ চলবে।
    if text in MENU_BUTTON_TEXTS:
        reset_all_input_state(context)

    # 2) Admin-এর text input flow (আগের অগ্রাধিকার বজায় রাখা হয়েছে)
    if await admin_text_input(update, context):
        return

    # 3) মূল মেনু — যেকোনো মেনু চাপলে user-এর submission state পরিষ্কার হবে
    if text == BTN_SUBMIT:
        clear_submission_state(context)
        await submit_start(update, context)
        return
    if text == BTN_PROFILE:
        clear_submission_state(context)
        await profile(update, context)
        return
    if text == BTN_PAYMENTS:
        clear_submission_state(context)
        await payments(update, context)
        return
    if text == BTN_PRICES:
        clear_submission_state(context)
        await prices(update, context)
        return
    if text == BTN_HOWTO:
        clear_submission_state(context)
        await howto(update, context)
        return
    if text == BTN_SUPPORT:
        clear_submission_state(context)
        await support(update, context)
        return
    if text == BTN_PAYISSUE:
        clear_submission_state(context)
        await update.message.reply_text(get_payment_issue_message())
        return
    if text == BTN_ADMIN:
        clear_submission_state(context)
        if is_admin(update.effective_user.id):
            await admin_panel(update, context)
        else:
            await update.message.reply_text(
                "এই অপশনটি আপনার জন্য উপলব্ধ নয়।",
                reply_markup=main_keyboard(update.effective_user.id))
        return

    # 4) User-এর Payment Details সংগ্রহ
    if context.user_data.get("awaiting_payment_details"):
        await handle_payment_details(update, context)
        return

    # 5) Fallback
    await update.message.reply_text(
        "মেনু থেকে একটি অপশন নির্বাচন করুন।",
        reply_markup=main_keyboard(update.effective_user.id))


async def reset_state_on_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """যেকোনো /command এলে আগের ইনপুট স্টেট বাতিল। এরপর কমান্ডের নিজস্ব হ্যান্ডলার স্বাভাবিকভাবে চলে।"""
    reset_all_input_state(context)


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    error = context.error
    if isinstance(error, (NetworkError, TimedOut)):
        log.warning("Temporary Telegram network problem (বট চালু আছে, নিজে ঠিক হবে): %s", error)
        return
    if error is not None:
        log.error(
            "Unhandled Telegram update exception",
            exc_info=(type(error), error, error.__traceback__)
        )
    else:
        log.error("Telegram update failed with an unknown exception")


def main():
    if not BOT_TOKEN:
        raise SystemExit(
            "BOT TOKEN MISSING: Hosting panel > Startup/Environment Variables-এ "
            "BOT_TOKEN নামে BotFather থেকে পাওয়া সম্পূর্ণ token সেট করুন। "
            "অথবা TELEGRAM_BOT_TOKEN নামে সেট করতে পারেন।"
        )

    # Telegram bot token সাধারণত numeric bot-id:secret format-এর হয়।
    # এতে ভুল/অসম্পূর্ণ value থাকলে Application তৈরি করার আগেই পরিষ্কার error দেখাবে।
    if not re.fullmatch(r"\d+:[A-Za-z0-9_-]{20,}", BOT_TOKEN):
        raise SystemExit(
            "BOT TOKEN FORMAT INVALID: BOT_TOKEN-এ BotFather-এর সম্পূর্ণ token দিন। "
            "শুরু/শেষের quote, space, অতিরিক্ত লেখা বা পুরোনো/revoked token ব্যবহার করবেন না।"
        )

    if not ADMIN_IDS:
        log.warning("ADMIN_IDS is empty. Set your numeric Telegram admin user ID.")
    UPLOAD_DIR.mkdir(exist_ok=True)
    init_db()
    _EMOJI_READY[0] = True

    try:
        app = (
            Application.builder()
            .token(BOT_TOKEN)
            .request(HTTPXRequest(connection_pool_size=32, connect_timeout=20, read_timeout=30,
                                  write_timeout=60, pool_timeout=20))
            .get_updates_request(HTTPXRequest(connection_pool_size=2, connect_timeout=20,
                                              read_timeout=40, pool_timeout=20))
            .concurrent_updates(True)
            .post_init(seed_default_premium_emojis)
            .build()
        )
    except InvalidToken as exc:
        raise SystemExit(
            "TELEGRAM REJECTED BOT TOKEN: Hosting Environment Variables-এ সঠিক "
            "BotFather token দিন। BotFather-এ token revoke করা হয়ে থাকলে নতুন token ব্যবহার করুন।"
        ) from exc

    install_premium_message_wrappers()

    # সবার আগে: যেকোনো কমান্ডে ইনপুট স্টেট রিসেট (কিছু block করে না)
    app.add_handler(MessageHandler(filters.COMMAND, reset_state_on_command), group=-2)

    # Force-to-Join gate runs before normal commands/messages for private users.
    app.add_handler(TypeHandler(Update, force_join_guard), group=-1)

    # --- পুরনো command গুলো backward compatibility-এর জন্য রাখা হয়েছে ---
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("admin", admin_panel))
    app.add_handler(CommandHandler("addcategory", addcategory))
    app.add_handler(CommandHandler("categories", categories_cmd))
    app.add_handler(CommandHandler("setrate", setrate))
    app.add_handler(CommandHandler("toggle", toggle_category))
    app.add_handler(CommandHandler("pending", pending_cmd))
    app.add_handler(CommandHandler("review", review_cmd))
    app.add_handler(CommandHandler("settle", settle_cmd))
    app.add_handler(CommandHandler("reject", reject_cmd))
    app.add_handler(CommandHandler("balance", balance_cmd))
    app.add_handler(CommandHandler("setgroup", setgroup_cmd))

    app.add_handler(CallbackQueryHandler(admin_callback, pattern=r"^adm:"))
    app.add_handler(CallbackQueryHandler(category_callback, pattern=r"^(cat:\d+|pm:\d+|cancel)$"))
    app.add_handler(CallbackQueryHandler(group_callback, pattern=r"^grp:"))

    app.add_handler(MessageHandler(filters.VIDEO, handle_video))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_router))

    # Broadcast: সব ধরনের মেসেজ (ছবি, ভিডিও, ভয়েস, স্টিকার ইত্যাদি) ধরার জন্য আলাদা group
    app.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, broadcast_capture), group=1)

    app.add_error_handler(error_handler)

    log.info("%s starting...", BOT_NAME)
    app.run_polling(allowed_updates=Update.ALL_TYPES, bootstrap_retries=-1)


if __name__ == "__main__":
    main()
