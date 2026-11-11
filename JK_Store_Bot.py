#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
JK Store — HEADLESS BOT (no UI)
================================
Runs the complete Telegram store bot (same logic as JK_Store_v6.pyw) WITHOUT any
window / tkinter / Windows-only module. Works on Windows, Linux, VPS, Docker,
Railway, Render, Replit, Termux ... anywhere Python 3.9+ runs.

Install:   pip install requests pandas openpyxl
Run:       python JK_Store_Bot.py

Settings (first one found wins):
  1) Environment:   BOT_TOKEN=...   MASTER_CHAT_ID=...   JK_DATA_DIR=/path/to/data
  2) <data dir>/config.json   {"token": "...", "master_chat_id": "..."}
  3) <data dir>/data_backup_v20.json  (the file the desktop app already creates)

Railway: add a Volume (mount path /data) + Variables BOT_TOKEN, MASTER_CHAT_ID.
         Then send /restore to the bot and upload your old customers.db once.

Data dir (default: ./bot_data next to this file) holds:
  customers.db            users, balances, deposits, purchases, settings, button names
  data_backup_v20.json    the stock lists (61xx / 1000x / newid) + token + chat id
To move from the desktop app, copy those two files from
  %APPDATA%\DataExtractorTool\   into the data dir.  (Run only ONE copy of the
bot per token, otherwise Telegram will answer conflicts.)

Admin CONTROL PANEL lives inside Telegram (no desktop UI needed):
  tap the "Control Panel" button (or send /panel) -> prices, button names, balances,
  discounts, shop ON/OFF, user access, stock, stats, deposits.
  Shortcuts still work: /names  /setname <1000x|61xx|newid> <name>  /setprice <type> <amt> [BDT|USD]
"""
import os
import re
import io
import sys
import json
import math
import time
import queue
import signal
import sqlite3
import logging
import tempfile
import threading
import traceback
from datetime import datetime, timezone, timedelta
import base64

import requests
import pandas as pd

# ============================================================
# Data folder / files
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))
# Railway: attach a Volume (mount path /data) -> RAILWAY_VOLUME_MOUNT_PATH is set automatically.
# Without a Volume the Railway disk is wiped on every deploy/restart and all data is lost!
APP_FOLDER = (os.environ.get("JK_DATA_DIR")
              or os.environ.get("RAILWAY_VOLUME_MOUNT_PATH")
              or os.path.join(_HERE, "bot_data"))
os.makedirs(APP_FOLDER, exist_ok=True)
BACKUP_FILE = os.path.join(APP_FOLDER, "data_backup_v20.json")
DB_FILE = os.path.join(APP_FOLDER, "customers.db")
LOG_FILE = os.path.join(APP_FOLDER, "error_log.txt")
CONFIG_FILE = os.path.join(APP_FOLDER, "config.json")

# ============================================================
# Logging (file + console, so server logs show what is happening)
# ============================================================
logger = logging.getLogger("ExtractorApp")
logger.setLevel(logging.DEBUG)
_fmt = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
try:
    _fh = logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8')
    _fh.setFormatter(_fmt)
    logger.addHandler(_fh)
except Exception:
    pass
_sh = logging.StreamHandler(sys.stdout)
_sh.setFormatter(_fmt)
_sh.setLevel(logging.INFO)
logger.addHandler(_sh)

BD_TZ = timezone(timedelta(hours=6))          # Bangladesh time (no DST) - server may be UTC
DEFAULT_AUTO_NUMBERS = "01954054330"          # receiving bKash/Nagad/Rocket numbers (editable in panel)
OCR_MODEL = os.environ.get("OCR_MODEL", "claude-sonnet-5-5")

# ============================================================
# Per-user Topic Chat (v31) - background mirror worker
# One single worker keeps the message order correct inside every topic
# and keeps the bot fast (mirroring never blocks the bot).
# ============================================================
_MIRROR_Q = queue.Queue()
_MIRROR_STARTED = [False]
_MIRROR_START_LOCK = threading.Lock()
_TOPIC_LOCK = threading.RLock()


def _mirror_worker():
    while True:
        fn = _MIRROR_Q.get()
        try:
            fn()
        except Exception as e:
            logger.error(f"[mirror worker] {e}")


def mirror_submit(fn):
    with _MIRROR_START_LOCK:
        if not _MIRROR_STARTED[0]:
            threading.Thread(target=_mirror_worker, daemon=True).start()
            _MIRROR_STARTED[0] = True
    _MIRROR_Q.put(fn)



# ============================================================
# Policy text (English)
# ============================================================
POLICY_TEXT = """⏱️ Replacement Guarantee

• 1–10 pieces: 2 hours
• 11–30 pieces: 6 hours
• 31–50 pieces: 8 hours
• 51+ pieces: maximum 12 hours

⚠️ After the specified time, no replacement requests will be accepted.

📌 Important Rules

• Check the ID immediately after purchase.
• For replacement, send as "UID PASSWORD COOKIES".
• Screenshots, images, or files are not acceptable.
• Replacement is not applicable if password is changed or 2FA is enabled.
• No replacement if the ID is still working or issues are due to your IP/browser.
• For multiple IDs, provide info per line, e.g. what problem each ID has.
• Verify quickly and complete required actions due to Used ID risk in the market."""

# ============================================================
# Payment info
# ============================================================
PAYMENT_INFO = """💳 Deposit / Payment Methods

🔹 Binance Pay ID: `1234153609`
🔹 BEP-20 (USDT) Network:
`0x028dc4c629380a54805066b3ca95539eacd3b766`

🔹 TRX-20 (USDT) Network:
`TK2sRJYpftBMsBHThKVTCk8PyCFedtB61Q`

🔹 bKash: `01954054330`
🔹 Nagad: `01954054330`
🔹 Rocket: `01954054330`

👆 Tap any number/address to copy it.

📸 After payment, send the payment screenshot here.
Admin will verify and credit your balance.

⚠️ Only send real payment screenshots. Fake attempts will be rejected."""


PAYMENT_INFO_USD = """💳 Deposit / Payment Methods (USD)

🔹 Binance Pay ID: `1234153609`
🔹 BEP-20 (USDT) Network:
`0x028dc4c629380a54805066b3ca95539eacd3b766`
🔹 TRX-20 (USDT) Network:
`TK2sRJYpftBMsBHThKVTCk8PyCFedtB61Q`

👆 Tap any ID/address to copy it.

📸 After payment, send the payment screenshot here.
Admin will verify and credit your balance in USD.

⚠️ Only send real payment screenshots. Fake attempts will be rejected."""


def parse_line_to_columns(line):
    line = line.strip()
    if not line:
        return ["", "", ""]
    if '\t' in line:
        parts = line.split('\t')
    else:
        parts = line.split(None, 2)
    if len(parts) == 1:
        return [parts[0], "", ""]
    elif len(parts) == 2:
        return [parts[0], parts[1], ""]
    else:
        return parts


# ============================================================
# LANGUAGE (English / বাংলা) — translation layer for customers
# ============================================================
LANG_BTN = "🌐 Language / ভাষা"
LANG_EN_BTN = "🇬🇧 English"
LANG_BN_BTN = "🇧🇩 বাংলা"

# ---- Product types: key -> display label ----
NEWID_KEY = "newid"
NEWID_BTN_LABEL = "New Number id"   # name shown on the Buy button (plain text, short enough for phones)
NEWID_LABEL = "New Number id"
# ---- Buy-button names (editable from Control Panel -> Bot Users & Prices) ----
DEFAULT_BUTTON_NAMES = {"1000x": "1000x", "61xx": "61xx", NEWID_KEY: NEWID_BTN_LABEL}
BUTTON_NAMES = dict(DEFAULT_BUTTON_NAMES)   # live names; refreshed from the DB
BUTTON_NAME_MAX = 24

def load_button_names(db):
    """Read custom Buy-button names from the DB (falls back to defaults)."""
    for key, default in DEFAULT_BUTTON_NAMES.items():
        try:
            val = db.get_setting("btn_name_" + key, "").strip()
        except Exception:
            val = ""
        BUTTON_NAMES[key] = val or default

def button_name(key):
    return BUTTON_NAMES.get(key) or DEFAULT_BUTTON_NAMES.get(key, key)

def type_label(p_type):
    if p_type in BUTTON_NAMES:
        return button_name(p_type)
    return p_type

BN_KEYS = {
    "🛒 Buy IDs": "🛒 আইডি কিনুন",
    "💳 Deposit": "💳 ডিপোজিট",
    "💰 My Balance": "💰 আমার ব্যালেন্স",
    "📜 My Orders": "📜 আমার অর্ডার",
    "🔁 Replacement": "🔁 রিপ্লেসমেন্ট",
    "📋 Policy": "📋 নিয়মাবলী",
    "📞 Support": "📞 সাপোর্ট",
    "📄 Text File (.txt)": "📄 টেক্সট ফাইল (.txt)",
    "📊 Excel File (.xlsx)": "📊 এক্সেল ফাইল (.xlsx)",
    "❌ Cancel": "❌ বাতিল",
    "❌ Close": "❌ বন্ধ করুন",
    # admin menu
    "📤 Upload 61xx": "📤 61xx আপলোড",
    "📤 Upload 1000x": "📤 1000x আপলোড",
    "📤 Upload newid": "📤 New create id আপলোড",
    "📊 Check Status": "📊 স্টক দেখুন",
    "👥 Customers": "👥 কাস্টমার",
    "💳 Deposit Info": "💳 ডিপোজিট তথ্য",
    "📈 Purchases": "📈 ক্রয়ের তালিকা",
    "📑 Report": "📑 রিপোর্ট",
    "🗄️ Backup": "🗄️ ব্যাকআপ",
    "💱 Currency": "💱 কারেন্সি",
    "📊 Stats": "📊 পরিসংখ্যান",
    "📢 Announce": "📢 ঘোষণা দিন",
    "⚙️ Control Panel": "⚙️ কন্ট্রোল প্যানেল",
}
BN_KEYS_REV = {v: k for k, v in BN_KEYS.items()}

BN_POLICY = """⏱️ রিপ্লেসমেন্ট গ্যারান্টি

• ১–১০ পিস: ২ ঘণ্টা
• ১১–৩০ পিস: ৬ ঘণ্টা
• ৩১–৫০ পিস: ৮ ঘণ্টা
• ৫১+ পিস: সর্বোচ্চ ১২ ঘণ্টা

⚠️ নির্দিষ্ট সময়ের পরে কোনো রিপ্লেসমেন্ট রিকোয়েস্ট গ্রহণ করা হবে না।

📌 গুরুত্বপূর্ণ নিয়ম

• কেনার সাথে সাথেই আইডি চেক করুন।
• রিপ্লেসমেন্টের জন্য "UID PASSWORD COOKIES" ফরম্যাটে পাঠান।
• স্ক্রিনশট, ছবি বা ফাইল গ্রহণযোগ্য নয়।
• পাসওয়ার্ড বদলানো বা 2FA চালু থাকলে রিপ্লেসমেন্ট প্রযোজ্য নয়।
• আইডি এখনও কাজ করলে, অথবা সমস্যা আপনার IP/ব্রাউজারের কারণে হলে রিপ্লেসমেন্ট নেই।
• একাধিক আইডির ক্ষেত্রে প্রতি লাইনে প্রতিটি আইডির সমস্যা লিখুন।
• মার্কেটে Used ID-এর ঝুঁকি থাকায় দ্রুত যাচাই করে প্রয়োজনীয় পদক্ষেপ নিন।"""

BN_PAYMENT = """💳 ডিপোজিট / পেমেন্ট মাধ্যম

🔹 Binance Pay ID: `1234153609`
🔹 BEP-20 নেটওয়ার্ক:
`0x028dc4c629380a54805066b3ca95539eacd3b766`

🔹 TRX-20 নেটওয়ার্ক:
`TK2sRJYpftBMsBHThKVTCk8PyCFedtB61Q`

🔹 bKash: `01954054330`
🔹 Nagad: `01954054330`
🔹 Rocket: `01954054330`

👆 কপি করতে যেকোনো নম্বর/অ্যাড্রেসে ট্যাপ করুন।

📸 পেমেন্টের পর এখানে পেমেন্টের স্ক্রিনশট পাঠান।
অ্যাডমিন যাচাই করে আপনার ব্যালেন্সে যোগ করবে।

⚠️ শুধু আসল পেমেন্টের স্ক্রিনশট পাঠান। ভুয়া চেষ্টা বাতিল করা হবে।"""

BN_PAYMENT_USD = """💳 ডিপোজিট / পেমেন্ট মাধ্যম (USD)

🔹 Binance Pay ID: `1234153609`
🔹 BEP-20 (USDT) নেটওয়ার্ক:
`0x028dc4c629380a54805066b3ca95539eacd3b766`

🔹 TRX-20 (USDT) নেটওয়ার্ক:
`TK2sRJYpftBMsBHThKVTCk8PyCFedtB61Q`

👆 কপি করতে যেকোনো আইডি/অ্যাড্রেসে ট্যাপ করুন।

📸 পেমেন্টের পর এখানে পেমেন্টের স্ক্রিনশট পাঠান।
অ্যাডমিন যাচাই করে আপনার ব্যালেন্সে USD-তে যোগ করবে।

⚠️ শুধু আসল পেমেন্টের স্ক্রিনশট পাঠান। ভুয়া চেষ্টা বাতিল করা হবে।"""

REGION_BD_BTN = "🇧🇩 Bangladesh / বাংলাদেশ"
REGION_OTHER_BTN = "🌍 Other Country (USD)"

BN_REGEX = [
    (r"❌ Only (\d+) IDs in stock for (\S+?)\.", r"❌ \2-এ মাত্র \1টি আইডি স্টকে আছে।"),
    (r"❌ Only (\d+) IDs available! Refunding\.\.\.", r"❌ মাত্র \1টি আইডি আছে! টাকা ফেরত দেওয়া হচ্ছে..."),
    (r"❌ Deposit #(\d+) (?:was )?rejected\.", r"❌ ডিপোজিট #\1 বাতিল করা হয়েছে।"),
    (r"✅ Replacement ticket #(\d+) submitted\.", r"✅ রিপ্লেসমেন্ট টিকেট #\1 জমা হয়েছে।"),
    (r"\+([\d.]+) (BDT|USD) (?:has been )?added to your balance\.", r"+\1 \2 আপনার ব্যালেন্সে যোগ হয়েছে।"),
]

BN_PHRASES = [
    ("✨ *Welcome, ", "✨ *স্বাগতম, "),
    ("Premium ID Store — fast, automatic delivery 24/7.", "প্রিমিয়াম আইডি স্টোর — দ্রুত, স্বয়ংক্রিয় ডেলিভারি ২৪/৭।"),
    ("Use the menu below to buy, deposit or request support.", "কিনতে, ডিপোজিট করতে বা সাপোর্ট পেতে নিচের মেনু ব্যবহার করুন।"),
    ("Your Chat ID:", "আপনার চ্যাট আইডি:"),
    ("Your ID:", "আপনার আইডি:"),
    ("Balance left:", "অবশিষ্ট ব্যালেন্স:"),
    ("Your balance:", "আপনার ব্যালেন্স:"),
    ("Balance:", "ব্যালেন্স:"),
    ("*Your Wallet*", "*আপনার ওয়ালেট*"),
    ("Status:", "স্ট্যাটাস:"),
    ("Total purchased:", "মোট কেনা:"),
    (" pcs", " পিস"),
    ("🥉 Member", "🥉 মেম্বার"),
    ("📭 No orders yet.", "📭 এখনও কোনো অর্ডার নেই।"),
    ("*Order History* (last 10)", "*অর্ডার হিস্ট্রি* (শেষ ১০টি)"),
    ("🏪 Shop is temporarily closed. Please try later.", "🏪 দোকান সাময়িকভাবে বন্ধ আছে। অনুগ্রহ করে পরে চেষ্টা করুন।"),
    ("*Premium Store*", "*প্রিমিয়াম স্টোর*"),
    ("Select an ID type:", "আইডির ধরন বেছে নিন:"),
    ("👇 Tap a type below to buy your IDs", "👇 আইডি কিনতে নিচ থেকে একটি ধরন চাপুন"),
    (" left |", " টি স্টকে |"),
    ("📝 How many ", "📝 কতটি "),
    (" IDs do you want?", " আইডি লাগবে?"),
    ("Price: ", "দাম: "),
    (" per piece", " প্রতি পিস"),
    ("Send the number only:", "শুধু সংখ্যা পাঠান:"),
    ("❌ Insufficient balance!", "❌ ব্যালেন্স অপর্যাপ্ত!"),
    ("💳 Tap *Deposit* to top up.", "💳 টপ-আপ করতে *ডিপোজিট* চাপুন।"),
    ("Required:", "প্রয়োজন:"),
    ("🧾 *Order Summary*", "🧾 *অর্ডার সামারি*"),
    ("Total:", "মোট:"),
    ("Choose delivery format:", "ডেলিভারি ফরম্যাট বেছে নিন:"),
    ("❌ Please send a valid number:", "❌ সঠিক সংখ্যা পাঠান:"),
    ("⚠️ Please use the buttons above.", "⚠️ অনুগ্রহ করে উপরের বাটনগুলো ব্যবহার করুন।"),
    ("❌ Cancelled.", "❌ বাতিল করা হয়েছে।"),
    ("❌ I didn't understand. Send /start.", "❌ বুঝতে পারিনি। /start পাঠান।"),
    ("✅ Order Confirmed", "✅ অর্ডার নিশ্চিত হয়েছে"),
    ("Order:", "অর্ডার:"),
    ("Paid:", "পরিশোধ:"),
    ("Format:", "ফরম্যাট:"),
    ("↩️ Your balance has been refunded.", "↩️ আপনার ব্যালেন্স ফেরত দেওয়া হয়েছে।"),
    ("❌ Error: ", "❌ ত্রুটি: "),
    ("✅ Deposit screenshot received!", "✅ ডিপোজিট স্ক্রিনশট পেয়েছি!"),
    ("Deposit ID:", "ডিপোজিট আইডি:"),
    ("Admin will verify and add balance shortly.", "অ্যাডমিন যাচাই করে শীঘ্রই ব্যালেন্স যোগ করবে।"),
    ("✅ Deposit approved!", "✅ ডিপোজিট অনুমোদিত হয়েছে!"),
    ("Please contact admin for details.", "বিস্তারিত জানতে অ্যাডমিনের সাথে যোগাযোগ করুন।"),
    ("❌ Deposit save error.", "❌ ডিপোজিট সেভ করতে সমস্যা হয়েছে।"),
    ("Admin has been notified. Please contact @Jihad777.", "অ্যাডমিনকে জানানো হয়েছে। অনুগ্রহ করে @Jihad777-এর সাথে যোগাযোগ করুন।"),
    ("📞 Contact admin: @Jihad777", "📞 অ্যাডমিনের সাথে যোগাযোগ: @Jihad777"),
    ("Please share your deposit screenshot + Chat ID.", "অনুগ্রহ করে আপনার ডিপোজিট স্ক্রিনশট ও চ্যাট আইডি শেয়ার করুন।"),
    ("🔁 *Replacement Request*", "🔁 *রিপ্লেসমেন্ট রিকোয়েস্ট*"),
    ("Send in one message:", "এক মেসেজে পাঠান:"),
    ("+ problem (one ID per line).", "+ সমস্যা (প্রতি লাইনে একটি আইডি)।"),
    ("Send ❌ Cancel to abort.", "বাতিল করতে ❌ Cancel পাঠান।"),
    ("Admin will review it shortly.", "অ্যাডমিন শীঘ্রই দেখবে।"),
    ("📩 *Message from Admin*", "📩 *অ্যাডমিনের বার্তা*"),
    ("🚫 You are blocked. Contact @Jihad777.", "🚫 আপনাকে ব্লক করা হয়েছে। @Jihad777-এর সাথে যোগাযোগ করুন।"),
    ("🆕 New Stock Added!", "🆕 নতুন স্টক যোগ হয়েছে!"),
    ("IDs (now ", "টি আইডি (এখন "),
    (" available)", " টি স্টকে)"),
    ("🔎 Screenshot received! Verifying your payment automatically...", "🔎 স্ক্রিনশট পেয়েছি! আপনার পেমেন্ট স্বয়ংক্রিয়ভাবে যাচাই করা হচ্ছে..."),
    ("This takes a few seconds.", "কয়েক সেকেন্ড সময় লাগবে।"),
    ("✅ Deposit approved automatically!", "✅ ডিপোজিট স্বয়ংক্রিয়ভাবে অনুমোদিত হয়েছে!"),
    ("Transaction ID:", "ট্রানজেকশন আইডি:"),
    ("⏳ Could not verify automatically. Sent to admin for manual review.", "⏳ স্বয়ংক্রিয়ভাবে যাচাই করা যায়নি। অ্যাডমিনের কাছে ম্যানুয়াল রিভিউয়ের জন্য পাঠানো হয়েছে।"),
    ("You will be notified shortly.", "শীঘ্রই আপনাকে জানানো হবে।"),
    ("❌ This Transaction ID was already used. Duplicate deposits are not allowed.", "❌ এই ট্রানজেকশন আইডি আগেই ব্যবহার করা হয়েছে। ডুপ্লিকেট ডিপোজিট গ্রহণযোগ্য নয়।"),
    ("⏳ This Transaction ID is already submitted and waiting for admin review.", "⏳ এই ট্রানজেকশন আইডি আগেই জমা হয়েছে এবং অ্যাডমিনের রিভিউয়ের অপেক্ষায় আছে।"),
    ("⚠️ Your deposit was reversed by admin.", "⚠️ আপনার ডিপোজিট অ্যাডমিন বাতিল করেছেন।"),
    ("🛒 Tap Buy IDs to order now.", "🛒 এখনই অর্ডার করতে \"আইডি কিনুন\" চাপুন।"),
    ("📢 Announcement", "📢 ঘোষণা"),
    ("🌐 Language set to English.", "🌐 ভাষা বাংলা করা হয়েছে।"),
]
BN_PHRASES.sort(key=lambda p: len(p[0]), reverse=True)
_BN_REGEX_C = [(re.compile(a), b) for a, b in BN_REGEX]


def tr_bn(text):
    """Translate an English bot message to Bengali (known phrases only)."""
    if not text:
        return text
    if text == POLICY_TEXT:
        return BN_POLICY
    if text == PAYMENT_INFO:
        return BN_PAYMENT
    if text == PAYMENT_INFO_USD:
        return BN_PAYMENT_USD
    for rx, rep in _BN_REGEX_C:
        text = rx.sub(rep, text)
    for en, bn in BN_PHRASES:
        if en in text:
            text = text.replace(en, bn)
    return text


def style_markup(markup):
    """Add Telegram button colors (Bot API 'style' field):
       green = success, red = danger, blue = primary."""
    if not markup:
        return markup
    import copy
    m = copy.deepcopy(markup)
    rows = m.get("keyboard") or m.get("inline_keyboard") or []
    for row in rows:
        for btn in row:
            if not isinstance(btn, dict) or "style" in btn:
                continue
            t = btn.get("text", "")
            if "\u274c" in t or "Reject" in t:          # ❌  -> red
                btn["style"] = "danger"
            elif "Buy IDs" in t or t.startswith("\u2705"):  # 🛒 Buy IDs / ✅ -> green
                btn["style"] = "success"
            else:                                          # everything else -> blue
                btn["style"] = "primary"
    return m


def tr_bn_markup(markup, keys_only=False):
    """Translate button labels of a reply/inline keyboard to Bengali."""
    if not markup:
        return markup
    import copy
    m = copy.deepcopy(markup)
    rows = m.get("keyboard") or m.get("inline_keyboard") or []
    for row in rows:
        for btn in row:
            t = btn.get("text", "")
            btn["text"] = BN_KEYS.get(t) or (t if keys_only else tr_bn(t))
    return m


# ============================================================
# DATABASE
# ============================================================
class CustomerDB:
    def __init__(self, db_path):
        self.db_path = db_path
        self._lock = threading.Lock()
        self._init_db()

    def _init_db(self):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            try:
                conn.execute("PRAGMA journal_mode=WAL")   # faster, no UI/bot lock clashes
                conn.execute("PRAGMA synchronous=NORMAL")
            except Exception:
                pass
            conn.execute("""CREATE TABLE IF NOT EXISTS customers(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                phone TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
            conn.execute("""CREATE TABLE IF NOT EXISTS transactions(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_id INTEGER NOT NULL,
                type TEXT NOT NULL CHECK(type IN ('taken','payment')),
                quantity INTEGER DEFAULT 0,
                unit_price REAL DEFAULT 0,
                amount REAL DEFAULT 0,
                note TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(customer_id) REFERENCES customers(id))""")
            conn.execute("""CREATE TABLE IF NOT EXISTS daily_sales(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                amount REAL NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
            conn.execute("""CREATE TABLE IF NOT EXISTS bot_users(
                chat_id TEXT PRIMARY KEY,
                name TEXT,
                balance REAL DEFAULT 0,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
            conn.execute("""CREATE TABLE IF NOT EXISTS settings(
                key TEXT PRIMARY KEY,
                value TEXT)""")
            conn.execute("""CREATE TABLE IF NOT EXISTS user_topics(
                chat_id TEXT PRIMARY KEY,
                thread_id INTEGER)""")

            # migrations (safe for existing DBs)
            for alter in [
                "ALTER TABLE bot_users ADD COLUMN currency TEXT DEFAULT 'BDT'",
                "ALTER TABLE bot_users ADD COLUMN username TEXT",
            ]:
                try:
                    conn.execute(alter)
                except:
                    pass

            conn.execute("""CREATE TABLE IF NOT EXISTS deposits(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chat_id TEXT NOT NULL,
                username TEXT,
                name TEXT,
                amount REAL DEFAULT 0,
                currency TEXT DEFAULT 'BDT',
                method TEXT,
                screenshot_file_id TEXT,
                status TEXT DEFAULT 'pending',
                note TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
            conn.execute("""CREATE TABLE IF NOT EXISTS user_purchases(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chat_id TEXT NOT NULL,
                username TEXT,
                panel_type TEXT,
                qty INTEGER,
                unit_price REAL,
                cost REAL,
                currency TEXT DEFAULT 'BDT',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")

            conn.execute("""CREATE TABLE IF NOT EXISTS balance_adds(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chat_id TEXT NOT NULL,
                username TEXT,
                amount REAL DEFAULT 0,
                currency TEXT DEFAULT 'BDT',
                source TEXT DEFAULT 'manual',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")

            try:
                conn.execute("ALTER TABLE bot_users ADD COLUMN banned INTEGER DEFAULT 0")
            except:
                pass
            try:
                conn.execute("ALTER TABLE bot_users ADD COLUMN lang TEXT")
            except:
                pass
            try:
                conn.execute("ALTER TABLE bot_users ADD COLUMN region_set INTEGER DEFAULT 0")
                # users that existed before this feature keep their current currency
                conn.execute("UPDATE bot_users SET region_set=1")
            except:
                pass
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('low_stock_limit', '20')")
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('disc_50', '5')")
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('disc_100', '10')")
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('shop_open', '1')")
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('price_61xx', '12')")
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('price_1000x', '20')")
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('price_61xx_usd', '0.15')")
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('price_1000x_usd', '0.25')")
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('price_newid', '12')")
            conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('price_newid_usd', '0.15')")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_tx_cust ON transactions(customer_id)")
            conn.execute("""CREATE TABLE IF NOT EXISTS deposit_txns(
                trx_id TEXT PRIMARY KEY,
                deposit_id INTEGER,
                chat_id TEXT,
                amount REAL DEFAULT 0,
                status TEXT DEFAULT 'pending',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")

    # -------- PRICES --------
    def get_price(self, id_type, currency='BDT'):
        key = f"price_{id_type}" if currency.upper() == 'BDT' else f"price_{id_type}_{currency.lower()}"
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
            if row:
                return float(row[0])
            row = conn.execute("SELECT value FROM settings WHERE key=?", (f"price_{id_type}",)).fetchone()
            return float(row[0]) if row else 0.0

    def set_price(self, id_type, price, currency='BDT'):
        key = f"price_{id_type}" if currency.upper() == 'BDT' else f"price_{id_type}_{currency.lower()}"
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT OR REPLACE INTO settings(key, value) VALUES(?, ?)",
                         (key, str(price)))

    # -------- SETTINGS / DISCOUNT / BAN --------
    def get_setting(self, key, default=""):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
            return row[0] if row else default

    def set_setting(self, key, value):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT OR REPLACE INTO settings(key, value) VALUES(?, ?)", (key, str(value)))

    # -------- USER TOPICS (one Telegram topic per user) --------
    def get_topic(self, chat_id):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT thread_id FROM user_topics WHERE chat_id=?",
                               (str(chat_id),)).fetchone()
            return int(row[0]) if row else None

    def set_topic(self, chat_id, thread_id):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT OR REPLACE INTO user_topics(chat_id, thread_id) VALUES(?,?)",
                         (str(chat_id), int(thread_id)))

    def del_topic(self, chat_id):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("DELETE FROM user_topics WHERE chat_id=?", (str(chat_id),))

    def get_chat_by_topic(self, thread_id):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT chat_id FROM user_topics WHERE thread_id=?",
                               (int(thread_id),)).fetchone()
            return row[0] if row else None

    def clear_topics(self):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("DELETE FROM user_topics")

    def get_discount_percent(self, qty):
        try:
            if qty >= 100:
                return float(self.get_setting("disc_100", "0"))
            if qty >= 50:
                return float(self.get_setting("disc_50", "0"))
        except ValueError:
            pass
        return 0.0

    def get_vip(self, chat_id):
        with sqlite3.connect(self.db_path) as conn:
            q = conn.execute("SELECT COALESCE(SUM(qty),0) FROM user_purchases WHERE chat_id=?", (chat_id,)).fetchone()[0]
        if q >= 1000:
            return ("💎 Diamond", 6.0, q)
        if q >= 500:
            return ("🥇 Gold", 4.0, q)
        if q >= 100:
            return ("🥈 Silver", 2.0, q)
        return ("🥉 Member", 0.0, q)

    def last_purchase_id(self, chat_id):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT id FROM user_purchases WHERE chat_id=? ORDER BY id DESC LIMIT 1", (chat_id,)).fetchone()
            return row[0] if row else 0

    def lifetime_stats(self):
        with sqlite3.connect(self.db_path) as conn:
            rows = conn.execute("""SELECT panel_type, COALESCE(currency,'BDT'), COUNT(*),
                                          COALESCE(SUM(qty),0), COALESCE(SUM(cost),0)
                                   FROM user_purchases GROUP BY panel_type, COALESCE(currency,'BDT')""").fetchall()
            dep = conn.execute("""SELECT COUNT(*),
                                         COALESCE(SUM(CASE WHEN currency='BDT' THEN amount ELSE 0 END),0),
                                         COALESCE(SUM(CASE WHEN currency='USD' THEN amount ELSE 0 END),0)
                                  FROM deposits WHERE status='approved'""").fetchone()
        blank = lambda: {"orders": 0, "qty": 0, "BDT": 0.0, "USD": 0.0}
        stats = {"61xx": blank(), "1000x": blank(), "newid": blank()}
        for panel, cur, orders, qty, cost in rows:
            d = stats.setdefault(panel, blank())
            d["orders"] += orders
            d["qty"] += qty
            d[cur] = d.get(cur, 0.0) + cost
        total = blank()
        for d in stats.values():
            for k in total:
                total[k] += d.get(k, 0)
        return stats, total, dep

    def recent_events(self, limit=80):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""
                SELECT 'ORDER', p.id, p.created_at, p.chat_id, COALESCE(p.username,''),
                       p.panel_type || ' x' || p.qty, p.cost, COALESCE(p.currency,'BDT'), 'done'
                FROM user_purchases p
                UNION ALL
                SELECT 'DEPOSIT', d.id, d.created_at, d.chat_id, COALESCE(d.username,''),
                       'Deposit #' || d.id, COALESCE(d.amount,0), COALESCE(d.currency,'BDT'), d.status
                FROM deposits d
                ORDER BY 3 DESC, 2 DESC LIMIT ?""", (limit,)).fetchall()

    def admin_currency(self):
        c = self.get_setting("admin_currency", "BDT").upper()
        return c if c in ("BDT", "USD") else "BDT"

    def region_done(self, chat_id):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT COALESCE(region_set,0) FROM bot_users WHERE chat_id=?", (chat_id,)).fetchone()
            return bool(row and row[0])

    def set_region(self, chat_id, currency):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE bot_users SET currency=?, region_set=1 WHERE chat_id=?",
                         (currency, chat_id))

    def get_lang(self, chat_id):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT lang FROM bot_users WHERE chat_id=?", (chat_id,)).fetchone()
            return row[0] if row and row[0] else None

    def set_lang(self, chat_id, lang):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE bot_users SET lang=? WHERE chat_id=?", (lang, chat_id))

    def is_banned(self, chat_id):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT COALESCE(banned,0) FROM bot_users WHERE chat_id=?", (chat_id,)).fetchone()
            return bool(row and row[0])

    def set_banned(self, chat_id, flag):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE bot_users SET banned=? WHERE chat_id=?", (1 if flag else 0, chat_id))

    def all_user_ids(self):
        with sqlite3.connect(self.db_path) as conn:
            return [r[0] for r in conn.execute("SELECT chat_id FROM bot_users WHERE COALESCE(banned,0)=0").fetchall()]

    def sales_report(self, days=1):
        mod = f"-{days-1} day"
        with sqlite3.connect(self.db_path) as conn:
            orders, qty, bdt, usd = conn.execute(
                """SELECT COUNT(*), COALESCE(SUM(qty),0),
                          COALESCE(SUM(CASE WHEN currency='BDT' THEN cost ELSE 0 END),0),
                          COALESCE(SUM(CASE WHEN currency='USD' THEN cost ELSE 0 END),0)
                   FROM user_purchases WHERE date(created_at) >= date('now', ?)""", (mod,)).fetchone()
            dep = conn.execute(
                """SELECT COALESCE(SUM(CASE WHEN currency='BDT' THEN amount ELSE 0 END),0),
                          COALESCE(SUM(CASE WHEN currency='USD' THEN amount ELSE 0 END),0)
                   FROM deposits WHERE status='approved' AND date(created_at) >= date('now', ?)""", (mod,)).fetchone()
            top = conn.execute(
                """SELECT COALESCE(MAX(username),''), chat_id, SUM(qty) q FROM user_purchases
                   WHERE date(created_at) >= date('now', ?) GROUP BY chat_id
                   ORDER BY q DESC LIMIT 3""", (mod,)).fetchall()
        return {"orders": orders, "qty": qty, "bdt": bdt, "usd": usd,
                "dep_bdt": dep[0], "dep_usd": dep[1], "top": top}

    # -------- USERS --------
    def get_user(self, chat_id):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute(
                "SELECT chat_id, name, balance, COALESCE(currency,'BDT'), COALESCE(username,'') "
                "FROM bot_users WHERE chat_id=?",
                (chat_id,)).fetchone()

    def register_user(self, chat_id, name, username=""):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT OR IGNORE INTO bot_users(chat_id, name, username) VALUES(?,?,?)",
                         (chat_id, name, username))
            conn.execute("UPDATE bot_users SET username=? WHERE chat_id=? AND (username IS NULL OR username='')",
                         (username, chat_id))
            conn.execute("UPDATE bot_users SET name=? WHERE chat_id=?", (name, chat_id))

    def add_balance(self, chat_id, amount, currency=None, source="manual"):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            if currency:
                conn.execute("UPDATE bot_users SET balance = balance + ?, currency=? WHERE chat_id=?",
                             (float(amount), currency.upper(), chat_id))
            else:
                conn.execute("UPDATE bot_users SET balance = balance + ? WHERE chat_id=?",
                             (float(amount), chat_id))
            # log it, so a mistaken add can be deleted later from the Control Panel
            row = conn.execute("SELECT COALESCE(username,''), COALESCE(currency,'BDT') "
                               "FROM bot_users WHERE chat_id=?", (chat_id,)).fetchone()
            if row:
                conn.execute("INSERT INTO balance_adds(chat_id, username, amount, currency, source) "
                             "VALUES(?,?,?,?,?)",
                             (chat_id, row[0], float(amount), (currency or row[1]).upper(), source))

    def deduct_balance(self, chat_id, amount):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE bot_users SET balance = balance - ? WHERE chat_id=?",
                         (float(amount), chat_id))

    def set_currency(self, chat_id, currency):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE bot_users SET currency=? WHERE chat_id=?",
                         (currency.upper(), chat_id))

    def list_users(self):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT chat_id, name, COALESCE(username,''), balance,
                                   COALESCE(currency,'BDT'), created_at
                                   FROM bot_users ORDER BY created_at DESC""").fetchall()

    # -------- DELETE / RESET --------
    def delete_user(self, chat_id):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("DELETE FROM bot_users WHERE chat_id=?", (chat_id,))
            conn.execute("DELETE FROM user_purchases WHERE chat_id=?", (chat_id,))
            conn.execute("DELETE FROM deposits WHERE chat_id=?", (chat_id,))
        return True

    def reset_balance(self, chat_id):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE bot_users SET balance=0 WHERE chat_id=?", (chat_id,))
        return True

    def clear_all_balances(self):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE bot_users SET balance=0")
        return True

    # -------- DEPOSITS (WITH AUTO-REPAIR) --------
    def add_deposit(self, chat_id, username, name, file_id, method="Screenshot"):
        try:
            with self._lock, sqlite3.connect(self.db_path) as conn:
                cur = conn.execute(
                    "INSERT INTO deposits(chat_id, username, name, screenshot_file_id, method) "
                    "VALUES(?,?,?,?,?)",
                    (chat_id, username, name, file_id, method))
                dep_id = cur.lastrowid
                conn.commit()
                logger.info(f"[add_deposit] OK #{dep_id} chat={chat_id}")
                return dep_id
        except sqlite3.OperationalError as e:
            logger.error(f"[add_deposit] FAILED: {e}")
            logger.error(traceback.format_exc())
            try:
                self._repair_deposits_table()
                with self._lock, sqlite3.connect(self.db_path) as conn:
                    cur = conn.execute(
                        "INSERT INTO deposits(chat_id, username, name, screenshot_file_id, method) "
                        "VALUES(?,?,?,?,?)",
                        (chat_id, username, name, file_id, method))
                    dep_id = cur.lastrowid
                    conn.commit()
                    logger.info(f"[add_deposit] OK after repair #{dep_id}")
                    return dep_id
            except Exception as e2:
                logger.error(f"[add_deposit] Repair failed: {e2}")
                logger.error(traceback.format_exc())
                raise e2

    def _repair_deposits_table(self):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            try:
                info = conn.execute("PRAGMA table_info(deposits)").fetchall()
                cols = {c[1] for c in info}
                required = {"id", "chat_id", "username", "name", "amount",
                            "currency", "method", "screenshot_file_id",
                            "status", "note", "created_at"}
                missing = required - cols
                if not missing:
                    return

                conn.execute("DROP TABLE IF EXISTS deposits_backup")
                conn.execute("ALTER TABLE deposits RENAME TO deposits_backup")
                conn.execute("""CREATE TABLE deposits(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    chat_id TEXT NOT NULL,
                    username TEXT,
                    name TEXT,
                    amount REAL DEFAULT 0,
                    currency TEXT DEFAULT 'BDT',
                    method TEXT,
                    screenshot_file_id TEXT,
                    status TEXT DEFAULT 'pending',
                    note TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
                common = list(cols & required)
                common = [c for c in common if c != "id"]
                if common:
                    col_list = ",".join(common)
                    try:
                        conn.execute(
                            f"INSERT INTO deposits ({col_list}) "
                            f"SELECT {col_list} FROM deposits_backup")
                    except Exception:
                        pass
                conn.commit()
            except Exception:
                pass

    def get_deposit(self, dep_id):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("SELECT * FROM deposits WHERE id=?", (dep_id,)).fetchone()

    def approve_deposit(self, dep_id, amount, currency):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT chat_id FROM deposits WHERE id=?", (dep_id,)).fetchone()
            if not row:
                return False
            chat_id = row[0]
            conn.execute("UPDATE deposits SET status='approved', amount=?, currency=? WHERE id=?",
                         (float(amount), currency.upper(), dep_id))
            conn.execute("UPDATE bot_users SET balance = balance + ?, currency=? WHERE chat_id=?",
                         (float(amount), currency.upper(), chat_id))
            conn.execute("UPDATE deposit_txns SET status='approved' WHERE deposit_id=?", (dep_id,))
            return True

    def reject_deposit(self, dep_id, note=""):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE deposits SET status='rejected', note=? WHERE id=?", (note, dep_id))
            conn.execute("DELETE FROM deposit_txns WHERE deposit_id=?", (dep_id,))   # TrxID free again

    # -------- AUTO-DEPOSIT helpers --------
    def reserve_trx(self, trx_id, dep_id, chat_id, amount):
        """Atomically claim a Transaction ID. Returns (True, 'pending', dep_id) or
        (False, existing_status, existing_deposit_id) when it was used before."""
        with self._lock, sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT status, deposit_id FROM deposit_txns WHERE trx_id=?",
                               (trx_id,)).fetchone()
            if row:
                return False, row[0], row[1]
            conn.execute("INSERT INTO deposit_txns(trx_id, deposit_id, chat_id, amount, status) "
                         "VALUES(?,?,?,?, 'pending')", (trx_id, dep_id, chat_id, float(amount)))
            return True, "pending", dep_id

    def set_deposit_note(self, dep_id, note):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE deposits SET note=? WHERE id=?", (note, dep_id))

    def undo_deposit(self, dep_id):
        """Reverse an approved deposit (takes balance back), keep its TrxID blocked.
        Returns (chat_id, amount, currency, not_recovered) or None."""
        with self._lock, sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT chat_id, COALESCE(amount,0), COALESCE(currency,'BDT'), status "
                               "FROM deposits WHERE id=?", (dep_id,)).fetchone()
            if not row or row[3] != "approved":
                return None
            short = self._reverse_credit(conn, row[0], row[1])
            conn.execute("UPDATE deposits SET status='rejected', note='Reversed by admin' WHERE id=?", (dep_id,))
            conn.execute("UPDATE deposit_txns SET status='blocked' WHERE deposit_id=?", (dep_id,))
            return (row[0], row[1], row[2], short)

    def get_pending_deposits(self):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT id, chat_id, COALESCE(username,''), COALESCE(name,''),
                                          amount, COALESCE(currency,'BDT'), method, created_at
                                          FROM deposits WHERE status='pending' ORDER BY id DESC""").fetchall()

    def get_all_deposits(self, limit=100):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT id, chat_id, COALESCE(username,''), amount,
                                          COALESCE(currency,'BDT'), status, created_at
                                   FROM deposits ORDER BY id DESC LIMIT ?""", (limit,)).fetchall()

    def deposit_stats(self):
        """Approved deposits: (people, deposits, total BDT, total USD)."""
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT COUNT(DISTINCT chat_id), COUNT(*),
                                          COALESCE(SUM(CASE WHEN currency='BDT' THEN amount ELSE 0 END),0),
                                          COALESCE(SUM(CASE WHEN currency='USD' THEN amount ELSE 0 END),0)
                                   FROM deposits WHERE status='approved'""").fetchone()

    def manual_add_stats(self):
        """Manually added balance: (people, times, total BDT, total USD)."""
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT COUNT(DISTINCT chat_id), COUNT(*),
                                          COALESCE(SUM(CASE WHEN currency='BDT' THEN amount ELSE 0 END),0),
                                          COALESCE(SUM(CASE WHEN currency='USD' THEN amount ELSE 0 END),0)
                                   FROM balance_adds""").fetchone()

    def get_balance_adds(self, limit=100):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT id, chat_id, COALESCE(username,''), amount,
                                          COALESCE(currency,'BDT'), source, created_at
                                   FROM balance_adds ORDER BY id DESC LIMIT ?""", (limit,)).fetchall()

    @staticmethod
    def _reverse_credit(conn, chat_id, amount):
        """Take `amount` back from a user's balance (never below 0).
        Returns how much could NOT be taken back (user already spent it)."""
        row = conn.execute("SELECT balance FROM bot_users WHERE chat_id=?", (chat_id,)).fetchone()
        if not row:
            return 0.0
        new = float(row[0]) - float(amount)
        short = 0.0
        if new < 0:
            short, new = -new, 0.0
        conn.execute("UPDATE bot_users SET balance=? WHERE chat_id=?", (new, chat_id))
        return short

    def delete_balance_add(self, add_id):
        """Delete a mistaken manual balance add and take it back from the user.
        Returns (chat_id, amount, currency, not_recovered) or None."""
        with self._lock, sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT chat_id, amount, currency FROM balance_adds WHERE id=?",
                               (add_id,)).fetchone()
            if not row:
                return None
            short = self._reverse_credit(conn, row[0], row[1])
            conn.execute("DELETE FROM balance_adds WHERE id=?", (add_id,))
            return (row[0], row[1], row[2], short)

    def delete_deposit(self, dep_id):
        """Delete a deposit record. If it was approved, its amount is taken back
        from the user's balance. Returns (chat_id, status, amount, currency, not_recovered) or None."""
        with self._lock, sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT chat_id, status, COALESCE(amount,0), COALESCE(currency,'BDT') "
                               "FROM deposits WHERE id=?", (dep_id,)).fetchone()
            if not row:
                return None
            short = 0.0
            if row[1] == "approved":
                short = self._reverse_credit(conn, row[0], row[2])
            conn.execute("DELETE FROM deposits WHERE id=?", (dep_id,))
            return (row[0], row[1], row[2], row[3], short)

    # -------- PURCHASES --------
    def log_purchase(self, chat_id, username, panel_type, qty, unit_price, cost, currency):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("""INSERT INTO user_purchases
                            (chat_id, username, panel_type, qty, unit_price, cost, currency)
                            VALUES(?,?,?,?,?,?,?)""",
                         (chat_id, username, panel_type, qty, float(unit_price),
                          float(cost), currency))

    def get_user_purchases(self, chat_id):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT panel_type, qty, unit_price, cost, currency, created_at
                                   FROM user_purchases WHERE chat_id=? ORDER BY id DESC""",
                                (chat_id,)).fetchall()

    def get_purchases_summary(self):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT p.chat_id, COALESCE(u.name,''), COALESCE(p.username,''),
                                          COALESCE(SUM(p.qty),0) AS total_qty,
                                          COALESCE(SUM(CASE WHEN p.currency='BDT' THEN p.cost ELSE 0 END),0) AS total_bdt,
                                          COALESCE(SUM(CASE WHEN p.currency='USD' THEN p.cost ELSE 0 END),0) AS total_usd
                                   FROM user_purchases p
                                   LEFT JOIN bot_users u ON u.chat_id=p.chat_id
                                   GROUP BY p.chat_id ORDER BY total_qty DESC""").fetchall()

    def get_all_purchases(self, limit=200):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT chat_id, COALESCE(username,''), panel_type,
                                          qty, cost, currency, created_at
                                   FROM user_purchases ORDER BY id DESC LIMIT ?""", (limit,)).fetchall()

    # -------- LEDGER --------
    def _get_or_create(self, username, phone=""):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT id FROM customers WHERE username=?", (username,)).fetchone()
            if row:
                return row[0]
            return conn.execute("INSERT INTO customers(username,phone) VALUES(?,?)",
                                (username, phone)).lastrowid

    def add_taken(self, username, quantity, unit_price=0, note=""):
        cid = self._get_or_create(username)
        amount = float(quantity) * float(unit_price)
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("""INSERT INTO transactions(customer_id,type,quantity,unit_price,amount,note)
                            VALUES(?,'taken',?,?,?,?)""",
                         (cid, int(quantity), float(unit_price), amount, note))

    def add_payment(self, username, amount, note=""):
        cid = self._get_or_create(username)
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("""INSERT INTO transactions(customer_id,type,amount,note)
                            VALUES(?,'payment',?,?)""", (cid, float(amount), note))

    def delete_transaction(self, tx_id):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("DELETE FROM transactions WHERE id=?", (tx_id,))

    def delete_customer(self, username):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT id FROM customers WHERE username=?", (username,)).fetchone()
            if row:
                cid = row[0]
                conn.execute("DELETE FROM transactions WHERE customer_id=?", (cid,))
                conn.execute("DELETE FROM customers WHERE id=?", (cid,))

    def get_summary(self, username):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("""
                SELECT COALESCE(SUM(CASE WHEN t.type='taken'   THEN t.quantity ELSE 0 END),0),
                       COALESCE(SUM(CASE WHEN t.type='taken'   THEN t.amount   ELSE 0 END),0),
                       COALESCE(SUM(CASE WHEN t.type='payment' THEN t.amount   ELSE 0 END),0)
                FROM transactions t JOIN customers c ON c.id=t.customer_id
                WHERE c.username=?""", (username,)).fetchone()
        ids, bill, paid = row
        return {"ids": ids, "bill": bill, "paid": paid, "due": bill - paid}

    def get_history(self, username):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""
                SELECT t.id, t.created_at, t.type, t.quantity, t.unit_price, t.amount, t.note
                FROM transactions t JOIN customers c ON c.id=t.customer_id
                WHERE c.username=? ORDER BY t.id DESC""", (username,)).fetchall()

    def list_all(self):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""
                SELECT c.username,
                  COALESCE(SUM(CASE WHEN t.type='taken'   THEN t.quantity ELSE 0 END),0),
                  COALESCE(SUM(CASE WHEN t.type='taken'   THEN t.amount   ELSE 0 END),0),
                  COALESCE(SUM(CASE WHEN t.type='payment' THEN t.amount   ELSE 0 END),0)
                FROM customers c LEFT JOIN transactions t ON t.customer_id=c.id
                GROUP BY c.id ORDER BY c.username""").fetchall()

    def add_sale(self, name, amount):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT INTO daily_sales(name, amount) VALUES(?,?)",
                         (name, float(amount)))

    def get_today_sales(self):
        today = datetime.now().strftime('%Y-%m-%d')
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute("""SELECT id, name, amount, created_at FROM daily_sales
                                   WHERE date(created_at) = ? ORDER BY id DESC""", (today,)).fetchall()

    def get_today_sales_total(self):
        today = datetime.now().strftime('%Y-%m-%d')
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT COALESCE(SUM(amount), 0) FROM daily_sales WHERE date(created_at) = ?",
                (today,)).fetchone()
            return row[0]

    def delete_sale(self, sale_id):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute("DELETE FROM daily_sales WHERE id=?", (sale_id,))




# ============================================================
# HEADLESS STAND-INS  (replace the desktop window pieces)
# The bot code below was written for the GUI app and talks to
# "panels", "entries", "labels" and "root.after". These tiny
# classes give it the same interface with no window at all.
# ============================================================
_STOCK_LOCK = threading.RLock()


class StockPanel:
    """ID stock list (one ID per line). Same interface as the GUI DataPanel."""
    def __init__(self, id_type):
        self.id_type = id_type
        self._data = ""

    def get_data(self):
        with _STOCK_LOCK:
            return self._data

    def set_data(self, data):
        with _STOCK_LOCK:
            self._data = data or ""

    def append_data(self, new_lines):
        with _STOCK_LOCK:
            current = self._data.strip()
            add = "\n".join(new_lines)
            self._data = (current + "\n" + add) if current else add


class _Value:
    def __init__(self, value=""):
        self._v = str(value or "")

    def get(self):
        return self._v

    def set(self, v):
        self._v = str(v or "")


class _NoopLabel:
    def config(self, *a, **k):
        pass
    configure = config


class _Root:
    """Replacement for the Tk root: after(ms, fn, *args) runs fn in the background."""
    def after(self, ms, fn, *args):
        def run():
            try:
                with _STOCK_LOCK:
                    fn(*args)
            except Exception as e:
                logger.error(f"[after] {e}")
                logger.error(traceback.format_exc())
        if not ms or ms <= 0:
            threading.Thread(target=run, daemon=True).start()
        else:
            t = threading.Timer(ms / 1000.0, run)
            t.daemon = True
            t.start()


# ============================================================
# BOT APPLICATION
# ============================================================
class BotApp:
    def __init__(self, token, master_chat_id):
        self.root = _Root()
        self.token_entry = _Value(token)
        self.chat_id_entry = _Value(master_chat_id)
        self.status_label = _NoopLabel()

        self.bot_running = False
        self.bot_thread = None
        self.last_update_id = 0
        self.user_states = {}
        self.pending_extract = {}
        self.pending_forward_files = {}

        self.customer_db = CustomerDB(DB_FILE)
        load_button_names(self.customer_db)
        self.tg_session = requests.Session()

        self.panel_61xx = StockPanel("61xx")
        self.panel_1000x = StockPanel("1000x")
        self.panel_newid = StockPanel("newid")
        self.load_backup()
        # token / chat id given at startup always win over the old backup file
        self.token_entry.set(token)
        self.chat_id_entry.set(master_chat_id)

    @property
    def users_enabled(self):
        """User-access switch, stored in the DB (settings key 'users_enabled')."""
        try:
            return self.customer_db.get_setting("users_enabled", "1") == "1"
        except Exception:
            return True

    # -------- start / stop --------
    def start(self):
        token = self.token_entry.get().strip()
        cid = self.chat_id_entry.get().strip()
        if not token:
            raise SystemExit("❌ Bot token missing. Set BOT_TOKEN or config.json")
        try:
            int(cid)
        except Exception:
            raise SystemExit("❌ MASTER_CHAT_ID missing or not a number (e.g. 7538066844)")
        self.bot_running = True
        self.bot_thread = threading.Thread(target=self.bot_polling_loop,
                                           args=(token,), daemon=True)
        self.bot_thread.start()
        logger.info("🟢 Bot online. Press Ctrl+C to stop.")
        try:
            n_users = len(self.customer_db.list_users())
            n_stock = sum(self._cp_count(k) for k in self._CP_TYPES)
            logger.info(f"📂 Data folder: {APP_FOLDER} | users: {n_users} | stock IDs: {n_stock}")
            logger.info(self._auto_status_text())
            if "NOT ACTIVE" in self._auto_status_text():
                self._send_text_direct(token, cid, "⚠️ " + self._auto_status_text()
                                       + "\nAdd GEMINI_API_KEY in Railway -> Variables, then redeploy.")
            if n_users == 0:
                self._send_text_direct(
                    token, cid,
                    "⚠️ Bot started with an EMPTY database (0 users).\n"
                    f"Data folder: {APP_FOLDER}\n\n"
                    "To bring your old data: send /restore, then send your customers.db "
                    "(and data_backup_v20.json for the stock).")
        except Exception as e:
            logger.error(f"[startup report] {e}")

    def stop(self):
        self.bot_running = False
        self.silent_auto_save()
        logger.info("⚪ Bot stopped.")

    def run_forever(self):
        self.start()
        stop_evt = threading.Event()

        def _sig(*_):
            stop_evt.set()
        for s in (signal.SIGINT, getattr(signal, "SIGTERM", None)):
            if s is not None:
                try:
                    signal.signal(s, _sig)
                except Exception:
                    pass
        while not stop_evt.is_set() and self.bot_thread.is_alive():
            stop_evt.wait(1.0)
        self.stop()

    def get_admin_keyboard(self):
        return {"keyboard": [
            [{"text": "🛒 Buy IDs"}, {"text": "⚙️ Control Panel"}],
            [{"text": "📊 Check Status"}, {"text": "👥 Customers"}],
            [{"text": "💳 Deposit Info"}, {"text": "📋 Policy"}],
            [{"text": "💰 My Balance"}, {"text": "📈 Purchases"}],
            [{"text": "📑 Report"}, {"text": "🗄️ Backup"}],
            [{"text": "💱 Currency"}, {"text": "📊 Stats"}],
            [{"text": "📢 Announce"}, {"text": LANG_BTN}]
        ], "resize_keyboard": True, "one_time_keyboard": False, "is_persistent": True}

    def get_user_keyboard(self):
        return {"keyboard": [
            [{"text": "🛒 Buy IDs"}],
            [{"text": "💳 Deposit"}, {"text": "💰 My Balance"}],
            [{"text": "📜 My Orders"}, {"text": "📋 Policy"}],
            [{"text": "📞 Support"}, {"text": LANG_BTN}]
        ], "resize_keyboard": True, "one_time_keyboard": False, "is_persistent": True}

    def get_currency_keyboard(self):
        return {"keyboard": [
            [{"text": "💵 USD ($)"}, {"text": "৳ BDT (Taka)"}],
            [{"text": "❌ Cancel"}]
        ], "resize_keyboard": True, "one_time_keyboard": True}

    def get_format_keyboard(self):
        return {"keyboard": [
            [{"text": "📄 Text File (.txt)"}, {"text": "📊 Excel File (.xlsx)"}],
            [{"text": "❌ Cancel"}]
        ], "resize_keyboard": True, "one_time_keyboard": True}

    def get_panel_choice_keyboard(self):
        return {"keyboard": [
            [{"text": "📥 Add to 61xx"}, {"text": "📥 Add to 1000x"}],
            [{"text": "📥 Add to newid"}],
            [{"text": "❌ Cancel"}]
        ], "resize_keyboard": True, "one_time_keyboard": True}

    def bot_polling_loop(self, token):
        while self.bot_running:
            self._maybe_auto_backup(token)
            try:
                url = (f"https://api.telegram.org/bot{token}/getUpdates?timeout=30"
                       f"&allowed_updates=%5B%22message%22%2C%22callback_query%22%5D")
                if self.last_update_id > 0:
                    url += f"&offset={self.last_update_id + 1}"
                response = self.tg_session.get(url, timeout=35)
                if response.status_code == 200:
                    data = response.json()
                    if data.get("ok") and data["result"]:
                        for update in data["result"]:
                            self.last_update_id = update["update_id"]
                            try:
                                if "callback_query" in update:
                                    self.process_callback(update["callback_query"], token)
                                else:
                                    self.process_update(update, token)
                            except Exception as e:
                                logger.error(f"[Update error] {e}")
                                logger.error(traceback.format_exc())
                else:
                    time.sleep(1)
            except:
                time.sleep(2)

    # -------- Callback queries --------
    def process_callback(self, query, token):
        chat_id = str(query["message"]["chat"]["id"])
        message_id = query["message"]["message_id"]
        data = query["data"]
        master_id = self.chat_id_entry.get().strip()
        is_admin = (chat_id == master_id)
        if not is_admin and not self.users_enabled:
            try:
                self.tg_session.post(
                    f"https://api.telegram.org/bot{token}/answerCallbackQuery",
                    data={"callback_query_id": query.get("id")}, timeout=5)
            except Exception:
                pass
            return
        if data == "soldout":
            try:
                self.tg_session.post(
                    f"https://api.telegram.org/bot{token}/answerCallbackQuery",
                    data={"callback_query_id": query.get("id"),
                          "text": "❌ Sold Out — this ID type is out of stock.",
                          "show_alert": "true"}, timeout=5)
            except Exception:
                pass
            return
        if data.startswith("cp:"):
            if is_admin:
                self._cp_callback(token, chat_id, message_id, data, query.get("id"))
            return
        if not is_admin:
            self._mirror_out_text(token, chat_id, str(data), label="🔘 Button tapped")
        kb = self.get_admin_keyboard() if is_admin else self.get_user_keyboard()

        if data == "close":
            try:
                requests.post(f"https://api.telegram.org/bot{token}/deleteMessage",
                              data={"chat_id": chat_id, "message_id": message_id},
                              timeout=5)
            except:
                pass
            return

        if data.startswith("buy_"):
            p_type = data.replace("buy_", "")
            if not [l for l in self._get_panel(p_type).get_data().splitlines() if l.strip()]:
                self.send_message(token, chat_id,
                                  f"❌ {type_label(p_type)} is Sold Out right now.", reply_markup=kb)
                return
            user = self.customer_db.get_user(chat_id)
            cur = self.customer_db.admin_currency() if is_admin else (user[3] if user else "BDT")
            price = self.customer_db.get_price(p_type, cur)
            self.user_states[chat_id] = f"waiting_{p_type}"
            sym = "৳" if cur == "BDT" else "$"
            # FIX: Use :g to remove .00 from price
            self.send_message(token, chat_id,
                              f"📝 How many {type_label(p_type)} IDs do you want?\n"
                              f"Price: {sym}{price:g} per piece\n\n"
                              f"Send the number only:",
                              reply_markup=kb)
            return

        if data.startswith("dep_custom_") and is_admin:
            dep_id = int(data.split("_")[2])
            self.user_states[chat_id] = f"waiting_dep_amt_{dep_id}"
            self.send_message(token, chat_id,
                              f"✏️ Send the amount and currency for Deposit #{dep_id}.\n\n"
                              f"Format: `<amount> <BDT|USD>`\n"
                              f"Example: `1500 BDT` or `10.5 USD`",
                              reply_markup=kb)
            return

        if data.startswith("dep_a_") and is_admin:
            parts = data.split("_")
            dep_id = int(parts[2])
            amt = float(parts[3])
            cur = parts[4]
            ok = self.customer_db.approve_deposit(dep_id, amt, cur)
            if ok:
                dep = self.customer_db.get_deposit(dep_id)
                user_chat = dep[1]
                self._notify_payment(token, dep_id, amt, cur)
                self.send_message(token, user_chat,
                                  f"✅ Deposit approved!\n"
                                  f"+{amt} {cur} has been added to your balance.")
                self._edit_admin_photo(token, chat_id, message_id,
                                       f"✅ APPROVED: {amt} {cur}\n"
                                       f"Deposit #{dep_id} → {user_chat}")
            return

        if data.startswith("dep_u_") and is_admin:
            dep_id = int(data.split("_")[2])
            res = self.customer_db.undo_deposit(dep_id)
            if res:
                u_chat, u_amt, u_cur, short = res
                self.send_message(token, u_chat, "⚠️ Your deposit was reversed by admin.")
                extra = f"\n(user had already spent {short:g}, not recoverable)" if short else ""
                self._edit_admin_photo(token, chat_id, message_id,
                                       f"↩️ UNDONE: Deposit #{dep_id} ({u_amt:g} {u_cur}) taken back from {u_chat}."
                                       f"\nTransaction ID blocked.{extra}")
            else:
                self._edit_admin_photo(token, chat_id, message_id,
                                       f"ℹ️ Deposit #{dep_id} is not in approved state (nothing to undo).")
            return

        if data.startswith("dep_r_") and is_admin:
            dep_id = int(data.split("_")[2])
            self.customer_db.reject_deposit(dep_id, "Rejected by admin")
            dep = self.customer_db.get_deposit(dep_id)
            if dep:
                self.send_message(token, dep[1],
                                  f"❌ Deposit #{dep_id} was rejected.\n"
                                  f"Please contact admin for details.")
            self._edit_admin_photo(token, chat_id, message_id,
                                   f"❌ REJECTED: Deposit #{dep_id}")
            return

    def _edit_admin_photo(self, token, chat_id, message_id, caption):
        try:
            requests.post(f"https://api.telegram.org/bot{token}/editMessageCaption",
                          data={"chat_id": chat_id,
                                "message_id": message_id,
                                "caption": caption}, timeout=5)
        except:
            pass

    # -------- Main update --------
    def process_update(self, update, token):
        if "message" not in update:
            return
        msg = update["message"]
        # v31: messages from the Topics group (admin <-> user chat boxes)
        if msg.get("chat", {}).get("type") in ("group", "supergroup"):
            self._handle_group_message(msg, token)
            return
        chat_id = str(msg["chat"]["id"])
        master_id = self.chat_id_entry.get().strip()
        is_admin = (chat_id == master_id)

        # Users switch OFF -> completely ignore everyone except Master Admin
        if not is_admin and not self.users_enabled:
            return

        name = msg["chat"].get("first_name", "Unknown")
        uname = msg["chat"].get("username", "")
        if not is_admin:
            self.customer_db.register_user(chat_id, name, uname)
            self._mirror_in(token, msg)          # v31: copy user's message into their topic

        if not is_admin and self.customer_db.is_banned(chat_id):
            self.send_message(token, chat_id, "🚫 You are blocked. Contact @Jihad777.")
            return

        # ---------- PHOTO / DEPOSIT SCREENSHOT ----------
        if "photo" in msg:
            if is_admin:
                if self.user_states.get(chat_id) == "waiting_announce":
                    self.user_states.pop(chat_id, None)
                    self._broadcast(token, chat_id, photo=msg["photo"][-1]["file_id"],
                                    caption="📢 " + (msg.get("caption") or ""))
                else:
                    self.send_message(token, chat_id,
                                      "ℹ️ Only users can submit deposit screenshots.\n"
                                      "To announce with a photo, tap 📢 Announce first.")
                return

            try:
                largest = msg["photo"][-1]
                file_id = largest["file_id"]
                dep_id = self.customer_db.add_deposit(chat_id, uname, name, file_id)
                logger.info(f"[DEPOSIT] Saved #{dep_id} chat={chat_id} name='{name}'")
            except Exception as e:
                err_text = f"{type(e).__name__}: {str(e)}"
                logger.error(f"[DEPOSIT FAIL] {err_text}")
                logger.error(traceback.format_exc())
                self.send_message(token, chat_id,
                                  f"❌ Deposit save error.\n\n"
                                  f"Error: {err_text[:200]}\n\n"
                                  f"Admin has been notified. Please contact @Jihad777.")
                if master_id:
                    self._send_text_direct(
                        token, master_id,
                        f"⚠️ DEPOSIT SAVE ERROR\n"
                        f"👤 User: {name} (@{uname or 'N/A'})\n"
                        f"🆔 Chat ID: {chat_id}\n"
                        f"❌ Error: {err_text}\n\n"
                        f"Check log: {LOG_FILE}")
                return

            user_info = self.customer_db.get_user(chat_id)
            cur_bal = user_info[2] if user_info else 0
            cur = user_info[3] if user_info else "BDT"
            sym = "৳" if cur == "BDT" else "$"

            auto_mode = self._auto_enabled() and cur == "BDT"
            if auto_mode:
                self.send_message(token, chat_id,
                                  f"🔎 Screenshot received! Verifying your payment automatically...\n\n"
                                  f"Deposit ID: #{dep_id}\n"
                                  f"This takes a few seconds.")
            else:
                self.send_message(token, chat_id,
                                  f"✅ Deposit screenshot received!\n\n"
                                  f"Deposit ID: #{dep_id}\n"
                                  f"Admin will verify and add balance shortly.\n\n"
                                  f"Your Chat ID: {chat_id}")

            if not master_id:
                logger.error("[DEPOSIT ALERT] Master Chat ID is EMPTY!")
                return

            if auto_mode:
                threading.Thread(
                    target=self._auto_deposit_job,
                    args=(token, dep_id, chat_id, name, uname, file_id,
                          msg.get("caption") or "", msg.get("date") or int(time.time())),
                    daemon=True).start()
                return

            self._send_admin_deposit_card(token, master_id, dep_id, name, uname, chat_id, file_id)
            return

        # ---------- DOCUMENT uploads ----------
        if "document" in msg:
            if not is_admin:
                self.send_message(token, chat_id,
                                  "❌ Sorry, you cannot upload/forward files.")
                return
            state = self.user_states.get(chat_id)
            fid = msg["document"]["file_id"]
            fname = msg["document"].get("file_name", "file.txt")
            if state == "waiting_restore":
                self.send_message(token, chat_id, f"⏳ Restoring from {fname} ...")
                threading.Thread(target=self._restore_from_file,
                                 args=(token, chat_id, fid, fname), daemon=True).start()
                return          # stay in restore mode so you can send the 2nd file too
            if state in ["upload_61xx", "upload_1000x", "upload_newid"]:
                p_type = state.replace("upload_", "")
                self.send_message(token, chat_id,
                                  f"⏳ Downloading & appending to {type_label(p_type)}...")
                threading.Thread(target=self.download_and_append_file,
                                 args=(token, chat_id, fid, p_type, fname),
                                 daemon=True).start()
                del self.user_states[chat_id]
            else:
                self.pending_forward_files[chat_id] = {"fid": fid, "fname": fname}
                self.send_message(token, chat_id,
                                  f"📥 File received: {fname}\n\nWhich series do you want to add it to?",
                                  reply_markup=self.get_panel_choice_keyboard())
            return

        if "text" not in msg:
            return
        text = msg["text"].strip()
        kb = self.get_admin_keyboard() if is_admin else self.get_user_keyboard()
        text = BN_KEYS_REV.get(text, text)   # Bengali button label -> internal English label
        if is_admin:
            _cmd0 = text.split()[0].split("@")[0].lower() if text.startswith("/") else ""
            if text == "⚙️ Control Panel" or _cmd0 in ("/panel", "/cp"):
                self.user_states.pop(chat_id, None)
                self._cp_main(token, chat_id)
                return
            _st = self.user_states.get(chat_id)
            if _st and _st.startswith("cp_") and self._cp_text(token, chat_id, _st, text):
                return

        if is_admin and text == LANG_BTN:
            self.send_message(token, chat_id, "🌐 Choose menu language / মেনুর ভাষা বেছে নিন:",
                              reply_markup=self.get_lang_keyboard())
            return
        if is_admin and text in (LANG_EN_BTN, LANG_BN_BTN):
            new_l = "bn" if text == LANG_BN_BTN else "en"
            self.customer_db.set_setting("admin_lang", new_l)
            self.send_message(token, chat_id,
                              "✅ মেনুর ভাষা বাংলা করা হয়েছে।\nনিয়মাবলী ও পেমেন্ট তথ্য এখন বাংলায় দেখাবে।"
                              if new_l == "bn" else "✅ Menu language set to English.",
                              reply_markup=self.get_admin_keyboard())
            return
        if is_admin and text.startswith("/policy"):
            arg = text.split()[1].lower() if len(text.split()) > 1 else ""
            if arg in ("bn", "en"):
                self._send_exact(token, chat_id, BN_POLICY if arg == "bn" else POLICY_TEXT)
            else:
                self.send_message(token, chat_id, "Usage: /policy bn  |  /policy en", reply_markup=kb)
            return
        if is_admin and text.startswith("/pay"):
            arg = text.split()[1].lower() if len(text.split()) > 1 else ""
            table = {"bn": BN_PAYMENT, "en": PAYMENT_INFO, "usd": PAYMENT_INFO_USD, "usdbn": BN_PAYMENT_USD}
            if arg in table:
                self._send_exact(token, chat_id, table[arg])
            else:
                self.send_message(token, chat_id,
                                  "Usage: /pay bn | en | usd | usdbn\n"
                                  "bn/en = full (BDT + crypto), usd/usdbn = crypto only (foreign buyers)",
                                  reply_markup=kb)
            return

        if is_admin and text == "📢 Announce":
            self.user_states[chat_id] = "waiting_announce"
            self.send_message(token, chat_id,
                              "📢 Send the announcement now (text, or a photo with caption).\n"
                              "It will go to ALL users. Send ❌ Cancel to abort.")
            return
        if is_admin and self.user_states.get(chat_id) == "waiting_announce" and not text.startswith("/"):
            self.user_states.pop(chat_id, None)
            if text.lower() in ("cancel", "❌ cancel"):
                self.send_message(token, chat_id, "❌ Cancelled.", reply_markup=kb)
            else:
                self._broadcast(token, chat_id, text="📢 Announcement\n━━━━━━━━━━━━━━\n" + text)
            return
        if not is_admin and text == LANG_BTN:
            self.send_message(token, chat_id, "🌐 Choose your language / আপনার ভাষা বেছে নিন:",
                              reply_markup=self.get_lang_keyboard())
            return
        if not is_admin and text in (LANG_EN_BTN, LANG_BN_BTN):
            new_lang = "bn" if text == LANG_BN_BTN else "en"
            self.customer_db.set_lang(chat_id, new_lang)
            self.__dict__.setdefault("_lang_cache", {})[chat_id] = new_lang
            text = "/start"
        if not is_admin and text in (REGION_BD_BTN, REGION_OTHER_BTN):
            if not self.customer_db.region_done(chat_id):
                cur_new = "BDT" if text == REGION_BD_BTN else "USD"
                self.customer_db.set_region(chat_id, cur_new)
                self._notify_admin(token,
                                   f"🆕 New user\n👤 {name}" + (f" (@{uname})" if uname else "") +
                                   f"\n🆔 {chat_id}\n🌍 Region: "
                                   f"{'Bangladesh (BDT)' if cur_new == 'BDT' else 'Other country (USD)'}")
            text = "/start"
        if not is_admin and not self.customer_db.region_done(chat_id):
            if self.customer_db.get_lang(chat_id) is None:
                self.send_message(token, chat_id, "🌐 Choose your language / আপনার ভাষা বেছে নিন:",
                                  reply_markup=self.get_lang_keyboard())
            else:
                self.send_message(token, chat_id,
                                  "🌍 Select your country / আপনার দেশ বেছে নিন:\n\n"
                                  "🇧🇩 Bangladesh → prices in ৳ (bKash / Nagad / Rocket / Binance)\n"
                                  "🌍 Other countries → prices in $ (Binance Pay / USDT)",
                                  reply_markup=self.get_region_keyboard())
            return
        if (not is_admin and text in ("/start", "/help")
                and self.customer_db.get_lang(chat_id) is None):
            self.send_message(token, chat_id, "🌐 Choose your language / আপনার ভাষা বেছে নিন:",
                              reply_markup=self.get_lang_keyboard())
            return

        # --- Custom Deposit Amount Handler ---
        if is_admin:
            state = self.user_states.get(chat_id)
            if state and state.startswith("waiting_dep_amt_"):
                if text.lower() in ("cancel", "❌ cancel"):
                    self.user_states.pop(chat_id, None)
                    self.send_message(token, chat_id, "❌ Custom amount cancelled.", reply_markup=kb)
                    return
                
                if not text.startswith("/"):
                    dep_id = int(state.replace("waiting_dep_amt_", ""))
                    parts = text.split()
                    if len(parts) == 2:
                        try:
                            amt = float(parts[0])
                            cur = parts[1].upper()
                            if cur not in ("BDT", "USD"):
                                raise ValueError
                            
                            # Approve deposit
                            if self.customer_db.approve_deposit(dep_id, amt, cur):
                                dep = self.customer_db.get_deposit(dep_id)
                                self._notify_payment(token, dep_id, amt, cur)
                                if dep:
                                    self.send_message(token, dep[1],
                                                      f"✅ Deposit approved!\n"
                                                      f"+{amt} {cur} has been added to your balance.")
                                self.send_message(token, chat_id,
                                                  f"✅ Deposit #{dep_id} approved: {amt} {cur}",
                                                  reply_markup=kb)
                            else:
                                self.send_message(token, chat_id, "❌ Deposit not found.")
                            self.user_states.pop(chat_id, None)
                        except ValueError:
                            self.send_message(token, chat_id,
                                              "❌ Invalid format. Send `<amount> <BDT|USD>` (e.g., `1500 BDT`)")
                    else:
                        self.send_message(token, chat_id,
                                          "❌ Invalid format. Send `<amount> <BDT|USD>` (e.g., `1500 BDT`)")
                    return

        if is_admin and text == "❌ Cancel" and chat_id in self.pending_forward_files:
            self.pending_forward_files.pop(chat_id, None)
            self.send_message(token, chat_id, "❌ Cancelled.", reply_markup=kb)
            return

        if is_admin and text in ("📥 Add to 61xx", "📥 Add to 1000x", "📥 Add to newid"):
            if chat_id in self.pending_forward_files:
                data = self.pending_forward_files.pop(chat_id)
                p_type = "61xx" if "61xx" in text else ("newid" if "newid" in text else "1000x")
                self.send_message(token, chat_id, f"⏳ Downloading & appending to {type_label(p_type)}...",
                                  reply_markup=kb)
                threading.Thread(target=self.download_and_append_file,
                                 args=(token, chat_id, data["fid"], p_type, data["fname"]),
                                 daemon=True).start()
            return

        if text in ("/start", "/help"):
            user = self.customer_db.get_user(chat_id)
            bal = user[2] if user else 0
            cur = user[3] if user else "BDT"
            sym = "৳" if cur == "BDT" else "$"
            help_text = (f"✨ *Welcome, {name}!*\n"
                         f"Premium ID Store — fast, automatic delivery 24/7.\n\n"
                         f"💳 Balance: {sym}{bal:.2f} ({cur})\n"
                         f"🆔 Your ID: `{chat_id}`\n\n"
                         f"Use the menu below to buy, deposit or request support.")
            if is_admin:
                help_text += ("\n\n👑 Admin Commands:\n"
                              "/users — list users\n"
                              "/deposits — pending deposits\n"
                              "/purchases — purchase log\n"
                              "/approve <id> <amt> <BDT|USD>\n"
                              "/reject <id>\n"
                              "/addbal <chat_id> <amt> <BDT|USD>\n"
                              "/setprice <61xx|1000x|newid> <amt> <BDT|USD>\n"
                              "/setcurrency <chat_id> <BDT|USD>\n"
                              "/deluser <chat_id> — delete user permanently\n"
                              "/resetbal <chat_id> — reset balance to 0\n"
                              "/test — verify bot connection\n"
                              "/report [days] — sales report\n"
                              "/ban <id> | /unban <id>\n"
                              "/broadcast <text>\n"
                              "/setdisc <50|100> <percent>\n"
                              "/setlow <count> — low stock alert\n"
                              "/shop on|off\n"
                              "📢 Announce button — message/photo to all users\n"
                              "🌐 Language button — switch your menu to বাংলা\n"
                              "/policy bn|en — policy to share (forwardable)\n"
                              "/pay bn|en|usd|usdbn — payment info to share\n"
                              "/stats — lifetime sales totals\n"
                              "/setlog <group_id>|off — copy orders/payments to a group\n"
                              "/mycur BDT|USD — your own sales currency\n"
                              "/find <id> — user profile\n"
                              "/reply <id> <text> — message a user\n"
                              "/backup — DB + stock backup (auto every 24h)\n"
                              "/restore — load old customers.db / stock json\n"
                              "⚙️ Control Panel → 🤖 Auto Deposit — auto-approve settings")
            self.send_message(token, chat_id, help_text, reply_markup=kb)
            return

        if text == "/myid":
            self.send_message(token, chat_id, f"🆔 Your Chat ID: {chat_id}")
            return

        if text == "💰 My Balance":
            user = self.customer_db.get_user(chat_id)
            if user:
                bal, cur = user[2], user[3]
                sym = "৳" if cur == "BDT" else "$"
                tier, _, tq = self.customer_db.get_vip(chat_id)
                self.send_message(token, chat_id,
                                  f"💎 *Your Wallet*\n━━━━━━━━━━━━━━\n"
                                  f"💰 Balance: {sym}{bal:.2f} ({cur})\n"
                                  f"🏅 Status: {tier}\n📦 Total purchased: {tq} pcs\n"
                                  f"━━━━━━━━━━━━━━",
                                  reply_markup=kb)
            else:
                self.send_message(token, chat_id, "💰 Balance: 0", reply_markup=kb)
            return

        if text == "📋 Policy":
            self.send_message(token, chat_id, POLICY_TEXT, reply_markup=kb)
            return

        if text in ("💳 Deposit", "💳 Deposit Info"):
            u_dep = self.customer_db.get_user(chat_id)
            if not is_admin and u_dep and u_dep[3] == "USD":
                self.send_message(token, chat_id, PAYMENT_INFO_USD, reply_markup=kb)
            else:
                self.send_message(token, chat_id, PAYMENT_INFO, reply_markup=kb)
            return

        if text == "📞 Support":
            self.send_message(token, chat_id,
                              "📞 Contact admin: @Jihad777\n"
                              "Please share your deposit screenshot + Chat ID.",
                              reply_markup=kb)
            return

        if text == "📜 My Orders":
            rows = self.customer_db.get_user_purchases(chat_id)[:10]
            if not rows:
                self.send_message(token, chat_id, "📭 No orders yet.", reply_markup=kb)
            else:
                lines = ["📜 *Order History* (last 10)\n"]
                for panel, qty, up, cost, cur, ts in rows:
                    sym = "৳" if cur == "BDT" else "$"
                    lines.append(f"{ts[:16]} | {panel} x{qty} | {sym}{cost:g}")
                self.send_message(token, chat_id, "\n".join(lines), reply_markup=kb)
            return

        if is_admin and text in ("📑 Report", "/report"):
            self._bot_report(token, chat_id, 1)
            return
        if is_admin and text.startswith("/report "):
            try:
                self._bot_report(token, chat_id, max(1, min(int(text.split()[1]), 90)))
            except ValueError:
                self.send_message(token, chat_id, "❌ Usage: /report <days>")
            return

        if is_admin and text.startswith("/reply "):
            p = text.split(" ", 2)
            if len(p) == 3:
                self.send_message(token, p[1], f"📩 *Message from Admin*\n\n{p[2]}")
                self.send_message(token, chat_id, "✅ Sent.")
            else:
                self.send_message(token, chat_id, "❌ Usage: /reply <chat_id> <message>")
            return
        if is_admin and text.startswith("/find "):
            p = text.split()
            u = self.customer_db.get_user(p[1]) if len(p) == 2 else None
            if not u:
                self.send_message(token, chat_id, "❌ User not found.")
            else:
                tier, _, tq = self.customer_db.get_vip(u[0])
                self.send_message(token, chat_id,
                                  f"👤 {u[1]} (@{u[4] or 'N/A'})\n🆔 {u[0]}\n💰 {u[2]:.2f} {u[3]}\n"
                                  f"🏅 {tier} | 📦 {tq} pcs\n🚫 Banned: {self.customer_db.is_banned(u[0])}")
            return
        if is_admin and text in ("🗄️ Backup", "/backup"):
            self.send_message(token, chat_id, "⏳ Creating backup...")
            threading.Thread(target=self._send_backup, args=(token, chat_id), daemon=True).start()
            return
        if is_admin and text == "/restore":
            self.user_states[chat_id] = "waiting_restore"
            self.send_message(token, chat_id,
                              "♻️ RESTORE MODE\n\n"
                              "Send your old files now (one by one):\n"
                              "• customers.db → users, balances, deposits, purchases, prices, settings\n"
                              "• data_backup_v20.json (or stock_backup_*.json) → stock IDs\n\n"
                              "⚠️ The database file REPLACES the current data "
                              "(a safety copy of the current one is kept).\n"
                              "Send /done when finished or ❌ Cancel to leave.",
                              reply_markup=kb)
            return
        if (is_admin and self.user_states.get(chat_id) == "waiting_restore"
                and text.lower() in ("/done", "done", "cancel", "❌ cancel", "/cancel")):
            self.user_states.pop(chat_id, None)
            self.send_message(token, chat_id, "✅ Restore mode closed. Use /test or 📊 Stats to check.",
                              reply_markup=kb)
            return
        if is_admin and text.startswith("/ban "):
            p = text.split()
            if len(p) == 2 and self.customer_db.get_user(p[1]):
                self.customer_db.set_banned(p[1], True)
                self.send_message(token, chat_id, f"🚫 {p[1]} banned.")
            else:
                self.send_message(token, chat_id, "❌ Usage: /ban <chat_id> (user must exist)")
            return
        if is_admin and text.startswith("/unban "):
            p = text.split()
            if len(p) == 2:
                self.customer_db.set_banned(p[1], False)
                self.send_message(token, chat_id, f"✅ {p[1]} unbanned.")
            return
        if is_admin and text.startswith("/broadcast "):
            msg_text = text.split(" ", 1)[1]
            ids = self.customer_db.all_user_ids()
            self.send_message(token, chat_id, f"📢 Sending to {len(ids)} users...")
            def _bc():
                for uid in ids:
                    self._send_text_direct(token, uid, f"📢 {msg_text}")
                    time.sleep(0.05)
                self._send_text_direct(token, chat_id, "✅ Broadcast finished.")
            threading.Thread(target=_bc, daemon=True).start()
            return
        if is_admin and text.startswith("/setdisc "):
            p = text.split()
            if len(p) == 3 and p[1] in ("50", "100"):
                try:
                    float(p[2])
                    self.customer_db.set_setting(f"disc_{p[1]}", p[2])
                    self.send_message(token, chat_id, f"✅ {p[1]}+ discount = {p[2]}%")
                except ValueError:
                    self.send_message(token, chat_id, "❌ Invalid percent.")
            else:
                self.send_message(token, chat_id, "❌ Usage: /setdisc <50|100> <percent>")
            return
        if is_admin and text.startswith("/setlow "):
            p = text.split()
            if len(p) == 2 and p[1].isdigit():
                self.customer_db.set_setting("low_stock_limit", p[1])
                self.send_message(token, chat_id, f"✅ Low-stock alert at {p[1]}")
            return
        if is_admin and text == "💱 Currency":
            cur_now = self.customer_db.admin_currency()
            self.send_message(token, chat_id,
                              f"💱 Select your sales currency\nCurrent: {cur_now}",
                              reply_markup=self.get_currency_keyboard())
            return
        if is_admin and (text in ("💵 USD ($)", "৳ BDT (Taka)") or text.startswith("/mycur")):
            p = text.split()
            if text == "💵 USD ($)":
                new = "USD"
            elif text == "৳ BDT (Taka)":
                new = "BDT"
            elif len(p) == 2 and p[1].upper() in ("BDT", "USD"):
                new = p[1].upper()
            else:
                self.send_message(token, chat_id, "❌ Usage: /mycur BDT|USD", reply_markup=kb)
                return
            self.customer_db.set_setting("admin_currency", new)
            sym = "৳" if new == "BDT" else "$"
            self.send_message(token, chat_id,
                              f"✅ Currency set to {new} ({sym})\n"
                              f"Prices & receipts will now use {sym}.",
                              reply_markup=kb)
            return
        if is_admin and text in ("📊 Stats", "/stats"):
            self._bot_stats(token, chat_id)
            return
        if is_admin and text.startswith("/setlog"):
            p = text.split()
            if len(p) == 2 and p[1].lower() == "off":
                self.customer_db.set_setting("log_chat_id", "")
                self.send_message(token, chat_id, "✅ Log group disabled.")
            elif len(p) == 2 and p[1].lstrip("-").isdigit():
                self.customer_db.set_setting("log_chat_id", p[1])
                self.send_message(token, chat_id, f"✅ Orders & payments will also be posted to {p[1]}")
            else:
                self.send_message(token, chat_id, "Usage: /setlog <group_chat_id> | /setlog off")
            return
        if is_admin and text in ("/shop off", "/shop on"):
            self.customer_db.set_setting("shop_open", "1" if text.endswith("on") else "0")
            self.send_message(token, chat_id, f"🏪 Shop is now {'OPEN' if text.endswith('on') else 'CLOSED'}.")
            return

        if text == "📈 Purchases" and is_admin:
            self._bot_show_purchases(token, chat_id)
            return

        # Buy flow (FIXED TEXT TRUNCATION)
        if text == "🛒 Buy IDs":
            if not is_admin and self.customer_db.get_setting("shop_open", "1") != "1":
                self.send_message(token, chat_id, "🏪 Shop is temporarily closed. Please try later.", reply_markup=kb)
                return
            load_button_names(self.customer_db)   # pick up any name change made in Control Panel
            c1 = len([l for l in self.panel_61xx.get_data().splitlines() if l.strip()])
            c2 = len([l for l in self.panel_1000x.get_data().splitlines() if l.strip()])
            user = self.customer_db.get_user(chat_id)
            cur = self.customer_db.admin_currency() if is_admin else (user[3] if user else "BDT")
            p1 = self.customer_db.get_price("61xx", cur)
            p2 = self.customer_db.get_price("1000x", cur)
            c3 = len([l for l in self.panel_newid.get_data().splitlines() if l.strip()])
            p3 = self.customer_db.get_price("newid", cur)
            sym = "৳" if cur == "BDT" else "$"
            
            # Using :g format to drop trailing .00 and shortening the text
            def stock_btn(label, count, price, key):
                if count <= 0:   # empty stock -> red "Sold Out" button, not purchasable
                    return [{"text": f"❌ {label} | Sold Out",
                             "callback_data": "soldout", "style": "danger"}]
                return [{"text": f"✅ {label} | {count} left | {sym}{price:g}",
                         "callback_data": f"buy_{key}"}]
            reply_markup = {"inline_keyboard": [
                stock_btn(button_name("1000x"), c2, p2, "1000x"),
                stock_btn(button_name("61xx"), c1, p1, "61xx"),
                stock_btn(button_name(NEWID_KEY), c3, p3, "newid"),
                [{"text": "❌ Close", "callback_data": "close"}]
            ]}
            self.send_message(token, chat_id, "🛍️ *Premium Store*\nSelect an ID type:\n👇 Tap a type below to buy your IDs",
                              reply_markup=reply_markup)
            return

        # Admin commands
        if is_admin:
            if text == "📤 Upload 61xx":
                self.user_states[chat_id] = "upload_61xx"
                self.send_message(token, chat_id, "📁 Send 61xx series file:")
                return
            if text == "📤 Upload newid":
                self.user_states[chat_id] = "upload_newid"
                self.send_message(token, chat_id, f"📁 Send {NEWID_LABEL} file:")
                return
            if text == "📤 Upload 1000x":
                self.user_states[chat_id] = "upload_1000x"
                self.send_message(token, chat_id, "📁 Send 1000x series file:")
                return
            if text == "📊 Check Status":
                c1 = len([l for l in self.panel_61xx.get_data().splitlines() if l.strip()])
                c2 = len([l for l in self.panel_1000x.get_data().splitlines() if l.strip()])
                c3 = len([l for l in self.panel_newid.get_data().splitlines() if l.strip()])
                self.send_message(token, chat_id,
                                  f"📊 Status:\n61xx: {c1}\n1000x: {c2}\n{NEWID_LABEL}: {c3}",
                                  reply_markup=kb)
                return
            if text == "👥 Customers":
                self._bot_customers_list(token, chat_id)
                return
            if text == "/users":
                self._bot_users_list(token, chat_id)
                return
            if text == "/deposits":
                self._bot_pending_deposits(token, chat_id)
                return
            if text == "/purchases":
                self._bot_show_purchases(token, chat_id)
                return
            if text == "/test":
                user_info = self.customer_db.get_user(chat_id)
                bal = user_info[2] if user_info else 0
                cur = user_info[3] if user_info else "BDT"
                self.send_message(token, chat_id,
                                  f"✅ Bot is working!\n\n"
                                  f"Master Chat ID: {master_id}\n"
                                  f"Your Chat ID: {chat_id}\n"
                                  f"Match: {is_admin}\n"
                                  f"Your Balance: {bal:.2f} {cur}\n"
                                  f"Time (BD): {datetime.now(BD_TZ).strftime('%Y-%m-%d %H:%M:%S')}\n"
                                  f"{self._auto_status_text()}")
                return
            if text.startswith("/approve "):
                parts = text.split()
                if len(parts) == 4:
                    try:
                        dep_id = int(parts[1])
                        amt = float(parts[2])
                        cur = parts[3].upper()
                        if cur not in ("BDT", "USD"):
                            raise ValueError
                        if self.customer_db.approve_deposit(dep_id, amt, cur):
                            dep = self.customer_db.get_deposit(dep_id)
                            self._notify_payment(token, dep_id, amt, cur)
                            self.send_message(token, dep[1],
                                              f"✅ Deposit approved!\n"
                                              f"+{amt} {cur} added to your balance.")
                            self.send_message(token, chat_id,
                                              f"✅ Deposit #{dep_id} approved: {amt} {cur}")
                        else:
                            self.send_message(token, chat_id, "❌ Deposit not found.")
                    except:
                        self.send_message(token, chat_id,
                                          "❌ Usage: /approve <id> <amt> <BDT|USD>")
                return
            if text.startswith("/reject "):
                parts = text.split()
                if len(parts) == 2:
                    try:
                        dep_id = int(parts[1])
                        self.customer_db.reject_deposit(dep_id, "Rejected by admin")
                        dep = self.customer_db.get_deposit(dep_id)
                        if dep:
                            self.send_message(token, dep[1],
                                              f"❌ Deposit #{dep_id} rejected.")
                            self.send_message(token, chat_id,
                                              f"❌ Deposit #{dep_id} rejected.")
                    except:
                        pass
                return
            if text.startswith("/deluser "):
                parts = text.split()
                if len(parts) == 2:
                    cid = parts[1].strip()
                    user = self.customer_db.get_user(cid)
                    if not user:
                        self.send_message(token, chat_id, f"❌ User {cid} not found.")
                        return
                    self.customer_db.delete_user(cid)
                    self.send_message(token, chat_id,
                                      f"✅ User {cid} ({user[1]}) deleted permanently.\n"
                                      f"Balance, purchases & deposits cleared.")
                else:
                    self.send_message(token, chat_id,
                                      "❌ Usage: /deluser <chat_id>")
                return
            if text.startswith("/resetbal "):
                parts = text.split()
                if len(parts) == 2:
                    cid = parts[1].strip()
                    user = self.customer_db.get_user(cid)
                    if not user:
                        self.send_message(token, chat_id, f"❌ User {cid} not found.")
                        return
                    self.customer_db.reset_balance(cid)
                    self.send_message(token, chat_id,
                                      f"✅ Balance reset to 0 for {cid} ({user[1]}).")
                else:
                    self.send_message(token, chat_id,
                                      "❌ Usage: /resetbal <chat_id>")
                return
            if text.startswith("/addbal "):
                parts = text.split()
                if len(parts) >= 3:
                    try:
                        cid = parts[1]
                        amt = float(parts[2])
                        cur = parts[3].upper() if len(parts) > 3 else "BDT"
                        self.customer_db.add_balance(cid, amt, cur)
                        self.send_message(token, chat_id,
                                          f"✅ Added {amt} {cur} to {cid}.")
                    except:
                        self.send_message(token, chat_id,
                                          "❌ Usage: /addbal <chat_id> <amt> [BDT|USD]")
                return
            if text == "/names":
                load_button_names(self.customer_db)
                self.send_message(token, chat_id,
                                  "🏷️ Buy button names:\n"
                                  f"1000x → {button_name('1000x')}\n"
                                  f"61xx → {button_name('61xx')}\n"
                                  f"newid → {button_name(NEWID_KEY)}\n\n"
                                  "Change: /setname <1000x|61xx|newid> <new name>\n"
                                  "Reset:  /setname <1000x|61xx|newid> default")
                return
            if text.startswith("/setname "):
                parts = text.split(None, 2)
                key = parts[1].lower() if len(parts) > 1 else ""
                if key not in DEFAULT_BUTTON_NAMES or len(parts) < 3:
                    self.send_message(token, chat_id,
                                      "❌ Usage: /setname <1000x|61xx|newid> <new name>\n"
                                      "Example: /setname newid Fresh Number ID")
                    return
                name = parts[2].strip()
                if name.lower() == "default":
                    name = ""
                if len(name) > BUTTON_NAME_MAX:
                    self.send_message(token, chat_id,
                                      f"❌ Name too long (max {BUTTON_NAME_MAX} characters).")
                    return
                self.customer_db.set_setting("btn_name_" + key, name)
                load_button_names(self.customer_db)
                self.send_message(token, chat_id,
                                  f"✅ Button '{key}' is now: {button_name(key)}")
                return
            if text.startswith("/setprice "):
                parts = text.split()
                if len(parts) >= 3:
                    try:
                        key = parts[1].lower()
                        val = float(parts[2])
                        cur = parts[3].upper() if len(parts) > 3 else "BDT"
                        self.customer_db.set_price(key, val, cur)
                        self.send_message(token, chat_id,
                                          f"✅ {key} price = {val} {cur}")
                    except:
                        self.send_message(token, chat_id,
                                          "❌ Usage: /setprice <61xx|1000x|newid> <amt> [BDT|USD]")
                return
            if text.startswith("/setcurrency "):
                parts = text.split()
                if len(parts) == 3:
                    try:
                        cid = parts[1]
                        cur = parts[2].upper()
                        if cur not in ("BDT", "USD"):
                            raise ValueError
                        self.customer_db.set_currency(cid, cur)
                        self.send_message(token, chat_id,
                                          f"✅ Currency for {cid} set to {cur}.")
                    except:
                        self.send_message(token, chat_id,
                                          "❌ Usage: /setcurrency <chat_id> <BDT|USD>")
                return
            if text.startswith("/sale "):
                parts = text.split()
                if len(parts) == 3:
                    try:
                        self.customer_db.add_sale(parts[1], float(parts[2]))
                        self.send_message(token, chat_id, "✅ Sale saved.")
                    except:
                        pass
                return
            if text.startswith("/due "):
                self._bot_due(token, chat_id, text.split(" ", 1)[1].strip())
                return
            if text.startswith("/taken"):
                self._bot_add_taken(token, chat_id, text)
                return
            if text.startswith("/paid"):
                self._bot_add_paid(token, chat_id, text)
                return

        # Buy quantity
        state = self.user_states.get(chat_id)
        if state in ["waiting_61xx", "waiting_1000x", "waiting_newid"]:
            p_type = state.replace("waiting_", "")
            try:
                qty = int(text)
                if qty <= 0:
                    raise ValueError
                user = self.customer_db.get_user(chat_id)
                cur = self.customer_db.admin_currency() if is_admin else (user[3] if user else "BDT")
                uname_u = user[4] if user else ""
                sym = "৳" if cur == "BDT" else "$"
                price = self.customer_db.get_price(p_type, cur)
                panel_obj = self._get_panel(p_type)
                stock = len([l for l in panel_obj.get_data().splitlines() if l.strip()])
                if qty > stock:
                    self.send_message(token, chat_id,
                                      f"❌ Only {stock} IDs in stock for {type_label(p_type)}.", reply_markup=kb)
                    self.user_states.pop(chat_id, None)
                    return
                disc = 0 if is_admin else (self.customer_db.get_discount_percent(qty) + self.customer_db.get_vip(chat_id)[1])
                cost = round(qty * price * (1 - disc / 100), 4)
                if not is_admin:
                    if not user or user[2] < cost:
                        self.send_message(token, chat_id,
                                          f"❌ Insufficient balance!\n💳 Tap *Deposit* to top up.\n"
                                          f"Required: {sym}{cost:g}\n"
                                          f"Your balance: {sym}{(user[2] if user else 0):g}",
                                          reply_markup=kb)
                        del self.user_states[chat_id]
                        return
                self.pending_extract[chat_id] = {
                    "panel": p_type, "qty": qty, "cost": cost,
                    "price": price, "currency": cur,
                    "is_admin": is_admin, "username": uname_u
                }
                self.user_states[chat_id] = "waiting_format"
                self.send_message(token, chat_id,
                                  f"🧾 *Order Summary*\n━━━━━━━━━━━━━━\n"
                                  f"📦 {type_label(p_type)} × {qty}\n💰 Total: {sym}{cost:g}\n"
                                  f"━━━━━━━━━━━━━━\nChoose delivery format:",
                                  reply_markup=self.get_format_keyboard())
            except ValueError:
                self.send_message(token, chat_id,
                                  "❌ Please send a valid number:")
            return

        # Format choice
        if state == "waiting_format":
            info = self.pending_extract.get(chat_id)
            if text == "❌ Cancel":
                self.user_states.pop(chat_id, None)
                self.pending_extract.pop(chat_id, None)
                self.send_message(token, chat_id, "❌ Cancelled.", reply_markup=kb)
                return
            if info and text in ("📄 Text File (.txt)", "📊 Excel File (.xlsx)"):
                fmt = "txt" if "Text" in text else "xlsx"
                self.user_states.pop(chat_id, None)
                self.pending_extract.pop(chat_id, None)
                if not info["is_admin"]:
                    self.customer_db.deduct_balance(chat_id, info["cost"])
                    self.customer_db.log_purchase(chat_id, info["username"],
                                                  info["panel"], info["qty"],
                                                  info["price"], info["cost"],
                                                  info["currency"])
                oid = self.customer_db.last_purchase_id(chat_id) if not info["is_admin"] else 0
                u2 = self.customer_db.get_user(chat_id)
                sym2 = "৳" if info["currency"] == "BDT" else "$"
                rec = (f"✅ Order Confirmed\n━━━━━━━━━━━━━━\n"
                       f"🧾 Order: #{oid or 'ADMIN'}\n📦 {type_label(info['panel'])} × {info['qty']}\n"
                       f"💰 Paid: {sym2}{info['cost']:g}\n"
                       f"📄 Format: {fmt.upper()}\n")
                if u2 and not info["is_admin"]:
                    rec += f"💳 Balance left: {sym2}{u2[2]:g}\n"
                rec += f"🕒 {datetime.now().strftime('%Y-%m-%d %H:%M')}\n━━━━━━━━━━━━━━"
                if not info["is_admin"]:
                    buyer = (u2[1] if u2 else "User") + (f" (@{u2[4]})" if u2 and u2[4] else "")
                    self._notify_admin(token,
                                       f"🛒 New Sale\n━━━━━━━━━━━━━━\n"
                                       f"🧾 Order: #{oid}\n👤 Buyer: {buyer}\n🆔 Chat ID: {chat_id}\n"
                                       f"🎁 Product: PC CLONE {type_label(info['panel'])}\n"
                                       f"📦 Quantity: {info['qty']} pcs\n"
                                       f"💲 Per piece: {info['cost'] / info['qty']:.2f} {info['currency']}\n"
                                       f"💲 Total: {info['cost']:.2f} {info['currency']}\n"
                                       f"🕒 {datetime.now().strftime('%Y-%m-%d %H:%M')}")
                self.root.after(0, self._execute_remote_extraction,
                                info["panel"], info["qty"], chat_id, token, fmt, rec)
                return
            self.send_message(token, chat_id,
                              "⚠️ Please use the buttons above.",
                              reply_markup=self.get_format_keyboard())
            return

        self.send_message(token, chat_id,
                          "❌ I didn't understand. Send /start.", reply_markup=kb)

    # -------- Bot helpers --------
    def _bot_users_list(self, token, chat_id):
        rows = self.customer_db.list_users()
        if not rows:
            return self.send_message(token, chat_id, "📭 No users yet.")
        lines = ["👥 Bot Users:\n"]
        for r in rows[:40]:
            cid, name, uname, bal, cur, _ = r
            sym = "৳" if cur == "BDT" else "$"
            lines.append(f"🆔 {cid} | {name} | @{uname or 'N/A'} | {sym}{bal:.2f}")
        self.send_message(token, chat_id, "\n".join(lines),
                          reply_markup=self.get_admin_keyboard())

    def _bot_customers_list(self, token, chat_id):
        rows = self.customer_db.list_all()
        if not rows:
            return self.send_message(token, chat_id, "📭 Empty.")
        lines = ["👥 Customer Due List\n"]
        for u, ids, bill, paid in rows[:30]:
            due = bill - paid
            lines.append(f"{'🟢' if due <= 0 else '🔴'} {u} — Due {due:.0f}৳")
        self.send_message(token, chat_id, "\n".join(lines),
                          reply_markup=self.get_admin_keyboard())

    def _bot_pending_deposits(self, token, chat_id):
        pending = self.customer_db.get_pending_deposits()
        if not pending:
            self.send_message(token, chat_id, "✅ No pending deposits.",
                              reply_markup=self.get_admin_keyboard())
            return
        lines = ["💳 Pending Deposits:\n"]
        for d in pending:
            did, cid, uname, name, amt, cur, method, ts = d
            lines.append(f"#{did} | {cid} | @{uname or 'N/A'} | {name} | {ts[:16]}")
        lines.append("\nUse /approve <id> <amt> <BDT|USD> or /reject <id>")
        self.send_message(token, chat_id, "\n".join(lines),
                          reply_markup=self.get_admin_keyboard())

    def _bot_show_purchases(self, token, chat_id):
        summary = self.customer_db.get_purchases_summary()
        if not summary:
            self.send_message(token, chat_id, "📭 No purchases yet.",
                              reply_markup=self.get_admin_keyboard())
            return
        lines = ["📈 User Purchases:\n"]
        for cid, name, uname, qty, bdt, usd in summary[:25]:
            lines.append(f"🆔 {cid} | {name} | @{uname or 'N/A'} | "
                         f"{qty} pcs | {bdt:.0f}৳ | ${usd:.2f}")
        self.send_message(token, chat_id, "\n".join(lines),
                          reply_markup=self.get_admin_keyboard())

    def _bot_due(self, token, chat_id, username):
        s = self.customer_db.get_summary(username)
        msg = f"👤 {username} | Bill: {s['bill']:.0f}৳ | Due: {s['due']:.0f}৳"
        self.send_message(token, chat_id, msg,
                          reply_markup=self.get_admin_keyboard())

    def _bot_add_taken(self, token, chat_id, text):
        parts = text.split()
        if len(parts) >= 3:
            try:
                self.customer_db.add_taken(parts[1], int(parts[2]),
                                            float(parts[3]) if len(parts) > 3 else 0)
                self.send_message(token, chat_id, "✅ Saved.",
                                  reply_markup=self.get_admin_keyboard())
            except:
                pass

    def _bot_add_paid(self, token, chat_id, text):
        parts = text.split()
        if len(parts) >= 3:
            try:
                self.customer_db.add_payment(parts[1], float(parts[2]))
                self.send_message(token, chat_id, "✅ Payment saved.",
                                  reply_markup=self.get_admin_keyboard())
            except:
                pass

    # -------- File downloads --------
    def download_and_append_file(self, token, chat_id, file_id, panel_type, file_name):
        try:
            res = self.tg_session.get(
                f"https://api.telegram.org/bot{token}/getFile?file_id={file_id}",
                timeout=10)
            data = res.json()
            if not data.get("ok"):
                return
            fpath = data["result"]["file_path"]
            content = self.tg_session.get(
                f"https://api.telegram.org/file/bot{token}/{fpath}",
                timeout=20).content
            ext = file_name.lower().split('.')[-1]
            if ext in ['xlsx', 'xls']:
                df = pd.read_excel(io.BytesIO(content), header=None)
                if not df.empty:
                    first_row = ' '.join(df.iloc[0].dropna().astype(str)).lower()
                    if 'uid' in first_row or 'password' in first_row or 'cookies' in first_row:
                        df = df.iloc[1:]
                lines = df.apply(lambda row: '\t'.join(row.dropna().astype(str)),
                                 axis=1).tolist()
            else:
                lines = content.decode('utf-8', errors='ignore').splitlines()
            self.root.after(0, self._execute_bot_upload,
                            panel_type, lines, chat_id, token)
        except Exception as e:
            self.send_message(token, chat_id, f"❌ Error: {str(e)}")

    def _get_panel(self, p_type):
        if p_type == "61xx":
            return self.panel_61xx
        if p_type == "newid":
            return self.panel_newid
        return self.panel_1000x

    def _execute_bot_upload(self, panel_type, lines, chat_id, token):
        panel = self._get_panel(panel_type)
        valid = [l.strip() for l in lines if l.strip()]
        panel.append_data(valid)
        self.silent_auto_save()
        self.send_message(token, chat_id,
                          f"✅ {len(valid)} IDs added to {type_label(panel_type)}!",
                          reply_markup=self.get_admin_keyboard())

    def _execute_remote_extraction(self, panel_type, qty, chat_id, token, fmt="txt", receipt=None):
        panel = self._get_panel(panel_type)
        lines = [l for l in panel.get_data().strip().splitlines() if l.strip()]
        kb = (self.get_admin_keyboard() if chat_id == self.chat_id_entry.get().strip()
              else self.get_user_keyboard())
        if qty > len(lines):
            self.send_message(token, chat_id,
                              f"❌ Only {len(lines)} IDs available! Refunding...",
                              reply_markup=kb)
            self._refund_last(chat_id, token)
            return
        extract, remaining = lines[:qty], lines[qty:]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            if fmt == "txt":
                fname = f"{panel_type}_IDs_{qty}_{ts}.txt"
                fpath = os.path.join(tempfile.gettempdir(), fname)
                with open(fpath, 'w', encoding='utf-8') as f:
                    for line in extract:
                        f.write(line + "\n")
            else:
                fname = f"{panel_type}_IDs_{qty}_{ts}.xlsx"
                fpath = os.path.join(tempfile.gettempdir(), fname)
                parsed_data = [parse_line_to_columns(l) for l in extract]
                pd.DataFrame(parsed_data,
                             columns=["UID", "Password", "Cookies"]
                             ).to_excel(fpath, index=False)
            panel.set_data("\n".join(remaining))
            self.silent_auto_save()
            caption = receipt or f"✅ Type: {type_label(panel_type)} | Qty: {qty} pcs | Format: {fmt.upper()}"
            try:
                limit = int(self.customer_db.get_setting("low_stock_limit", "20"))
                master = self.chat_id_entry.get().strip()
                if master and len(remaining) <= limit:
                    self._send_text_direct(token, master,
                                           f"⚠️ LOW STOCK: {type_label(panel_type)} has only {len(remaining)} IDs left!")
            except Exception:
                pass
            threading.Thread(target=self._send_file_thread,
                             args=(token, chat_id, fpath, caption),
                             daemon=True).start()
        except Exception as e:
            self.send_message(token, chat_id, f"❌ Error: {str(e)}")
            self._refund_last(chat_id, token)

    def _refund_last(self, chat_id, token):
        """Refund the most recent purchase of this user (used when delivery fails)."""
        try:
            if chat_id == self.chat_id_entry.get().strip():
                return
            with sqlite3.connect(self.customer_db.db_path) as conn:
                row = conn.execute("SELECT id, cost FROM user_purchases WHERE chat_id=? ORDER BY id DESC LIMIT 1",
                                   (chat_id,)).fetchone()
                if row:
                    conn.execute("DELETE FROM user_purchases WHERE id=?", (row[0],))
                    conn.execute("UPDATE bot_users SET balance = balance + ? WHERE chat_id=?", (row[1], chat_id))
                    conn.commit()
            self.send_message(token, chat_id, "↩️ Your balance has been refunded.")
            if row:
                self._notify_admin(token, f"↩️ Order #{row[0]} refunded (delivery failed)\nChat ID: {chat_id}")
        except Exception as e:
            logger.error(f"[refund] {e}")

    def _notify_admin(self, token, text, master=True):
        targets = []
        m = self.chat_id_entry.get().strip()
        if master and m:
            targets.append(m)
        log_id = self.customer_db.get_setting("log_chat_id", "").strip()
        if log_id and log_id not in targets:
            targets.append(log_id)
        for t in targets:
            threading.Thread(target=self._send_text_direct, args=(token, t, text), daemon=True).start()

    def _notify_payment(self, token, dep_id, amt, cur):
        try:
            dep = self.customer_db.get_deposit(dep_id)
            if not dep:
                return
            chat, uname, name = dep[1], dep[2] or "", dep[3] or ""
            user = self.customer_db.get_user(chat)
            bal = user[2] if user else 0
            sym = "৳" if cur == "BDT" else "$"
            self._notify_admin(token,
                               f"💰 Payment Received\n━━━━━━━━━━━━━━\n"
                               f"🆔 Deposit: #{dep_id}\n👤 {name}" + (f" (@{uname})" if uname else "") + "\n"
                               f"💬 Chat ID: {chat}\n💵 Amount: {amt:g} {cur}\n"
                               f"💳 New balance: {sym}{bal:.2f}\n"
                               f"🕒 {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                               master=False)
        except Exception as e:
            logger.error(f"[notify_payment] {e}")

    def _bot_stats(self, token, chat_id):
        stats, total, dep = self.customer_db.lifetime_stats()
        def line(name, d):
            return f"{name}: {d['orders']} orders | {d['qty']} pcs | {d['BDT']:.2f}৳ | ${d['USD']:.2f}"
        self.send_message(token, chat_id,
                          f"📊 Lifetime Stats\n━━━━━━━━━━━━━━\n"
                          f"📦 {line('61xx', stats['61xx'])}\n"
                          f"📦 {line('1000x', stats['1000x'])}\n"
                          f"📦 {line(NEWID_LABEL, stats['newid'])}\n━━━━━━━━━━━━━━\n"
                          f"🏆 {line('TOTAL', total)}\n"
                          f"💳 Deposits approved: {dep[0]} | {dep[1]:.2f}৳ | ${dep[2]:.2f}",
                          reply_markup=self.get_admin_keyboard())

    def _broadcast(self, token, admin_chat, text=None, photo=None, caption=None):
        ids = self.customer_db.all_user_ids()
        self._send_text_direct(token, admin_chat, f"📢 Sending to {len(ids)} users...")

        def _run():
            ok = fail = 0
            for uid in ids:
                try:
                    if photo:
                        cap = self._tr(uid, caption or "")
                        r = self.tg_session.post(f"https://api.telegram.org/bot{token}/sendPhoto",
                                                 data={"chat_id": uid, "photo": photo, "caption": cap},
                                                 timeout=20)
                        good = r.status_code == 200
                    else:
                        good = self._send_text_direct(token, uid, text, mirror=False)
                except Exception:
                    good = False
                ok += 1 if good else 0
                fail += 0 if good else 1
                time.sleep(0.05)
            self._send_text_direct(token, admin_chat, f"✅ Announcement finished.\nDelivered: {ok} | Failed: {fail}")
        threading.Thread(target=_run, daemon=True).start()

    def _send_backup(self, token, chat_id):
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            tmp = os.path.join(tempfile.gettempdir(), f"customers_backup_{ts}.db")
            src = sqlite3.connect(DB_FILE)
            dst = sqlite3.connect(tmp)
            src.backup(dst)
            dst.close()
            src.close()
            with open(tmp, "rb") as f:
                requests.post(f"https://api.telegram.org/bot{token}/sendDocument",
                              files={"document": f},
                              data={"chat_id": chat_id, "caption": f"🗄️ Database backup {ts}"},
                              timeout=60)
            os.remove(tmp)
            # stock lists (ID/password/cookies) - without token / chat id
            with _STOCK_LOCK:
                stock = {"panel_61xx": self.panel_61xx.get_data(),
                         "panel_1000x": self.panel_1000x.get_data(),
                         "panel_newid": self.panel_newid.get_data()}
            stmp = os.path.join(tempfile.gettempdir(), f"stock_backup_{ts}.json")
            with open(stmp, "w", encoding="utf-8") as f:
                json.dump(stock, f)
            with open(stmp, "rb") as f:
                requests.post(f"https://api.telegram.org/bot{token}/sendDocument",
                              files={"document": f},
                              data={"chat_id": chat_id, "caption": f"📦 Stock backup {ts}"},
                              timeout=60)
            os.remove(stmp)
        except Exception as e:
            logger.error(f"[backup] {e}")

    # ========================================================
    # AUTO DEPOSIT  (screenshot -> OCR -> checks -> approve)
    # Every check must pass, otherwise the deposit goes to the admin
    # for manual approval exactly like before.
    # ========================================================
    def _auto_enabled(self):
        try:
            return (self.customer_db.get_setting("auto_dep", "1") == "1"
                    and self._ocr_provider() is not None)
        except Exception:
            return False

    def _auto_status_text(self):
        on = self.customer_db.get_setting("auto_dep", "1") == "1"
        prov = self._ocr_provider()
        if on and prov:
            mdl = (os.environ.get("GEMINI_MODEL", "").strip() or "gemini-2.5-flash") if prov == "gemini" else OCR_MODEL
            return f"🤖 Auto deposit: ON ({prov}, model {mdl})"
        if on and not prov:
            return ("🤖 Auto deposit: NOT ACTIVE - GEMINI_API_KEY (or ANTHROPIC_API_KEY) is missing "
                    "(every deposit needs manual approval)")
        return "🤖 Auto deposit: OFF (switched off in Control Panel)"

    def _download_tg_file(self, token, file_id):
        res = self.tg_session.get(
            f"https://api.telegram.org/bot{token}/getFile?file_id={file_id}", timeout=15).json()
        if not res.get("ok"):
            raise RuntimeError(f"getFile failed: {res.get('description')}")
        return self.tg_session.get(
            f"https://api.telegram.org/file/bot{token}/{res['result']['file_path']}", timeout=60).content

    def _ocr_provider(self):
        """'gemini' | 'anthropic' | None  (OCR_PROVIDER env can force one)."""
        gem = bool(os.environ.get("GEMINI_API_KEY", "").strip())
        ant = bool(os.environ.get("ANTHROPIC_API_KEY", "").strip())
        want = os.environ.get("OCR_PROVIDER", "").strip().lower()
        if want == "gemini" and gem:
            return "gemini"
        if want == "anthropic" and ant:
            return "anthropic"
        return "anthropic" if ant else ("gemini" if gem else None)

    def _ocr_prompt(self):
        now_bd = datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M")
        return (
            "You read mobile-money payment screenshots from Bangladesh (bKash, Nagad, Rocket, Upay...). "
            f"Current Bangladesh date/time: {now_bd}. Dates on these screenshots are usually day-first "
            "(10/11/26 = 10 Nov 2026). Times are Bangladesh local time.\n"
            "The screenshot may be in ENGLISH or BENGALI (বাংলা), e.g. 'টাকা/পরিমাণ/খরচ/সর্বমোট/ট্রানজেকশন আইডি/"
            "সময়/চার্জ/সেন্ড মানি/সফল'. Read both. ALWAYS answer with English (0-9) digits and 24-hour time: convert "
            "Bengali digits ০১২৩৪৫৬৭৮৯ to 0-9, and Bengali time words (ভোর/সকাল = AM, দুপুর/বিকাল/সন্ধ্যা/রাত = PM) "
            "and Bengali month names (জানুয়ারি, ফেব্রুয়ারি, মার্চ, এপ্রিল, মে, জুন, জুলাই, আগস্ট, সেপ্টেম্বর, অক্টোবর, "
            "নভেম্বর/নভে, ডিসেম্বর) into the JSON format below. The Transaction ID stays exactly as printed "
            "(Latin letters/digits).\n"
            "Reply with ONE JSON object only, no other text, with keys:\n"
            '{"is_payment_screenshot": bool, "status_success": bool, '
            '"provider": "bkash|nagad|rocket|upay|other", '
            '"amount": number|null, "charge": number|null, "total": number|null, '
            '"trx_id": string|null, "recipient_number": string|null, '
            '"datetime": "YYYY-MM-DD HH:MM" (24-hour)|null, '
            '"confidence": "high|medium|low", "looks_edited": bool, "notes": string}\n'
            "Rules: 'amount' is the money the RECIPIENT receives, WITHOUT the fee. If the screen shows "
            "Amount, Charge and Total (e.g. 245 + 5 = 250, or Bengali পরিমাণ/খরচ/সর্বমোট 100 + 0.85 = 100.85), "
            "amount=245, charge=5, total=250. Never put the "
            "Total into 'amount'. If only one figure is shown, put it in 'amount'. "
            "'trx_id' is the Transaction ID exactly as printed (letters+digits, no spaces). "
            "'recipient_number' is the number money was sent to, digits only. "
            "Set looks_edited=true only if you see clear signs of editing (mismatched fonts, "
            "misaligned text, pasted regions). Set confidence=low if any key field is blurry or cropped. "
            "Use null for anything not visible - never guess.")

    @staticmethod
    def _parse_json_reply(txt):
        txt = (txt or "").strip()
        if txt.startswith("```"):
            txt = txt.strip("`")
        i, j = txt.find("{"), txt.rfind("}")
        if i < 0 or j < 0:
            raise ValueError(f"no JSON in OCR reply: {txt[:120]}")
        return json.loads(txt[i:j + 1])

    def _ocr_extract(self, img_bytes):
        """Read the payment screenshot with Gemini or Claude (vision) -> dict."""
        prov = self._ocr_provider()
        if not prov:
            raise RuntimeError("no OCR key (GEMINI_API_KEY / ANTHROPIC_API_KEY)")
        prompt = self._ocr_prompt()
        b64 = base64.b64encode(img_bytes).decode()
        if prov == "gemini":
            key = os.environ["GEMINI_API_KEY"].strip()
            first = os.environ.get("GEMINI_MODEL", "").strip() or "gemini-2.5-flash"
            models = [first] + [m_ for m_ in ("gemini-3.5-flash", "gemini-2.5-flash") if m_ != first]
            last = ""
            for model in models:
                r = requests.post(
                    f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                    headers={"x-goog-api-key": key, "Content-Type": "application/json"},
                    json={"contents": [{"parts": [
                              {"inline_data": {"mime_type": "image/jpeg", "data": b64}},
                              {"text": prompt}]}],
                          "generationConfig": {"temperature": 0, "responseMimeType": "application/json"}},
                    timeout=60)
                if r.status_code == 200:
                    parts = (r.json().get("candidates") or [{}])[0].get("content", {}).get("parts", [])
                    return self._parse_json_reply("".join(p_.get("text", "") for p_ in parts))
                last = f"Gemini {model} -> {r.status_code}: {r.text[:200]}"
                if r.status_code not in (400, 404):      # only try the next model on "unknown model" style errors
                    break
            raise RuntimeError(last)
        key = os.environ["ANTHROPIC_API_KEY"].strip()
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": key, "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json={"model": OCR_MODEL, "max_tokens": 500,
                  "messages": [{"role": "user", "content": [
                      {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                      {"type": "text", "text": prompt}]}]},
            timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"Claude API {r.status_code}: {r.text[:200]}")
        return self._parse_json_reply("".join(
            b_.get("text", "") for b_ in r.json().get("content", []) if b_.get("type") == "text"))

    @staticmethod
    def _bn2en(v):
        """Bengali digits (০-৯) -> 0-9, Bengali comma/decimal marks -> ASCII."""
        return str(v if v is not None else "").translate(
            str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")).replace("٫", ".")

    def _to_num(self, v):
        try:
            return float(self._bn2en(v).replace(",", "").replace("৳", "").replace("Tk", "").strip())
        except Exception:
            return None

    def _auto_decide(self, d, caption, upload_dt, dep_id, chat_id):
        """Returns (status, reason, info). status: ok | manual | dup | dup_pending"""
        db = self.customer_db
        info = {}
        if not isinstance(d, dict):
            return "manual", "OCR result unreadable", info
        if not d.get("is_payment_screenshot"):
            return "manual", "not a payment screenshot", info
        if d.get("status_success") is False:
            return "manual", "payment not shown as successful", info
        trx = re.sub(r"[^A-Za-z0-9]", "", self._bn2en(d.get("trx_id"))).upper()
        if not re.fullmatch(r"[A-Z0-9]{8,12}", trx):
            return "manual", "Transaction ID not readable", info
        info["trx"] = trx
        amt, chg, tot = (self._to_num(d.get("amount")), self._to_num(d.get("charge")),
                         self._to_num(d.get("total")))
        if amt is not None and chg and tot and abs(amt - tot) < 0.01:
            amt = round(tot - chg, 2)             # OCR picked "Total" -> remove the fee
        if amt is None and tot is not None:
            amt = round(tot - (chg or 0), 2)
        if amt is None or amt <= 0:
            return "manual", "amount not readable", info
        if chg and tot and abs((amt + chg) - tot) > 0.5:
            return "manual", f"Amount {amt:g} + Charge {chg:g} does not equal Total {tot:g}", info
        info.update(amt=amt, chg=chg, tot=tot)

        # duplicate TrxID (atomic claim)
        ok, st, other = db.reserve_trx(trx, dep_id, chat_id, amt)
        if not ok:
            kind = "dup_pending" if st == "pending" else "dup"
            return kind, f"TrxID {trx} already {st} (deposit #{other})", info

        if str(d.get("confidence", "")).lower() == "low":
            return "manual", "screenshot not clear enough (low confidence)", info
        if d.get("looks_edited"):
            return "manual", "screenshot looks edited", info
        rec = re.sub(r"\D", "", self._bn2en(d.get("recipient_number")))[-11:]
        allowed = {re.sub(r"\D", "", x)[-11:]
                   for x in db.get_setting("auto_numbers", DEFAULT_AUTO_NUMBERS).split(",") if x.strip()}
        if rec not in allowed:
            return "manual", f"receiver number {rec or 'unknown'} is not one of your numbers", info
        shot = None
        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
            try:
                shot = datetime.strptime(self._bn2en(d.get("datetime")).strip(), fmt).replace(tzinfo=BD_TZ)
                break
            except Exception:
                pass
        if not shot:
            return "manual", "payment time not readable", info
        try:
            window = float(db.get_setting("auto_window_min", "5"))
        except ValueError:
            window = 5.0
        delta = (upload_dt - shot).total_seconds() / 60.0
        info["shot"] = shot.strftime("%Y-%m-%d %H:%M")
        info["delta"] = delta
        if delta > window:
            return "manual", f"screenshot payment time is {delta:.0f} min before upload (limit {window:g})", info
        if delta < -2:
            return "manual", "payment time is in the future", info
        m = re.search(r"\d+(?:[.,]\d+)?", self._bn2en(caption))
        if m:
            user_amt = float(m.group(0).replace(",", "."))
            if abs(user_amt - amt) > 0.01:
                return "manual", f"user wrote {user_amt:g} but screenshot amount is {amt:g}", info
        try:
            mx = float(db.get_setting("auto_max_bdt", "3000"))
        except ValueError:
            mx = 3000.0
        if amt > mx:
            return "manual", f"amount {amt:g} is above the auto limit {mx:g}", info
        return "ok", "", info

    def _auto_deposit_job(self, token, dep_id, chat_id, name, uname, file_id, caption, upload_ts):
        master_id = self.chat_id_entry.get().strip()
        db = self.customer_db

        def to_manual(reason):
            logger.info(f"[AUTO] #{dep_id} -> manual: {reason}")
            try:
                db.set_deposit_note(dep_id, "Auto-check: " + reason)
            except Exception:
                pass
            self.send_message(token, chat_id,
                              "⏳ Could not verify automatically. Sent to admin for manual review.\n"
                              "You will be notified shortly.")
            self._send_admin_deposit_card(token, master_id, dep_id, name, uname, chat_id, file_id,
                                          extra=reason)
        try:
            d = self._ocr_extract(self._download_tg_file(token, file_id))
            upload_dt = datetime.fromtimestamp(int(upload_ts), BD_TZ)
            status, reason, info = self._auto_decide(d, caption, upload_dt, dep_id, chat_id)
        except Exception as e:
            logger.error(f"[AUTO] #{dep_id} error: {e}")
            logger.error(traceback.format_exc())
            to_manual(f"OCR/API error: {str(e)[:120]}")
            return

        if status == "dup":
            db.reject_deposit(dep_id, "Duplicate TrxID")
            self.send_message(token, chat_id,
                              "❌ This Transaction ID was already used. Duplicate deposits are not allowed.")
            self._send_text_direct(token, master_id,
                                   f"🚫 DUPLICATE TrxID blocked\nDeposit #{dep_id} | {name}"
                                   + (f" (@{uname})" if uname else "") + f" | {chat_id}\n{reason}")
            return
        if status == "dup_pending":
            db.reject_deposit(dep_id, "Duplicate (pending) TrxID")
            self.send_message(token, chat_id,
                              "⏳ This Transaction ID is already submitted and waiting for admin review.")
            return
        if status != "ok":
            to_manual(reason)
            return

        amt, trx = info["amt"], info["trx"]
        if not db.approve_deposit(dep_id, amt, "BDT"):
            to_manual("deposit record not found")
            return
        db.set_deposit_note(dep_id, f"AUTO-APPROVED trx={trx} shot={info['shot']} delta={info['delta']:.1f}min")
        logger.info(f"[AUTO] #{dep_id} approved {amt:g} BDT trx={trx}")
        self._notify_payment(token, dep_id, amt, "BDT")
        self.send_message(token, chat_id,
                          f"✅ Deposit approved automatically!\n"
                          f"+{amt:g} BDT has been added to your balance.\n"
                          f"Transaction ID: {trx}")
        u = db.get_user(chat_id)
        cap = (f"🤖 AUTO-APPROVED\n━━━━━━━━━━━━━━━━━━━━━\n"
               f"🆔 Deposit #{dep_id} | 💵 {amt:g} BDT\n"
               f"👤 {name}" + (f" (@{uname})" if uname else "") + f"\n💬 Chat ID: {chat_id}\n"
               f"🔖 TrxID: {trx}\n🕒 Payment time: {info['shot']} ({info['delta']:.1f} min before upload)\n"
               f"💳 New balance: ৳{(u[2] if u else 0):.2f}\n━━━━━━━━━━━━━━━━━━━━━\n"
               f"Wrong? Tap Undo - balance is taken back and TrxID stays blocked.")
        self.send_photo(token, master_id, file_id, cap,
                        {"inline_keyboard": [[{"text": "❌ Undo (take back)",
                                               "callback_data": f"dep_u_{dep_id}"}]]})

    def _send_admin_deposit_card(self, token, master_id, dep_id, name, uname, chat_id, file_id, extra=""):
        user_info = self.customer_db.get_user(chat_id)
        cur_bal = user_info[2] if user_info else 0
        cur = user_info[3] if user_info else "BDT"
        sym = "৳" if cur == "BDT" else "$"
        admin_caption = (
            f"💰 DEPOSIT REQUEST\n"
            f"━━━━━━━━━━━━━━━━━━━━━\n"
            f"🆔 Deposit ID: #{dep_id}\n"
            f"👤 Name: {name}\n"
            f"🏷️ Username: @{uname if uname else 'N/A'}\n"
            f"💬 Chat ID: {chat_id}\n"
            f"💳 Current Balance: {sym}{cur_bal:.2f} {cur}\n"
            f"🕒 Time (BD): {datetime.now(BD_TZ).strftime('%Y-%m-%d %H:%M:%S')}\n"
            + (f"🤖 Auto-check failed: {extra}\n" if extra else "") +
            f"━━━━━━━━━━━━━━━━━━━━━\n"
            f"Use buttons below, or:\n"
            f"/approve {dep_id} <amt> <BDT|USD>\n"
            f"/reject {dep_id}"
        )
        admin_kb = {"inline_keyboard": [
            [{"text": "✅ 100৳", "callback_data": f"dep_a_{dep_id}_100_BDT"},
             {"text": "✅ 500৳", "callback_data": f"dep_a_{dep_id}_500_BDT"},
             {"text": "✅ 1000৳", "callback_data": f"dep_a_{dep_id}_1000_BDT"}],
            [{"text": "💵 $5", "callback_data": f"dep_a_{dep_id}_5_USD"},
             {"text": "💵 $10", "callback_data": f"dep_a_{dep_id}_10_USD"},
             {"text": "💵 $50", "callback_data": f"dep_a_{dep_id}_50_USD"}],
            [{"text": "✏️ Custom Amount / Currency", "callback_data": f"dep_custom_{dep_id}"}],
            [{"text": "❌ Reject", "callback_data": f"dep_r_{dep_id}"}]
        ]}
        self.send_photo(token, master_id, file_id, admin_caption, admin_kb)
        self._notify_admin(token,
                           f"💰 New Deposit Request\n━━━━━━━━━━━━━━\n🆔 Deposit: #{dep_id}\n"
                           f"👤 {name}" + (f" (@{uname})" if uname else "") + f"\n💬 Chat ID: {chat_id}\n"
                           f"🕒 {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                           master=False)
        logger.info(f"[DEPOSIT ALERT] Sent to master_id={master_id} for #{dep_id}")

    def _restore_from_file(self, token, chat_id, file_id, file_name):
        """Admin sent a file in /restore mode: customers.db or a stock json."""
        try:
            res = self.tg_session.get(
                f"https://api.telegram.org/bot{token}/getFile?file_id={file_id}", timeout=15).json()
            if not res.get("ok"):
                self._send_text_direct(token, chat_id,
                                       f"❌ Cannot download file: {res.get('description')}\n"
                                       "(Telegram bots can download files up to 20 MB only.)")
                return
            content = self.tg_session.get(
                f"https://api.telegram.org/file/bot{token}/{res['result']['file_path']}",
                timeout=120).content
            low = (file_name or "").lower()
            if low.endswith((".db", ".sqlite", ".sqlite3")) or content[:15] == b"SQLite format 3":
                info = self._restore_db(content)
                self._send_text_direct(
                    token, chat_id,
                    "✅ Database restored!\n━━━━━━━━━━━━━━\n"
                    f"👥 Users: {info['users']}\n💳 Deposits: {info['deposits']}\n"
                    f"🛒 Purchases: {info['purchases']}\n"
                    f"💰 Total balances: {info['bdt']:.2f}৳ | ${info['usd']:.2f}\n━━━━━━━━━━━━━━\n"
                    "Now send the stock json too (if you have it), then /done.",
                    reply_markup=self.get_admin_keyboard())
            elif low.endswith(".json"):
                info = self._restore_stock(content)
                self._send_text_direct(
                    token, chat_id,
                    "✅ Stock restored!\n━━━━━━━━━━━━━━\n"
                    f"61xx: {info['61xx']}\n1000x: {info['1000x']}\n"
                    f"{NEWID_LABEL}: {info['newid']}\n━━━━━━━━━━━━━━\nSend /done when finished.",
                    reply_markup=self.get_admin_keyboard())
            else:
                self._send_text_direct(token, chat_id,
                                       "❌ Unsupported file. Send customers.db or a .json stock file.")
        except Exception as e:
            logger.error(f"[restore] {e}")
            logger.error(traceback.format_exc())
            self._send_text_direct(token, chat_id, f"❌ Restore failed: {e}")

    def _restore_db(self, content):
        """Validate an uploaded customers.db and copy it over the live database."""
        tmp = os.path.join(tempfile.gettempdir(), "restore_incoming.db")
        with open(tmp, "wb") as f:
            f.write(content)
        src = sqlite3.connect(tmp)
        try:
            ok = src.execute("PRAGMA integrity_check").fetchone()[0]
            if ok != "ok":
                raise ValueError(f"file is damaged ({ok})")
            tables = {r[0] for r in src.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if "bot_users" not in tables:
                raise ValueError("no user data inside this file. If you copied customers.db while the "
                                 "desktop app was running, close the app first and copy again "
                                 "(or use the app's Backup button).")
            cnt = lambda t: (src.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                             if t in tables else 0)
            with self.customer_db._lock:
                # keep a safety copy of whatever is there now
                cur = sqlite3.connect(DB_FILE)
                keep = sqlite3.connect(DB_FILE + ".before_restore")
                cur.backup(keep)
                keep.close()
                cur.close()
                live = sqlite3.connect(DB_FILE)
                src.backup(live)
                live.close()
        finally:
            src.close()
            try:
                os.remove(tmp)
            except Exception:
                pass
        self.customer_db._init_db()             # add any missing newer columns / settings
        load_button_names(self.customer_db)
        self.__dict__.pop("_lang_cache", None)
        with sqlite3.connect(DB_FILE) as conn:
            bal = dict(conn.execute("SELECT COALESCE(currency,'BDT'), COALESCE(SUM(balance),0) "
                                    "FROM bot_users GROUP BY COALESCE(currency,'BDT')").fetchall())
            info = {"users": conn.execute("SELECT COUNT(*) FROM bot_users").fetchone()[0],
                    "deposits": conn.execute("SELECT COUNT(*) FROM deposits").fetchone()[0],
                    "purchases": conn.execute("SELECT COUNT(*) FROM user_purchases").fetchone()[0]}
        info["bdt"] = float(bal.get("BDT", 0))
        info["usd"] = float(bal.get("USD", 0))
        return info

    def _restore_stock(self, content):
        d = json.loads(content.decode("utf-8", errors="ignore"))
        keys = {"panel_61xx": self.panel_61xx, "panel_1000x": self.panel_1000x,
                "panel_newid": self.panel_newid}
        if not any(k in d for k in keys):
            raise ValueError("no stock lists found in this json")
        with _STOCK_LOCK:
            for k, panel in keys.items():
                if k in d:
                    panel.set_data(d.get(k, ""))
        self.silent_auto_save()
        return {k: self._cp_count(k) for k in self._CP_TYPES}

    def _maybe_auto_backup(self, token):
        try:
            master = self.chat_id_entry.get().strip()
            last = float(self.customer_db.get_setting("last_backup", "0"))
            if master and time.time() - last > 86400:
                self.customer_db.set_setting("last_backup", str(time.time()))
                threading.Thread(target=self._send_backup, args=(token, master), daemon=True).start()
        except Exception:
            pass

    def _bot_report(self, token, chat_id, days):
        r = self.customer_db.sales_report(days)
        top = "\n".join(f"  {i+1}. @{u or 'N/A'} ({c}) — {q} pcs" for i, (u, c, q) in enumerate(r["top"])) or "  -"
        self.send_message(token, chat_id,
                          f"📑 Report (last {days} day{'s' if days > 1 else ''})\n\n"
                          f"🛒 Orders: {r['orders']}\n📦 IDs sold: {r['qty']}\n"
                          f"💵 Sales: {r['bdt']:.0f}৳ | ${r['usd']:.2f}\n"
                          f"💳 Deposits approved: {r['dep_bdt']:.0f}৳ | ${r['dep_usd']:.2f}\n\n"
                          f"🏆 Top buyers:\n{top}",
                          reply_markup=self.get_admin_keyboard())

    def _send_file_thread(self, token, chat_id, file_path, caption):
        try:
            kb = (self.get_admin_keyboard() if chat_id == self.chat_id_entry.get().strip()
                  else self.get_user_keyboard())
            with open(file_path, "rb") as f:
                requests.post(f"https://api.telegram.org/bot{token}/sendDocument",
                              files={"document": f},
                              data={"chat_id": chat_id, "caption": self._tr(chat_id, caption),
                                    "reply_markup": json.dumps(self._trm(chat_id, kb))},
                              timeout=30)
            # v31: keep a copy of the delivered file inside the user's topic
            self._mirror_out_file(token, chat_id, file_path, caption)
        except:
            pass

    # -------- Send helpers --------
    def get_region_keyboard(self):
        return {"keyboard": [[{"text": REGION_BD_BTN}], [{"text": REGION_OTHER_BTN}]],
                "resize_keyboard": True, "one_time_keyboard": True}

    def get_lang_keyboard(self):
        return {"keyboard": [[{"text": LANG_EN_BTN}, {"text": LANG_BN_BTN}]],
                "resize_keyboard": True, "one_time_keyboard": True}

    def _is_master(self, chat_id):
        return str(chat_id) == self.chat_id_entry.get().strip()

    def _get_lang(self, chat_id):
        if self._is_master(chat_id):
            return self.customer_db.get_setting("admin_lang", "en")
        cache = self.__dict__.setdefault("_lang_cache", {})
        if chat_id not in cache:
            cache[chat_id] = self.customer_db.get_lang(str(chat_id)) or "en"
        return cache[chat_id]

    def _tr(self, chat_id, text):
        if self._get_lang(chat_id) != "bn":
            return text
        if self._is_master(chat_id):
            # admin: only shareable texts are translated, reports/notifications stay as written
            return tr_bn(text) if text in (POLICY_TEXT, PAYMENT_INFO, PAYMENT_INFO_USD) else text
        return tr_bn(text)

    def _trm(self, chat_id, markup):
        markup = style_markup(markup)
        if self._get_lang(chat_id) != "bn":
            return markup
        return tr_bn_markup(markup, keys_only=self._is_master(chat_id))

    def _send_exact(self, token, chat_id, text):
        """Send text exactly as given (no auto-translation), Markdown for tap-to-copy."""
        self._mirror_out_text(token, chat_id, text)
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        def _go():
            try:
                r = self.tg_session.post(url, data={"chat_id": chat_id, "text": text,
                                                    "parse_mode": "Markdown"}, timeout=10)
                if r.status_code != 200:
                    self.tg_session.post(url, data={"chat_id": chat_id, "text": text}, timeout=10)
            except Exception:
                pass
        threading.Thread(target=_go, daemon=True).start()

    def send_message(self, token, chat_id, text, reply_markup=None):
        self._mirror_out_text(token, chat_id, text)
        text = self._tr(chat_id, text)
        reply_markup = self._trm(chat_id, reply_markup)
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        payload = {"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}
        if reply_markup:
            payload["reply_markup"] = json.dumps(reply_markup)

        def _send():
            try:
                res = self.tg_session.post(url, data=payload, timeout=10)
                if res.status_code != 200:
                    payload.pop("parse_mode", None)
                    self.tg_session.post(url, data=payload, timeout=10)
            except:
                pass
        threading.Thread(target=_send, daemon=True).start()

    def send_photo(self, token, chat_id, file_id, caption, reply_markup=None):
        self._mirror_out_photo(token, chat_id, file_id, caption)
        url = f"https://api.telegram.org/bot{token}/sendPhoto"

        def _send():
            try:
                payload = {"chat_id": chat_id, "photo": file_id, "caption": caption}
                if reply_markup:
                    payload["reply_markup"] = json.dumps(style_markup(reply_markup))
                res = self.tg_session.post(url, data=payload, timeout=20)
                if res.status_code == 200:
                    logger.info(f"[send_photo OK] chat={chat_id}")
                    return
                err_body = res.text[:300]
                logger.error(f"[send_photo FAIL] status={res.status_code} body={err_body}")
                self._send_text_direct(token, chat_id,
                                       f"📸 (Photo failed to attach)\n\n{caption}",
                                       reply_markup)
            except Exception as e:
                logger.error(f"[send_photo EXC] {e}")
                logger.error(traceback.format_exc())
                try:
                    self._send_text_direct(token, chat_id,
                                           f"📸 (Photo delivery error)\n\n{caption}",
                                           reply_markup)
                except:
                    pass
        threading.Thread(target=_send, daemon=True).start()

    def _send_text_direct(self, token, chat_id, text, reply_markup=None, mirror=True):
        try:
            if mirror:
                self._mirror_out_text(token, chat_id, text)
            text = self._tr(chat_id, text)
            reply_markup = self._trm(chat_id, reply_markup)
            url = f"https://api.telegram.org/bot{token}/sendMessage"
            payload = {"chat_id": chat_id, "text": text}
            if reply_markup:
                payload["reply_markup"] = json.dumps(reply_markup)
            res = self.tg_session.post(url, data=payload, timeout=15)
            if res.status_code == 200:
                logger.info(f"[send_text_direct OK] chat={chat_id}")
                return True
            logger.error(f"[send_text_direct FAIL] status={res.status_code} body={res.text[:300]}")
            return False
        except Exception as e:
            logger.error(f"[send_text_direct EXC] {e}")
            return False

    # ============================================================
    # v31: PER-USER TOPIC CHAT  (Telegram forum group)
    #  - every user gets their own topic (chat box) in the admin group
    #  - user messages / bot replies / delivered files / orders are shown there
    #  - anything the admin types inside a topic is sent ONLY to that user
    # ============================================================
    def _topics_gid(self):
        try:
            return self.customer_db.get_setting("topics_group_id", "").strip()
        except Exception:
            return ""

    def _tg_call(self, token, method, data=None, files=None, timeout=20):
        try:
            r = self.tg_session.post(f"https://api.telegram.org/bot{token}/{method}",
                                     data=data, files=files, timeout=timeout)
            return r.json()
        except Exception as e:
            logger.error(f"[tg_call {method}] {e}")
            return {"ok": False, "description": str(e)}

    def _group_reply(self, token, gid, tid, text):
        data = {"chat_id": gid, "text": text[:4000]}
        if tid:
            data["message_thread_id"] = tid
        return self._tg_call(token, "sendMessage", data)

    def _user_card_text(self, chat_id):
        u = self.customer_db.get_user(chat_id)
        if not u:
            return f"💬 Chat ID: {chat_id}"
        cur = u[3] or "BDT"
        sym = "৳" if cur == "BDT" else "$"
        purchases = self.customer_db.get_user_purchases(chat_id)
        total_qty = sum(p[1] for p in purchases)
        lines = [
            "👤 USER CHAT BOX",
            "━━━━━━━━━━━━━━━━━━━━━",
            f"Name: {u[1]}",
            f"Username: @{u[4] if len(u) > 4 and u[4] else 'N/A'}",
            f"Chat ID: {chat_id}",
            f"Balance: {sym}{u[2]:.2f} {cur}",
            f"Orders: {len(purchases)} | IDs bought: {total_qty}",
            "━━━━━━━━━━━━━━━━━━━━━",
            "✍️ Type here -> goes to this user only.",
            "/info -> full order history",
        ]
        return "\n".join(lines)

    def _ensure_topic(self, token, chat_id, force_new=False, create=True):
        gid = self._topics_gid()
        if not gid:
            return None
        chat_id = str(chat_id)
        with _TOPIC_LOCK:
            if not force_new:
                t = self.customer_db.get_topic(chat_id)
                if t:
                    return t
            if not create:
                return None
            u = self.customer_db.get_user(chat_id)
            name = (u[1] if u else "") or "User"
            uname = (u[4] if u and len(u) > 4 and u[4] else "")
            title = f"{name}" + (f" @{uname}" if uname else "") + f" | {chat_id}"
            res = self._tg_call(token, "createForumTopic",
                                {"chat_id": gid, "name": title[:120]})
            if not res.get("ok"):
                logger.error(f"[topic create FAIL] {res.get('description')}")
                return None
            tid = res["result"]["message_thread_id"]
            self.customer_db.set_topic(chat_id, tid)
            self._tg_call(token, "sendMessage",
                          {"chat_id": gid, "message_thread_id": tid,
                           "text": self._user_card_text(chat_id)})
            return tid

    def _topic_send(self, token, chat_id, method, data, file_path=None,
                    file_field="document", create=True):
        """Send something into the user's topic (re-creates the topic if it was deleted)."""
        gid = self._topics_gid()
        chat_id = str(chat_id)
        if not gid or chat_id == gid:
            return None
        res = None
        for attempt in (0, 1, 2):
            tid = self._ensure_topic(token, chat_id, force_new=(attempt == 1), create=create)
            if not tid:
                return None
            d = dict(data)
            d["chat_id"] = gid
            d["message_thread_id"] = tid
            if file_path:
                try:
                    with open(file_path, "rb") as f:
                        res = self._tg_call(token, method, d, files={file_field: f}, timeout=60)
                except Exception as e:
                    logger.error(f"[topic_send file] {e}")
                    return None
            else:
                res = self._tg_call(token, method, d)
            if res.get("ok"):
                return res
            desc = (res.get("description") or "").lower()
            if res.get("error_code") == 429:
                time.sleep(int((res.get("parameters") or {}).get("retry_after", 3)) + 1)
                continue
            if "thread" in desc or "topic" in desc:
                # topic was deleted by admin -> forget it
                self.customer_db.del_topic(chat_id)
                if not create:
                    return None
                continue
            logger.error(f"[topic_send {method}] {res.get('description')}")
            return res
        return res

    def _mirror_skip(self, chat_id):
        """True when this chat must NOT be mirrored (admin, log chat, the group itself)."""
        gid = self._topics_gid()
        if not gid:
            return True
        chat_id = str(chat_id)
        if chat_id == gid or self._is_master(chat_id):
            return True
        if chat_id == self.customer_db.get_setting("log_chat_id", "").strip():
            return True
        return False

    def _mirror_in(self, token, msg):
        """User -> topic (exact copy of what the user sent: text, photo, file, voice...)."""
        try:
            chat_id = str(msg["chat"]["id"])
            if self._mirror_skip(chat_id):
                return
            mid = msg["message_id"]
            mirror_submit(lambda: self._topic_send(
                token, chat_id, "copyMessage",
                {"from_chat_id": chat_id, "message_id": mid}, create=True))
        except Exception as e:
            logger.error(f"[mirror_in] {e}")

    def _mirror_out_text(self, token, chat_id, text, label="🤖 Bot"):
        try:
            if not text or self._mirror_skip(chat_id):
                return
            chat_id = str(chat_id)
            body = f"{label}:\n{text}"[:4000]
            mirror_submit(lambda: self._topic_send(
                token, chat_id, "sendMessage", {"text": body}, create=False))
        except Exception as e:
            logger.error(f"[mirror_out_text] {e}")

    def _mirror_out_photo(self, token, chat_id, file_id, caption):
        try:
            if self._mirror_skip(chat_id):
                return
            chat_id = str(chat_id)
            cap = f"🤖 Bot:\n{caption or ''}"[:1000]
            mirror_submit(lambda: self._topic_send(
                token, chat_id, "sendPhoto", {"photo": file_id, "caption": cap}, create=False))
        except Exception as e:
            logger.error(f"[mirror_out_photo] {e}")

    def _mirror_out_file(self, token, chat_id, file_path, caption):
        """Delivered order file -> topic. The temp file is deleted by the worker after upload."""
        try:
            if self._mirror_skip(chat_id) or not os.path.exists(file_path):
                if os.path.exists(file_path):
                    os.remove(file_path)
                return
            chat_id = str(chat_id)
            cap = f"📦 FILE DELIVERED TO USER\n{caption or ''}"[:1000]

            def _job():
                try:
                    self._topic_send(token, chat_id, "sendDocument", {"caption": cap},
                                     file_path=file_path, create=False)
                finally:
                    try:
                        os.remove(file_path)
                    except Exception:
                        pass
            mirror_submit(_job)
        except Exception as e:
            logger.error(f"[mirror_out_file] {e}")
            try:
                os.remove(file_path)
            except Exception:
                pass

    def _handle_group_message(self, msg, token):
        """Messages typed inside the Topics group by the admin."""
        try:
            gid = str(msg["chat"]["id"])
            frm = msg.get("from") or {}
            if frm.get("is_bot"):
                return
            master_id = self.chat_id_entry.get().strip()
            if str(frm.get("id", "")) != master_id:
                return                                    # only the master admin can control it
            tid = msg.get("message_thread_id")
            text = (msg.get("text") or "").strip()
            cmd = text.split()[0].split("@")[0].lower() if text.startswith("/") else ""

            # ----- link / unlink the group -----
            if cmd == "/setgroup":
                info = self._tg_call(token, "getChat", {"chat_id": gid})
                if not (info.get("ok") and info["result"].get("is_forum")):
                    self._group_reply(token, gid, tid,
                        "❌ This group must be a Supergroup with Topics enabled.\n"
                        "Group settings -> Topics -> turn ON. Then send /setgroup again.")
                    return
                test = self._tg_call(token, "createForumTopic", {"chat_id": gid, "name": "setup check"})
                if not test.get("ok"):
                    self._group_reply(token, gid, tid,
                        "❌ Bot cannot create topics.\n"
                        "Make the bot ADMIN and enable 'Manage Topics'. Then send /setgroup again.\n"
                        f"({test.get('description')})")
                    return
                self._tg_call(token, "deleteForumTopic",
                              {"chat_id": gid, "message_thread_id": test["result"]["message_thread_id"]})
                if self._topics_gid() != gid:
                    self.customer_db.clear_topics()
                self.customer_db.set_setting("topics_group_id", gid)
                self._group_reply(token, gid, tid,
                    "✅ User Chat Boxes are ON.\n"
                    "Each user will get their own topic here. Whatever you type inside a "
                    "topic is sent only to that user.\n\n"
                    "/info (inside a topic) -> user details + orders\n"
                    "/unsetgroup -> turn this system OFF")
                return
            if cmd == "/unsetgroup":
                self.customer_db.set_setting("topics_group_id", "")
                self._group_reply(token, gid, tid, "⛔ User Chat Boxes are OFF.")
                return

            if gid != self._topics_gid() or not tid:
                return                                    # General topic / other groups: ignore

            user_chat = self.customer_db.get_chat_by_topic(tid)
            if not user_chat:
                return

            # ----- /info inside a topic -----
            if cmd == "/info":
                lines = [self._user_card_text(user_chat), "", "🧾 ORDER HISTORY (latest 25)"]
                rows = self.customer_db.get_user_purchases(user_chat)[:25]
                if not rows:
                    lines.append("No orders yet.")
                for p_type, qty, unit, cost, cur, ts in rows:
                    sym = "৳" if cur == "BDT" else "$"
                    lines.append(f"• {str(ts)[:16]} | {p_type} x{qty} | {sym}{cost:g}")
                self._group_reply(token, gid, tid, "\n".join(lines))
                return
            if cmd:
                return                                    # other commands are not sent to users

            # ----- anything else -> copy to that user only -----
            content_keys = ("text", "photo", "document", "video", "audio", "voice",
                            "sticker", "animation", "video_note", "contact", "location")
            if not any(k in msg for k in content_keys):
                return
            res = self._tg_call(token, "copyMessage",
                                {"chat_id": user_chat, "from_chat_id": gid,
                                 "message_id": msg["message_id"]})
            if res.get("ok"):
                self._tg_call(token, "setMessageReaction",
                              {"chat_id": gid, "message_id": msg["message_id"],
                               "reaction": json.dumps([{"type": "emoji", "emoji": "👍"}])})
            else:
                self._group_reply(token, gid, tid,
                                  f"❌ Not delivered: {res.get('description')}")
        except Exception as e:
            logger.error(f"[group msg] {e}")
            logger.error(traceback.format_exc())

    # -------- Auto-save --------

    # ========================================================
    # TELEGRAM CONTROL PANEL (admin only)
    # Replaces the desktop Control Panel: prices, button names, balances,
    # discounts, shop switches, stock, stats - all from inside Telegram.
    # Open it with the "⚙️ Control Panel" button or /panel
    # ========================================================
    _CP_TYPES = ("1000x", "61xx", "newid")

    @staticmethod
    def _cp_sym(cur):
        return "৳" if cur == "BDT" else "$"

    def _cp_api(self, token, method, data):
        try:
            return self.tg_session.post(f"https://api.telegram.org/bot{token}/{method}",
                                        data=data, timeout=10)
        except Exception as e:
            logger.error(f"[panel api] {e}")
            return None

    def _cp_show(self, token, chat_id, message_id, text, rows):
        """Edit the panel message in place (or send a new one)."""
        markup = json.dumps({"inline_keyboard": rows})
        if message_id:
            r = self._cp_api(token, "editMessageText",
                             {"chat_id": chat_id, "message_id": message_id,
                              "text": text, "reply_markup": markup})
            if r is not None and (r.status_code == 200 or "not modified" in (r.text or "")):
                return
        self._cp_api(token, "sendMessage",
                     {"chat_id": chat_id, "text": text, "reply_markup": markup})

    @staticmethod
    def _cp_b(text, data):
        return {"text": text, "callback_data": data}

    def _cp_count(self, key):
        return len([l for l in self._get_panel(key).get_data().splitlines() if l.strip()])

    # ---------- screens ----------
    def _cp_main(self, token, chat_id, message_id=None):
        load_button_names(self.customer_db)
        shop = self.customer_db.get_setting("shop_open", "1") == "1"
        access = self.users_enabled
        b = self._cp_b
        stock = " | ".join(f"{button_name(k)}: {self._cp_count(k)}" for k in self._CP_TYPES)
        text = ("⚙️ CONTROL PANEL\n━━━━━━━━━━━━━━\n"
                f"🏪 Shop: {'OPEN' if shop else 'CLOSED'}\n"
                f"👥 User access: {'ON' if access else 'OFF'}\n"
                f"📦 Stock → {stock}\n━━━━━━━━━━━━━━\nChoose an option:")
        rows = [
            [b("💲 Prices", "cp:prices"), b("🏷️ Button Names", "cp:names")],
            [b("👥 Users & Balance", "cp:users"), b("🎟 Discounts", "cp:disc")],
            [b(f"🏪 Shop: {'OPEN ✅' if shop else 'CLOSED ⛔'}", "cp:shop"),
             b(f"👥 Access: {'ON ✅' if access else 'OFF ⛔'}", "cp:access")],
            [b("📦 Stock", "cp:stock"), b("📊 Stats", "cp:stats")],
            [b("💳 Pending Deposits", "cp:deps"), b("📈 Purchases", "cp:purch")],
            [b("🤖 Auto Deposit", "cp:auto"), b("🔔 Low-stock alert", "cp:ask:low")],
            [b("❌ Close", "close")],
        ]
        self._cp_show(token, chat_id, message_id, text, rows)

    def _cp_prices(self, token, chat_id, message_id):
        b = self._cp_b
        lines = ["💲 PRICES (per piece)\n━━━━━━━━━━━━━━"]
        rows = []
        for k in self._CP_TYPES:
            bdt = self.customer_db.get_price(k, "BDT")
            usd = self.customer_db.get_price(k, "USD")
            lines.append(f"{button_name(k)}:  ৳{bdt:g}  |  ${usd:g}")
            rows.append([b(f"✏️ {button_name(k)} ৳", f"cp:ask:price:{k}:BDT"),
                         b(f"✏️ {button_name(k)} $", f"cp:ask:price:{k}:USD")])
        rows.append([b("⬅️ Back", "cp:main")])
        lines.append("\nTap a button, then send the new price.")
        self._cp_show(token, chat_id, message_id, "\n".join(lines), rows)

    def _cp_names(self, token, chat_id, message_id):
        b = self._cp_b
        load_button_names(self.customer_db)
        lines = ["🏷️ BUY BUTTON NAMES\n━━━━━━━━━━━━━━"]
        rows = []
        for k in self._CP_TYPES:
            lines.append(f"{k} → {button_name(k)}")
            rows.append([b(f"✏️ Rename: {button_name(k)}", f"cp:ask:name:{k}")])
        rows.append([b("↩ Reset all to default", "cp:nmreset")])
        rows.append([b("⬅️ Back", "cp:main")])
        lines.append(f"\nMax {BUTTON_NAME_MAX} characters. Customers see the change instantly.")
        self._cp_show(token, chat_id, message_id, "\n".join(lines), rows)

    def _cp_users(self, token, chat_id, message_id):
        b = self._cp_b
        rows = [
            [b("📋 Users list", "cp:ulist"), b("➕ Add Balance", "cp:ask:bal")],
            [b("♻️ Reset Balance", "cp:ask:resetbal"), b("🗑 Delete User", "cp:ask:deluser")],
            [b("🚫 Ban", "cp:ask:ban"), b("✅ Unban", "cp:ask:unban")],
            [b("⬅️ Back", "cp:main")],
        ]
        self._cp_show(token, chat_id, message_id, "👥 USERS & BALANCE\n━━━━━━━━━━━━━━", rows)

    def _cp_disc(self, token, chat_id, message_id):
        b = self._cp_b
        d50 = self.customer_db.get_setting("disc_50", "0")
        d100 = self.customer_db.get_setting("disc_100", "0")
        low = self.customer_db.get_setting("low_stock_limit", "20")
        text = (f"🎟 DISCOUNTS\n━━━━━━━━━━━━━━\n50+ pcs: {d50}%\n100+ pcs: {d100}%\n"
                f"🔔 Low-stock alert at: {low}")
        rows = [[b("✏️ 50+ discount", "cp:ask:disc50"), b("✏️ 100+ discount", "cp:ask:disc100")],
                [b("🔔 Low-stock alert", "cp:ask:low")],
                [b("⬅️ Back", "cp:main")]]
        self._cp_show(token, chat_id, message_id, text, rows)

    def _cp_auto(self, token, chat_id, message_id):
        db, b = self.customer_db, self._cp_b
        on = db.get_setting("auto_dep", "1") == "1"
        prov = self._ocr_provider()
        win = db.get_setting("auto_window_min", "5")
        mx = db.get_setting("auto_max_bdt", "3000")
        nums = db.get_setting("auto_numbers", DEFAULT_AUTO_NUMBERS)
        text = ("🤖 AUTO DEPOSIT\n━━━━━━━━━━━━━━\n"
                f"Status: {'ON ✅' if on else 'OFF ⛔'}\n"
                f"OCR: {('✅ ' + prov) if prov else '❌ GEMINI_API_KEY / ANTHROPIC_API_KEY missing -> manual mode'}\n"
                f"⏱ Time window: {win} min\n"
                f"💰 Max auto amount: ৳{mx}\n"
                f"📱 Receiving numbers: {nums}\n━━━━━━━━━━━━━━\n"
                "BDT users only. If any check fails, the deposit comes to you for manual approval.")
        rows = [[b(f"🤖 Auto: {'ON ✅' if on else 'OFF ⛔'} (tap to switch)", "cp:autotog")],
                [b("⏱ Time window", "cp:ask:autowin"), b("💰 Max amount", "cp:ask:automax")],
                [b("📱 Receiving numbers", "cp:ask:autonum")],
                [b("⬅️ Back", "cp:main")]]
        self._cp_show(token, chat_id, message_id, text, rows)

    def _cp_stock(self, token, chat_id, message_id):
        lines = ["📦 STOCK\n━━━━━━━━━━━━━━"]
        for k in self._CP_TYPES:
            lines.append(f"{button_name(k)}: {self._cp_count(k)} IDs")
        self._cp_show(token, chat_id, message_id, "\n".join(lines),
                      [[self._cp_b("⬅️ Back", "cp:main")]])

    _CP_PROMPTS = {
        "bal": "➕ ADD BALANCE\nSend:  <chat_id> <amount> [BDT|USD]\nExample:  7538066844 500 BDT",
        "resetbal": "♻️ RESET BALANCE\nSend the user's chat_id:",
        "deluser": "🗑 DELETE USER\nSend the user's chat_id (you will be asked to confirm):",
        "ban": "🚫 BAN USER\nSend the user's chat_id:",
        "unban": "✅ UNBAN USER\nSend the user's chat_id:",
        "disc50": "🎟 Send the discount percent for 50+ pcs (0 = none). Example: 5",
        "disc100": "🎟 Send the discount percent for 100+ pcs (0 = none). Example: 10",
        "low": "🔔 Send the low-stock alert number. Example: 20",
        "autowin": "⏱ Send the allowed minutes between payment time and upload. Example: 5",
        "automax": "💰 Send the maximum amount (৳) that can be auto-approved. Above it = manual. Example: 3000",
        "autonum": "📱 Send your receiving numbers (bKash/Nagad/Rocket), comma separated.\nExample: 01954054330,01711111111",
    }

    def _cp_ask(self, token, chat_id, message_id, parts):
        """parts = ['price', key, cur] | ['name', key] | [what]"""
        what = parts[0]
        back = "cp:main"
        if what == "price":
            key, cur = parts[1], parts[2]
            self.user_states[chat_id] = f"cp_price:{key}:{cur}"
            cur_val = self.customer_db.get_price(key, cur)
            text = (f"💲 {button_name(key)} price in {cur}\nNow: {self._cp_sym(cur)}{cur_val:g}\n\n"
                    f"Send the NEW price (number only):")
            back = "cp:prices"
        elif what == "name":
            key = parts[1]
            self.user_states[chat_id] = f"cp_name:{key}"
            text = (f"🏷️ Rename button '{key}'\nNow: {button_name(key)}\n\n"
                    f"Send the NEW name (max {BUTTON_NAME_MAX} chars), or 'default':")
            back = "cp:names"
        elif what in self._CP_PROMPTS:
            self.user_states[chat_id] = f"cp_{what}"
            text = self._CP_PROMPTS[what]
            back = ("cp:users" if what in ("bal", "resetbal", "deluser", "ban", "unban")
                    else "cp:auto" if what.startswith("auto") else "cp:disc")
        else:
            return
        self._cp_show(token, chat_id, message_id, text + "\n\n(or tap Cancel)",
                      [[self._cp_b("❌ Cancel", "cp:cancel:" + back[3:])]])

    # ---------- callbacks ----------
    def _cp_callback(self, token, chat_id, message_id, data, query_id=None):
        if query_id:
            self._cp_api(token, "answerCallbackQuery", {"callback_query_id": query_id})
        p = data.split(":")
        act = p[1] if len(p) > 1 else "main"
        if act == "cancel":
            self.user_states.pop(chat_id, None)
            act = p[2] if len(p) > 2 else "main"
        if act == "main":
            self._cp_main(token, chat_id, message_id)
        elif act == "prices":
            self._cp_prices(token, chat_id, message_id)
        elif act == "names":
            self._cp_names(token, chat_id, message_id)
        elif act == "nmreset":
            for k in self._CP_TYPES:
                self.customer_db.set_setting("btn_name_" + k, "")
            load_button_names(self.customer_db)
            self._cp_names(token, chat_id, message_id)
        elif act == "users":
            self._cp_users(token, chat_id, message_id)
        elif act == "disc":
            self._cp_disc(token, chat_id, message_id)
        elif act == "stock":
            self._cp_stock(token, chat_id, message_id)
        elif act == "auto":
            self._cp_auto(token, chat_id, message_id)
        elif act == "autotog":
            now = self.customer_db.get_setting("auto_dep", "1") == "1"
            self.customer_db.set_setting("auto_dep", "0" if now else "1")
            self._cp_auto(token, chat_id, message_id)
        elif act == "shop":
            now = self.customer_db.get_setting("shop_open", "1") == "1"
            self.customer_db.set_setting("shop_open", "0" if now else "1")
            self._cp_main(token, chat_id, message_id)
        elif act == "access":
            self.customer_db.set_setting("users_enabled", "0" if self.users_enabled else "1")
            self._cp_main(token, chat_id, message_id)
        elif act == "ask":
            self._cp_ask(token, chat_id, message_id, p[2:])
        elif act == "delyes" and len(p) > 2:
            user = self.customer_db.get_user(p[2])
            if user:
                self.customer_db.delete_user(p[2])
                self._cp_show(token, chat_id, message_id,
                              f"✅ User {p[2]} ({user[1]}) deleted permanently.",
                              [[self._cp_b("⬅️ Panel", "cp:main")]])
            else:
                self._cp_show(token, chat_id, message_id, f"❌ User {p[2]} not found.",
                              [[self._cp_b("⬅️ Panel", "cp:main")]])
        elif act == "stats":
            self._bot_stats(token, chat_id)
        elif act == "deps":
            self._bot_pending_deposits(token, chat_id)
        elif act == "purch":
            self._bot_show_purchases(token, chat_id)
        elif act == "ulist":
            self._bot_users_list(token, chat_id)

    # ---------- typed answers (after tapping a panel button) ----------
    def _cp_done(self, token, chat_id, text, back="cp:main"):
        self._cp_api(token, "sendMessage",
                     {"chat_id": chat_id, "text": text,
                      "reply_markup": json.dumps({"inline_keyboard": [
                          [self._cp_b("⬅️ Back to Panel", back)]]})})

    def _cp_text(self, token, chat_id, state, text):
        """Handle the admin's typed reply for a panel action. True = handled."""
        t = text.strip()
        if t.lower() in ("cancel", "❌ cancel", "/cancel"):
            self.user_states.pop(chat_id, None)
            self._cp_done(token, chat_id, "❌ Cancelled.")
            return True
        if t.startswith("/"):            # a normal command: leave the panel flow
            self.user_states.pop(chat_id, None)
            return False
        parts = state.split(":")
        what = parts[0][3:]
        db = self.customer_db

        def retry(msg):
            self._cp_done(token, chat_id, msg + "\nTry again, or tap ❌ Cancel / send 'cancel'.")
            return True

        if what == "price":
            key, cur = parts[1], parts[2]
            try:
                val = float(t)
                if val < 0:
                    raise ValueError
            except ValueError:
                return retry("❌ Send a valid number.")
            db.set_price(key, val, cur)
            self.user_states.pop(chat_id, None)
            self._cp_done(token, chat_id,
                          f"✅ {button_name(key)} price = {self._cp_sym(cur)}{val:g} ({cur})", "cp:prices")
            return True
        if what == "name":
            key = parts[1]
            name = "" if t.lower() == "default" else t
            if len(name) > BUTTON_NAME_MAX:
                return retry(f"❌ Too long (max {BUTTON_NAME_MAX} characters).")
            db.set_setting("btn_name_" + key, name)
            load_button_names(db)
            self.user_states.pop(chat_id, None)
            self._cp_done(token, chat_id, f"✅ Button '{key}' is now: {button_name(key)}", "cp:names")
            return True
        if what == "bal":
            ps = t.split()
            try:
                cid, amt = ps[0], float(ps[1])
                cur = ps[2].upper() if len(ps) > 2 else None
                if cur not in (None, "BDT", "USD") or not db.get_user(cid):
                    raise ValueError
            except (ValueError, IndexError):
                return retry("❌ Format: <chat_id> <amount> [BDT|USD] (user must exist).")
            db.add_balance(cid, amt, cur)
            self.user_states.pop(chat_id, None)
            u = db.get_user(cid)
            self._cp_done(token, chat_id,
                          f"✅ Added {amt:g} {cur or u[3]} to {cid} ({u[1]}).\nNew balance: {u[2]:.2f} {u[3]}",
                          "cp:users")
            return True
        if what in ("resetbal", "deluser", "ban", "unban"):
            cid = t.split()[0]
            user = db.get_user(cid)
            if not user:
                return retry(f"❌ User {cid} not found.")
            self.user_states.pop(chat_id, None)
            if what == "resetbal":
                db.reset_balance(cid)
                self._cp_done(token, chat_id, f"✅ Balance reset to 0 for {cid} ({user[1]}).", "cp:users")
            elif what == "ban":
                db.set_banned(cid, True)
                self._cp_done(token, chat_id, f"🚫 {cid} ({user[1]}) banned.", "cp:users")
            elif what == "unban":
                db.set_banned(cid, False)
                self._cp_done(token, chat_id, f"✅ {cid} ({user[1]}) unbanned.", "cp:users")
            else:
                self._cp_api(token, "sendMessage", {
                    "chat_id": chat_id,
                    "text": f"⚠️ Delete {cid} ({user[1]}) permanently?\nBalance, purchases & deposits will be erased.",
                    "reply_markup": json.dumps({"inline_keyboard": [[
                        self._cp_b("🗑 Yes, delete", f"cp:delyes:{cid}"),
                        self._cp_b("❌ No", "cp:users")]]})})
            return True
        if what in ("disc50", "disc100"):
            try:
                val = float(t)
                if not 0 <= val <= 100:
                    raise ValueError
            except ValueError:
                return retry("❌ Send a percent between 0 and 100.")
            db.set_setting("disc_" + what[4:], f"{val:g}")
            self.user_states.pop(chat_id, None)
            self._cp_done(token, chat_id, f"✅ {what[4:]}+ discount = {val:g}%", "cp:disc")
            return True
        if what in ("autowin", "automax"):
            try:
                val = float(t)
                if val <= 0:
                    raise ValueError
            except ValueError:
                return retry("❌ Send a number greater than 0.")
            db.set_setting("auto_window_min" if what == "autowin" else "auto_max_bdt", f"{val:g}")
            self.user_states.pop(chat_id, None)
            self._cp_done(token, chat_id, f"✅ Saved: {val:g}", "cp:auto")
            return True
        if what == "autonum":
            nums = [re.sub(r"\D", "", x)[-11:] for x in t.replace("\n", ",").split(",")]
            nums = [n for n in nums if len(n) == 11]
            if not nums:
                return retry("❌ Send 11-digit numbers, comma separated.")
            db.set_setting("auto_numbers", ",".join(nums))
            self.user_states.pop(chat_id, None)
            self._cp_done(token, chat_id, f"✅ Receiving numbers: {', '.join(nums)}", "cp:auto")
            return True
        if what == "low":
            if not t.isdigit():
                return retry("❌ Send a whole number.")
            db.set_setting("low_stock_limit", t)
            self.user_states.pop(chat_id, None)
            self._cp_done(token, chat_id, f"✅ Low-stock alert at {t}", "cp:disc")
            return True
        self.user_states.pop(chat_id, None)
        return False


    # -------- stock persistence --------
    def silent_auto_save(self):
        try:
            with _STOCK_LOCK:
                data = {"panel_61xx": self.panel_61xx.get_data(),
                        "panel_1000x": self.panel_1000x.get_data(),
                        "panel_newid": self.panel_newid.get_data(),
                        "tg_token": self.token_entry.get(),
                        "tg_chat_id": self.chat_id_entry.get()}
                tmp = BACKUP_FILE + ".tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(data, f)
                os.replace(tmp, BACKUP_FILE)   # atomic: never leaves a half-written file
        except Exception as e:
            logger.error(f"[autosave] {e}")

    def load_backup(self):
        if os.path.exists(BACKUP_FILE):
            try:
                with open(BACKUP_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.panel_61xx.set_data(data.get("panel_61xx", ""))
                self.panel_1000x.set_data(data.get("panel_1000x", ""))
                self.panel_newid.set_data(data.get("panel_newid", ""))
            except Exception as e:
                logger.error(f"[load_backup] {e}")


# ============================================================
# এখানে আপনার আসল মান বসান (শুধু এই ৩ লাইন বদলাবেন)
# ============================================================
DEFAULT_BOT_TOKEN = "9174114244:AAGzpvUJYlZOwZ7vzYgH6dysnzuvsy8Y3qc"      # <- আপনার বট টোকেন (BotFather থেকে)
DEFAULT_MASTER_CHAT_ID = "5504053909"                                      # <- আপনার টেলিগ্রাম চ্যাট আইডি
DEFAULT_GEMINI_API_KEY = "AQ.Ab8RN6JXAvU1xtq3F_OVlz-iPVpxffRHBhEvqPCJTKVgTATTHg"         # <- আপনার Gemini API কী


def load_settings():
    """token + master chat id: environment > config.json > desktop backup json > default above."""
    token = os.environ.get("BOT_TOKEN", "9174114244:AAGzpvUJYlZOwZ7vzYgH6dysnzuvsy8Y3qc").strip()
    cid = os.environ.get("MASTER_CHAT_ID", "5504053909").strip()
    for path, tk_key, id_key in ((CONFIG_FILE, "token", "master_chat_id"),
                                 (BACKUP_FILE, "tg_token", "tg_chat_id")):
        if token and cid:
            break
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    d = json.load(f)
                token = token or str(d.get(tk_key, "")).strip()
                cid = cid or str(d.get(id_key, "")).strip()
            except Exception as e:
                logger.error(f"[settings] cannot read {path}: {e}")
    return token or DEFAULT_BOT_TOKEN, cid or DEFAULT_MASTER_CHAT_ID


def _load_api_key_from_config():
    """Gemini / Anthropic key: environment > config.json > default above."""
    try:
        has_env = {n: bool(os.environ.get(n, "").strip()) for n in ("GEMINI_API_KEY", "ANTHROPIC_API_KEY")}
        logger.info(f"[keys] env GEMINI_API_KEY={'set' if has_env['GEMINI_API_KEY'] else 'missing'}, "
                    f"ANTHROPIC_API_KEY={'set' if has_env['ANTHROPIC_API_KEY'] else 'missing'} | "
                    f"config.json: {CONFIG_FILE} ({'found' if os.path.exists(CONFIG_FILE) else 'NOT FOUND'})")
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            logger.info("[keys] config.json fields: " + ", ".join(sorted(cfg.keys())))
            for env_name, cfg_name in (("ANTHROPIC_API_KEY", "anthropic_api_key"),
                                       ("GEMINI_API_KEY", "gemini_api_key")):
                val = str(cfg.get(cfg_name, "")).strip()
                if val and not os.environ.get(env_name, "").strip():
                    os.environ[env_name] = val
    except Exception as e:
        logger.error(f"[settings] keys from config.json: {e}")
    # last fallback: the key written in this file
    if not os.environ.get("GEMINI_API_KEY", "").strip() and DEFAULT_GEMINI_API_KEY:
        os.environ["GEMINI_API_KEY"] = DEFAULT_GEMINI_API_KEY


def main():
    _load_api_key_from_config()
    token, cid = load_settings()
    logger.info(f"Data folder: {APP_FOLDER}")
    if not token or not cid:
        print("Bot token / master chat id not found.\n"
              "Set env BOT_TOKEN and MASTER_CHAT_ID, or create:\n"
              f"  {CONFIG_FILE}\n"
              '  {"token": "123:ABC...", "master_chat_id": "7061528394"}')
        sys.exit(1)
    BotApp(token, cid).run_forever()


if __name__ == "__main__":
    main()
