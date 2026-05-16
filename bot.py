# ============================================================
# GGOD FATHERR VCF MAKER BOT  ⚡  PREMIUM EDITION
# Admin: @ggod_fatherr
# ============================================================

import os
import io
import csv
import logging
import asyncio
import openpyxl
from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup,
    ReplyKeyboardMarkup, KeyboardButton
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)

# ── Logging ───────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────
BOT_TOKEN      = os.environ.get("BOT_TOKEN", "8603833578:AAFT3eB6HiDMyq8t_litDCw3Z2rFzB4LD4Q")
PREMIUM_CODE   = os.environ.get("PREMIUM_CODE", "#000#")
ADMIN_USERNAME = "@ggod_fatherr"
ADMIN_URL      = "https://t.me/ggod_fatherr"

# ── In-memory stores ──────────────────────────────────────────
premium_users: set = set()

# ── Branding ──────────────────────────────────────────────────
FOOTER  = "\n\n⚡ Powered by @ggod_fatherr"
DIVIDER = "▰▰▰▰▰▰▰▰▰▰▰▰▰▰▰▰▰▰▰▰"
THIN    = "─────────────────────"

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def is_premium(user_id: int) -> bool:
    return user_id in premium_users


def file_stem(fname: str) -> str:
    if "." in fname:
        return fname.rsplit(".", 1)[0]
    return fname or "output"


def premium_required_msg() -> str:
    return (
        "🔒 Premium Access Required\n"
        + THIN + "\n"
        "This feature is locked for free users.\n\n"
        "💎 Get Premium: " + ADMIN_USERNAME
        + FOOTER
    )


# ─────────────────────────────────────────────────────────────
# ANIMATED LOADING HELPER
# ─────────────────────────────────────────────────────────────

async def show_progress(msg, label: str = "Processing"):
    """Animate a progress bar by editing an existing message."""
    frames = [
        label + "  [▓▓░░░░░░░░] 20%",
        label + "  [▓▓▓▓░░░░░░] 40%",
        label + "  [▓▓▓▓▓▓░░░░] 60%",
        label + "  [▓▓▓▓▓▓▓▓░░] 80%",
        label + "  [▓▓▓▓▓▓▓▓▓▓] 100% ✅",
    ]
    for frame in frames:
        await asyncio.sleep(0.35)
        try:
            await msg.edit_text("⏳ " + frame)
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────
# KEYBOARDS
# ─────────────────────────────────────────────────────────────

def main_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📁 TXT → VCF",       callback_data="cmd_txt_to_vcf"),
            InlineKeyboardButton("📊 TXT/VCF → CSV",   callback_data="cmd_txtvcf_to_csv"),
        ],
        [
            InlineKeyboardButton("📇 CSV → VCF",       callback_data="cmd_csv_to_vcf"),
            InlineKeyboardButton("📄 VCF → TXT",       callback_data="cmd_vcf_to_txt"),
        ],
        [
            InlineKeyboardButton("💬 MSG → TXT",       callback_data="cmd_msg_to_txt"),
            InlineKeyboardButton("📝 Rename File",     callback_data="cmd_rename_file"),
        ],
        [
            InlineKeyboardButton("✏️ Rename Contacts", callback_data="cmd_rename_ctc"),
            InlineKeyboardButton("🧩 Merge VCFs",      callback_data="cmd_merge_vcf"),
        ],
        [
            InlineKeyboardButton("📚 Merge TXTs",      callback_data="cmd_merge_txt"),
            InlineKeyboardButton("✂️ Split File",      callback_data="cmd_split_file"),
        ],
        [
            InlineKeyboardButton("🛡️ Navy Format",     callback_data="cmd_admin_navy"),
            InlineKeyboardButton("🔄 Restart",         callback_data="cmd_reset"),
        ],
        [
            InlineKeyboardButton("ℹ️ Help & Guide",    callback_data="help"),
            InlineKeyboardButton("👤 Contact Admin",   url=ADMIN_URL),
        ],
    ])


def shortcut_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("📁 TXT→VCF"),  KeyboardButton("📄 VCF→TXT"),  KeyboardButton("📇 CSV→VCF")],
            [KeyboardButton("📊 →CSV"),     KeyboardButton("📝 Rename"),    KeyboardButton("✏️ Ren.CTC")],
            [KeyboardButton("🧩 MergeVCF"), KeyboardButton("📚 MergeTXT"), KeyboardButton("✂️ Split")],
            [KeyboardButton("🛡️ Navy"),     KeyboardButton("ℹ️ Help"),      KeyboardButton("🔄 Reset")],
        ],
        resize_keyboard=True,
        input_field_placeholder="Choose a feature or type a command..."
    )


def back_keyboard(feature_cb: str = None) -> InlineKeyboardMarkup:
    rows = []
    if feature_cb:
        rows.append([InlineKeyboardButton("◀️ Back", callback_data=feature_cb)])
    rows.append([InlineKeyboardButton("🏠 Main Menu", callback_data="back_to_menu")])
    return InlineKeyboardMarkup(rows)


def output_fmt_keyboard(feature_cb: str = None) -> InlineKeyboardMarkup:
    rows = [
        [
            InlineKeyboardButton("📄 TXT",  callback_data="fmt_txt"),
            InlineKeyboardButton("📇 VCF",  callback_data="fmt_vcf"),
        ],
        [
            InlineKeyboardButton("📊 CSV",  callback_data="fmt_csv"),
            InlineKeyboardButton("📗 XLSX", callback_data="fmt_xlsx"),
        ],
    ]
    if feature_cb:
        rows.append([InlineKeyboardButton("◀️ Back", callback_data=feature_cb)])
    rows.append([InlineKeyboardButton("🏠 Main Menu", callback_data="back_to_menu")])
    return InlineKeyboardMarkup(rows)


def done_keyboard(feature_cb: str = None) -> InlineKeyboardMarkup:
    rows = []
    if feature_cb:
        rows.append([InlineKeyboardButton("🔁 Use Again", callback_data=feature_cb)])
    rows.append([InlineKeyboardButton("🏠 Main Menu", callback_data="back_to_menu")])
    return InlineKeyboardMarkup(rows)


# ─────────────────────────────────────────────────────────────
# MESSAGE TEMPLATES
# ─────────────────────────────────────────────────────────────

def welcome_text() -> str:
    return (
        "╔══════════════════════════════╗\n"
        "║   👑  GGOD FATHERR BOT  👑  ║\n"
        "║   ⚡  VCF MAKER PREMIUM  ⚡  ║\n"
        "╚══════════════════════════════╝\n\n"
        "🔐 Premium Features Locked\n"
        + THIN + "\n"
        "🌟 Convert, merge, split & rename\n"
        "   contact files with ease.\n\n"
        "💎 Buy Access: " + ADMIN_USERNAME + "\n"
        + THIN + "\n"
        "👇 Tap below to unlock your access"
        + FOOTER
    )


def dashboard_text() -> str:
    return (
        "╔══════════════════════════════╗\n"
        "║  ⚡  GGOD FATHERR PREMIUM  ⚡ ║\n"
        "║  ✅  ALL FEATURES UNLOCKED   ║\n"
        "╚══════════════════════════════╝\n\n"
        + DIVIDER + "\n"
        "🚀 What would you like to do?\n"
        + DIVIDER
        + FOOTER
    )


def help_text() -> str:
    return (
        "╔══════════════════════════════╗\n"
        "║       ℹ️   HELP  GUIDE       ║\n"
        "╚══════════════════════════════╝\n\n"
        + THIN + "\n"
        "📌 Available Commands\n"
        + THIN + "\n"
        "▸ /start — Launch the bot\n"
        "▸ /txt_to_vcf — Numbers → VCF\n"
        "▸ /txtvcf_to_csv — TXT/VCF → CSV\n"
        "▸ /csv_to_vcf — CSV → VCF\n"
        "▸ /vcf_to_txt — VCF → Numbers\n"
        "▸ /msg_to_txt — Message → TXT\n"
        "▸ /rename_file — Rename any file\n"
        "▸ /rename_ctc — Rename VCF contacts\n"
        "▸ /merge_vcf — Merge VCF files\n"
        "▸ /merge_txt — Merge TXT files\n"
        "▸ /split_file — Split contacts\n"
        "▸ /admin_navy_file — Navy format\n"
        "▸ /done — Finalize merge\n"
        "▸ /reset — Clear session\n"
        + THIN + "\n"
        "👤 Admin: " + ADMIN_USERNAME
        + FOOTER
    )


# ─────────────────────────────────────────────────────────────
# CONVERSION LOGIC
# ─────────────────────────────────────────────────────────────

def txt_to_vcf(text: str) -> str:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    vcf_entries = []
    for i, number in enumerate(lines, 1):
        vcf_entries.append(
            "BEGIN:VCARD\nVERSION:3.0\n"
            "FN:Contact " + str(i) + "\nTEL:" + number + "\nEND:VCARD"
        )
    return "\n".join(vcf_entries)


def vcf_to_txt(text: str) -> str:
    numbers = []
    for line in text.splitlines():
        line = line.strip()
        if line.upper().startswith("TEL"):
            parts = line.split(":")
            if len(parts) >= 2:
                numbers.append(parts[-1].strip())
    return "\n".join(numbers)


def parse_vcf(text: str) -> list:
    entries = []
    name, phone = "Unknown", ""
    for line in text.splitlines():
        line = line.strip()
        if line.upper().startswith("FN:"):
            name = line[3:].strip()
        elif line.upper().startswith("TEL"):
            parts = line.split(":")
            if len(parts) >= 2:
                phone = parts[-1].strip()
        elif line.upper() == "END:VCARD":
            if phone:
                entries.append((name, phone))
            name, phone = "Unknown", ""
    return entries


def csv_to_vcf(text: str) -> str:
    reader = csv.DictReader(io.StringIO(text))
    vcf_entries = []
    i = 1
    for row in reader:
        name  = (row.get("Name") or row.get("name") or
                 row.get("FN") or ("Contact " + str(i)))
        phone = (row.get("Phone") or row.get("phone") or
                 row.get("TEL") or row.get("Number") or "")
        if phone:
            vcf_entries.append(
                "BEGIN:VCARD\nVERSION:3.0\n"
                "FN:" + name + "\nTEL:" + phone + "\nEND:VCARD"
            )
        i += 1
    return "\n".join(vcf_entries)


def rename_vcf_contacts(text: str, prefix: str) -> str:
    lines = text.splitlines()
    result = []
    counter = 1
    for line in lines:
        if line.strip().upper().startswith("FN:"):
            result.append("FN:" + prefix + " " + str(counter))
            counter += 1
        else:
            result.append(line)
    return "\n".join(result)


def split_contacts(text: str, count: int, ext: str) -> list:
    chunks = []
    if ext == "vcf":
        blocks = []
        current = []
        for line in text.splitlines():
            current.append(line)
            if line.strip().upper() == "END:VCARD":
                blocks.append("\n".join(current))
                current = []
        for i in range(0, len(blocks), count):
            chunks.append("\n".join(blocks[i:i + count]))
    else:
        lines = [l for l in text.splitlines() if l.strip()]
        for i in range(0, len(lines), count):
            chunks.append("\n".join(lines[i:i + count]))
    return chunks


def admin_navy_format(text: str) -> str:
    entries = parse_vcf(text)
    vcf_entries = []
    for i, (name, phone) in enumerate(entries, 1):
        vcf_entries.append(
            "BEGIN:VCARD\nVERSION:3.0\n"
            "FN:NAVY " + str(i).zfill(4) + " " + name + "\n"
            "TEL:" + phone + "\nEND:VCARD"
        )
    return "\n".join(vcf_entries)


def content_to_csv_bytes(content: str, source_ext: str) -> bytes:
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Name", "Phone"])
    if source_ext == "vcf":
        for name, phone in parse_vcf(content):
            writer.writerow([name, phone])
    else:
        for i, line in enumerate(
            [l.strip() for l in content.splitlines() if l.strip()], 1
        ):
            writer.writerow(["Contact " + str(i), line])
    return output.getvalue().encode("utf-8")


def content_to_xlsx_bytes(content: str, source_ext: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contacts"
    ws.append(["Name", "Phone"])
    if source_ext == "vcf":
        for name, phone in parse_vcf(content):
            ws.append([name, phone])
    else:
        for i, line in enumerate(
            [l.strip() for l in content.splitlines() if l.strip()], 1
        ):
            ws.append(["Contact " + str(i), line])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def convert_content(content: str, source_ext: str, target_ext: str) -> tuple:
    if target_ext == "txt":
        data = vcf_to_txt(content).encode("utf-8") if source_ext == "vcf" else content.encode("utf-8")
        return data, "text/plain"
    elif target_ext == "vcf":
        if source_ext == "txt":
            data = txt_to_vcf(content).encode("utf-8")
        elif source_ext in ("csv", "xlsx"):
            data = csv_to_vcf(content).encode("utf-8")
        else:
            data = content.encode("utf-8")
        return data, "text/vcard"
    elif target_ext == "csv":
        return content_to_csv_bytes(content, source_ext), "text/csv"
    elif target_ext == "xlsx":
        return content_to_xlsx_bytes(content, source_ext), \
               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return content.encode("utf-8"), "application/octet-stream"


def count_contacts(content: str, norm_ext: str) -> int:
    """Count total contacts/numbers in any normalised content."""
    if norm_ext == "vcf":
        return content.upper().count("BEGIN:VCARD")
    elif norm_ext in ("csv", "xlsx"):
        # Count non-header rows that have a phone column
        rows = list(csv.reader(io.StringIO(content)))
        if len(rows) <= 1:
            return 0
        return len([r for r in rows[1:] if any(c.strip() for c in r)])
    else:
        # txt — count non-empty lines
        return len([l for l in content.splitlines() if l.strip()])


# ─────────────────────────────────────────────────────────────
# UNIVERSAL FILE READER  (.txt / .csv / .vcf / .xlsx)
# ─────────────────────────────────────────────────────────────

ALL_FORMATS = ("txt", "csv", "vcf", "xlsx")

def read_file_content(raw: bytes, ext: str) -> tuple:
    """
    Read raw file bytes and return (text_content, normalised_ext).
    XLSX is converted to CSV-style text so all downstream logic works.
    Returns (content_str, ext) or raises ValueError on bad format.
    """
    if ext == "xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        ws = wb.active
        out = io.StringIO()
        writer = csv.writer(out)
        for row in ws.iter_rows(values_only=True):
            writer.writerow([str(c) if c is not None else "" for c in row])
        # treat xlsx as csv for all conversion logic
        return out.getvalue(), "csv"
    else:
        return raw.decode("utf-8", errors="ignore"), ext


# ─────────────────────────────────────────────────────────────
# FEATURE PROMPTS
# ─────────────────────────────────────────────────────────────

FEATURE_PROMPTS = {
    "txt_to_vcf":    ("cmd_txt_to_vcf",    "📁 Any → VCF\n\nSend a .txt / .csv / .vcf / .xlsx file.\nPhone numbers will be converted to VCF contacts."),
    "txtvcf_to_csv": ("cmd_txtvcf_to_csv", "📊 Any → CSV/XLSX\n\nSend a .txt / .csv / .vcf / .xlsx file to convert."),
    "csv_to_vcf":    ("cmd_csv_to_vcf",    "📇 Any → VCF\n\nSend a .txt / .csv / .vcf / .xlsx file to convert to VCF contacts."),
    "vcf_to_txt":    ("cmd_vcf_to_txt",    "📄 Any → TXT\n\nSend a .txt / .csv / .vcf / .xlsx file to extract phone numbers."),
    "msg_to_txt":    ("cmd_msg_to_txt",    "💬 MSG → TXT\n\nType or paste your message below."),
    "rename_file":   ("cmd_rename_file",   "📝 Rename File\n\nSend any file (.txt / .csv / .vcf / .xlsx) to rename."),
    "rename_ctc":    ("cmd_rename_ctc",    "✏️ Rename Contacts\n\nSend a .vcf / .csv / .xlsx file to rename all contacts."),
    "merge_vcf":     ("cmd_merge_vcf",     "🧩 Merge → VCF\n\nSend .txt / .csv / .vcf / .xlsx files one by one.\nTap Done or type /done when finished."),
    "merge_txt":     ("cmd_merge_txt",     "📚 Merge → TXT\n\nSend .txt / .csv / .vcf / .xlsx files one by one.\nTap Done or type /done when finished."),
    "split_file":    ("cmd_split_file",    "✂️ Split File\n\nSend a .txt / .csv / .vcf / .xlsx file to split."),
    "admin_navy":    ("cmd_admin_navy",    "🛡️ Navy Format\n\nSend a .vcf / .csv / .xlsx file to apply admin navy numbering."),
}


# ─────────────────────────────────────────────────────────────
# COMMAND HANDLERS
# ─────────────────────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user    = update.effective_user
    user_id = user.id
    context.user_data.pop("state", None)
    context.user_data.pop("awaiting_code", None)

    if is_premium(user_id):
        first = user.first_name or "there"
        await update.message.reply_text(
            "👋 Welcome back, " + first + "!\n"
            "⌨️ Shortcut keyboard is ready." + FOOTER,
            reply_markup=shortcut_keyboard()
        )
        await update.message.reply_text(
            dashboard_text(),
            reply_markup=main_menu_keyboard()
        )
        return

    await update.message.reply_text(
        welcome_text(),
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🔑 Enter Premium Code", callback_data="enter_code")],
            [InlineKeyboardButton("👤 Contact Admin",      url=ADMIN_URL)],
            [InlineKeyboardButton("ℹ️ Help & Guide",       callback_data="help")],
        ])
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    back_cb = "back_to_welcome"
    if update.callback_query and is_premium(update.callback_query.from_user.id):
        back_cb = "back_to_menu"

    if update.message:
        await update.message.reply_text(
            help_text(),
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🏠 Main Menu", callback_data="back_to_menu")]
            ])
        )
    elif update.callback_query:
        await update.callback_query.edit_message_text(
            help_text(),
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ Back", callback_data=back_cb)]
            ])
        )


async def reset_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    msg = (
        "🔄 Session Cleared!\n"
        + THIN + "\n"
        "All temporary data has been wiped.\n"
        "You're ready for a fresh start."
        + FOOTER
    )
    if update.message:
        kb = done_keyboard() if is_premium(update.effective_user.id) else None
        await update.message.reply_text(msg, reply_markup=kb)
    elif update.callback_query:
        kb = done_keyboard() if is_premium(update.callback_query.from_user.id) else None
        await update.callback_query.edit_message_text(msg, reply_markup=kb)


async def done_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_premium(user_id):
        await update.message.reply_text(premium_required_msg())
        return

    state = context.user_data.get("state")
    fcb   = context.user_data.get("active_feature_cb")

    if state == "merge_vcf":
        files = context.user_data.get("merge_vcf_files", [])
        if not files:
            await update.message.reply_text("❌ No VCF files received yet.")
            return
        stems = context.user_data.get("merge_vcf_stems", [])
        context.user_data["pending_result"]     = "\n".join(files)
        context.user_data["pending_stem"]       = "_".join(stems[:3]) if stems else "merged"
        context.user_data["pending_source_ext"] = "vcf"
        context.user_data["state"]              = "ask_output_fmt"
        context.user_data.pop("merge_vcf_files", None)
        context.user_data.pop("merge_vcf_stems", None)
        await update.message.reply_text(
            "✅ " + str(len(files)) + " VCF files queued!\n📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )

    elif state == "merge_txt":
        files = context.user_data.get("merge_txt_files", [])
        if not files:
            await update.message.reply_text("❌ No TXT files received yet.")
            return
        stems = context.user_data.get("merge_txt_stems", [])
        context.user_data["pending_result"]     = "\n".join(files)
        context.user_data["pending_stem"]       = "_".join(stems[:3]) if stems else "merged"
        context.user_data["pending_source_ext"] = "txt"
        context.user_data["state"]              = "ask_output_fmt"
        context.user_data.pop("merge_txt_files", None)
        context.user_data.pop("merge_txt_stems", None)
        await update.message.reply_text(
            "✅ " + str(len(files)) + " TXT files queued!\n📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )

    else:
        await update.message.reply_text(
            "⚠️ Nothing to finalize. Start a merge operation first.",
            reply_markup=main_menu_keyboard()
        )


# ─────────────────────────────────────────────────────────────
# FEATURE COMMAND SHORTCUTS
# ─────────────────────────────────────────────────────────────

async def _feature(update: Update, context: ContextTypes.DEFAULT_TYPE, state: str):
    if not is_premium(update.effective_user.id):
        await update.message.reply_text(premium_required_msg())
        return
    fcb, prompt = FEATURE_PROMPTS[state]
    context.user_data["state"]             = state
    context.user_data["active_feature_cb"] = fcb
    await update.message.reply_text(prompt + FOOTER, reply_markup=back_keyboard(fcb))


async def cmd_txt_to_vcf(u, c):     await _feature(u, c, "txt_to_vcf")
async def cmd_txtvcf_to_csv(u, c):  await _feature(u, c, "txtvcf_to_csv")
async def cmd_csv_to_vcf(u, c):     await _feature(u, c, "csv_to_vcf")
async def cmd_vcf_to_txt(u, c):     await _feature(u, c, "vcf_to_txt")
async def cmd_msg_to_txt(u, c):     await _feature(u, c, "msg_to_txt")
async def cmd_rename_file(u, c):    await _feature(u, c, "rename_file")
async def cmd_rename_ctc(u, c):     await _feature(u, c, "rename_ctc")
async def cmd_merge_vcf(u, c):      await _feature(u, c, "merge_vcf")
async def cmd_merge_txt(u, c):      await _feature(u, c, "merge_txt")
async def cmd_split_file(u, c):     await _feature(u, c, "split_file")
async def cmd_admin_navy_file(u, c): await _feature(u, c, "admin_navy")


# ─────────────────────────────────────────────────────────────
# CALLBACK QUERY HANDLER
# ─────────────────────────────────────────────────────────────

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query   = update.callback_query
    await query.answer()
    data    = query.data
    user_id = query.from_user.id

    if data == "enter_code":
        context.user_data["awaiting_code"] = True
        await query.edit_message_text(
            "🔑 Premium Code Verification\n"
            + THIN + "\n"
            "Please type and send your premium code below 👇"
            + FOOTER,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ Back", callback_data="back_to_welcome")]
            ])
        )
        return

    if data == "help":
        await help_command(update, context)
        return

    if data == "cmd_reset":
        await reset_command(update, context)
        return

    if data == "back_to_menu":
        context.user_data.pop("state", None)
        context.user_data.pop("active_feature_cb", None)
        if is_premium(user_id):
            await query.edit_message_text(
                dashboard_text(),
                reply_markup=main_menu_keyboard()
            )
        else:
            await query.edit_message_text(
                "🔒 Premium Required\n"
                "💎 Get access: " + ADMIN_USERNAME + FOOTER,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔑 Enter Code", callback_data="enter_code")]
                ])
            )
        return

    if data == "back_to_welcome":
        context.user_data.pop("awaiting_code", None)
        context.user_data.pop("state", None)
        await query.edit_message_text(
            welcome_text(),
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔑 Enter Premium Code", callback_data="enter_code")],
                [InlineKeyboardButton("👤 Contact Admin",      url=ADMIN_URL)],
                [InlineKeyboardButton("ℹ️ Help & Guide",       callback_data="help")],
            ])
        )
        return

    if data == "merge_vcf_done":
        files = context.user_data.get("merge_vcf_files", [])
        if not files:
            await query.answer("⚠️ No files received yet!", show_alert=True)
            return
        stems = context.user_data.get("merge_vcf_stems", [])
        context.user_data["pending_result"]     = "\n".join(files)
        context.user_data["pending_stem"]       = "_".join(stems[:3]) if stems else "merged"
        context.user_data["pending_source_ext"] = "vcf"
        context.user_data["state"]              = "ask_output_fmt"
        fcb = context.user_data.get("active_feature_cb")
        context.user_data.pop("merge_vcf_files", None)
        context.user_data.pop("merge_vcf_stems", None)
        await query.edit_message_text(
            "✅ " + str(len(files)) + " VCF files ready!\n📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )
        return

    if data == "merge_txt_done":
        files = context.user_data.get("merge_txt_files", [])
        if not files:
            await query.answer("⚠️ No files received yet!", show_alert=True)
            return
        stems = context.user_data.get("merge_txt_stems", [])
        context.user_data["pending_result"]     = "\n".join(files)
        context.user_data["pending_stem"]       = "_".join(stems[:3]) if stems else "merged"
        context.user_data["pending_source_ext"] = "txt"
        context.user_data["state"]              = "ask_output_fmt"
        fcb = context.user_data.get("active_feature_cb")
        context.user_data.pop("merge_txt_files", None)
        context.user_data.pop("merge_txt_stems", None)
        await query.edit_message_text(
            "✅ " + str(len(files)) + " TXT files ready!\n📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )
        return

    if data in ("fmt_txt", "fmt_vcf", "fmt_csv", "fmt_xlsx"):
        chosen_ext = data.split("_", 1)[1]
        fcb        = context.user_data.get("active_feature_cb")
        cur_state  = context.user_data.get("state")

        if cur_state == "split_ask_fmt":
            split_data = context.user_data.get("split_file_data")
            split_stem = context.user_data.get("split_file_stem", "split")
            count      = context.user_data.get("split_count", 100)
            orig_ext   = context.user_data.get("split_file_ext", "vcf")
            if not split_data:
                await query.edit_message_text(
                    "❌ File not found. Please start again.",
                    reply_markup=back_keyboard(fcb)
                )
                return
            prog = await query.edit_message_text("⏳ Splitting  [▓▓░░░░░░░░] 20%")
            await asyncio.sleep(0.4)
            chunks = split_contacts(split_data, count, orig_ext)
            await prog.edit_text("⏳ Splitting  [▓▓▓▓▓▓░░░░] 60%")
            await asyncio.sleep(0.3)
            for i, chunk in enumerate(chunks, 1):
                file_bytes, _ = convert_content(chunk, orig_ext, chosen_ext)
                oname = split_stem + "_part" + str(i) + "." + chosen_ext
                buf = io.BytesIO(file_bytes)
                buf.name = oname
                await query.message.reply_document(document=buf, filename=oname)
            await prog.edit_text("✅ Split complete!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            context.user_data.pop("state", None)
            context.user_data.pop("split_file_data", None)
            context.user_data.pop("split_file_ext", None)
            context.user_data.pop("split_file_stem", None)
            context.user_data.pop("split_count", None)
            await query.message.reply_text(
                "🎉 Done! Split into " + str(len(chunks)) + " files as ." + chosen_ext + FOOTER,
                reply_markup=done_keyboard(fcb)
            )
            return

        result     = context.user_data.get("pending_result", "")
        stem       = context.user_data.get("pending_stem", "output")
        source_ext = context.user_data.get("pending_source_ext", "vcf")

        prog = await query.edit_message_text(
            "⏳ Converting to ." + chosen_ext + "  [▓▓▓▓▓░░░░░] 50%"
        )
        await asyncio.sleep(0.5)
        file_bytes, _ = convert_content(result, source_ext, chosen_ext)
        oname = stem + "." + chosen_ext
        buf   = io.BytesIO(file_bytes)
        buf.name = oname
        await prog.edit_text("✅ Conversion complete!  [▓▓▓▓▓▓▓▓▓▓] 100%")
        await query.message.reply_document(
            document=buf,
            filename=oname,
            caption=(
                "✅ File Ready!\n"
                + THIN + "\n"
                "📄 " + oname + "\n"
                "📦 Format: ." + chosen_ext
                + FOOTER
            )
        )
        context.user_data.pop("state", None)
        context.user_data.pop("pending_result", None)
        context.user_data.pop("pending_stem", None)
        context.user_data.pop("pending_source_ext", None)
        context.user_data.pop("fmt_options", None)
        await query.message.reply_text(
            "🎉 All done! What's next?",
            reply_markup=done_keyboard(fcb)
        )
        return

    feature_map = {
        "cmd_txt_to_vcf":    "txt_to_vcf",
        "cmd_txtvcf_to_csv": "txtvcf_to_csv",
        "cmd_csv_to_vcf":    "csv_to_vcf",
        "cmd_vcf_to_txt":    "vcf_to_txt",
        "cmd_msg_to_txt":    "msg_to_txt",
        "cmd_rename_file":   "rename_file",
        "cmd_rename_ctc":    "rename_ctc",
        "cmd_merge_vcf":     "merge_vcf",
        "cmd_merge_txt":     "merge_txt",
        "cmd_split_file":    "split_file",
        "cmd_admin_navy":    "admin_navy",
    }

    if data in feature_map:
        if not is_premium(user_id):
            await query.edit_message_text(premium_required_msg())
            return
        state = feature_map[data]
        _, prompt = FEATURE_PROMPTS[state]
        context.user_data["state"]             = state
        context.user_data["active_feature_cb"] = data
        context.user_data.pop("merge_vcf_files", None)
        context.user_data.pop("merge_txt_files", None)
        await query.edit_message_text(
            prompt + FOOTER,
            reply_markup=back_keyboard(data)
        )


# ─────────────────────────────────────────────────────────────
# TEXT MESSAGE HANDLER
# ─────────────────────────────────────────────────────────────

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text    = update.message.text.strip()
    state   = context.user_data.get("state")

    # Shortcut keyboard button mapping
    shortcut_map = {
        "📁 TXT→VCF":  "txt_to_vcf",
        "📄 VCF→TXT":  "vcf_to_txt",
        "📇 CSV→VCF":  "csv_to_vcf",
        "📊 →CSV":     "txtvcf_to_csv",
        "📝 Rename":   "rename_file",
        "✏️ Ren.CTC":  "rename_ctc",
        "🧩 MergeVCF": "merge_vcf",
        "📚 MergeTXT": "merge_txt",
        "✂️ Split":    "split_file",
        "🛡️ Navy":     "admin_navy",
        "ℹ️ Help":     "__help__",
        "🔄 Reset":    "__reset__",
    }

    if text in shortcut_map:
        mapped = shortcut_map[text]
        if mapped == "__help__":
            await help_command(update, context)
            return
        if mapped == "__reset__":
            await reset_command(update, context)
            return
        if not is_premium(user_id):
            await update.message.reply_text(premium_required_msg())
            return
        fcb, prompt = FEATURE_PROMPTS[mapped]
        context.user_data["state"]             = mapped
        context.user_data["active_feature_cb"] = fcb
        await update.message.reply_text(prompt + FOOTER, reply_markup=back_keyboard(fcb))
        return

    # Premium code verification
    if context.user_data.get("awaiting_code"):
        context.user_data["awaiting_code"] = False
        if text == PREMIUM_CODE:
            premium_users.add(user_id)
            first = update.effective_user.first_name or "there"
            msg = await update.message.reply_text("🔓 Verifying code...")
            await asyncio.sleep(0.5)
            await msg.edit_text("✅ Code Accepted!\n🚀 Unlocking premium features...")
            await asyncio.sleep(0.8)
            await msg.edit_text("🎉 Premium Access Granted!\n\nWelcome, " + first + "! 👑")
            await asyncio.sleep(0.5)
            await update.message.reply_text(
                "⌨️ Shortcut keyboard activated!" + FOOTER,
                reply_markup=shortcut_keyboard()
            )
            await update.message.reply_text(
                dashboard_text(),
                reply_markup=main_menu_keyboard()
            )
        else:
            await update.message.reply_text(
                "❌ Invalid Code\n"
                + THIN + "\n"
                "⚠️ Contact admin: " + ADMIN_USERNAME + FOOTER,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔑 Try Again",     callback_data="enter_code")],
                    [InlineKeyboardButton("👤 Contact Admin", url=ADMIN_URL)],
                ])
            )
        return

    if not is_premium(user_id):
        await update.message.reply_text(
            premium_required_msg(),
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔑 Enter Code", callback_data="enter_code")]
            ])
        )
        return

    fcb = context.user_data.get("active_feature_cb")

    if state == "msg_to_txt":
        buf = io.BytesIO(text.encode("utf-8"))
        buf.name = "message.txt"
        await update.message.reply_document(
            document=buf,
            filename="message.txt",
            caption="✅ Message saved as TXT!" + FOOTER
        )
        context.user_data.pop("state", None)
        await update.message.reply_text("🎉 Done! What's next?", reply_markup=done_keyboard(fcb))
        return

    if state == "rename_newname":
        new_name   = text
        file_bytes = context.user_data.get("rename_file_bytes")
        if not file_bytes:
            await update.message.reply_text("❌ File not found. Please start again.", reply_markup=back_keyboard(fcb))
            return
        buf = io.BytesIO(file_bytes)
        buf.name = new_name
        await update.message.reply_document(
            document=buf,
            filename=new_name,
            caption="✅ File renamed to: " + new_name + FOOTER
        )
        context.user_data.pop("state", None)
        context.user_data.pop("rename_file_bytes", None)
        context.user_data.pop("rename_file_stem", None)
        await update.message.reply_text("🎉 Done! What's next?", reply_markup=done_keyboard(fcb))
        return

    if state == "rename_ctc_prefix":
        prefix    = text
        vcf_bytes = context.user_data.get("rename_ctc_bytes")
        stem      = context.user_data.get("rename_ctc_stem", "contacts")
        if not vcf_bytes:
            await update.message.reply_text("❌ VCF not found. Please start again.", reply_markup=back_keyboard(fcb))
            return
        result = rename_vcf_contacts(vcf_bytes.decode("utf-8", errors="ignore"), prefix)
        fname  = stem + "_" + prefix + "_renamed.vcf"
        buf    = io.BytesIO(result.encode("utf-8"))
        buf.name = fname
        await update.message.reply_document(
            document=buf,
            filename=fname,
            caption="✅ Contacts renamed!\n🏷️ Prefix: " + prefix + "\n📄 " + fname + FOOTER
        )
        context.user_data.pop("state", None)
        context.user_data.pop("rename_ctc_bytes", None)
        context.user_data.pop("rename_ctc_stem", None)
        await update.message.reply_text("🎉 Done! What's next?", reply_markup=done_keyboard(fcb))
        return

    if state == "split_count":
        try:
            count = int(text)
            if count <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text(
                "❌ Please enter a valid positive number.\nExample: 100",
                reply_markup=back_keyboard(fcb)
            )
            return
        split_data = context.user_data.get("split_file_data")
        if not split_data:
            await update.message.reply_text("❌ File not found. Please start again.", reply_markup=back_keyboard(fcb))
            return
        context.user_data["split_count"] = count
        context.user_data["state"]       = "split_ask_fmt"
        total = count_contacts(split_data, context.user_data.get("split_file_ext", "txt"))
        files_count = (total + count - 1) // count if total > 0 else "?"
        await update.message.reply_text(
            "✅ Got it!\n"
            + THIN + "\n"
            "📋 Total Contacts: " + str(total) + "\n"
            "✂️ Per File: " + str(count) + "\n"
            "📦 Files to create: " + str(files_count) + "\n\n"
            "📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )
        return

    await update.message.reply_text(
        "💡 Select a feature from the menu to get started.",
        reply_markup=main_menu_keyboard()
    )


# ─────────────────────────────────────────────────────────────
# FILE / DOCUMENT HANDLER
# ─────────────────────────────────────────────────────────────

async def file_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    if not is_premium(user_id):
        await update.message.reply_text(
            premium_required_msg(),
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔑 Enter Code", callback_data="enter_code")]
            ])
        )
        return

    state = context.user_data.get("state")
    doc   = update.message.document

    if not doc:
        await update.message.reply_text("❌ Please send a valid file.")
        return

    fname = doc.file_name or "file"
    ext   = fname.rsplit(".", 1)[-1].lower() if "." in fname else ""
    fcb   = context.user_data.get("active_feature_cb")

    prog = await update.message.reply_text("⏳ Receiving file  [▓▓░░░░░░░░] 20%")

    try:
        tg_file = await doc.get_file()
        raw     = await tg_file.download_as_bytearray()
        raw     = bytes(raw)

        await prog.edit_text("⏳ Processing  [▓▓▓▓▓░░░░░] 50%")
        await asyncio.sleep(0.3)

        # ── Validate extension is one of the 4 accepted formats ──
        if ext not in ALL_FORMATS:
            await prog.edit_text(
                "❌ Unsupported format: ." + ext + "\n"
                "✅ Accepted: .txt  .csv  .vcf  .xlsx"
            )
            await update.message.reply_text("↩️ Try again:", reply_markup=back_keyboard(fcb))
            return

        # ── Read & normalise the file ─────────────────────────
        content, norm_ext = read_file_content(raw, ext)

        # ── Route by active state ─────────────────────────────

        if state == "txt_to_vcf":
            context.user_data["pending_result"]     = content
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = norm_ext
            context.user_data["state"]              = "ask_output_fmt"
            await prog.edit_text("✅ File received!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "📤 Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        elif state == "txtvcf_to_csv":
            context.user_data["pending_result"]     = content
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = norm_ext
            context.user_data["state"]              = "ask_output_fmt"
            await prog.edit_text("✅ File received!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "📤 Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        elif state == "csv_to_vcf":
            context.user_data["pending_result"]     = content
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = norm_ext
            context.user_data["state"]              = "ask_output_fmt"
            await prog.edit_text("✅ File received!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "📤 Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        elif state == "vcf_to_txt":
            context.user_data["pending_result"]     = content
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = norm_ext
            context.user_data["state"]              = "ask_output_fmt"
            await prog.edit_text("✅ File received!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "📤 Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        elif state == "rename_file":
            context.user_data["rename_file_bytes"] = raw
            context.user_data["rename_file_stem"]  = file_stem(fname)
            context.user_data["state"] = "rename_newname"
            await prog.edit_text("✅ File received!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "📝 Got: " + fname + "\n\n"
                "Now send the new filename with extension.\n"
                "Example: contacts_backup.vcf" + FOOTER,
                reply_markup=back_keyboard(fcb)
            )

        elif state == "rename_ctc":
            # For rename contacts, normalise to VCF text first
            if norm_ext == "vcf":
                vcf_text = content
            elif norm_ext in ("csv", "xlsx"):
                vcf_text = csv_to_vcf(content)
            else:
                # txt — treat each line as a phone number, build vcf
                vcf_text = txt_to_vcf(content)
            context.user_data["rename_ctc_bytes"] = vcf_text.encode("utf-8")
            context.user_data["rename_ctc_stem"]  = file_stem(fname)
            context.user_data["state"] = "rename_ctc_prefix"
            await prog.edit_text("✅ File received!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "✏️ Got: " + fname + "\n\n"
                "Now send the contact name prefix.\n"
                "Example: GGOD → GGOD 1, GGOD 2..." + FOOTER,
                reply_markup=back_keyboard(fcb)
            )

        elif state == "merge_vcf":
            files = context.user_data.setdefault("merge_vcf_files", [])
            stems = context.user_data.setdefault("merge_vcf_stems", [])
            files.append(content)
            stems.append(file_stem(fname))
            await prog.edit_text("✅ File " + str(len(files)) + " added!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "📎 " + fname + " added ✓\n"
                "📦 Total files: " + str(len(files)) + "\n\n"
                "Send more files or tap Done to merge." + FOOTER,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("✅ Done — Merge Now", callback_data="merge_vcf_done")],
                    [InlineKeyboardButton("◀️ Cancel",           callback_data="back_to_menu")],
                ])
            )

        elif state == "merge_txt":
            files = context.user_data.setdefault("merge_txt_files", [])
            stems = context.user_data.setdefault("merge_txt_stems", [])
            files.append(content)
            stems.append(file_stem(fname))
            await prog.edit_text("✅ File " + str(len(files)) + " added!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "📎 " + fname + " added ✓\n"
                "📦 Total files: " + str(len(files)) + "\n\n"
                "Send more files or tap Done to merge." + FOOTER,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("✅ Done — Merge Now", callback_data="merge_txt_done")],
                    [InlineKeyboardButton("◀️ Cancel",           callback_data="back_to_menu")],
                ])
            )

        elif state == "split_file":
            context.user_data["split_file_data"] = content
            context.user_data["split_file_ext"]  = norm_ext
            context.user_data["split_file_stem"] = file_stem(fname)
            context.user_data["state"] = "split_count"
            total = count_contacts(content, norm_ext)
            await prog.edit_text("✅ File received!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "✂️ Got: " + fname + "\n\n"
                "📊 Analysis Complete\n"
                + THIN + "\n"
                "📋 Total Contacts: " + str(total) + "\n\n"
                "🔢 Enter how many contacts per file?\n"
                "Example: 100" + FOOTER,
                reply_markup=back_keyboard(fcb)
            )

        elif state == "admin_navy":
            # Normalise to VCF entries first
            if norm_ext == "vcf":
                vcf_text = content
            elif norm_ext in ("csv", "xlsx"):
                vcf_text = csv_to_vcf(content)
            else:
                vcf_text = txt_to_vcf(content)
            context.user_data["pending_result"]     = admin_navy_format(vcf_text)
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = "vcf"
            context.user_data["state"]              = "ask_output_fmt"
            await prog.edit_text("✅ Navy format applied!  [▓▓▓▓▓▓▓▓▓▓] 100%")
            await update.message.reply_text(
                "📤 Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        else:
            await prog.edit_text("⚠️ No active feature selected. Please choose from the menu first.")
            await update.message.reply_text("👇 Select a feature:", reply_markup=main_menu_keyboard())

    except Exception as e:
        logger.error("File handler error: %s", e)
        try:
            await prog.edit_text("❌ Error: " + str(e))
        except Exception:
            pass
        await update.message.reply_text(
            "❌ Something went wrong. Please try again.",
            reply_markup=back_keyboard(fcb)
        )


# ─────────────────────────────────────────────────────────────
# ERROR HANDLER
# ─────────────────────────────────────────────────────────────

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logger.error("Unhandled exception: %s", context.error, exc_info=context.error)


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main():
    print("=" * 40)
    print("  ⚡ GGOD FATHERR BOT STARTING...  ")
    print("  👑 PREMIUM SYSTEM ACTIVE         ")
    print("  🟢 STATUS: ONLINE                ")
    print("=" * 40)

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start",           start))
    app.add_handler(CommandHandler("help",            help_command))
    app.add_handler(CommandHandler("reset",           reset_command))
    app.add_handler(CommandHandler("done",            done_command))
    app.add_handler(CommandHandler("txt_to_vcf",      cmd_txt_to_vcf))
    app.add_handler(CommandHandler("txtvcf_to_csv",   cmd_txtvcf_to_csv))
    app.add_handler(CommandHandler("csv_to_vcf",      cmd_csv_to_vcf))
    app.add_handler(CommandHandler("vcf_to_txt",      cmd_vcf_to_txt))
    app.add_handler(CommandHandler("msg_to_txt",      cmd_msg_to_txt))
    app.add_handler(CommandHandler("rename_file",     cmd_rename_file))
    app.add_handler(CommandHandler("rename_ctc",      cmd_rename_ctc))
    app.add_handler(CommandHandler("merge_vcf",       cmd_merge_vcf))
    app.add_handler(CommandHandler("merge_txt",       cmd_merge_txt))
    app.add_handler(CommandHandler("split_file",      cmd_split_file))
    app.add_handler(CommandHandler("admin_navy_file", cmd_admin_navy_file))

    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, file_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))

    app.add_error_handler(error_handler)

    print("🤖 Bot is polling for updates...")
    app.run_polling(allowed_updates=["message", "callback_query"])


if __name__ == "__main__":
    main()
