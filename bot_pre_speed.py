# ============================================================
# GGOD FATHERR VCF MAKER BOT
# Admin: @ggod_fatherr
# ============================================================

import os
import io
import csv
import logging
import asyncio
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
from telegram.constants import ParseMode

# ── Logging setup ─────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ── Bot config ────────────────────────────────────────────────
BOT_TOKEN      = "8603833578:AAFT3eB6HiDMyq8t_litDCw3Z2rFzB4LD4Q"   # <-- Replace with your token
PREMIUM_CODE   = "#000#"
ADMIN_USERNAME = "@ggod_fatherr"

# ── In-memory stores ──────────────────────────────────────────
premium_users: set = set()   # verified user IDs

# ── Shared footer ─────────────────────────────────────────────
FOOTER = "\n\n👑 Powered By @ggod_fatherr"

# ─────────────────────────────────────────────────────────────
# UTILITY FUNCTIONS
# ─────────────────────────────────────────────────────────────

def is_premium(user_id: int) -> bool:
    """Check if user has premium access."""
    return user_id in premium_users


def file_stem(fname: str) -> str:
    """Return filename without extension, e.g. 'contacts.vcf' → 'contacts'."""
    if "." in fname:
        return fname.rsplit(".", 1)[0]
    return fname or "output"


def premium_required_msg() -> str:
    """Standard message for non-premium users."""
    return (
        "🚫 Premium Access Required\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"💎 Buy Premium: {ADMIN_USERNAME}"
        + FOOTER
    )


def main_menu_keyboard() -> InlineKeyboardMarkup:
    """Build the main premium dashboard keyboard."""
    buttons = [
        [
            InlineKeyboardButton("📁 Text → VCF",      callback_data="cmd_txt_to_vcf"),
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
            InlineKeyboardButton("✏️ Rename CTC VCF",  callback_data="cmd_rename_ctc"),
            InlineKeyboardButton("🧩 Merge VCFs",      callback_data="cmd_merge_vcf"),
        ],
        [
            InlineKeyboardButton("📚 Merge TXTs",      callback_data="cmd_merge_txt"),
            InlineKeyboardButton("✂️ Split File",      callback_data="cmd_split_file"),
        ],
        [
            InlineKeyboardButton("🛡️ Admin Navy File", callback_data="cmd_admin_navy"),
            InlineKeyboardButton("🔄 Restart",         callback_data="cmd_reset"),
        ],
        [
            InlineKeyboardButton("ℹ️ Help",            callback_data="help"),
            InlineKeyboardButton("👤 Contact Admin",   url="https://t.me/ggod_fatherr"),
        ],
    ]
    return InlineKeyboardMarkup(buttons)


def shortcut_keyboard() -> ReplyKeyboardMarkup:
    """Persistent bottom keyboard with command shortcuts."""
    buttons = [
        [KeyboardButton("/txt_to_vcf"),    KeyboardButton("/vcf_to_txt"),    KeyboardButton("/csv_to_vcf")],
        [KeyboardButton("/txtvcf_to_csv"), KeyboardButton("/rename_file"),   KeyboardButton("/rename_ctc")],
        [KeyboardButton("/merge_vcf"),     KeyboardButton("/merge_txt"),     KeyboardButton("/split_file")],
        [KeyboardButton("/admin_navy_file"),KeyboardButton("/help"),          KeyboardButton("/reset")],
    ]
    return ReplyKeyboardMarkup(
        buttons,
        resize_keyboard=True,
        input_field_placeholder="Choose a feature or type a command..."
    )


def back_keyboard(feature_cb: str = None) -> InlineKeyboardMarkup:
    """Back button — goes back to the feature prompt if known, else main menu."""
    if feature_cb:
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("◀️ Back to Feature", callback_data=feature_cb)],
            [InlineKeyboardButton("🏠 Main Menu",        callback_data="back_to_menu")],
        ])
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏠 Main Menu", callback_data="back_to_menu")]
    ])


def output_fmt_keyboard(feature_cb: str = None) -> InlineKeyboardMarkup:
    """Ask user to choose output format: TXT, VCF, CSV, XLSX."""
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
        rows.append([InlineKeyboardButton("◀️ Back to Feature", callback_data=feature_cb)])
    rows.append([InlineKeyboardButton("🏠 Main Menu", callback_data="back_to_menu")])
    return InlineKeyboardMarkup(rows)


def done_keyboard(feature_cb: str = None) -> InlineKeyboardMarkup:
    """Shown after a feature completes."""
    if feature_cb:
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("🔁 Use Again",  callback_data=feature_cb)],
            [InlineKeyboardButton("🏠 Main Menu",  callback_data="back_to_menu")],
        ])
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏠 Main Menu", callback_data="back_to_menu")]
    ])


def dashboard_text() -> str:
    """Premium dashboard header text."""
    return (
        "╔══════════════════════════════╗\n"
        "║  ⚡ GGOD FATHERR PREMIUM ⚡  ║\n"
        "╚══════════════════════════════╝\n\n"
        "✅ Access Granted — All Features Unlocked\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "Select a feature below 👇"
        + FOOTER
    )


async def show_dashboard(update: Update):
    """Send the premium dashboard to the user."""
    if update.message:
        await update.message.reply_text(
            dashboard_text(),
            reply_markup=main_menu_keyboard()
        )
    elif update.callback_query:
        await update.callback_query.edit_message_text(
            dashboard_text(),
            reply_markup=main_menu_keyboard()
        )


# ─────────────────────────────────────────────────────────────
# CONVERSION LOGIC
# ─────────────────────────────────────────────────────────────

def txt_to_vcf(text: str) -> str:
    """Convert phone numbers (one per line) to VCF format."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    vcf_entries = []
    for i, number in enumerate(lines, 1):
        vcf_entries.append(
            f"BEGIN:VCARD\nVERSION:3.0\n"
            f"FN:Contact {i}\nTEL:{number}\nEND:VCARD"
        )
    return "\n".join(vcf_entries)


def txtvcf_to_csv(text: str, ext: str) -> str:
    """Convert TXT (numbers) or VCF to CSV."""
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Name", "Phone"])
    if ext == "vcf":
        entries = parse_vcf(text)
        for name, phone in entries:
            writer.writerow([name, phone])
    else:
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        for i, number in enumerate(lines, 1):
            writer.writerow([f"Contact {i}", number])
    return output.getvalue()


def csv_to_vcf(text: str) -> str:
    """Convert CSV to VCF contacts."""
    reader = csv.DictReader(io.StringIO(text))
    vcf_entries = []
    i = 1
    for row in reader:
        # Try common column name variations
        name  = row.get("Name") or row.get("name") or row.get("FN") or f"Contact {i}"
        phone = row.get("Phone") or row.get("phone") or row.get("TEL") or row.get("Number") or ""
        if phone:
            vcf_entries.append(
                f"BEGIN:VCARD\nVERSION:3.0\n"
                f"FN:{name}\nTEL:{phone}\nEND:VCARD"
            )
        i += 1
    return "\n".join(vcf_entries)


def vcf_to_txt(text: str) -> str:
    """Extract phone numbers from VCF."""
    numbers = []
    for line in text.splitlines():
        line = line.strip()
        if line.upper().startswith("TEL"):
            # Handle TEL;TYPE=...:number and TEL:number
            parts = line.split(":")
            if len(parts) >= 2:
                numbers.append(parts[-1].strip())
    return "\n".join(numbers)


def parse_vcf(text: str) -> list:
    """Parse VCF and return list of (name, phone) tuples."""
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


def rename_vcf_contacts(text: str, prefix: str) -> str:
    """Rename all VCF contact FN fields with prefix + number."""
    lines = text.splitlines()
    result = []
    counter = 1
    for line in lines:
        if line.strip().upper().startswith("FN:"):
            result.append(f"FN:{prefix} {counter}")
            counter += 1
        else:
            result.append(line)
    return "\n".join(result)


def split_contacts(text: str, count: int, ext: str) -> list:
    """Split contacts file into chunks of `count` contacts each."""
    chunks = []
    if ext == "vcf":
        # Split by VCARD blocks
        blocks = []
        current = []
        for line in text.splitlines():
            current.append(line)
            if line.strip().upper() == "END:VCARD":
                blocks.append("\n".join(current))
                current = []
        # Group into chunks
        for i in range(0, len(blocks), count):
            chunks.append("\n".join(blocks[i:i+count]))
    else:
        # Split by lines
        lines = [l for l in text.splitlines() if l.strip()]
        for i in range(0, len(lines), count):
            chunks.append("\n".join(lines[i:i+count]))
    return chunks


def admin_navy_format(text: str) -> str:
    """Generate admin navy numbered VCF format."""
    entries = parse_vcf(text)
    vcf_entries = []
    for i, (name, phone) in enumerate(entries, 1):
        vcf_entries.append(
            f"BEGIN:VCARD\nVERSION:3.0\n"
            f"FN:NAVY {i:04d} {name}\n"
            f"TEL:{phone}\nEND:VCARD"
        )
    return "\n".join(vcf_entries)


def content_to_csv_bytes(content: str, source_ext: str) -> bytes:
    """Convert any supported content to CSV bytes."""
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Name", "Phone"])
    if source_ext == "vcf":
        for name, phone in parse_vcf(content):
            writer.writerow([name, phone])
    else:
        for i, line in enumerate([l.strip() for l in content.splitlines() if l.strip()], 1):
            writer.writerow([f"Contact {i}", line])
    return output.getvalue().encode("utf-8")


def content_to_xlsx_bytes(content: str, source_ext: str) -> bytes:
    """Convert any supported content to XLSX bytes."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contacts"
    ws.append(["Name", "Phone"])
    if source_ext == "vcf":
        for name, phone in parse_vcf(content):
            ws.append([name, phone])
    else:
        for i, line in enumerate([l.strip() for l in content.splitlines() if l.strip()], 1):
            ws.append([f"Contact {i}", line])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def convert_content(content: str, source_ext: str, target_ext: str) -> tuple[bytes, str]:
    """
    Convert content to target_ext.
    Returns (file_bytes, mime_hint).
    source_ext: original file extension (vcf / txt / csv)
    target_ext: desired output extension (txt / vcf / csv / xlsx)
    """
    if target_ext == "txt":
        if source_ext == "vcf":
            data = vcf_to_txt(content).encode("utf-8")
        else:
            data = content.encode("utf-8")
        return data, "text/plain"

    elif target_ext == "vcf":
        if source_ext == "txt":
            data = txt_to_vcf(content).encode("utf-8")
        elif source_ext == "csv":
            data = csv_to_vcf(content).encode("utf-8")
        else:
            data = content.encode("utf-8")   # already vcf
        return data, "text/vcard"

    elif target_ext == "csv":
        return content_to_csv_bytes(content, source_ext), "text/csv"

    elif target_ext == "xlsx":
        return content_to_xlsx_bytes(content, source_ext), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # fallback
    return content.encode("utf-8"), "application/octet-stream"


# ─────────────────────────────────────────────────────────────
# COMMAND HANDLERS
# ─────────────────────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /start command."""
    user_id = update.effective_user.id

    # Always clear any stuck state on /start
    context.user_data.pop("state", None)
    context.user_data.pop("awaiting_code", None)

    # If already premium, show dashboard directly
    if is_premium(user_id):
        await update.message.reply_text(
            "⌨️ Shortcut keyboard enabled!" + FOOTER,
            reply_markup=shortcut_keyboard()
        )
        await update.message.reply_text(
            dashboard_text(),
            reply_markup=main_menu_keyboard()
        )
        return

    # Show welcome + verification prompt
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔑 Enter Premium Code", callback_data="enter_code")],
        [InlineKeyboardButton("👤 Contact Admin",      url="https://t.me/ggod_fatherr")],
        [InlineKeyboardButton("ℹ️ Help",               callback_data="help")],
    ])
    await update.message.reply_text(
        "╔══════════════════════════════╗\n"
        "║   👑 GGOD FATHERR BOT 👑    ║\n"
        "╚══════════════════════════════╝\n\n"
        "⚡ VCF MAKER BOT — PREMIUM EDITION\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "🔐 Premium Features Locked\n"
        f"💎 Buy Premium Access: {ADMIN_USERNAME}\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "👇 Press below to verify your code"
        + FOOTER,
        reply_markup=keyboard
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /help command."""
    text = (
        "╔══════════════════════════════╗\n"
        "║        ℹ️  HELP MENU         ║\n"
        "╚══════════════════════════════╝\n\n"
        "📌 Available Commands:\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "▸ /start — Start the bot\n"
        "▸ /txt_to_vcf — Text → VCF\n"
        "▸ /txtvcf_to_csv — TXT/VCF → CSV\n"
        "▸ /csv_to_vcf — CSV → VCF\n"
        "▸ /vcf_to_txt — VCF → TXT\n"
        "▸ /msg_to_txt — Message → TXT\n"
        "▸ /rename_file — Rename a file\n"
        "▸ /rename_ctc — Rename VCF contacts\n"
        "▸ /merge_vcf — Merge VCF files\n"
        "▸ /merge_txt — Merge TXT files\n"
        "▸ /split_file — Split contacts\n"
        "▸ /admin_navy_file — Navy format\n"
        "▸ /reset — Restart bot\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 Admin: {ADMIN_USERNAME}"
        + FOOTER
    )
    if update.message:
        await update.message.reply_text(text, reply_markup=back_keyboard())
    elif update.callback_query:
        await update.callback_query.edit_message_text(
            text,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ Back", callback_data="back_to_welcome" if not is_premium(update.callback_query.from_user.id) else "back_to_menu")]
            ])
        )


async def reset_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /reset command — clear session data."""
    context.user_data.clear()
    msg = (
        "🔄 Bot Restarted Successfully!\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "All temporary data cleared."
        + FOOTER
    )
    if update.message:
        await update.message.reply_text(msg)
        if is_premium(update.effective_user.id):
            await update.message.reply_text("What would you like to do next?", reply_markup=done_keyboard())
    elif update.callback_query:
        await update.callback_query.edit_message_text(msg, reply_markup=done_keyboard() if is_premium(update.callback_query.from_user.id) else None)


async def done_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /done — finalize merge operations."""
    user_id = update.effective_user.id
    if not is_premium(user_id):
        await update.message.reply_text(premium_required_msg())
        return

    state = context.user_data.get("state")

    if state == "merge_vcf":
        files = context.user_data.get("merge_vcf_files", [])
        if not files:
            await update.message.reply_text("❌ No VCF files received yet.")
            return
        stems = context.user_data.get("merge_vcf_stems", [])
        merged_stem = "_".join(stems[:3]) if stems else "merged"
        context.user_data["pending_result"]     = "\n".join(files)
        context.user_data["pending_stem"]       = merged_stem
        context.user_data["pending_source_ext"] = "vcf"
        context.user_data["state"]              = "ask_output_fmt"
        fcb = context.user_data.get("active_feature_cb")
        context.user_data.pop("merge_vcf_files", None)
        context.user_data.pop("merge_vcf_stems", None)
        await update.message.reply_text(
            f"✅ {len(files)} VCF files ready!\n📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )

    elif state == "merge_txt":
        files = context.user_data.get("merge_txt_files", [])
        if not files:
            await update.message.reply_text("❌ No TXT files received yet.")
            return
        stems = context.user_data.get("merge_txt_stems", [])
        merged_stem = "_".join(stems[:3]) if stems else "merged"
        context.user_data["pending_result"]     = "\n".join(files)
        context.user_data["pending_stem"]       = merged_stem
        context.user_data["pending_source_ext"] = "txt"
        context.user_data["state"]              = "ask_output_fmt"
        fcb = context.user_data.get("active_feature_cb")
        context.user_data.pop("merge_txt_files", None)
        context.user_data.pop("merge_txt_stems", None)
        await update.message.reply_text(
            f"✅ {len(files)} TXT files ready!\n📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )

    else:
        await update.message.reply_text(
            "⚠️ Nothing to finalize. Use the menu to start a merge operation."
        )


# ─────────────────────────────────────────────────────────────
# FEATURE COMMAND SHORTCUTS
# ─────────────────────────────────────────────────────────────

async def _feature(update: Update, context: ContextTypes.DEFAULT_TYPE, state: str, prompt: str, feature_cb: str = None):
    """Generic feature activator — checks premium and sets state."""
    if not is_premium(update.effective_user.id):
        await update.message.reply_text(premium_required_msg())
        return
    context.user_data["state"] = state
    if feature_cb:
        context.user_data["active_feature_cb"] = feature_cb
    await update.message.reply_text(prompt + FOOTER, reply_markup=back_keyboard(feature_cb))


async def cmd_txt_to_vcf(u, c):
    await _feature(u, c, "txt_to_vcf", "📁 Send me a .txt file with phone numbers (one per line).")

async def cmd_txtvcf_to_csv(u, c):
    await _feature(u, c, "txtvcf_to_csv", "📊 Send me a .txt or .vcf file to convert to CSV.")

async def cmd_csv_to_vcf(u, c):
    await _feature(u, c, "csv_to_vcf", "📇 Send me a .csv file to convert to VCF.")

async def cmd_vcf_to_txt(u, c):
    await _feature(u, c, "vcf_to_txt", "📄 Send me a .vcf file to extract phone numbers.")

async def cmd_msg_to_txt(u, c):
    await _feature(u, c, "msg_to_txt", "💬 Type or paste your message and I'll save it as a TXT file.")

async def cmd_rename_file(u, c):
    await _feature(u, c, "rename_file", "📝 Send me the file you want to rename.")

async def cmd_rename_ctc(u, c):
    await _feature(u, c, "rename_ctc", "✏️ Send me a .vcf file to rename all contacts.")

async def cmd_merge_vcf(u, c):
    await _feature(u, c, "merge_vcf", "🧩 Send me .vcf files one by one, then type /done to merge.")

async def cmd_merge_txt(u, c):
    await _feature(u, c, "merge_txt", "📚 Send me .txt files one by one, then type /done to merge.")

async def cmd_split_file(u, c):
    await _feature(u, c, "split_file", "✂️ Send me a .vcf or .txt file to split.")

async def cmd_admin_navy_file(u, c):
    await _feature(u, c, "admin_navy", "🛡️ Send me a .vcf file to generate admin navy format.")


# ─────────────────────────────────────────────────────────────
# CALLBACK QUERY HANDLER (Inline Buttons)
# ─────────────────────────────────────────────────────────────

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle all inline keyboard button presses."""
    query = update.callback_query
    await query.answer()
    data    = query.data
    user_id = query.from_user.id

    # ── Enter premium code ────────────────────────────────────
    if data == "enter_code":
        await query.edit_message_text(
            "🔑 Premium Code Verification\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "📥 Please type and send your premium code:"
            + FOOTER,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ Back", callback_data="back_to_welcome")]
            ])
        )
        context.user_data["awaiting_code"] = True
        return

    # ── Help ──────────────────────────────────────────────────
    if data == "help":
        await help_command(update, context)
        return

    # ── Reset ─────────────────────────────────────────────────
    if data == "cmd_reset":
        await reset_command(update, context)
        return

    # ── Back to main menu (premium users) ─────────────────────
    if data == "back_to_menu":
        if is_premium(user_id):
            await query.edit_message_text(
                dashboard_text(),
                reply_markup=main_menu_keyboard()
            )
        else:
            await query.edit_message_text(
                "🚫 Premium Access Required\n"
                f"💎 Buy Premium: {ADMIN_USERNAME}" + FOOTER,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔑 Enter Premium Code", callback_data="enter_code")]
                ])
            )
        context.user_data.pop("state", None)
        context.user_data.pop("active_feature_cb", None)
        return

    # ── Back to welcome screen (non-premium) ──────────────────
    if data == "back_to_welcome":
        context.user_data.pop("awaiting_code", None)
        context.user_data.pop("state", None)
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("🔑 Enter Premium Code", callback_data="enter_code")],
            [InlineKeyboardButton("👤 Contact Admin",      url="https://t.me/ggod_fatherr")],
            [InlineKeyboardButton("ℹ️ Help",               callback_data="help")],
        ])
        await query.edit_message_text(
            "╔══════════════════════════════╗\n"
            "║   👑 GGOD FATHERR BOT 👑    ║\n"
            "╚══════════════════════════════╝\n\n"
            "⚡ VCF MAKER BOT — PREMIUM EDITION\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "🔐 Premium Features Locked\n"
            f"💎 Buy Premium Access: {ADMIN_USERNAME}\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "👇 Press below to verify your code"
            + FOOTER,
            reply_markup=keyboard
        )
        return

    # ── Merge done via button ─────────────────────────────────
    if data == "merge_vcf_done":
        files = context.user_data.get("merge_vcf_files", [])
        if not files:
            await query.answer("No files received yet!", show_alert=True)
            return
        stems = context.user_data.get("merge_vcf_stems", [])
        merged_stem = "_".join(stems[:3]) if stems else "merged"
        context.user_data["pending_result"]     = "\n".join(files)
        context.user_data["pending_stem"]       = merged_stem
        context.user_data["pending_source_ext"] = "vcf"
        context.user_data["state"]              = "ask_output_fmt"
        fcb = context.user_data.get("active_feature_cb")
        await query.edit_message_text(
            f"✅ {len(files)} VCF files ready to merge!\n📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )
        context.user_data.pop("merge_vcf_files", None)
        context.user_data.pop("merge_vcf_stems", None)
        return

    if data == "merge_txt_done":
        files = context.user_data.get("merge_txt_files", [])
        if not files:
            await query.answer("No files received yet!", show_alert=True)
            return
        stems = context.user_data.get("merge_txt_stems", [])
        merged_stem = "_".join(stems[:3]) if stems else "merged"
        context.user_data["pending_result"]     = "\n".join(files)
        context.user_data["pending_stem"]       = merged_stem
        context.user_data["pending_source_ext"] = "txt"
        context.user_data["state"]              = "ask_output_fmt"
        fcb = context.user_data.get("active_feature_cb")
        await query.edit_message_text(
            f"✅ {len(files)} TXT files ready to merge!\n📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )
        context.user_data.pop("merge_txt_files", None)
        context.user_data.pop("merge_txt_stems", None)
        return
    # ── Output format choice (fmt_txt / fmt_vcf / fmt_csv / fmt_xlsx) ──
    if data in ("fmt_txt", "fmt_vcf", "fmt_csv", "fmt_xlsx"):
        chosen_ext = data.split("_", 1)[1]   # "txt" / "vcf" / "csv" / "xlsx"
        fcb        = context.user_data.get("active_feature_cb")
        cur_state  = context.user_data.get("state")

        # ── Split format choice ───────────────────────────────
        if cur_state == "split_ask_fmt":
            split_data = context.user_data.get("split_file_data")
            split_stem = context.user_data.get("split_file_stem", "split")
            count      = context.user_data.get("split_count", 100)
            orig_ext   = context.user_data.get("split_file_ext", "vcf")
            if not split_data:
                await query.edit_message_text("❌ No file found. Please start again.", reply_markup=back_keyboard(fcb))
                return
            await query.edit_message_text("⏳ Processing File..." + FOOTER)
            chunks = split_contacts(split_data, count, orig_ext)
            for i, chunk in enumerate(chunks, 1):
                file_bytes, _ = convert_content(chunk, orig_ext, chosen_ext)
                oname = f"{split_stem}_part{i}.{chosen_ext}"
                buf = io.BytesIO(file_bytes)
                buf.name = oname
                await query.message.reply_document(document=buf, filename=oname)
            context.user_data.pop("state", None)
            context.user_data.pop("split_file_data", None)
            context.user_data.pop("split_file_ext", None)
            context.user_data.pop("split_file_stem", None)
            context.user_data.pop("split_count", None)
            await query.message.reply_text(
                f"✅ Split into {len(chunks)} files as .{chosen_ext}!" + FOOTER,
                reply_markup=done_keyboard(fcb)
            )
            return

        # ── Generic pending result ────────────────────────────
        result     = context.user_data.get("pending_result", "")
        stem       = context.user_data.get("pending_stem", "output")
        source_ext = context.user_data.get("pending_source_ext", "vcf")
        file_bytes, _ = convert_content(result, source_ext, chosen_ext)
        oname = f"{stem}.{chosen_ext}"
        buf   = io.BytesIO(file_bytes)
        buf.name = oname
        await query.edit_message_text(f"⏳ Sending as .{chosen_ext}..." + FOOTER)
        await query.message.reply_document(
            document=buf, filename=oname,
            caption=f"✅ File Ready!\n📄 {oname}" + FOOTER
        )
        context.user_data.pop("state", None)
        context.user_data.pop("pending_result", None)
        context.user_data.pop("pending_stem", None)
        context.user_data.pop("pending_source_ext", None)
        context.user_data.pop("fmt_options", None)
        await query.message.reply_text("What would you like to do next?", reply_markup=done_keyboard(fcb))
        return

    feature_map = {
        "cmd_txt_to_vcf":    ("txt_to_vcf",    "📁 Send me a .txt file with phone numbers (one per line)."),
        "cmd_txtvcf_to_csv": ("txtvcf_to_csv", "📊 Send me a .txt or .vcf file to convert to CSV."),
        "cmd_csv_to_vcf":    ("csv_to_vcf",    "📇 Send me a .csv file to convert to VCF."),
        "cmd_vcf_to_txt":    ("vcf_to_txt",    "📄 Send me a .vcf file to extract phone numbers."),
        "cmd_msg_to_txt":    ("msg_to_txt",    "💬 Type or paste your message and I'll save it as TXT."),
        "cmd_rename_file":   ("rename_file",   "📝 Send me the file you want to rename."),
        "cmd_rename_ctc":    ("rename_ctc",    "✏️ Send me a .vcf file to rename all contacts."),
        "cmd_merge_vcf":     ("merge_vcf",     "🧩 Send .vcf files one by one, then type /done to merge."),
        "cmd_merge_txt":     ("merge_txt",     "📚 Send .txt files one by one, then type /done to merge."),
        "cmd_split_file":    ("split_file",    "✂️ Send me a .vcf or .txt file to split."),
        "cmd_admin_navy":    ("admin_navy",    "🛡️ Send me a .vcf file to generate admin navy format."),
    }

    if data in feature_map:
        if not is_premium(user_id):
            await query.edit_message_text(premium_required_msg())
            return
        state, prompt = feature_map[data]
        context.user_data["state"] = state
        context.user_data["active_feature_cb"] = data   # remember which feature is active
        # Clear any leftover merge buffers
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
    """Handle all incoming text messages."""
    user_id = update.effective_user.id
    text    = update.message.text.strip()
    state   = context.user_data.get("state")

    # ── Premium code verification ─────────────────────────────
    if context.user_data.get("awaiting_code"):
        context.user_data["awaiting_code"] = False
        if text == PREMIUM_CODE:
            premium_users.add(user_id)
            await update.message.reply_text(
                "✅ Premium Access Verified!\n"
                "🚀 Unlocking Premium Features..."
                + FOOTER
            )
            await asyncio.sleep(1)
            await update.message.reply_text(
                "⌨️ Shortcut keyboard enabled!" + FOOTER,
                reply_markup=shortcut_keyboard()
            )
            await update.message.reply_text(
                dashboard_text(),
                reply_markup=main_menu_keyboard()
            )
        else:
            await update.message.reply_text(
                "❌ Invalid Premium Code\n"
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                f"⚠️ Contact Admin: {ADMIN_USERNAME}"
                + FOOTER
            )
        return

    # ── Block non-premium users ───────────────────────────────
    if not is_premium(user_id):
        await update.message.reply_text(premium_required_msg())
        return

    # ── MSG → TXT ─────────────────────────────────────────────
    if state == "msg_to_txt":
        fcb = context.user_data.get("active_feature_cb")
        buf = io.BytesIO(text.encode("utf-8"))
        buf.name = "message.txt"
        await update.message.reply_document(
            document=buf, filename="message.txt",
            caption="✅ Message saved as TXT file!" + FOOTER
        )
        context.user_data.pop("state", None)
        await update.message.reply_text("What would you like to do next?", reply_markup=done_keyboard(fcb))
        return

    # ── Rename file — waiting for new name ────────────────────
    if state == "rename_newname":
        new_name   = text
        file_bytes = context.user_data.get("rename_file_bytes")
        fcb        = context.user_data.get("active_feature_cb")
        if not file_bytes:
            await update.message.reply_text("❌ No file found. Please start again.", reply_markup=back_keyboard(fcb))
            return
        buf = io.BytesIO(file_bytes)
        buf.name = new_name
        await update.message.reply_document(
            document=buf, filename=new_name,
            caption=f"✅ File renamed to: `{new_name}`" + FOOTER
        )
        context.user_data.pop("state", None)
        context.user_data.pop("rename_file_bytes", None)
        context.user_data.pop("rename_file_stem", None)
        await update.message.reply_text("What would you like to do next?", reply_markup=done_keyboard(fcb))
        return

    # ── Rename CTC — waiting for prefix ──────────────────────
    if state == "rename_ctc_prefix":
        prefix    = text
        vcf_bytes = context.user_data.get("rename_ctc_bytes")
        stem      = context.user_data.get("rename_ctc_stem", "contacts")
        fcb       = context.user_data.get("active_feature_cb")
        if not vcf_bytes:
            await update.message.reply_text("❌ No VCF found. Please start again.", reply_markup=back_keyboard(fcb))
            return
        result = rename_vcf_contacts(vcf_bytes.decode("utf-8", errors="ignore"), prefix)
        buf = io.BytesIO(result.encode("utf-8"))
        fname = f"{stem}_{prefix}_renamed.vcf"
        buf.name = fname
        await update.message.reply_document(
            document=buf, filename=fname,
            caption=f"✅ Contacts renamed with prefix: {prefix}\n📄 {fname}" + FOOTER
        )
        context.user_data.pop("state", None)
        context.user_data.pop("rename_ctc_bytes", None)
        context.user_data.pop("rename_ctc_stem", None)
        await update.message.reply_text("What would you like to do next?", reply_markup=done_keyboard(fcb))
        return

    # ── Split — waiting for count ─────────────────────────────
    if state == "split_count":
        try:
            count = int(text)
            if count <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("❌ Please enter a valid positive number.", reply_markup=back_keyboard(context.user_data.get("active_feature_cb")))
            return
        split_data = context.user_data.get("split_file_data")
        split_ext  = context.user_data.get("split_file_ext", "vcf")
        split_stem = context.user_data.get("split_file_stem", "split")
        fcb        = context.user_data.get("active_feature_cb")
        if not split_data:
            await update.message.reply_text("❌ No file found. Please start again.", reply_markup=back_keyboard(fcb))
            return
        # Store count and ask for output format
        context.user_data["split_count"]    = count
        context.user_data["state"]          = "split_ask_fmt"
        context.user_data["fmt_options"]    = (split_ext, "txt" if split_ext == "vcf" else "vcf")
        await update.message.reply_text(
            f"✅ Will split into chunks of {count}.\n📤 Choose output format:" + FOOTER,
            reply_markup=output_fmt_keyboard(fcb)
        )
        return

    # ── Fallback ──────────────────────────────────────────────
    await update.message.reply_text(
        "⚠️ Please select a feature from the menu.",
        reply_markup=main_menu_keyboard()
    )


# ─────────────────────────────────────────────────────────────
# FILE / DOCUMENT HANDLER
# ─────────────────────────────────────────────────────────────

async def file_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle all incoming document/file uploads."""
    user_id = update.effective_user.id

    if not is_premium(user_id):
        await update.message.reply_text(premium_required_msg())
        return

    state = context.user_data.get("state")
    doc   = update.message.document

    if not doc:
        await update.message.reply_text("❌ Please send a valid file.")
        return

    fname = doc.file_name or ""
    ext   = fname.rsplit(".", 1)[-1].lower() if "." in fname else ""
    fcb   = context.user_data.get("active_feature_cb")

    await update.message.reply_text("⏳ Processing File...")

    try:
        tg_file = await doc.get_file()
        raw     = await tg_file.download_as_bytearray()
        content = bytes(raw)

        # ── TXT → VCF ────────────────────────────────────────
        if state == "txt_to_vcf":
            if ext != "txt":
                await update.message.reply_text("❌ Invalid File Format — send a .txt file", reply_markup=back_keyboard(fcb))
                return
            context.user_data["pending_result"]     = content.decode("utf-8", errors="ignore")
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = "txt"
            context.user_data["state"]              = "ask_output_fmt"
            await update.message.reply_text(
                "✅ File processed!\n📤 Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        # ── TXT/VCF → CSV ────────────────────────────────────
        elif state == "txtvcf_to_csv":
            if ext not in ("txt", "vcf"):
                await update.message.reply_text("❌ Invalid File Format — send .txt or .vcf", reply_markup=back_keyboard(fcb))
                return
            context.user_data["pending_result"]     = content.decode("utf-8", errors="ignore")
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = ext
            context.user_data["state"]              = "ask_output_fmt"
            await update.message.reply_text(
                "✅ File processed!\n� Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        # ── CSV → VCF ────────────────────────────────────────
        elif state == "csv_to_vcf":
            if ext != "csv":
                await update.message.reply_text("❌ Invalid File Format — send a .csv file", reply_markup=back_keyboard(fcb))
                return
            context.user_data["pending_result"]     = content.decode("utf-8", errors="ignore")
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = "csv"
            context.user_data["state"]              = "ask_output_fmt"
            await update.message.reply_text(
                "✅ File processed!\n📤 Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        # ── VCF → TXT ────────────────────────────────────────
        elif state == "vcf_to_txt":
            if ext != "vcf":
                await update.message.reply_text("❌ Invalid File Format — send a .vcf file", reply_markup=back_keyboard(fcb))
                return
            context.user_data["pending_result"]     = content.decode("utf-8", errors="ignore")
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = "vcf"
            context.user_data["state"]              = "ask_output_fmt"
            await update.message.reply_text(
                "✅ File processed!\n📤 Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        # ── Rename File ──────────────────────────────────────
        elif state == "rename_file":
            context.user_data["rename_file_bytes"] = content
            context.user_data["rename_file_stem"]  = file_stem(fname)
            context.user_data["state"] = "rename_newname"
            await update.message.reply_text(
                f"📝 File received! (`{fname}`)\nNow send the new filename with extension.\nExample: contacts_new.vcf" + FOOTER,
                reply_markup=back_keyboard(fcb)
            )

        # ── Rename CTC ───────────────────────────────────────
        elif state == "rename_ctc":
            if ext != "vcf":
                await update.message.reply_text("❌ Invalid File Format — send a .vcf file", reply_markup=back_keyboard(fcb))
                return
            context.user_data["rename_ctc_bytes"] = content
            context.user_data["rename_ctc_stem"]  = file_stem(fname)
            context.user_data["state"] = "rename_ctc_prefix"
            await update.message.reply_text(
                f"✏️ VCF received! (`{fname}`)\nNow send the contact name prefix.\nExample: GGOD → GGOD 1, GGOD 2..." + FOOTER,
                reply_markup=back_keyboard(fcb)
            )

        # ── Merge VCF ────────────────────────────────────────
        elif state == "merge_vcf":
            if ext != "vcf":
                await update.message.reply_text("❌ Send .vcf files only. Type /done when finished.", reply_markup=back_keyboard(fcb))
                return
            files = context.user_data.setdefault("merge_vcf_files", [])
            stems = context.user_data.setdefault("merge_vcf_stems", [])
            files.append(content.decode("utf-8", errors="ignore"))
            stems.append(file_stem(fname))
            await update.message.reply_text(
                f"✅ File {len(files)} added! (`{fname}`)\nSend more or press Done to merge." + FOOTER,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("✅ Done — Merge Now", callback_data="merge_vcf_done")],
                    [InlineKeyboardButton("◀️ Cancel", callback_data="back_to_menu")],
                ])
            )

        # ── Merge TXT ────────────────────────────────────────
        elif state == "merge_txt":
            if ext != "txt":
                await update.message.reply_text("❌ Send .txt files only. Type /done when finished.", reply_markup=back_keyboard(fcb))
                return
            files = context.user_data.setdefault("merge_txt_files", [])
            stems = context.user_data.setdefault("merge_txt_stems", [])
            files.append(content.decode("utf-8", errors="ignore"))
            stems.append(file_stem(fname))
            await update.message.reply_text(
                f"✅ File {len(files)} added! (`{fname}`)\nSend more or press Done to merge." + FOOTER,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("✅ Done — Merge Now", callback_data="merge_txt_done")],
                    [InlineKeyboardButton("◀️ Cancel", callback_data="back_to_menu")],
                ])
            )

        # ── Split File ───────────────────────────────────────
        elif state == "split_file":
            if ext not in ("vcf", "txt"):
                await update.message.reply_text("❌ Invalid File Format — send .vcf or .txt", reply_markup=back_keyboard(fcb))
                return
            context.user_data["split_file_data"] = content.decode("utf-8", errors="ignore")
            context.user_data["split_file_ext"]  = ext
            context.user_data["split_file_stem"] = file_stem(fname)
            context.user_data["state"] = "split_count"
            await update.message.reply_text(
                f"✂️ File received! (`{fname}`)\nHow many contacts per file?\nExample: 100" + FOOTER,
                reply_markup=back_keyboard(fcb)
            )

        # ── Admin Navy File ───────────────────────────────────
        elif state == "admin_navy":
            if ext != "vcf":
                await update.message.reply_text("❌ Invalid File Format — send a .vcf file", reply_markup=back_keyboard(fcb))
                return
            context.user_data["pending_result"]     = admin_navy_format(content.decode("utf-8", errors="ignore"))
            context.user_data["pending_stem"]       = file_stem(fname)
            context.user_data["pending_source_ext"] = "vcf"
            context.user_data["state"]              = "ask_output_fmt"
            await update.message.reply_text(
                "✅ File processed!\n📤 Choose output format:" + FOOTER,
                reply_markup=output_fmt_keyboard(fcb)
            )

        # ── No active state ───────────────────────────────────
        else:
            await update.message.reply_text(
                "⚠️ Please select a feature from the menu first." + FOOTER,
                reply_markup=main_menu_keyboard()
            )

    except Exception as e:
        logger.error(f"File handler error: {e}")
        await update.message.reply_text(
            f"❌ Error processing file: {e}" + FOOTER,
            reply_markup=back_keyboard(fcb)
        )


# ─────────────────────────────────────────────────────────────
# ERROR HANDLER
# ─────────────────────────────────────────────────────────────

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    """Log all errors."""
    logger.error(f"Exception while handling update: {context.error}")


# ─────────────────────────────────────────────────────────────
# MAIN — BOT STARTUP
# ─────────────────────────────────────────────────────────────

def main():
    """Start the bot."""
    print("=================================")
    print("  GGOD FATHERR BOT RUNNING...   ")
    print("  PREMIUM SYSTEM ACTIVE         ")
    print("  BOT STATUS: ONLINE            ")
    print("=================================")

    # Build application
    app = Application.builder().token(BOT_TOKEN).build()

    # ── Command handlers ──────────────────────────────────────
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

    # ── Callback query handler (inline buttons) ───────────────
    app.add_handler(CallbackQueryHandler(button_handler))

    # ── Message handlers ──────────────────────────────────────
    app.add_handler(MessageHandler(filters.Document.ALL, file_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))

    # ── Error handler ─────────────────────────────────────────
    app.add_error_handler(error_handler)

    # ── Start polling ─────────────────────────────────────────
    print("Bot is polling for updates...")
    app.run_polling(allowed_updates=["message", "callback_query"])


if __name__ == "__main__":
    main()
