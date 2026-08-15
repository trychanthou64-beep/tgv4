import logging
import sqlite3
import sys
import os
import json
import pandas as pd
import asyncio
from datetime import datetime, timezone, timedelta

# Timezone for Phnom Penh, Cambodia (UTC+7)
PHNOM_PENH_TZ = timezone(timedelta(hours=7))

def get_phnom_penh_time_str():
    return datetime.now(PHNOM_PENH_TZ).strftime("%Y-%m-%d %H:%M:%S")

# Ensure UTF-8 output encoding for Windows consoles
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

from telegram import ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup, Update, BotCommand
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes

# កំណត់កំណត់ហេតុ (Logging)
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# ស្ថានភាពសម្រាប់ Conversation (Admin Update)
EDIT_CHOOSE, EDIT_INPUT = range(2)

# ----------------- PATH SYSTEM -----------------
USER_BASE_DIR = os.environ.get("USER_BASE_DIR")
if USER_BASE_DIR:
    BASE_DIR = USER_BASE_DIR
elif getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if getattr(sys, 'frozen', False):
    RESOURCE_DIR = sys._MEIPASS
else:
    RESOURCE_DIR = BASE_DIR

# ----------------- ការគ្រប់គ្រង DATABASE -----------------
def init_db():
    db_path = os.path.join(BASE_DIR, 'bot_data.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            menu_name TEXT,
            username TEXT,
            password TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            tg_username TEXT,
            tg_name TEXT,
            menu_name TEXT,
            account_username TEXT,
            account_password TEXT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

# ----------------- CONFIGURATION SYSTEM -----------------
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')

def load_config():
    default_config = {
        "TOKEN": "8788148855:AAG3f3N_vCMZ7BEo1hjCcK721zIWRyij4pI",
        "ADMIN_IDS": [8558847170],
        "MENUS": ["F88$", "F88$_CH", "F88_CH_KH", "F88_KH"]
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if "ADMIN_IDS" not in data:
                    data["ADMIN_IDS"] = [8558847170]
                elif 8558847170 not in data["ADMIN_IDS"]:
                    data["ADMIN_IDS"].insert(0, 8558847170)
                return data
        except Exception as e:
            logging.error(f"Error loading config.json: {e}")
    return default_config

config_data = load_config()
TOKEN = os.environ.get("TOKEN", config_data.get("TOKEN", "8788148855:AAG3f3N_vCMZ7BEo1hjCcK721zIWRyij4pI"))
ADMIN_IDS = config_data.get("ADMIN_IDS", [8558847170]) 
MENUS = config_data.get("MENUS", ["F88$", "F88$_CH", "F88_CH_KH", "F88_KH"])
MAIN_MENUS = config_data.get("MAIN_MENUS", {})

# Migrate if needed
if not MAIN_MENUS:
    MAIN_MENUS = {"ទូទៅ": list(MENUS)}

# Sync flat MENUS with MAIN_MENUS
assigned_subs = set()
for main_m, subs in MAIN_MENUS.items():
    assigned_subs.update(subs)
for m in MENUS:
    if m not in assigned_subs:
        first_key = list(MAIN_MENUS.keys())[0]
        MAIN_MENUS[first_key].append(m)
        assigned_subs.add(m)

# Update config_data just in case
config_data["MAIN_MENUS"] = MAIN_MENUS
config_data["MENUS"] = MENUS

def reload_config():
    global config_data, TOKEN, ADMIN_IDS, MENUS, MAIN_MENUS
    config_data = load_config()
    TOKEN = config_data.get("TOKEN", "8788148855:AAG3f3N_vCMZ7BEo1hjCcK721zIWRyij4pI")
    ADMIN_IDS = config_data.get("ADMIN_IDS", [8558847170]) 
    MENUS = config_data.get("MENUS", ["F88$", "F88$_CH", "F88_CH_KH", "F88_KH"])
    MAIN_MENUS = config_data.get("MAIN_MENUS", {})
    if not MAIN_MENUS:
        MAIN_MENUS = {"ទូទៅ": list(MENUS)}
    assigned_subs_local = set()
    for main_m, subs in MAIN_MENUS.items():
        assigned_subs_local.update(subs)
    for m in MENUS:
        if m not in assigned_subs_local:
            first_key = list(MAIN_MENUS.keys())[0]
            MAIN_MENUS[first_key].append(m)
            assigned_subs_local.add(m)
    config_data["MAIN_MENUS"] = MAIN_MENUS
    config_data["MENUS"] = MENUS


EMOJI_LIST = ["📂", "🔑", "💳", "💰", "🎮", "🎲", "🏆", "💎", "🌟", "⚡", "🔥", "🍀"]

# Helper function to chunk lists
def chunk_list(lst, n):
    return [lst[i:i + n] for i in range(0, len(lst), n)] 

# Function សម្រាប់បង្កើត Sample Excel File (Data_Entry, menu_name, និង All Data)
def create_sample_accounts_excel(filepath):
    db_path = os.path.join(BASE_DIR, 'bot_data.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    menu_data = []
    for menu in MENUS:
        try:
            cursor.execute("SELECT COUNT(*) FROM accounts WHERE menu_name = ?", (menu,))
            count = cursor.fetchone()[0]
        except Exception:
            count = 0
        menu_data.append({"menu_name": menu, "Amount": count})

    try:
        cursor.execute("SELECT menu_name, username, password FROM accounts")
        all_rows = cursor.fetchall()
    except Exception:
        all_rows = []
    conn.close()

    all_accounts_data = {
        "menu_name": [r[0] for r in all_rows] if all_rows else [],
        "username": [r[1] for r in all_rows] if all_rows else [],
        "password": [r[2] for r in all_rows] if all_rows else []
    }

    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        # Sheet 1: Data_Entry (Sample Template Data)
        data = {
            "menu_name": ["F88$", "F88$_CH", "F88_CH_KH", "F88_KH"],
            "username": ["user123", "user456", "user789", "user999"],
            "password": ["pass123", "pass456", "pass789", "pass999"]
        }
        df_sample = pd.DataFrame(data)
        df_sample.to_excel(writer, sheet_name="Data_Entry", index=False)

        # Sheet 2: menu_name (Menu summary & counts)
        df_menus = pd.DataFrame(menu_data)
        df_menus = df_menus.sort_values(by="Amount", ascending=True)
        df_menus.to_excel(writer, sheet_name="menu_name", index=False)

        # Sheet 3: All Data (All existing accounts in DB)
        df_all = pd.DataFrame(all_accounts_data)
        df_all.to_excel(writer, sheet_name="All Data", index=False)


# Function សម្រាប់អានគណនីពី Excel រួចបញ្ចូលទៅ Database
def import_accounts_from_excel(filepath, db_path=None):
    if db_path is None:
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
    try:
        try:
            df = pd.read_excel(filepath, sheet_name="Data_Entry")
        except Exception:
            try:
                df = pd.read_excel(filepath, sheet_name="All Data")
            except Exception:
                df = pd.read_excel(filepath, sheet_name=0)
            
        required_cols = {"menu_name", "username", "password"}
        if not required_cols.issubset(df.columns):
            # Check if it has at least menu_name column for menu-only import
            if "menu_name" in df.columns:
                df = df.dropna(subset=["menu_name"])
                if df.empty:
                    return False, "Excel file is empty or has incomplete data", []
                new_menus_found = set()
                for _, row in df.iterrows():
                    menu_name = str(row["menu_name"]).strip()
                    if menu_name:
                        new_menus_found.add(menu_name)
                return True, f"Successfully imported {len(new_menus_found)} menus", list(new_menus_found)
            else:
                return False, "Excel columns must contain at least: menu_name (or menu_name, username, password)", []
        
        df = df.dropna(subset=["menu_name", "username", "password"])
        if df.empty:
            return False, "Excel file is empty or has incomplete data", []
        
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        imported_count = 0
        skipped_count = 0
        new_menus_found = set()
        for _, row in df.iterrows():
            menu_name = str(row["menu_name"]).strip()
            username = str(row["username"]).strip()
            password = str(row["password"]).strip()
            
            # Check if this menu_name and username already exist in the database
            cursor.execute("SELECT 1 FROM accounts WHERE menu_name = ? AND username = ?", (menu_name, username))
            if cursor.fetchone():
                skipped_count += 1
                continue
                
            cursor.execute("INSERT INTO accounts (menu_name, username, password) VALUES (?, ?, ?)", (menu_name, username, password))
            imported_count += 1
            new_menus_found.add(menu_name)
            
        conn.commit()
        conn.close()
        
        msg = f"Successfully imported {imported_count} accounts"
        if skipped_count > 0:
            msg += f" (skipped {skipped_count} duplicates)"
            
        return True, msg, list(new_menus_found)
    except Exception as e:
        return False, str(e), [] 

# Helper function to generate user reply keyboard with updated counts and layout
def get_user_reply_keyboard(user_id, context=None):
    db_path = os.path.join(BASE_DIR, 'bot_data.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    current_main = None
    if context and 'current_main_menu' in context.user_data:
        current_main = context.user_data['current_main_menu']
        
    if current_main and current_main in MAIN_MENUS:
        # User is in a Main Menu, show its Sub Menus
        sub_menus = MAIN_MENUS[current_main]
        menu_buttons = []
        for i, menu in enumerate(sub_menus):
            try:
                cursor.execute("SELECT COUNT(*) FROM accounts WHERE menu_name = ?", (menu,))
                count = cursor.fetchone()[0]
            except Exception:
                count = 0
            emoji = EMOJI_LIST[i % len(EMOJI_LIST)]
            menu_buttons.append(f"{emoji} {menu} ({count})")
        conn.close()
        
        keyboard = chunk_list(menu_buttons, 2)
    else:
        # Show Main Menus
        main_menu_buttons = []
        for i, main_menu in enumerate(MAIN_MENUS.keys()):
            emoji = EMOJI_LIST[i % len(EMOJI_LIST)]
            main_menu_buttons.append(f"{emoji} {main_menu}")
        conn.close()
        
        keyboard = chunk_list(main_menu_buttons, 3)
        if user_id in ADMIN_IDS:
            keyboard.append(['🛠️ Admin Panel'])
            
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

# ----------------- FUNCTIONS សម្រាប់ USER -----------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    # Print to console for easy viewing
    print(f"User ID: {user_id} started the bot.")
    
    # Reset menu level to Main Menus
    context.user_data['current_main_menu'] = None
    
    reply_markup = get_user_reply_keyboard(user_id, context)
    msg = "សូមស្វាគមន៍! សូមជ្រើសរើស អាខោន តែស ខាងក្រោម៖"
    await update.message.reply_text(msg, reply_markup=reply_markup, parse_mode="HTML")

# User clicks on any of the configured menus
async def handle_menu_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    menu_text = update.message.text.strip()

    # Check if the user is in a state waiting for new main menu name input (from /addmainmenu)
    if context.user_data.get('state') == 'WAITING_ADD_MAIN_MENU_NAME':
        context.user_data['state'] = None
        new_main = menu_text
        
        if not new_main:
            await update.message.reply_text("❌ ឈ្មោះ Main Menu ថ្មីមិនអាចទទេឡើយ។")
            return
            
        if new_main in MAIN_MENUS:
            await update.message.reply_text(f"❌ Main Menu <b>{new_main}</b> មានរួចហើយ។", parse_mode="HTML")
            return
            
        MAIN_MENUS[new_main] = []
        config_data["MAIN_MENUS"] = MAIN_MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            await update.message.reply_text(f"✅ បានបន្ថែម Main Menu <b>{new_main}</b> ដោយជោគជ័យ! សូមចុច <code>/start</code> សារជាថ្មីដើម្បីមើលប៊ូតុង។", parse_mode="HTML")
        except Exception as e:
            logging.error(f"Error saving config.json on addmainmenu text input: {e}")
            await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")
        return

    # Check if the user is in a state waiting for new menu name input (from /addmenu)
    if context.user_data.get('state') == 'WAITING_ADD_MENU_NAME':
        context.user_data['state'] = None
        new_menu = menu_text
        
        if not new_menu:
            await update.message.reply_text("❌ ឈ្មោះ Menu ថ្មីមិនអាចទទេឡើយ។")
            return
            
        if new_menu in MENUS:
            await update.message.reply_text(f"❌ Menu <b>{new_menu}</b> មានរួចហើយ។", parse_mode="HTML")
            return
            
        target_main = context.user_data.get('selected_main_for_sub')
        if not target_main or target_main not in MAIN_MENUS:
            if not MAIN_MENUS:
                MAIN_MENUS["ទូទៅ"] = []
            target_main = list(MAIN_MENUS.keys())[0]
            
        MENUS.append(new_menu)
        MAIN_MENUS[target_main].append(new_menu)
        
        config_data["MENUS"] = MENUS
        config_data["MAIN_MENUS"] = MAIN_MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            
            context.user_data.pop('selected_main_for_sub', None)
            await update.message.reply_text(f"✅ បានបន្ថែម Menu <b>{new_menu}</b> ទៅក្នុង Main Menu <b>{target_main}</b> ដោយជោគជ័យ! សូមចុច <code>/start</code> សារជាថ្មីដើម្បីមើលប៊ូតុង។", parse_mode="HTML")
        except Exception as e:
            logging.error(f"Error saving config.json on addmenu text input: {e}")
            await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")
        return

    # Check if the user is in a state waiting for new main menu name input (from /editmainmenu)
    if context.user_data.get('state') == 'WAITING_EDIT_MAIN_MENU_NAME':
        context.user_data['state'] = None
        target_main = context.user_data.get('edit_main_menu_target')
        new_name = menu_text
        
        if not new_name:
            await update.message.reply_text("❌ ឈ្មោះ Main Menu ថ្មីមិនអាចទទេឡើយ។")
            return
            
        if new_name in MAIN_MENUS:
            await update.message.reply_text(f"❌ Main Menu <b>{new_name}</b> មានរួចហើយ។", parse_mode="HTML")
            return
            
        if target_main not in MAIN_MENUS:
            await update.message.reply_text(f"❌ មិនរកឃើញ Main Menu <b>{target_main}</b> ឡើយ។", parse_mode="HTML")
            return
            
        MAIN_MENUS[new_name] = MAIN_MENUS.pop(target_main)
        config_data["MAIN_MENUS"] = MAIN_MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            await update.message.reply_text(f"📝 បានកែប្រែ Main Menu ពី <b>{target_main}</b> ទៅជា <b>{new_name}</b> រួចរាល់!", parse_mode="HTML")
        except Exception as e:
            logging.error(f"Error renaming main menu {target_main} to {new_name}: {e}")
            await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")
        return

    # Check if the user is in a state waiting for new menu name input (from /editmenu)
    if context.user_data.get('state') == 'WAITING_EDIT_MENU_NAME':
        # Reset the state
        context.user_data['state'] = None
        target_menu = context.user_data.get('edit_menu_target')
        new_name = menu_text
        
        if not new_name:
            await update.message.reply_text("❌ ឈ្មោះ Menu ថ្មីមិនអាចទទេឡើយ។")
            return
            
        if new_name in MENUS:
            await update.message.reply_text(f"❌ Menu <b>{new_name}</b> មានរួចហើយ។", parse_mode="HTML")
            return
            
        if target_menu not in MENUS:
            await update.message.reply_text(f"❌ មិនរកឃើញ Menu <b>{target_menu}</b> ឡើយ។", parse_mode="HTML")
            return
            
        # Update dynamic list in memory
        idx = MENUS.index(target_menu)
        MENUS[idx] = new_name

        # Update in MAIN_MENUS as well
        for main_name, subs in MAIN_MENUS.items():
            if target_menu in subs:
                s_idx = subs.index(target_menu)
                subs[s_idx] = new_name

        # Save to config.json
        config_data["MENUS"] = MENUS
        config_data["MAIN_MENUS"] = MAIN_MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
                
            # Update database table accounts for renamed menu
            db_path = os.path.join(BASE_DIR, 'bot_data.db')
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("UPDATE accounts SET menu_name = ? WHERE menu_name = ?", (new_name, target_menu))
            conn.commit()
            conn.close()
            
            await update.message.reply_text(f"📝 បានកែប្រែ Menu ពី <b>{target_menu}</b> ទៅជា <b>{new_name}</b> រួចរាល់! គណនីក្នុង Database ក៏ត្រូវបានធ្វើបច្ចុប្បន្នភាពផងដែរ។", parse_mode="HTML")
        except Exception as e:
            logging.error(f"Error updating menu {target_menu} to {new_name}: {e}")
            await update.message.reply_text("❌ 有បញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")
        return

    # Check if user clicked Back or Refresh
    if menu_text in ['⬅️ ត្រឡប់ក្រោយ (Back)', '⬅️ ត្រលប់ក្រោយ (Back)', '⬅️ ត្រឡប់ក្រោយ', '⬅️ ត្រលប់ក្រោយ', '🔄 Refresh', '⬅️ Back', 'Back', 'ត្រឡប់ក្រោយ', 'ត្រលប់ក្រោយ']:
        context.user_data['current_main_menu'] = None
        reply_markup = get_user_reply_keyboard(user_id, context)
        await update.message.reply_text("សូមជ្រើសរើស អាខោន តែស ខាងក្រោម៖", reply_markup=reply_markup)
        return

    # Check if it matches a Main Menu name
    matched_main = None
    if not context.user_data.get('current_main_menu'):
        for main_menu in MAIN_MENUS.keys():
            if main_menu in menu_text:
                matched_main = main_menu
                break
            
    if matched_main:
        context.user_data['current_main_menu'] = matched_main
        reply_markup = get_user_reply_keyboard(user_id, context)
        await update.message.reply_text(f"📁 សូមជ្រើសរើសគណនី <b>{matched_main}</b> ខាងក្រោម៖", reply_markup=reply_markup, parse_mode="HTML")
        return

    # Extract menu name by matching the longest configured menu name in the text
    actual_menu_name = None
    for menu in sorted(MENUS, key=len, reverse=True):
        if menu in menu_text:
            actual_menu_name = menu
            break
            
    if not actual_menu_name:
        await start(update, context)
        return
    db_path = os.path.join(BASE_DIR, 'bot_data.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id, username, password FROM accounts WHERE menu_name = ? ORDER BY id ASC LIMIT 1", (actual_menu_name,))
        row = cursor.fetchone()
        
        if row:
            account_id, acc_user, acc_pass = row
            # Delete from Database immediately (One-time use)
            cursor.execute("DELETE FROM accounts WHERE id = ?", (account_id,))
            
            # Log to download history
            user = update.effective_user
            user_id = user.id
            tg_username = user.username or ""
            tg_name = f"{user.first_name or ''} {user.last_name or ''}".strip()
            now_phnom_penh = get_phnom_penh_time_str()
            
            cursor.execute("""
                INSERT INTO history (user_id, tg_username, tg_name, menu_name, account_username, account_password, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (user_id, tg_username, tg_name, actual_menu_name, acc_user, acc_pass, now_phnom_penh))
            conn.commit()
            
            # Send greeting/title first as Message 1
            msg1 = f"នេះជាគណនី {actual_menu_name} របស់អ្នក៖"
            await update.message.reply_text(msg1)
            
            # Message 2 contains credentials in plain normal text without labels or emojis
            response = f"{acc_user}\n{acc_pass}"
            
            # Re-generate menu buttons with updated counts and refresh keyboard
            reply_markup = get_user_reply_keyboard(user_id, context)
            await update.message.reply_text(response, reply_markup=reply_markup)
        else:
            await update.message.reply_text(f"❌ សុំទោស! គណនីក្នុង Menu <b>{actual_menu_name}</b> អស់ហើយ។ សូមរង់ចាំ Admin បញ្ចូលបន្ថែម។", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error getting account for {actual_menu_name}: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការទាញយកទិន្នន័យ។")
    finally:
        conn.close()

# ----------------- FUNCTIONS សម្រាប់ ADMIN (CRUD) -----------------

def get_admin_list_text():
    db_path = os.path.join(BASE_DIR, 'bot_data.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    lines = []
    for a in ADMIN_IDS:
        role = "Default Admin" if a == 8558847170 else "Admin"
        try:
            cursor.execute("SELECT tg_username, tg_name FROM history WHERE user_id = ? ORDER BY id DESC LIMIT 1", (a,))
            row = cursor.fetchone()
            if row:
                username, name = row
                details = []
                if name:
                    details.append(name)
                if username:
                    details.append(f"@{username}")
                info = f" ({', '.join(details)})" if details else ""
                lines.append(f"• <code>{a}</code>{info} - តួនាទី៖ <b>{role}</b>")
            else:
                lines.append(f"• <code>{a}</code> - តួនាទី៖ <b>{role}</b>")
        except Exception:
            lines.append(f"• <code>{a}</code> - តួនាទី៖ <b>{role}</b>")
    conn.close()
    return "\n".join(lines)

def make_inline_markup(keyboard):
    return InlineKeyboardMarkup(keyboard) if keyboard else None

# Helper to generate the main admin keyboard markup
def get_main_admin_keyboard():
    return [
        [InlineKeyboardButton("➕ បញ្ចូលគណនីថ្មី (Create)", callback_data="admin_add_help")],
        [InlineKeyboardButton("🗑️ លុបគណនីទាំងអស់ (Delete All)", callback_data="admin_delete_all_confirm")],
        [InlineKeyboardButton("⚙️ គ្រប់គ្រង Menu (Menus)", callback_data="admin_menus")],
        [InlineKeyboardButton("👤 គ្រប់គ្រង Admin (Admins)", callback_data="admin_admins")],
        [InlineKeyboardButton("📜 ប្រវត្តិទាញយក (History)", callback_data="admin_history")],
        [InlineKeyboardButton("📥 បម្រុងទុក (Backup)", callback_data="admin_backup")],
        [InlineKeyboardButton("📤 ស្តារឡើងវិញ (Restore)", callback_data="admin_restore")]
    ]

# បង្ហាញផ្ទាំងបញ្ជា Admin Panel
async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    keyboard = get_main_admin_keyboard()
    reply_markup = make_inline_markup(keyboard)
    await update.message.reply_text("🛠️ **ផ្ទាំងគ្រប់គ្រងសម្រាប់ Admin (CRUD Panel):**", reply_markup=reply_markup, parse_mode="Markdown")

# បង្ហាញផ្ទាំងគ្រប់គ្រង Menu
# បង្ហាញផ្ទាំងគ្រប់គ្រង Menu ធំ (Main Menus & Sub Menus configuration panel)
async def show_admin_menus_panel(query, success_msg=None):
    keyboard = [
        [InlineKeyboardButton("📁 គ្រប់គ្រង Main Menu", callback_data="admin_main_menus")],
        [InlineKeyboardButton("📂 គ្រប់គ្រង Sub Menu", callback_data="admin_sub_menus")],
        [InlineKeyboardButton("🔄 បំលាស់ទី Sub Menu (Move Menu)", callback_data="admin_move_menu_select_sub")],
        [InlineKeyboardButton("🎨 តម្រៀបដោយខ្លួនឯង (Sort Manual)", callback_data="admin_sort_manual")],
        [InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_back")]
    ]
    header = "⚙️ <b>គ្រប់គ្រង Menu ទាំងអស់ក្នុងប្រព័ន្ធ (Menus Config):</b>\n\n"
    if success_msg:
        header += f"✅ <b>{success_msg}</b>\n\n"
        
    main_count = len(MAIN_MENUS)
    sub_count = len(MENUS)
    
    await query.edit_message_text(
        f"{header}"
        f"📋 <b>ស្ថិតិបច្ចុប្បន្ន៖</b>\n"
        f"- ចំនួន Main Menu: <b>{main_count}</b>\n"
        f"- ចំនួន Sub Menu: <b>{sub_count}</b>\n\n"
        f"សូមជ្រើសរើសផ្នែកខាងក្រោម ដើម្បីគ្រប់គ្រង៖",
        reply_markup=make_inline_markup(keyboard),
        parse_mode="HTML"
    )

async def show_admin_main_menus_panel(query, success_msg=None):
    lines = []
    for i, (main_name, subs) in enumerate(MAIN_MENUS.items(), 1):
        lines.append(f"{i}. <b>{main_name}</b> (មាន Sub Menu ចំនួន <code>{len(subs)}</code>)")
    main_list_str = "\n".join(lines) if lines else "<i>(មិនទាន់មាន Main Menu នៅឡើយទេ)</i>"

    keyboard = [
        [InlineKeyboardButton("➕ បន្ថែម Main Menu", callback_data="admin_add_main_menu")],
        [InlineKeyboardButton("📝 កែឈ្មោះ Main Menu", callback_data="admin_edit_main_menu_list")],
        [InlineKeyboardButton("🗑️ លុប Main Menu", callback_data="admin_del_main_menu_list")],
        [InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_menus")]
    ]

    header = "📁 <b>គ្រប់គ្រង Main Menu ទាំងអស់ក្នុងប្រព័ន្ធ៖</b>\n\n"
    if success_msg:
        header += f"✅ <b>{success_msg}</b>\n\n"

    await query.edit_message_text(
        f"{header}"
        f"📋 <b>បញ្ជី Main Menu បច្ចុប្បន្ន៖</b>\n{main_list_str}\n\n"
        f"💬 <b>របៀបគ្រប់គ្រងតាម Command:</b>\n"
        f"- បន្ថែម Main Menu ថ្មី៖\n"
        f"<code>/addmainmenu [ឈ្មោះMain]</code>\n"
        f"- កែប្រែឈ្មោះ Main Menu៖\n"
        f"<code>/editmainmenu [ឈ្មោះចាស់] [ឈ្មោះថ្មី]</code>\n"
        f"- លុប Main Menu ចោល៖\n"
        f"<code>/delmainmenu [ឈ្មោះMain]</code>",
        reply_markup=make_inline_markup(keyboard),
        parse_mode="HTML"
    )

async def show_admin_sub_menus_panel(query, success_msg=None):
    lines = []
    for main_name, subs in MAIN_MENUS.items():
        lines.append(f"📁 <b>{main_name}</b>:")
        if not subs:
            lines.append("  <i>(គ្មាន Sub Menu)</i>")
        for sub in subs:
            db_path = os.path.join(BASE_DIR, 'bot_data.db')
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            try:
                cursor.execute("SELECT COUNT(*) FROM accounts WHERE menu_name = ?", (sub,))
                count = cursor.fetchone()[0]
            except Exception:
                count = 0
            finally:
                conn.close()
            lines.append(f"  ▪️ <b>{sub}</b> (មានគណនី៖ <code>{count}</code>)")
    sub_list_str = "\n".join(lines) if lines else "<i>(មិនទាន់មាន Sub Menu នៅឡើយទេ)</i>"

    keyboard = [
        [InlineKeyboardButton("➕ បន្ថែម Sub Menu", callback_data="admin_add_sub_menu")],
        [InlineKeyboardButton("📝 កែឈ្មោះ Sub Menu", callback_data="admin_edit_sub_menu_list")],
        [InlineKeyboardButton("🗑️ លុប Sub Menu", callback_data="admin_del_sub_menu_list")],
        [InlineKeyboardButton("📥 ទាញយកគំរូ Excel (Sub Menus)", callback_data="admin_download_menus_sample")],
        [InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_menus")]
    ]

    header = "📂 <b>គ្រប់គ្រង Sub Menu ទាំងអស់ក្នុងប្រព័ន្ធ៖</b>\n\n"
    if success_msg:
        header += f"✅ <b>{success_msg}</b>\n\n"

    await query.edit_message_text(
        f"{header}"
        f"📋 <b>បញ្ជី Sub Menu បច្ចុប្បន្ន៖</b>\n{sub_list_str}\n\n"
        f"💬 <b>របៀបគ្រប់គ្រងតាម Command:</b>\n"
        f"- បន្ថែម Sub Menu ថ្មី៖\n"
        f"<code>/addmenu [ឈ្មោះSub]</code>\n"
        f"- កែប្រែឈ្មោះ Sub Menu៖\n"
        f"<code>/editmenu [ឈ្មោះចាស់] [ឈ្មោះថ្មី]</code>\n"
        f"- លុប Sub Menu ចោល៖\n"
        f"<code>/delmenu [ឈ្មោះSub]</code>",
        reply_markup=make_inline_markup(keyboard),
        parse_mode="HTML"
    )

#  កាត់ផ្ទាំងបញ្ជា Inline Button Clicks
async def button_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    
    # Check if the callback is for normal user account confirmation/cancellation
    is_user_callback = query.data.startswith("user_get_acc_") or query.data == "user_get_acc_cancel"
    if user_id not in ADMIN_IDS and not is_user_callback: 
        return

    # Clear state when admin navigates through Admin Panel buttons
    if not is_user_callback:
        context.user_data['state'] = None

    if query.data == "user_get_acc_cancel":
        await query.delete_message()
        return

    elif query.data.startswith("user_get_acc_"):
        actual_menu_name = query.data.replace("user_get_acc_", "")
        
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, username, password FROM accounts WHERE menu_name = ? ORDER BY id ASC LIMIT 1", (actual_menu_name,))
        row = cursor.fetchone()
        
        if row:
            account_id, acc_user, acc_pass = row
            # Delete from Database immediately (One-time use)
            cursor.execute("DELETE FROM accounts WHERE id = ?", (account_id,))
            
            # Log to download history
            user = query.from_user
            user_id = user.id
            tg_username = user.username or ""
            tg_name = f"{user.first_name or ''} {user.last_name or ''}".strip()
            now_phnom_penh = get_phnom_penh_time_str()
            
            cursor.execute("""
                INSERT INTO history (user_id, tg_username, tg_name, menu_name, account_username, account_password, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (user_id, tg_username, tg_name, actual_menu_name, acc_user, acc_pass, now_phnom_penh))
            conn.commit()
            conn.close()
            
            # Delete the confirmation message completely
            await query.delete_message()
            
            # Send greeting/title first as Message 1
            msg1 = f"នេះជាគណនី {actual_menu_name} របស់អ្នក៖"
            await context.bot.send_message(chat_id=query.message.chat_id, text=msg1)
            
            # Message 2 contains credentials in plain normal text without labels or emojis
            response = f"{acc_user}\n{acc_pass}"
            
            # Re-generate menu buttons with updated counts and refresh keyboard
            reply_markup = get_user_reply_keyboard(user_id, context)
            await context.bot.send_message(chat_id=query.message.chat_id, text=response, reply_markup=reply_markup)
        else:
            conn.close()
            await query.edit_message_text(f"❌ សុំទោស! គណនីក្នុង Menu <b>{actual_menu_name}</b> អស់ហើយ។", parse_mode="HTML")
        return

    elif query.data == "admin_close_panel":
        try:
            await query.delete_message()
        except Exception:
            await query.edit_message_text("ផ្ទាំងគ្រប់គ្រងត្រូវបានបិទ។")
        return

    elif query.data.startswith("admin_refresh_"):
        # Reload configuration dynamically
        reload_config()
        refresh_target = query.data.replace("admin_refresh_", "")
        
        if refresh_target == "main":
            keyboard = get_main_admin_keyboard()
            reply_markup = make_inline_markup(keyboard)
            await query.edit_message_text("🛠️ **ផ្ទាំងគ្រប់គ្រងសម្រាប់ Admin (CRUD Panel):**\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!", reply_markup=reply_markup, parse_mode="Markdown")
            
        elif refresh_target == "admins":
            admin_list = get_admin_list_text()
            keyboard = []
            await query.edit_message_text(
                f"👤 <b>គ្រប់គ្រង Admin ទាំងអស់ក្នុងប្រព័ន្ធ (Admins Config)៖</b>\n\n"
                f"📋 <b>បញ្ជី ID Admin និងតួនាទី៖</b>\n{admin_list}\n\n"
                f"➕ <b>របៀបបន្ថែម Admin៖</b>\n"
                f"សូមវាយផ្ញើសារជាអក្សរធម្មតាមកកាន់ Bot តាមទម្រង់៖\n"
                f"<code>/addadmin [លេខID]</code>\n"
                f"<i>(ឧទាហរណ៍៖ <code>/addadmin 123456789</code>)</i>\n\n"
                f"🗑️ <b>របៀបលុប Admin៖</b>\n"
                f"<code>/deladmin [លេខID]</code>\n"
                f"<i>(ឧទាករណ៍៖ <code>/deladmin 123456789</code>)</i>\n\n"
                f"🔄 បានធ្វើឲ្យស្រស់រួចរាល់!",
                reply_markup=make_inline_markup(keyboard),
                parse_mode="HTML"
            )
            
        elif refresh_target == "backup":
            db_path = os.path.join(BASE_DIR, 'bot_data.db')
            config_path = os.path.join(BASE_DIR, 'config.json')
            
            status_msg = await query.message.reply_text("⏳ កំពុងរៀបចំឯកសារបម្រុងទុក (Backup)...")
            
            backup_files_sent = 0
            if os.path.exists(db_path):
                try:
                    with open(db_path, 'rb') as f:
                        await context.bot.send_document(
                            chat_id=query.message.chat_id,
                            document=f,
                            filename="bot_data.db",
                            caption="📦 ឯកសារបម្រុងទុក Database (SQLite Database Backup)"
                        )
                    backup_files_sent += 1
                except Exception as e:
                    logging.error(f"Error backup db: {e}")
                    await context.bot.send_message(chat_id=query.message.chat_id, text="❌ មិនអាចផ្ញើឯកសារ Database បានទេ។")
                    
                try:
                    conn = sqlite3.connect(db_path)
                    cursor = conn.cursor()
                    cursor.execute("SELECT menu_name, username, password FROM accounts")
                    rows = cursor.fetchall()
                    conn.close()
                    
                    excel_filepath = os.path.join(BASE_DIR, "backup_accounts.xlsx")
                    data = {
                        "menu_name": [r[0] for r in rows] if rows else [],
                        "username": [r[1] for r in rows] if rows else [],
                        "password": [r[2] for r in rows] if rows else []
                    }
                    df = pd.DataFrame(data)
                    with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name="Data_Entry", index=False)
                        
                    with open(excel_filepath, 'rb') as f:
                        await context.bot.send_document(
                            chat_id=query.message.chat_id,
                            document=f,
                            filename="backup_accounts.xlsx",
                            caption=f"📊 ឯកសារបម្រុងទុកគណនីជា Excel (Excel Accounts Backup - សរុប៖ {len(rows) if rows else 0} អាខោន)"
                        )
                    backup_files_sent += 1
                    if os.path.exists(excel_filepath):
                        os.remove(excel_filepath)
                except Exception as e:
                    logging.error(f"Error backup accounts excel: {e}")
                    await context.bot.send_message(chat_id=query.message.chat_id, text="❌ មិនអាចផ្ញើឯកសារបម្រុងទុកជា Excel បានទេ។")
                    
            if os.path.exists(config_path):
                try:
                    with open(config_path, 'rb') as f:
                        await context.bot.send_document(
                            chat_id=query.message.chat_id,
                            document=f,
                            filename="config.json",
                            caption="⚙️ ឯកសារបម្រុងទុកការកំណត់ (Config Backup)"
                        )
                    backup_files_sent += 1
                except Exception as e:
                    logging.error(f"Error backup config: {e}")
                    await context.bot.send_message(chat_id=query.message.chat_id, text="❌ មិនអាចផ្ញើឯកសារ Config បានទេ។")
                    
            keyboard = []
            reply_markup = make_inline_markup(keyboard)
            if backup_files_sent > 0:
                await status_msg.edit_text("✅ ការបម្រុងទុកទិន្នន័យត្រូវបានបញ្ចប់ដោយជោគជ័យ!\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!", reply_markup=reply_markup)
            else:
                await status_msg.edit_text("❌ គ្មានឯកសារត្រូវបានរកឃើញសម្រាប់បម្រុងទុកឡើយ。\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!", reply_markup=reply_markup)
            
        elif refresh_target == "view":
            db_path = os.path.join(BASE_DIR, 'bot_data.db')
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT menu_name, username, password FROM accounts")
            rows = cursor.fetchall()
            conn.close()
            keyboard = []
            reply_markup = make_inline_markup(keyboard)
            if not rows:
                await query.edit_message_text("📊 មិនទាន់មានទិន្នន័យនៅក្នុង Database ឡើយ。\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!", reply_markup=reply_markup)
            else:
                await query.edit_message_text("⏳ កំពុងបង្កើត និងនាំចេញឯកសារ Excel...", reply_markup=reply_markup)
                filepath = os.path.join(BASE_DIR, "exported_accounts.xlsx")
                try:
                    data = {
                        "menu_name": [r[0] for r in rows],
                        "username": [r[1] for r in rows],
                        "password": [r[2] for r in rows]
                    }
                    df = pd.DataFrame(data)
                    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name="Data_Entry", index=False)
                    
                    with open(filepath, 'rb') as f:
                        await context.bot.send_document(
                            chat_id=query.message.chat_id,
                            document=f,
                            filename="exported_accounts.xlsx",
                            caption=f"📊 **គណនីទាំងអស់ក្នុងប្រព័ន្ធ (សរុប៖ {len(rows)} អាខោន)**"
                        )
                    if os.path.exists(filepath):
                        os.remove(filepath)
                except Exception as e:
                    logging.error(f"Error exporting accounts to Excel: {e}")
                    await context.bot.send_message(
                        chat_id=query.message.chat_id,
                        text="❌ មានបញ្ហាក្នុងការនាំចេញទិន្នន័យជា Excel。"
                    )
                    if os.path.exists(filepath):
                        os.remove(filepath)
                        
        elif refresh_target == "history":
            keyboard = [
                [InlineKeyboardButton("📊 ទាញយកជា Excel", callback_data="admin_history_excel")],
                [InlineKeyboardButton("💬 មើលលើ Telegram", callback_data="admin_history_chat")]
            ]
            reply_markup = make_inline_markup(keyboard)
            await query.edit_message_text(
                "📜 **ប្រវត្តិទាញយក (Download History):**\n\n"
                "សូមជ្រើសរើសរបៀបមើលប្រវត្តិទាញយកខាងក្រោម៖\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!",
                reply_markup=reply_markup,
                parse_mode="Markdown"
            )
            
        elif refresh_target == "history_excel":
            db_path = os.path.join(BASE_DIR, 'bot_data.db')
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT id, user_id, tg_username, tg_name, menu_name, account_username, account_password, timestamp FROM history ORDER BY id DESC")
            rows = cursor.fetchall()
            conn.close()
            keyboard = []
            reply_markup = make_inline_markup(keyboard)
            if not rows:
                await query.edit_message_text("📜 មិនទាន់មានប្រវត្តិទាញយកគណនីនៅក្នុង Database ឡើយ。\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!", reply_markup=reply_markup)
            else:
                await query.edit_message_text("⏳ កំពុងបង្កើត និងនាំចេញឯកសារប្រវត្តិទាញយក (History)...", reply_markup=reply_markup)
                filepath = os.path.join(BASE_DIR, "download_history.xlsx")
                try:
                    data = {
                        "ID": [r[0] for r in rows],
                        "User_ID": [r[1] for r in rows],
                        "TG_Username": [r[2] for r in rows],
                        "TG_Name": [r[3] for r in rows],
                        "Menu_Name": [r[4] for r in rows],
                        "Account_Username": [r[5] for r in rows],
                        "Account_Password": [r[6] for r in rows],
                        "Timestamp": [r[7] for r in rows]
                    }
                    df = pd.DataFrame(data)
                    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name="Download_History", index=False)
                    with open(filepath, 'rb') as f:
                        await context.bot.send_document(
                            chat_id=query.message.chat_id,
                            document=f,
                            filename="download_history.xlsx",
                            caption=f"📜 **ប្រវត្តិអ្នកប្រើប្រាស់ទាញយកគណនី (សរុប៖ {len(rows)} ដង)**"
                        )
                    if os.path.exists(filepath):
                        os.remove(filepath)
                except Exception as e:
                    logging.error(f"Error exporting history to Excel: {e}")
                    await context.bot.send_message(
                        chat_id=query.message.chat_id,
                        text="❌ មានបញ្ហាក្នុងការនាំចេញប្រវត្តិទាញយកជា Excel。"
                    )
                    if os.path.exists(filepath):
                        os.remove(filepath)

        elif refresh_target == "history_chat":
            db_path = os.path.join(BASE_DIR, 'bot_data.db')
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT id, tg_name, menu_name, account_username, timestamp FROM history ORDER BY id DESC LIMIT 15")
            rows = cursor.fetchall()
            conn.close()
            keyboard = []
            reply_markup = make_inline_markup(keyboard)
            if not rows:
                await query.edit_message_text("📜 មិនទាន់មានប្រវត្តិទាញយកគណនីនៅក្នុង Database ឡើយ。\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!", reply_markup=reply_markup)
            else:
                txt = "📜 **ប្រវត្តិទាញយកគណនីថ្មីៗបំផុតចំនួន ១៥៖**\n\n"
                for r in rows:
                    txt += f"🆔 ID: `{r[0]}` | 👤 Name: *{r[1]}*\n📂 Menu: *{r[2]}* | 🔑 U: `{r[3]}`\n📅 ម៉ោង៖ `{r[4]}`\n------------------------\n"
                txt += "\n💡 *ចំណាំ៖* ដើម្បីទាញយកប្រវត្តិទាំងអស់ សូមប្រើប៊ូតុង 'ទាញយកជា Excel'。\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!"
                await query.edit_message_text(txt, reply_markup=reply_markup, parse_mode="Markdown")

        elif refresh_target == "add_help":
            keyboard = [
                [InlineKeyboardButton("📥 ទាញយកគំរូ Excel", callback_data="admin_download_sample")]
            ]
            reply_markup = make_inline_markup(keyboard)
            await query.edit_message_text(
                "➕ <b>របៀបបញ្ចូលទិន្នន័យ (Create):</b>\n\n"
                "<b>វិធីទី១៖ បញ្ចូលម្តងមួយៗ (Single Import)</b>\n"
                "សូមផ្ញើសារជាអក្សរធម្មតាមកកាន់ Bot តាមទម្រង់៖\n"
                "<code>/add [ឈ្មោះMenu] [Username] [Password]</code>\n"
                "<i>(ឧទាហរណ៍៖ <code>/add F88 user777 pass999</code>)</i>\n\n"
                "<b>វិធីទី២៖ បញ្ចូលជាគំនរតាមរយៈ Excel (Batch Import)</b>\n"
                "១. ចុចប៊ូតុងខាងក្រោម ឬផ្ញើសារ <code>/sample</code> ដើម្បីទាញយកឯកសារគំរូ Excel\n"
                "២. បំពេញទិន្នន័យគណនីចូលក្នុង Excel នោះរួចរក្សាទុក\n"
                "៣. ផ្ញើឯកសារ Excel (.xlsx) នោះមកកាន់ Bot នេះជាការស្រេច៖\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!",
                reply_markup=reply_markup,
                parse_mode="HTML"
            )

        elif refresh_target == "update_list":
            keyboard = []
            reply_markup = make_inline_markup(keyboard)
            await query.edit_message_text(
                "📝 **របៀបកែប្រែទិន្នន័យ (Update):**\n\n"
                "សូមប្រើប្រាស់ Command នេះដើម្បីកែប្រែតាមរយៈ ID គណនី៖\n"
                "`/edit [ID] [Usernameថ្មី] [Passwordថ្មី]`\n\n"
                "*(អ្នកអាចមើល ID បានដោយចុចប៊ូតុង 'មើលទិន្នន័យទាំងអស់')*\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!",
                reply_markup=reply_markup,
                parse_mode="Markdown"
            )

        elif refresh_target == "delete_all_confirm":
            keyboard = [
                [InlineKeyboardButton("✅ ព្រមលុបទាំងអស់", callback_data="admin_delete_all_exec")]
            ]
            await query.edit_message_text(
                "⚠️ **តើអ្នកពិតជាចង់លុបទិន្នន័យគណនីទាំងអស់ចេញពី Database មែនទេ?** Action នេះមិនអាចស្តារឡើងវិញបានឡើយ!\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!",
                reply_markup=make_inline_markup(keyboard),
                parse_mode="Markdown"
            )

        elif refresh_target == "delete_all_exec":
            keyboard = []
            await query.edit_message_text("🗑️ បានលុបទិន្នន័យទាំងអស់ចេញពី Database រួចរាល់!\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!", reply_markup=make_inline_markup(keyboard))

        elif refresh_target == "menus":
            await show_admin_menus_panel(query, "បានធ្វើឲ្យស្រស់រួចរាល់!")

        elif refresh_target == "main_menus":
            await show_admin_main_menus_panel(query, "បានធ្វើឲ្យស្រស់រួចរាល់!")

        elif refresh_target == "sub_menus":
            await show_admin_sub_menus_panel(query, "បានធ្វើឲ្យស្រស់រួចរាល់!")

        elif refresh_target == "edit_main_menu_list":
            keyboard = []
            for i, main_name in enumerate(MAIN_MENUS.keys()):
                keyboard.append([InlineKeyboardButton(f"📝 កែ {main_name}", callback_data=f"editmain_sel_{i}")])
            keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_main_menus")])
            await query.edit_message_text(
                "📝 <b>សូមជ្រើសរើស Main Menu ដែលអ្នកចង់កែប្រែឈ្មោះ៖</b>\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!",
                reply_markup=make_inline_markup(keyboard),
                parse_mode="HTML"
            )

        elif refresh_target == "del_main_menu_list":
            keyboard = []
            for i, main_name in enumerate(MAIN_MENUS.keys()):
                keyboard.append([InlineKeyboardButton(f"🗑️ លុប {main_name}", callback_data=f"delmain_sel_{i}")])
            keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_main_menus")])
            await query.edit_message_text(
                "🗑️ <b>សូមជ្រើសរើស Main Menu ដែលអ្នកចង់លុប៖</b>\n\n"
                "⚠️ <b>បញ្ជាក់៖</b> មិនអាចលុប Main Menu ដែលមិនទាន់ទទេ (មាន Sub Menus) បានឡើយ。\n\n🔄 បានធ្វើឲ្យស្រស់រួចរាល់!",
                reply_markup=make_inline_markup(keyboard),
                parse_mode="HTML"
            )
        return

    elif query.data == "admin_history":
        keyboard = [
            [InlineKeyboardButton("📊 ទាញយកជា Excel", callback_data="admin_history_excel")],
            [InlineKeyboardButton("💬 មើលលើ Telegram", callback_data="admin_history_chat")],
            [InlineKeyboardButton("🗑️ លុបប្រវត្តិទាំងអស់ (Clear History)", callback_data="admin_history_clear_confirm")]
        ]
        reply_markup = make_inline_markup(keyboard)
        await query.edit_message_text(
            "📜 **ប្រវត្តិទាញយក (Download History):**\n\n"
            "សូមជ្រើសរើសរបៀបមើល ឬលុបប្រវត្តិទាញយកខាងក្រោម៖",
            reply_markup=reply_markup,
            parse_mode="Markdown"
        )
 
    elif query.data == "admin_history_excel":
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, user_id, tg_username, tg_name, menu_name, account_username, account_password, timestamp FROM history ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()
        keyboard = []
        reply_markup = make_inline_markup(keyboard)
        if not rows:
            await query.edit_message_text("📜 មិនទាន់មានប្រវត្តិទាញយកគណនីនៅក្នុង Database ឡើយ。", reply_markup=reply_markup)
        else:
            await query.edit_message_text("⏳ កំពុងបង្កើត និងនាំចេញឯកសារប្រវត្តិទាញយក (History)...", reply_markup=reply_markup)
            
            filepath = os.path.join(BASE_DIR, "download_history.xlsx")
            try:
                data = {
                    "ID": [r[0] for r in rows],
                    "User_ID": [r[1] for r in rows],
                    "TG_Username": [r[2] for r in rows],
                    "TG_Name": [r[3] for r in rows],
                    "Menu_Name": [r[4] for r in rows],
                    "Account_Username": [r[5] for r in rows],
                    "Account_Password": [r[6] for r in rows],
                    "Timestamp": [r[7] for r in rows]
                }
                df = pd.DataFrame(data)
                
                # Write to Excel
                with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name="Download_History", index=False)
                
                # Send the Excel file
                with open(filepath, 'rb') as f:
                    await context.bot.send_document(
                        chat_id=query.message.chat_id,
                        document=f,
                        filename="download_history.xlsx",
                        caption=f"📜 **ប្រវត្តិអ្នកប្រើប្រាស់ទាញយកគណនី (សរុប៖ {len(rows)} ដង)**"
                    )
                if os.path.exists(filepath):
                    os.remove(filepath)
            except Exception as e:
                logging.error(f"Error exporting history to Excel: {e}")
                await context.bot.send_message(
                    chat_id=query.message.chat_id,
                    text="❌ មានបញ្ហាក្នុងការនាំចេញប្រវត្តិទាញយកជា Excel。"
                )
                if os.path.exists(filepath):
                    os.remove(filepath)
 
    elif query.data == "admin_history_chat":
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, tg_name, menu_name, account_username, timestamp FROM history ORDER BY id DESC LIMIT 15")
        rows = cursor.fetchall()
        conn.close()
        
        keyboard = []
        reply_markup = make_inline_markup(keyboard)
        
        if not rows:
            await query.edit_message_text("📜 មិនទាន់មានប្រវត្តិទាញយកគណនីនៅក្នុង Database ឡើយ。", reply_markup=reply_markup)
        else:
            txt = "📜 **ប្រវត្តិទាញយកគណនីថ្មីៗបំផុតចំនួន ១៥៖**\n\n"
            for r in rows:
                txt += f"🆔 ID: `{r[0]}` | 👤 Name: *{r[1]}*\n📂 Menu: *{r[2]}* | 🔑 U: `{r[3]}`\n📅 ម៉ោង៖ `{r[4]}`\n------------------------\n"
            txt += "\n💡 *ចំណាំ៖* ដើម្បីទាញយកប្រវត្តិទាំងអស់ សូមប្រើប៊ូតុង 'ទាញយកជា Excel'。"
            await query.edit_message_text(txt, reply_markup=reply_markup, parse_mode="Markdown")

    elif query.data == "admin_history_clear_confirm":
        keyboard = [
            [InlineKeyboardButton("✅ ព្រមលុបប្រវត្តិទាំងអស់", callback_data="admin_history_clear_exec")]
        ]
        await query.edit_message_text(
            "⚠️ **តើអ្នកពិតជាចង់លុបប្រវត្តិទាញយកទាំងអស់ចេញពី Database មែនទេ?**\n\n"
            "Action នេះមិនអាចស្តារឡើងវិញបានឡើយ!",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="Markdown"
        )

    elif query.data == "admin_history_clear_exec":
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM history")
            conn.commit()
        finally:
            conn.close()
        keyboard = []
        await query.edit_message_text("🗑️ បានលុបប្រវត្តិទាញយកទាំងអស់ចេញពី Database រួចរាល់!", reply_markup=make_inline_markup(keyboard))
 
    elif query.data == "admin_add_help":
        keyboard = [
            [InlineKeyboardButton("📥 ទាញយកគំរូ Excel", callback_data="admin_download_sample")]
        ]
        reply_markup = make_inline_markup(keyboard)
        await query.edit_message_text(
            "➕ <b>របៀបបញ្ចូលទិន្នន័យ (Create):</b>\n\n"
            "<b>វិធីទី១៖ បញ្ចូលម្តងមួយៗ (Single Import)</b>\n"
            "សូមផ្ញើសារជាអក្សរធម្មតាមកកាន់ Bot តាមទម្រង់៖\n"
            "<code>/add [ឈ្មោះMenu] [Username] [Password]</code>\n"
            "<i>(ឧទាហរណ៍៖ <code>/add F88 user777 pass999</code>)</i>\n\n"
            "<b>វិធីទី២៖ បញ្ចូលជាគំនរតាមរយៈ Excel (Batch Import)</b>\n"
            "១. ចុចប៊ូតុងខាងក្រោម ឬផ្ញើសារ <code>/sample</code> ដើម្បីទាញយកឯកសារគំរូ Excel\n"
            "២. បំពេញទិន្នន័យគណនីចូលក្នុង Excel នោះរួចរក្សាទុក\n"
            "៣. ផ្ញើឯកសារ Excel (.xlsx) នោះមកកាន់ Bot នេះជាការស្រេច。",
            reply_markup=reply_markup,
            parse_mode="HTML"
        )
 
    elif query.data == "admin_update_list":
        keyboard = []
        reply_markup = make_inline_markup(keyboard)
        await query.edit_message_text(
            "📝 **របៀបកែប្រែទិន្នន័យ (Update):**\n\n"
            "សូមប្រើប្រាស់ Command នេះដើម្បីកែប្រែតាមរយៈ ID គណនី៖\n"
            "`/edit [ID] [Usernameថ្មី] [Passwordថ្មី]`\n\n"
            "*(អ្នកអាចមើល ID បានដោយចុចប៊ូតុង 'មើលទិន្នន័យទាំងអស់')*",
            reply_markup=reply_markup,
            parse_mode="Markdown"
        )
 
    elif query.data == "admin_delete_all_confirm":
        keyboard = [
            [InlineKeyboardButton("✅ ព្រមលុបទាំងអស់", callback_data="admin_delete_all_exec")]
        ]
        await query.edit_message_text("⚠️ **តើអ្នកពិតជាចង់លុបទិន្នន័យគណនីទាំងអស់ចេញពី Database មែនទេ?** Action នេះមិនអាចស្តារឡើងវិញបានឡើយ!", reply_markup=make_inline_markup(keyboard), parse_mode="Markdown")
 
    elif query.data == "admin_delete_all_exec":
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM accounts")
            conn.commit()
        finally:
            conn.close()
        keyboard = []
        await query.edit_message_text("🗑️ បានលុបទិន្នន័យទាំងអស់ចេញពី Database រួចរាល់!", reply_markup=make_inline_markup(keyboard))
 
    elif query.data == "admin_menus":
        await show_admin_menus_panel(query)

    elif query.data == "admin_main_menus":
        await show_admin_main_menus_panel(query)

    elif query.data == "admin_sub_menus":
        await show_admin_sub_menus_panel(query)

    elif query.data == "admin_add_main_menu":
        context.user_data['state'] = 'WAITING_ADD_MAIN_MENU_NAME'
        keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_main_menus")]]
        await query.edit_message_text(
            "➕ <b>សូមវាយផ្ញើឈ្មោះ Main Menu ថ្មីដែលអ្នកចង់បង្កើត៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data == "admin_edit_main_menu_list":
        keyboard = []
        for i, main_name in enumerate(MAIN_MENUS.keys()):
            keyboard.append([InlineKeyboardButton(f"📝 កែ {main_name}", callback_data=f"editmain_sel_{i}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_main_menus")])
        await query.edit_message_text(
            "📝 <b>សូមជ្រើសរើស Main Menu ដែលអ្នកចង់កែប្រែឈ្មោះ៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("editmain_sel_"):
        idx = int(query.data.replace("editmain_sel_", ""))
        main_keys = list(MAIN_MENUS.keys())
        if idx >= len(main_keys):
            await query.answer("❌ មិនរកឃើញ Main Menu", show_alert=True)
            return
        target_main = main_keys[idx]
        context.user_data['state'] = 'WAITING_EDIT_MAIN_MENU_NAME'
        context.user_data['edit_main_menu_target'] = target_main
        keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_edit_main_menu_list")]]
        await query.edit_message_text(
            f"📝 <b>ឈ្មោះបច្ចុប្បន្ន៖</b> <code>{target_main}</code>\n\n"
            f"សូមវាយផ្ញើឈ្មោះថ្មីសម្រាប់ Main Menu នេះមកកាន់ Bot៖",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data == "admin_del_main_menu_list":
        keyboard = []
        for i, main_name in enumerate(MAIN_MENUS.keys()):
            keyboard.append([InlineKeyboardButton(f"🗑️ លុប {main_name}", callback_data=f"delmain_sel_{i}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_main_menus")])
        await query.edit_message_text(
            "🗑️ <b>សូមជ្រើសរើស Main Menu ដែលអ្នកចង់លុប៖</b>\n\n"
            "⚠️ <b>បញ្ជាក់៖</b> មិនអាចលុប Main Menu ដែលមិនទាន់ទទេ (មាន Sub Menus) បានឡើយ។",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("delmain_sel_"):
        idx = int(query.data.replace("delmain_sel_", ""))
        main_keys = list(MAIN_MENUS.keys())
        if idx >= len(main_keys):
            await query.answer("❌ មិនរកឃើញ Main Menu", show_alert=True)
            return
        target_main = main_keys[idx]
        subs = MAIN_MENUS[target_main]
        if subs:
            keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_del_main_menu_list")]]
            await query.edit_message_text(
                f"❌ មិនអាចលុប Main Menu <b>{target_main}</b> បានទេ ព្រោះវាមិនទាន់ទទេ (មាន Sub Menu ចំនួន <b>{len(subs)}</b> នៅក្នុងនោះ)។ សូមលុប ឬបំលាស់ទី Sub Menu ទាំងនោះជាមុនសិន។",
                reply_markup=make_inline_markup(keyboard),
                parse_mode="HTML"
            )
            return

        MAIN_MENUS.pop(target_main)
        if not MAIN_MENUS:
            MAIN_MENUS["ទូទៅ"] = []
        config_data["MAIN_MENUS"] = MAIN_MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            await show_admin_main_menus_panel(query, f"បានលុប Main Menu {target_main} រួចរាល់!")
        except Exception as e:
            logging.error(f"Error deleting main menu: {e}")
            await query.answer("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ", show_alert=True)

    elif query.data == "admin_add_sub_menu":
        keyboard = []
        for i, main_name in enumerate(MAIN_MENUS.keys()):
            keyboard.append([InlineKeyboardButton(f"📁 {main_name}", callback_data=f"admin_sel_main_sub_{i}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_sub_menus")])
        await query.edit_message_text(
            "📁 <b>សូមជ្រើសរើស Main Menu ដែលអ្នកចង់បន្ថែម Sub Menu ទៅក្នុងនោះ៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("admin_sel_main_sub_"):
        idx = int(query.data.replace("admin_sel_main_sub_", ""))
        main_keys = list(MAIN_MENUS.keys())
        if idx >= len(main_keys):
            await query.answer("❌ មិនរកឃើញ Main Menu", show_alert=True)
            return
        target_main = main_keys[idx]
            
        pending_name = context.user_data.get('pending_sub_menu_name')
        if pending_name:
            context.user_data.pop('pending_sub_menu_name', None)
            if pending_name in MENUS:
                keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_sub_menus")]]
                await query.edit_message_text(f"❌ Menu <b>{pending_name}</b> មានរួចហើយ។", reply_markup=make_inline_markup(keyboard), parse_mode="HTML")
                return
                
            MENUS.append(pending_name)
            MAIN_MENUS[target_main].append(pending_name)
            config_data["MENUS"] = MENUS
            config_data["MAIN_MENUS"] = MAIN_MENUS
            try:
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=2)
                await show_admin_sub_menus_panel(query, f"បានបន្ថែម Menu '{pending_name}' ទៅក្នុង Main Menu '{target_main}' រួចរាល់!")
            except Exception as e:
                logging.error(f"Error saving config.json on pending sub menu: {e}")
                await query.edit_message_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។", parse_mode="HTML")
            return

        context.user_data['state'] = 'WAITING_ADD_MENU_NAME'
        context.user_data['selected_main_for_sub'] = target_main
        keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_add_sub_menu")]]
        await query.edit_message_text(
            f"➕ <b>សូមវាយផ្ញើឈ្មោះ Sub Menu ថ្មីដែលអ្នកចង់បង្កើតនៅក្រោម Main Menu <i>{target_main}</i>៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data == "admin_edit_sub_menu_list":
        keyboard = []
        for i, m in enumerate(MENUS):
            keyboard.append([InlineKeyboardButton(f"📝 កែ {m}", callback_data=f"editmenu_sel_{i}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_sub_menus")])
        await query.edit_message_text(
            "📝 <b>សូមជ្រើសរើស Sub Menu ដែលអ្នកចង់កែប្រែឈ្មោះ៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("editmenu_sel_"):
        idx = int(query.data.replace("editmenu_sel_", ""))
        if idx >= len(MENUS):
            await query.answer("❌ មិនរកឃើញ Menu", show_alert=True)
            return
        target_menu = MENUS[idx]
        context.user_data['state'] = 'WAITING_EDIT_MENU_NAME'
        context.user_data['edit_menu_target'] = target_menu
        keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_edit_sub_menu_list")]]
        await query.edit_message_text(
            f"📝 <b>ឈ្មោះបច្ចុប្បន្ន៖</b> <code>{target_menu}</code>\n\n"
            f"សូមវាយផ្ញើឈ្មោះថ្មីសម្រាប់ Menu នេះមកកាន់ Bot៖", 
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data == "admin_del_sub_menu_list":
        keyboard = []
        for i, m in enumerate(MENUS):
            keyboard.append([InlineKeyboardButton(f"🗑️ លុប {m}", callback_data=f"delmenu_sel_{i}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_sub_menus")])
        await query.edit_message_text(
            "🗑️ <b>សូមជ្រើសរើស Sub Menu ដែលអ្នកចង់លុប៖</b>\n\n"
            "⚠️ <b>បញ្ជាក់៖</b> មិនអាចលុប Sub Menu ដែលមានគណនីសល់ក្នុង database បានទេ។",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("delmenu_sel_"):
        idx = int(query.data.replace("delmenu_sel_", ""))
        if idx >= len(MENUS):
            await query.answer("❌ មិនរកឃើញ Menu", show_alert=True)
            return
        target_menu = MENUS[idx]
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT COUNT(*) FROM accounts WHERE menu_name = ?", (target_menu,))
            count = cursor.fetchone()[0]
        except Exception:
            count = 0
        finally:
            conn.close()

        if count > 0:
            keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_del_sub_menu_list")]]
            await query.edit_message_text(
                f"❌ មិនអាចលុប Menu <b>{target_menu}</b> បានទេ ព្រោះវាមិនទាន់ទទេ (មានគណនីចំនួន <b>{count}</b> ក្នុង Database)។ សូមលុបទិន្នន័យគណនីទាំងនោះជាមុនសិន។",
                reply_markup=make_inline_markup(keyboard),
                parse_mode="HTML"
            )
            return

        if target_menu in MENUS:
            MENUS.remove(target_menu)
        for main_name, subs in MAIN_MENUS.items():
            if target_menu in subs:
                subs.remove(target_menu)
                
        config_data["MENUS"] = MENUS
        config_data["MAIN_MENUS"] = MAIN_MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            await show_admin_sub_menus_panel(query, f"បានលុប Menu {target_menu} រួចរាល់!")
        except Exception as e:
            logging.error(f"Error saving config.json on callback delmenu: {e}")
            await query.answer("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ", show_alert=True)

    elif query.data == "admin_move_menu_select_sub":
        keyboard = []
        for i, m in enumerate(MENUS):
            keyboard.append([InlineKeyboardButton(f"🔄 {m}", callback_data=f"admin_move_select_sub_{i}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_menus")])
        await query.edit_message_text(
            "🔄 <b>សូមជ្រើសរើស Sub Menu ដែលអ្នកចង់បំលាស់ទី៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("admin_move_select_sub_"):
        sub_idx = int(query.data.replace("admin_move_select_sub_", ""))
        if sub_idx >= len(MENUS):
            await query.answer("❌ Sub Menu index error", show_alert=True)
            return
        sub_name = MENUS[sub_idx]
        context.user_data['move_sub_menu_target'] = sub_name
        
        keyboard = []
        for i, main_name in enumerate(MAIN_MENUS.keys()):
            keyboard.append([InlineKeyboardButton(f"📁 {main_name}", callback_data=f"admin_move_select_main_{i}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_move_menu_select_sub")])
        await query.edit_message_text(
            f"🔄 <b>សូមជ្រើសរើស Main Menu ដើម្បីបំលាស់ទី Sub Menu '{sub_name}' ទៅកាន់៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("admin_move_select_main_"):
        main_idx = int(query.data.replace("admin_move_select_main_", ""))
        main_keys = list(MAIN_MENUS.keys())
        if main_idx >= len(main_keys):
            await query.answer("❌ Main Menu index error", show_alert=True)
            return
        main_name = main_keys[main_idx]
        sub_name = context.user_data.get('move_sub_menu_target')
        
        if not sub_name:
            await query.answer("❌ មិនរកឃើញ Sub Menu សម្រាប់ផ្ទេរឡើយ", show_alert=True)
            return
            
        for k, subs in MAIN_MENUS.items():
            if sub_name in subs:
                subs.remove(sub_name)
        
        MAIN_MENUS[main_name].append(sub_name)
        config_data["MAIN_MENUS"] = MAIN_MENUS
        
        MENUS.clear()
        for subs in MAIN_MENUS.values():
            MENUS.extend(subs)
        config_data["MENUS"] = MENUS
        
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            await show_admin_menus_panel(query, f"បានផ្ទេរ '{sub_name}' ទៅ '{main_name}' ដោយជោគជ័យ!")
        except Exception as e:
            logging.error(f"Error moving menu: {e}")
            await query.answer("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ", show_alert=True)

    elif query.data == "admin_sort_manual":
        keyboard = [
            [InlineKeyboardButton("📁 តម្រៀប Main Menu", callback_data="admin_sort_manual_main")],
            [InlineKeyboardButton("📂 តម្រៀប Sub Menu", callback_data="admin_sort_manual_sub_select_main")],
            [InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_menus")]
        ]
        await query.edit_message_text(
            "🎨 <b>តម្រៀប Menu ដោយខ្លួនឯង (Manual Reordering):</b>\n\nសូមជ្រើសរើសប្រភេទ Menu ដែលអ្នកចង់តម្រៀប៖",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data == "admin_sort_manual_main":
        keyboard = []
        for i, main_name in enumerate(MAIN_MENUS.keys()):
            keyboard.append([InlineKeyboardButton(f"📁 {main_name}", callback_data=f"sort_main_sel_{i}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_sort_manual")])
        await query.edit_message_text(
            "📁 <b>សូមជ្រើសរើស Main Menu ដែលចង់ផ្លាស់ទី/តម្រៀប៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("sort_main_sel_"):
        idx = int(query.data.replace("sort_main_sel_", ""))
        main_keys = list(MAIN_MENUS.keys())
        if idx >= len(main_keys):
            await query.answer("❌ Main Menu index error", show_alert=True)
            return
        target_main = main_keys[idx]
        lines = []
        for i, k in enumerate(main_keys):
            prefix = "👉 " if i == idx else "▫️ "
            lines.append(f"{prefix}<b>{k}</b>")
        list_str = "\n".join(lines)

        keyboard = [
            [
                InlineKeyboardButton("⬆️ ឡើងលើ (Move Up)", callback_data=f"sort_main_up_{idx}"),
                InlineKeyboardButton("⬇️ ចុះក្រោម (Move Down)", callback_data=f"sort_main_down_{idx}")
            ],
            [InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_sort_manual_main")]
        ]
        await query.edit_message_text(
            f"📁 <b>តម្រៀប Main Menu៖</b>\n\n"
            f"{list_str}\n\n"
            f"សូមប្រើប៊ូតុងខាងក្រោមដើម្បីផ្លាស់ទី <b>{target_main}</b> ឡើងលើ ឬចុះក្រោម៖",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("sort_main_up_") or query.data.startswith("sort_main_down_"):
        is_up = "sort_main_up_" in query.data
        idx = int(query.data.replace("sort_main_up_", "").replace("sort_main_down_", ""))
        main_keys = list(MAIN_MENUS.keys())
        if idx >= len(main_keys):
            return
            
        new_idx = idx
        if is_up and idx > 0:
            main_keys[idx], main_keys[idx-1] = main_keys[idx-1], main_keys[idx]
            new_idx = idx - 1
        elif not is_up and idx < len(main_keys) - 1:
            main_keys[idx], main_keys[idx+1] = main_keys[idx+1], main_keys[idx]
            new_idx = idx + 1
            
        new_main_menus = {}
        for k in main_keys:
            new_main_menus[k] = MAIN_MENUS[k]
        MAIN_MENUS.clear()
        MAIN_MENUS.update(new_main_menus)
        
        MENUS.clear()
        for subs in MAIN_MENUS.values():
            MENUS.extend(subs)
            
        config_data["MAIN_MENUS"] = MAIN_MENUS
        config_data["MENUS"] = MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            await query.answer("បានផ្លាស់ទីរួចរាល់")
        except Exception as e:
            logging.error(f"Error saving config.json on sort main: {e}")
            await query.answer("❌ មានបញ្ហាក្នុងការរក្សាទុក", show_alert=True)
            return

        target_main = main_keys[new_idx]
        lines = []
        for i, k in enumerate(main_keys):
            prefix = "👉 " if i == new_idx else "▫️ "
            lines.append(f"{prefix}<b>{k}</b>")
        list_str = "\n".join(lines)

        keyboard = [
            [
                InlineKeyboardButton("⬆️ ឡើងលើ (Move Up)", callback_data=f"sort_main_up_{new_idx}"),
                InlineKeyboardButton("⬇️ ចុះក្រោម (Move Down)", callback_data=f"sort_main_down_{new_idx}")
            ],
            [InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_sort_manual_main")]
        ]
        await query.edit_message_text(
            f"📁 <b>តម្រៀប Main Menu៖</b>\n\n"
            f"{list_str}\n\n"
            f"សូមប្រើប៊ូតុងខាងក្រោមដើម្បីផ្លាស់ទី <b>{target_main}</b> ឡើងលើ ឬចុះក្រោម៖",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data == "admin_sort_manual_sub_select_main":
        keyboard = []
        for i, main_name in enumerate(MAIN_MENUS.keys()):
            keyboard.append([InlineKeyboardButton(f"📁 {main_name}", callback_data=f"sort_sub_list_{i}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_sort_manual")])
        await query.edit_message_text(
            "📂 <b>សូមជ្រើសរើស Main Menu ដើម្បីមើល Sub Menus របស់វាសម្រាប់តម្រៀប៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("sort_sub_list_"):
        main_idx = int(query.data.replace("sort_sub_list_", ""))
        main_keys = list(MAIN_MENUS.keys())
        if main_idx >= len(main_keys):
            await query.answer("❌ Main Menu index error", show_alert=True)
            return
        target_main = main_keys[main_idx]
        subs = MAIN_MENUS[target_main]
        keyboard = []
        for sub_idx, sub in enumerate(subs):
            keyboard.append([InlineKeyboardButton(f"▪️ {sub}", callback_data=f"sort_sub_sel_{main_idx}_{sub_idx}")])
        keyboard.append([InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_sort_manual_sub_select_main")])
        await query.edit_message_text(
            f"📂 <b>សូមជ្រើសរើស Sub Menu ក្នុង '{target_main}' ដែលចង់ផ្លាស់ទី/តម្រៀប៖</b>",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("sort_sub_sel_"):
        parts = query.data.replace("sort_sub_sel_", "").split("_")
        if len(parts) != 2:
            return
        main_idx, sub_idx = int(parts[0]), int(parts[1])
        main_keys = list(MAIN_MENUS.keys())
        if main_idx >= len(main_keys):
            return
        target_main = main_keys[main_idx]
        subs = MAIN_MENUS[target_main]
        if sub_idx >= len(subs):
            return
        target_sub = subs[sub_idx]

        lines = []
        for i, s in enumerate(subs):
            prefix = "👉 " if i == sub_idx else "▫️ "
            lines.append(f"{prefix}<b>{s}</b>")
        list_str = "\n".join(lines)

        keyboard = [
            [
                InlineKeyboardButton("⬆️ ឡើងលើ (Move Up)", callback_data=f"sort_sub_up_{main_idx}_{sub_idx}"),
                InlineKeyboardButton("⬇️ ចុះក្រោម (Move Down)", callback_data=f"sort_sub_down_{main_idx}_{sub_idx}")
            ],
            [InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data=f"sort_sub_list_{main_idx}")]
        ]
        await query.edit_message_text(
            f"📂 <b>តម្រៀប Sub Menu ក្នុង '{target_main}'៖</b>\n\n"
            f"{list_str}\n\n"
            f"សូមប្រើប៊ូតុងខាងក្រោមដើម្បីផ្លាស់ទី <b>{target_sub}</b> ឡើងលើ ឬចុះក្រោម៖",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data.startswith("sort_sub_up_") or query.data.startswith("sort_sub_down_"):
        is_up = "sort_sub_up_" in query.data
        raw_parts = query.data.replace("sort_sub_up_", "").replace("sort_sub_down_", "").split("_")
        if len(raw_parts) != 2:
            return
        main_idx, sub_idx = int(raw_parts[0]), int(raw_parts[1])
        main_keys = list(MAIN_MENUS.keys())
        if main_idx >= len(main_keys):
            return
        target_main = main_keys[main_idx]
        subs = MAIN_MENUS[target_main]
        if sub_idx >= len(subs):
            return
            
        new_sub_idx = sub_idx
        if is_up and sub_idx > 0:
            subs[sub_idx], subs[sub_idx-1] = subs[sub_idx-1], subs[sub_idx]
            new_sub_idx = sub_idx - 1
        elif not is_up and sub_idx < len(subs) - 1:
            subs[sub_idx], subs[sub_idx+1] = subs[sub_idx+1], subs[sub_idx]
            new_sub_idx = sub_idx + 1
            
        MAIN_MENUS[target_main] = subs
        
        MENUS.clear()
        for s_list in MAIN_MENUS.values():
            MENUS.extend(s_list)
            
        config_data["MAIN_MENUS"] = MAIN_MENUS
        config_data["MENUS"] = MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            await query.answer("បានផ្លាស់ទីរួចរាល់")
        except Exception as e:
            logging.error(f"Error saving config.json on sort sub: {e}")
            await query.answer("❌ មានបញ្ហាក្នុងការរក្សាទុក", show_alert=True)
            return

        target_sub = subs[new_sub_idx]
        lines = []
        for i, s in enumerate(subs):
            prefix = "👉 " if i == new_sub_idx else "▫️ "
            lines.append(f"{prefix}<b>{s}</b>")
        list_str = "\n".join(lines)

        keyboard = [
            [
                InlineKeyboardButton("⬆️ ឡើងលើ (Move Up)", callback_data=f"sort_sub_up_{main_idx}_{new_sub_idx}"),
                InlineKeyboardButton("⬇️ ចុះក្រោម (Move Down)", callback_data=f"sort_sub_down_{main_idx}_{new_sub_idx}")
            ],
            [InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data=f"sort_sub_list_{main_idx}")]
        ]
        await query.edit_message_text(
            f"📂 <b>តម្រៀប Sub Menu ក្នុង '{target_main}'៖</b>\n\n"
            f"{list_str}\n\n"
            f"សូមប្រើប៊ូតុងខាងក្រោមដើម្បីផ្លាស់ទី <b>{target_sub}</b> ឡើងលើ ឬចុះក្រោម៖",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data == "admin_admins":
        admin_list = get_admin_list_text()
        keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_back")]]
        await query.edit_message_text(
            f"👤 <b>គ្រប់គ្រង Admin ទាំងអស់ក្នុងប្រព័ន្ធ (Admins Config)៖</b>\n\n"
            f"📋 <b>បញ្ជី ID Admin និងតួនាទី៖</b>\n{admin_list}\n\n"
            f"➕ <b>របៀបបន្ថែម Admin៖</b>\n"
            f"<code>/addadmin [លេខID]</code>\n\n"
            f"🗑️ <b>របៀបលុប Admin៖</b>\n"
            f"<code>/deladmin [លេខID]</code>", 
            reply_markup=make_inline_markup(keyboard),
            parse_mode="HTML"
        )

    elif query.data == "admin_backup":
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        config_path = os.path.join(BASE_DIR, 'config.json')
        
        status_msg = await query.message.reply_text("⏳ កំពុងរៀបចំឯកសារបម្រុងទុក (Backup)...")
        
        backup_files_sent = 0
        if os.path.exists(db_path):
            try:
                with open(db_path, 'rb') as f:
                    await context.bot.send_document(
                        chat_id=query.message.chat_id,
                        document=f,
                        filename="bot_data.db",
                        caption="📦 ឯកសារបម្រុងទុក Database (SQLite Database Backup)"
                    )
                backup_files_sent += 1
            except Exception as e:
                logging.error(f"Error backup db: {e}")
                await context.bot.send_message(chat_id=query.message.chat_id, text="❌ មិនអាចផ្ញើឯកសារ Database បានទេ។")
                
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT menu_name, username, password FROM accounts")
                rows = cursor.fetchall()
                conn.close()
                
                excel_filepath = os.path.join(BASE_DIR, "backup_accounts.xlsx")
                data = {
                    "menu_name": [r[0] for r in rows] if rows else [],
                    "username": [r[1] for r in rows] if rows else [],
                    "password": [r[2] for r in rows] if rows else []
                }
                df = pd.DataFrame(data)
                with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name="Data_Entry", index=False)
                    
                with open(excel_filepath, 'rb') as f:
                    await context.bot.send_document(
                        chat_id=query.message.chat_id,
                        document=f,
                        filename="backup_accounts.xlsx",
                        caption=f"📊 ឯកសារបម្រុងទុកគណនីជា Excel (Excel Accounts Backup - សរុប៖ {len(rows) if rows else 0} អាខោន)"
                    )
                backup_files_sent += 1
                if os.path.exists(excel_filepath):
                    os.remove(excel_filepath)
            except Exception as e:
                logging.error(f"Error backup accounts excel: {e}")
                await context.bot.send_message(chat_id=query.message.chat_id, text="❌ មិនអាចផ្ញើឯកសារបម្រុងទុកជា Excel បានទេ។")
                
        if os.path.exists(config_path):
            try:
                with open(config_path, 'rb') as f:
                    await context.bot.send_document(
                        chat_id=query.message.chat_id,
                        document=f,
                        filename="config.json",
                        caption="⚙️ ឯកសារបម្រុងទុកការកំណត់ (Config Backup)"
                    )
                backup_files_sent += 1
            except Exception as e:
                logging.error(f"Error backup config: {e}")
                await context.bot.send_message(chat_id=query.message.chat_id, text="❌ មិនអាចផ្ញើឯកសារ Config បានទេ។")
                
        keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_back")]]
        reply_markup = make_inline_markup(keyboard)
        if backup_files_sent > 0:
            await status_msg.edit_text("✅ ការបម្រុងទុកទិន្នន័យត្រូវបានបញ្ចប់ដោយជោគជ័យ!", reply_markup=reply_markup)
        else:
            await status_msg.edit_text("❌ គ្មានឯកសារត្រូវបានរកឃើញសម្រាប់បម្រុងទុកឡើយ。", reply_markup=reply_markup)

    elif query.data == "admin_restore":
        context.user_data['state'] = 'WAITING_RESTORE_FILE'
        keyboard = [[InlineKeyboardButton("⬅️ ត្រឡប់ក្រោយ (Back)", callback_data="admin_back")]]
        await query.edit_message_text(
            "📤 **សូមផ្ញើឯកសារ Backup (.db ឬ .json) មកកាន់ Bot នេះដើម្បីស្តារទិន្នន័យឡើងវិញ៖**\n\n"
            "💡 **ចំណាំ៖**\n"
            "- ប្រសិនបើអ្នកផ្ញើឯកសារ `.db` (ឧ. `bot_data.db`) វានឹងជំនួស Database គណនីទាំងអស់。\n"
            "- ប្រសិនបើអ្នកផ្ញើឯកសារ `.json` (ឧ. `config.json`) វានឹងជំនួសការកំណត់របស់ Bot (Tokens, Admins, Menus)៖",
            reply_markup=make_inline_markup(keyboard),
            parse_mode="Markdown"
        )

    elif query.data == "admin_download_menus_sample":
        filepath = os.path.join(BASE_DIR, "sample_menus.xlsx")
        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                data = {
                    "menu_name": ["F88$", "F88$_CH", "F88_CH_KH", "F88_KH", "Bet99", "Bet88"]
                }
                df = pd.DataFrame(data)
                df.to_excel(writer, sheet_name="Menus", index=False)
            
            with open(filepath, 'rb') as f:
                await context.bot.send_document(
                    chat_id=query.message.chat_id,
                    document=f,
                    filename="sample_menus.xlsx",
                    caption="📋 នេះជាគំរូឯកសារ Excel សម្រាប់បញ្ចូលតែឈ្មោះ Menu (Sample Menus-only Template)។\n\n💡 របៀបប្រើ៖ បំពេញឈ្មោះ Menu ក្នុងជួរឈរ menu_name រួចផ្ញើឯកសារនេះមកកាន់ Bot ដើម្បីបង្កើតប៊ូតុង Menu ទាំងនោះ។"
                )
            
            if os.path.exists(filepath):
                os.remove(filepath)
        except Exception as e:
            logging.error(f"Error generating sample menus excel callback: {e}")
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="❌ មានបញ្ហាក្នុងការបង្កើតឯកសារគំរូ。"
            )

    elif query.data.startswith("delmenu_select_"):
        target_menu = query.data.replace("delmenu_select_", "")
        
        if target_menu not in MENUS:
            keyboard = []
            await query.edit_message_text(f"❌ មិនរកឃើញ Menu <b>{target_menu}</b> ឡើយ។", reply_markup=make_inline_markup(keyboard), parse_mode="HTML")
            return
            
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT COUNT(*) FROM accounts WHERE menu_name = ?", (target_menu,))
            count = cursor.fetchone()[0]
        except Exception:
            count = 0
        finally:
            conn.close()

        if count > 0:
            keyboard = []
            await query.edit_message_text(
                f"❌ មិនអាចលុប Menu <b>{target_menu}</b> បានទេ ព្រោះវាមិនទាន់ទទេ (មានគណនីចំនួន <b>{count}</b> ក្នុង Database)។ សូមលុបទិន្នន័យគណនីទាំងនោះជាមុនសិន។",
                reply_markup=make_inline_markup(keyboard),
                parse_mode="HTML"
            )
            return

        MENUS.remove(target_menu)
        for main_name, subs in MAIN_MENUS.items():
            if target_menu in subs:
                subs.remove(target_menu)
                
        config_data["MENUS"] = MENUS
        config_data["MAIN_MENUS"] = MAIN_MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            
            keyboard = []
            await query.edit_message_text(
                f"🗑️ បានលុប Menu <b>{target_menu}</b> រួចរាល់! សូមចុច <code>/start</code> សារជាថ្មីដើម្បីមើលការផ្លាស់ប្តូរ។", 
                reply_markup=make_inline_markup(keyboard),
                parse_mode="HTML"
            )
        except Exception as e:
            logging.error(f"Error saving config.json on callback delmenu: {e}")
            keyboard = []
            await query.edit_message_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។", reply_markup=make_inline_markup(keyboard))

    elif query.data == "admin_back":
        keyboard = get_main_admin_keyboard()
        reply_markup = make_inline_markup(keyboard)
        await query.edit_message_text("🛠️ **ផ្ទាំងគ្រប់គ្រងសម្រាប់ Admin (CRUD Panel):**", reply_markup=reply_markup, parse_mode="Markdown")

    elif query.data == "admin_download_sample":
        # Generate and send the excel sample file
        filepath = os.path.join(BASE_DIR, "sample_accounts.xlsx")
        try:
            create_sample_accounts_excel(filepath)
            
            with open(filepath, 'rb') as f:
                await context.bot.send_document(
                    chat_id=query.message.chat_id,
                    document=f,
                    filename="sample_accounts.xlsx",
                    caption="📋 នេះជាគំរូឯកសារ Excel (Sample Excel Template) សម្រាប់បញ្ចូលគណនី។"
                )
            
            if os.path.exists(filepath):
                os.remove(filepath)
        except Exception as e:
            logging.error(f"Error generating sample excel callback: {e}")
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="❌ មានបញ្ហាក្នុងការបង្កើតឯកសារគំរូ。"
            )

    elif query.data == "admin_back":
        keyboard = get_main_admin_keyboard()
        reply_markup = make_inline_markup(keyboard)
        await query.edit_message_text("🛠️ **ផ្ទាំងគ្រប់គ្រងសម្រាប់ Admin (CRUD Panel):**", reply_markup=reply_markup, parse_mode="Markdown")

    elif query.data == "admin_admins":
        admin_list = get_admin_list_text()
        keyboard = []
        reply_markup = make_inline_markup(keyboard)
        await query.edit_message_text(
            f"👤 <b>គ្រប់គ្រង Admin ទាំងអស់ក្នុងប្រព័ន្ធ (Admins Config)៖</b>\n\n"
            f"📋 <b>បញ្ជី ID Admin និងតួនាទី៖</b>\n{admin_list}\n\n"
            f"➕ <b>របៀបបន្ថែម Admin៖</b>\n"
            f"<code>/addadmin [លេខID]</code>\n\n"
            f"🗑️ <b>របៀបលុប Admin៖</b>\n"
            f"<code>/deladmin [លេខID]</code>", 
            reply_markup=reply_markup,
            parse_mode="HTML"
        )

    elif query.data == "admin_backup":
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        config_path = os.path.join(BASE_DIR, 'config.json')
        
        status_msg = await query.message.reply_text("⏳ កំពុងរៀបចំឯកសារបម្រុងទុក (Backup)...")
        
        backup_files_sent = 0
        if os.path.exists(db_path):
            try:
                with open(db_path, 'rb') as f:
                    await context.bot.send_document(
                        chat_id=query.message.chat_id,
                        document=f,
                        filename="bot_data.db",
                        caption="📦 ឯកសារបម្រុងទុក Database (SQLite Database Backup)"
                    )
                backup_files_sent += 1
            except Exception as e:
                logging.error(f"Error backup db: {e}")
                await context.bot.send_message(chat_id=query.message.chat_id, text="❌ មិនអាចផ្ញើឯកសារ Database បានទេ។")
                
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT menu_name, username, password FROM accounts")
                rows = cursor.fetchall()
                conn.close()
                
                excel_filepath = os.path.join(BASE_DIR, "backup_accounts.xlsx")
                data = {
                    "menu_name": [r[0] for r in rows] if rows else [],
                    "username": [r[1] for r in rows] if rows else [],
                    "password": [r[2] for r in rows] if rows else []
                }
                df = pd.DataFrame(data)
                with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name="Data_Entry", index=False)
                    
                with open(excel_filepath, 'rb') as f:
                    await context.bot.send_document(
                        chat_id=query.message.chat_id,
                        document=f,
                        filename="backup_accounts.xlsx",
                        caption=f"📊 ឯកសារបម្រុងទុកគណនីជា Excel (Excel Accounts Backup - សរុប៖ {len(rows) if rows else 0} អាខោន)"
                    )
                backup_files_sent += 1
                if os.path.exists(excel_filepath):
                    os.remove(excel_filepath)
            except Exception as e:
                logging.error(f"Error backup accounts excel: {e}")
                await context.bot.send_message(chat_id=query.message.chat_id, text="❌ មិនអាចផ្ញើឯកសារបម្រុងទុកជា Excel បានទេ។")
                
        if os.path.exists(config_path):
            try:
                with open(config_path, 'rb') as f:
                    await context.bot.send_document(
                        chat_id=query.message.chat_id,
                        document=f,
                        filename="config.json",
                        caption="⚙️ ឯកសារបម្រុងទុកការកំណត់ (Config Backup)"
                    )
                backup_files_sent += 1
            except Exception as e:
                logging.error(f"Error backup config: {e}")
                await context.bot.send_message(chat_id=query.message.chat_id, text="❌ មិនអាចផ្ញើឯកសារ Config បានទេ។")
                
        keyboard = [
            [
                InlineKeyboardButton("🔄 ធ្វើឲ្យស្រស់ (Refresh)", callback_data="admin_refresh_backup"),
                InlineKeyboardButton("🏠 ទៅទំព័រដើម (Home)", callback_data="admin_back"),
                InlineKeyboardButton("⬅️ ត្រលប់ក្រោយ (Back)", callback_data="admin_back")
            ]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        if backup_files_sent > 0:
            await status_msg.edit_text("✅ ការបម្រុងទុកទិន្នន័យត្រូវបានបញ្ចប់ដោយជោគជ័យ!", reply_markup=reply_markup)
        else:
            await status_msg.edit_text("❌ គ្មានឯកសារត្រូវបានរកឃើញសម្រាប់បម្រុងទុកឡើយ。", reply_markup=reply_markup)

    elif query.data == "admin_restore":
        context.user_data['state'] = 'WAITING_RESTORE_FILE'
        keyboard = [
            [
                InlineKeyboardButton("🔄 ធ្វើឲ្យស្រស់ (Refresh)", callback_data="admin_refresh_restore"),
                InlineKeyboardButton("🏠 ទៅទំព័រដើម (Home)", callback_data="admin_back"),
                InlineKeyboardButton("⬅️ ត្រលប់ក្រោយ (Back)", callback_data="admin_back")
            ]
        ]
        await query.edit_message_text(
            "📤 **សូមផ្ញើឯកសារ Backup (.db ឬ .json) មកកាន់ Bot នេះដើម្បីស្តារទិន្នន័យឡើងវិញ៖**\n\n"
            "💡 **ចំណាំ៖**\n"
            "- ប្រសិនបើអ្នកផ្ញើឯកសារ `.db` (ឧ. `bot_data.db`) វានឹងជំនួស Database គណនីទាំងអស់。\n"
            "- ប្រសិនបើអ្នកផ្ញើឯកសារ `.json` (ឧ. `config.json`) វានឹងជំនួសការកំណត់របស់ Bot (Tokens, Admins, Menus)៖",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )

# បន្ថែម Menu ថ្មីទៅក្នុង config.json តាម Command /addmenu
async def add_menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if not args:
        keyboard = []
        for i, main_name in enumerate(MAIN_MENUS.keys()):
            keyboard.append([InlineKeyboardButton(f"📁 {main_name}", callback_data=f"admin_sel_main_sub_{i}")])
        keyboard.append([InlineKeyboardButton("❌ បោះបង់", callback_data="admin_sub_menus")])
        await update.message.reply_text(
            "📁 **សូមជ្រើសរើស Main Menu ដែលអ្នកចង់បន្ថែម Sub Menu ទៅក្នុងនោះ៖**",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )
        return

    new_menu = args[0].strip()
    if new_menu in MENUS:
        await update.message.reply_text(f"❌ Menu <b>{new_menu}</b> មានរួចហើយ។", parse_mode="HTML")
        return

    context.user_data['pending_sub_menu_name'] = new_menu
    keyboard = []
    for i, main_name in enumerate(MAIN_MENUS.keys()):
        keyboard.append([InlineKeyboardButton(f"📁 {main_name}", callback_data=f"admin_sel_main_sub_{i}")])
    keyboard.append([InlineKeyboardButton("❌ បោះបង់", callback_data="admin_sub_menus")])
    await update.message.reply_text(
        f"📁 **សូមជ្រើសរើស Main Menu ដើម្បីបង្កើត Sub Menu '{new_menu}' នៅក្រោមនោះ៖**",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown"
    )

# លុប Menu ពីក្នុង config.json តាម Command /delmenu
async def del_menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if not args:
        keyboard = []
        for i, m in enumerate(MENUS):
            keyboard.append([InlineKeyboardButton(f"🗑️ លុប {m}", callback_data=f"delmenu_sel_{i}")])
        keyboard.append([InlineKeyboardButton("❌ បោះបង់", callback_data="admin_sub_menus")])
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text("🗑️ **សូមជ្រើសរើស Menu ដែលអ្នកចង់លុប៖**", reply_markup=reply_markup, parse_mode="Markdown")
        return

    target_menu = args[0].strip()
    if target_menu not in MENUS:
        await update.message.reply_text(f"❌ រកមិនឃើញ Sub Menu <b>{target_menu}</b> ឡើយ។", parse_mode="HTML")
        return

    db_path = os.path.join(BASE_DIR, 'bot_data.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT COUNT(*) FROM accounts WHERE menu_name = ?", (target_menu,))
        count = cursor.fetchone()[0]
    except Exception:
        count = 0
    finally:
        conn.close()

    if count > 0:
        await update.message.reply_text(f"❌ មិនអាចលុប Menu <b>{target_menu}</b> បានទេ ព្រោះវាមិនទាន់ទទេ (មានគណនីចំនួន <b>{count}</b> ក្នុង Database)។ សូមលុបទិន្នន័យគណនីទាំងនោះជាមុនសិន។", parse_mode="HTML")
        return

    MENUS.remove(target_menu)
    for main_name, subs in MAIN_MENUS.items():
        if target_menu in subs:
            subs.remove(target_menu)

    config_data["MENUS"] = MENUS
    config_data["MAIN_MENUS"] = MAIN_MENUS
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
        await update.message.reply_text(f"🗑️ បានលុប Menu <b>{target_menu}</b> រួចរាល់! សូមចុច <code>/start</code> សារជាថ្មីដើម្បីមើលការផ្លាស់ប្តូរ។", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error saving config.json: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")

# កែប្រែឈ្មោះ Menu ក្នុង config.json និង Database តាម Command /editmenu
async def edit_menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if len(args) < 2:
        keyboard = []
        for i, m in enumerate(MENUS):
            keyboard.append([InlineKeyboardButton(f"📝 កែប្រែ {m}", callback_data=f"editmenu_sel_{i}")])
        keyboard.append([InlineKeyboardButton("❌ បោះបង់", callback_data="admin_sub_menus")])
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text("📝 **សូមជ្រើសរើស Menu ដែលអ្នកចង់កែប្រែឈ្មោះ៖**", reply_markup=reply_markup, parse_mode="Markdown")
        return

    target_menu, new_name = args[0].strip(), args[1].strip()
    if target_menu not in MENUS:
        await update.message.reply_text(f"❌ មិនរកឃើញ Menu <b>{target_menu}</b> ឡើយ។", parse_mode="HTML")
        return

    if new_name in MENUS:
        await update.message.reply_text(f"❌ Menu <b>{new_name}</b> មានរួចហើយ។", parse_mode="HTML")
        return

    idx = MENUS.index(target_menu)
    MENUS[idx] = new_name

    for main_name, subs in MAIN_MENUS.items():
        if target_menu in subs:
            s_idx = subs.index(target_menu)
            subs[s_idx] = new_name

    config_data["MENUS"] = MENUS
    config_data["MAIN_MENUS"] = MAIN_MENUS
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
            
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("UPDATE accounts SET menu_name = ? WHERE menu_name = ?", (new_name, target_menu))
        conn.commit()
        conn.close()
        
        await update.message.reply_text(f"📝 បានកែប្រែ Menu ពី <b>{target_menu}</b> ទៅជា <b>{new_name}</b> រួចរាល់! គណនីក្នុង Database ក៏ត្រូវបានធ្វើបច្ចុប្បន្នភាពផងដែរ។", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error updating menu {target_menu} to {new_name}: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")

# បំលាស់ទី Sub Menu ទៅ Main Menu តាម Command /movemenu
async def move_menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if len(args) < 2:
        keyboard = []
        for i, m in enumerate(MENUS):
            keyboard.append([InlineKeyboardButton(f"🔄 {m}", callback_data=f"admin_move_select_sub_{i}")])
        keyboard.append([InlineKeyboardButton("❌ បោះបង់", callback_data="admin_menus")])
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text("🔄 **សូមជ្រើសរើស Sub Menu ដែលអ្នកចង់បំលាស់ទី៖**", reply_markup=reply_markup, parse_mode="Markdown")
        return

    sub_name, main_name = args[0].strip(), args[1].strip()
    if sub_name not in MENUS:
        await update.message.reply_text(f"❌ មិនរកឃើញ Sub Menu <b>{sub_name}</b> ឡើយ។", parse_mode="HTML")
        return

    if main_name not in MAIN_MENUS:
        await update.message.reply_text(f"❌ មិនរកឃើញ Main Menu <b>{main_name}</b> ឡើយ។", parse_mode="HTML")
        return

    for k, subs in MAIN_MENUS.items():
        if sub_name in subs:
            subs.remove(sub_name)
    
    MAIN_MENUS[main_name].append(sub_name)
    config_data["MAIN_MENUS"] = MAIN_MENUS
    
    MENUS.clear()
    for subs in MAIN_MENUS.values():
        MENUS.extend(subs)
    config_data["MENUS"] = MENUS
    
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
        await update.message.reply_text(f"✅ បានផ្ទេរ Sub Menu <b>{sub_name}</b> ទៅកាន់ Main Menu <b>{main_name}</b> ដោយជោគជ័យ!", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error moving menu via command: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")


# បន្ថែម Admin ថ្មីទៅក្នុង config.json តាម Command /addadmin
async def add_admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if not args:
        await update.message.reply_text("⚠️ ទម្រង់ខុស! សូមប្រើ៖ `/addadmin [លេខID]`")
        return

    try:
        new_admin = int(args[0].strip())
    except ValueError:
        await update.message.reply_text("❌ លេខ ID ត្រូវតែជាលេខពិតប្រាកដ។")
        return

    if new_admin in ADMIN_IDS:
        await update.message.reply_text(f"❌ ID <code>{new_admin}</code> ជា Admin រួចហើយ។", parse_mode="HTML")
        return

    ADMIN_IDS.append(new_admin)
    config_data["ADMIN_IDS"] = ADMIN_IDS
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
        await update.message.reply_text(f"✅ បានបន្ថែម Admin ID: <code>{new_admin}</code> ដោយជោគជ័យ!", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error saving config.json: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")


# លុប Admin ពីក្នុង config.json តាម Command /deladmin
async def del_admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if not args:
        await update.message.reply_text("⚠️ ទម្រង់ខុស! សូមប្រើ៖ `/deladmin [លេខID]`")
        return

    try:
        target_admin = int(args[0].strip())
    except ValueError:
        await update.message.reply_text("❌ លេខ ID ត្រូវតែជាលេខពិតប្រាកដ។")
        return

    if target_admin == 8558847170:
        await update.message.reply_text("❌ មិនអាចលុប Default Admin ID (8558847170) បានឡើយ។")
        return

    if target_admin not in ADMIN_IDS:
        await update.message.reply_text(f"❌ រកមិនឃើញ ID <code>{target_admin}</code> ក្នុងបញ្ជី Admin ឡើយ។", parse_mode="HTML")
        return

    ADMIN_IDS.remove(target_admin)
    config_data["ADMIN_IDS"] = ADMIN_IDS
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
        await update.message.reply_text(f"🗑️ បានលុប Admin ID: <code>{target_admin}</code> រួចរាល់!", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error saving config.json: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")


# បន្ថែមគណនីថ្មីទៅក្នុង Database តាម Command /add [menu_name] [username] [password]
async def add_account_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if len(args) < 3:
        await update.message.reply_text("⚠️ ទម្រង់ខុស! សូមប្រើ៖ `/add [menu_name] [username] [password]`", parse_mode="Markdown")
        return

    menu_name, username, password = args[0].strip(), args[1].strip(), args[2].strip()

    if menu_name not in MENUS:
        MENUS.append(menu_name)
        if not MAIN_MENUS:
            MAIN_MENUS["ទូទៅ"] = []
        first_key = list(MAIN_MENUS.keys())[0]
        MAIN_MENUS[first_key].append(menu_name)
        config_data["MENUS"] = MENUS
        config_data["MAIN_MENUS"] = MAIN_MENUS
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
        except Exception as e:
            logging.error(f"Error saving config on add_account_command: {e}")

    db_path = os.path.join(BASE_DIR, 'bot_data.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO accounts (menu_name, username, password) VALUES (?, ?, ?)", (menu_name, username, password))
        conn.commit()
        await update.message.reply_text(f"✅ បានបន្ថែមគណនីថ្មីក្នុង Menu <b>{menu_name}</b> ដោយជោគជ័យ!", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error adding account: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការបន្ថែមគណនី។")
    finally:
        conn.close()


# កែប្រែគណនីតាម Command /edit [ID] [Userថ្មី] [Passថ្មី]
async def edit_account_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if len(args) < 3:
        await update.message.reply_text("⚠️ ទម្រង់ខុស! សូមប្រើ៖ `/edit [លេខID] [Userថ្មី] [Passថ្មី]`", parse_mode="Markdown")
        return

    try:
        acc_id = int(args[0].strip())
    except ValueError:
        await update.message.reply_text("❌ លេខ ID ត្រូវតែជាលេខពិតប្រាកដ (ឧ. `/edit 1 newuser newpass`)", parse_mode="Markdown")
        return

    new_user, new_pass = args[1].strip(), args[2].strip()

    db_path = os.path.join(BASE_DIR, 'bot_data.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id FROM accounts WHERE id = ?", (acc_id,))
        if not cursor.fetchone():
            await update.message.reply_text(f"❌ រកមិនឃើញគណនី ID <code>{acc_id}</code> ឡើយ។", parse_mode="HTML")
            return

        cursor.execute("UPDATE accounts SET username = ?, password = ? WHERE id = ?", (new_user, new_pass, acc_id))
        conn.commit()
        await update.message.reply_text(f"📝 បានកែប្រែគណនី ID <code>{acc_id}</code> ទៅជា User: <b>{new_user}</b> / Pass: <b>{new_pass}</b> ដោយជោគជ័យ!", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error editing account {acc_id}: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការកែប្រែគណនី។")
    finally:
        conn.close()


# បន្ថែម Main Menu ថ្មី តាម Command /addmainmenu
async def add_main_menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if not args:
        context.user_data['state'] = 'WAITING_ADD_MAIN_MENU_NAME'
        await update.message.reply_text("➕ **សូមវាយផ្ញើឈ្មោះ Main Menu ថ្មីដែលអ្នកចង់បង្កើត៖**", parse_mode="Markdown")
        return

    new_main = args[0].strip()
    if new_main in MAIN_MENUS:
        await update.message.reply_text(f"❌ Main Menu <b>{new_main}</b> មានរួចហើយ។", parse_mode="HTML")
        return

    MAIN_MENUS[new_main] = []
    config_data["MAIN_MENUS"] = MAIN_MENUS
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
        await update.message.reply_text(f"✅ បានបន្ថែម Main Menu <b>{new_main}</b> ដោយជោគជ័យ! សូមចុច <code>/start</code> សារជាថ្មីដើម្បីមើលប៊ូតុង។", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error saving config.json on addmainmenu command: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")


# កែប្រែឈ្មោះ Main Menu តាម Command /editmainmenu
async def edit_main_menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if len(args) < 2:
        await update.message.reply_text("⚠️ ទម្រង់ខុស! សូមប្រើ៖ `/editmainmenu [ឈ្មោះចាស់] [ឈ្មោះថ្មី]`", parse_mode="Markdown")
        return

    target_main, new_name = args[0].strip(), args[1].strip()
    if target_main not in MAIN_MENUS:
        await update.message.reply_text(f"❌ មិនរកឃើញ Main Menu <b>{target_main}</b> ឡើយ។", parse_mode="HTML")
        return

    if new_name in MAIN_MENUS:
        await update.message.reply_text(f"❌ Main Menu <b>{new_name}</b> មានរួចហើយ។", parse_mode="HTML")
        return

    MAIN_MENUS[new_name] = MAIN_MENUS.pop(target_main)
    config_data["MAIN_MENUS"] = MAIN_MENUS
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
        await update.message.reply_text(f"📝 បានកែប្រែ Main Menu ពី <b>{target_main}</b> ទៅជា <b>{new_name}</b> រួចរាល់!", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error renaming main menu via command: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")


# លុប Main Menu តាម Command /delmainmenu
async def del_main_menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    args = context.args
    if not args:
        await update.message.reply_text("⚠️ ទម្រង់ខុស! សូមប្រើ៖ `/delmainmenu [ឈ្មោះMain]`", parse_mode="Markdown")
        return

    target_main = args[0].strip()
    if target_main not in MAIN_MENUS:
        await update.message.reply_text(f"❌ មិនរកឃើញ Main Menu <b>{target_main}</b> ឡើយ។", parse_mode="HTML")
        return

    subs = MAIN_MENUS[target_main]
    if subs:
        await update.message.reply_text(f"❌ មិនអាចលុប Main Menu <b>{target_main}</b> បានទេ ព្រោះវាមិនទាន់ទទេ (មាន Sub Menu ចំនួន <b>{len(subs)}</b>)។", parse_mode="HTML")
        return

    MAIN_MENUS.pop(target_main)
    if not MAIN_MENUS:
        MAIN_MENUS["ទូទៅ"] = []
    config_data["MAIN_MENUS"] = MAIN_MENUS
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
        await update.message.reply_text(f"🗑️ បានលុប Main Menu <b>{target_main}</b> រួចរាល់!", parse_mode="HTML")
    except Exception as e:
        logging.error(f"Error deleting main menu via command: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យ។")


# មគ្គុទ្ទេសក៍ណែនាំ តាម Command /help
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    help_text = (
        "📖 <b>មគ្គុទ្ទេសក៍ណែនាំសម្រាប់ការប្រើប្រាស់ Bot (Help & Commands)</b>\n\n"
        "🟢 <b>ពាក្យបញ្ជាសម្រាប់អ្នកប្រើប្រាស់ទូទៅ (General Commands):</b>\n"
        "• <code>/start</code> - ចាប់ផ្ដើម ឬ Refresh ប៊ូតុង Menu របស់ Bot\n"
        "• <code>/help</code> - មើលមគ្គុទ្ទេសក៍ណែនាំនេះ\n\n"
    )
    if user_id in ADMIN_IDS:
        help_text += (
            "🛠️ <b>ពាក្យបញ្ជាសម្រាប់ Admin (Admin Commands):</b>\n"
            "• <code>/add [menu_name] [user] [pass]</code> - បន្ថែមគណនីថ្មី\n"
            "• <code>/edit [ID] [user] [pass]</code> - កែប្រែគណនីតាម ID\n"
            "• <code>/sample</code> - ទាញយកឯកសារគំរូ Excel សម្រាប់បញ្ចូលគណនី\n\n"
            "📁 <b>ការគ្រប់គ្រង Main Menu:</b>\n"
            "• <code>/addmainmenu [name]</code> - បន្ថែម Main Menu ថ្មី\n"
            "• <code>/editmainmenu [old] [new]</code> - កែប្រែឈ្មោះ Main Menu\n"
            "• <code>/delmainmenu [name]</code> - លុប Main Menu\n\n"
            "📂 <b>ការគ្រប់គ្រង Sub Menu:</b>\n"
            "• <code>/addmenu [name]</code> - បន្ថែម Sub Menu ថ្មី\n"
            "• <code>/editmenu [old] [new]</code> - កែប្រែឈ្មោះ Sub Menu\n"
            "• <code>/delmenu [name]</code> - លុប Sub Menu\n"
            "• <code>/movemenu [sub] [main]</code> - បំលាស់ទី Sub Menu\n\n"
            "👤 <b>ការគ្រប់គ្រង Admin:</b>\n"
            "• <code>/addadmin [ID]</code> - បន្ថែម Admin ID ថ្មី\n"
            "• <code>/deladmin [ID]</code> - លុប Admin ID\n"
        )
    await update.message.reply_text(help_text, parse_mode="HTML")


# បង្កើតឯកសារគំរូ Excel សម្រាប់ Admin ទាញយក
async def sample_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    filepath = os.path.join(BASE_DIR, "sample_accounts.xlsx")
    try:
        create_sample_accounts_excel(filepath)
        
        with open(filepath, 'rb') as f:
            await update.message.reply_document(
                document=f, 
                filename="sample_accounts.xlsx", 
                caption="📋 នេះជាគំរូឯកសារ Excel (Sample Excel Template) សម្រាប់បញ្ចូលគណនី។\n\n💡 ចំណាំ៖ ប្រសិនបើអ្នកវាយឈ្មោះ Menu ថ្មីដែលមិនទាន់មានក្នុងប្រព័ន្ធ នោះប្រព័ន្ធនឹងបន្ថែមវាដោយស្វ័យប្រវត្តិ។"
            )
        
        if os.path.exists(filepath):
            os.remove(filepath)
    except Exception as e:
        logging.error(f"Error generating sample excel: {e}")
        await update.message.reply_text("❌ មានបញ្ហាក្នុងការបង្កើតឯកសារគំរូ។")


# អានឯកសារ Excel ឬ Backup ដែល Admin ផ្ញើមក ដើម្បីបញ្ចូល ឬស្តារទិន្នន័យ
async def handle_document_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global config_data, TOKEN, ADMIN_IDS, MENUS
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return

    document = update.message.document
    file_name = document.file_name.lower()

    # Check if we are waiting for a restore file
    if context.user_data.get('state') == 'WAITING_RESTORE_FILE':
        context.user_data['state'] = None
        
        if file_name.endswith('.db'):
            status_msg = await update.message.reply_text("⏳ កំពុងស្តារ Database ឡើងវិញ...")
            temp_db_path = os.path.join(BASE_DIR, 'temp_restore_db.db')
            try:
                file = await context.bot.get_file(document.file_id)
                await file.download_to_drive(temp_db_path)
                
                # Perform transactional and safe backup copy to keep active locks valid
                db_path = os.path.join(BASE_DIR, 'bot_data.db')
                src = sqlite3.connect(temp_db_path)
                dst = sqlite3.connect(db_path)
                with dst:
                    src.backup(dst)
                dst.close()
                src.close()
                
                # Cleanup temp file
                if os.path.exists(temp_db_path):
                    os.remove(temp_db_path)

                init_db()
                    
                await status_msg.edit_text("✅ បានស្តារ Database ឡើងវិញដោយជោគជ័យ! សូមចុច <code>/start</code> សារជាថ្មីដើម្បីមើលប៊ូតុង។", parse_mode="HTML")
            except Exception as e:
                logging.error(f"Error restoring database: {e}")
                if os.path.exists(temp_db_path):
                    os.remove(temp_db_path)
                await status_msg.edit_text("❌ មានបញ្ហាក្នុងការស្តារ Database ឡើងវិញ។")
            return
            
        elif file_name.endswith('.json'):
            status_msg = await update.message.reply_text("⏳ កំពុងស្តារ Config ឡើងវិញ...")
            temp_filepath = os.path.join(BASE_DIR, "temp_restore_config.json")
            try:
                file = await context.bot.get_file(document.file_id)
                await file.download_to_drive(temp_filepath)
                
                with open(temp_filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                if not isinstance(data, dict) or "TOKEN" not in data or "MENUS" not in data or "ADMIN_IDS" not in data:
                    if os.path.exists(temp_filepath):
                        os.remove(temp_filepath)
                    await status_msg.edit_text("❌ ឯកសារ Config មិនត្រឹមត្រូវ (ត្រូវមាន TOKEN, MENUS, និង ADMIN_IDS)។")
                    return
                
                config_path = os.path.join(BASE_DIR, 'config.json')
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2)
                
                if os.path.exists(temp_filepath):
                    os.remove(temp_filepath)
                
                # Update global config data
                config_data = data
                TOKEN = data.get("TOKEN", TOKEN)
                ADMIN_IDS = data.get("ADMIN_IDS", ADMIN_IDS)
                MENUS = data.get("MENUS", MENUS)
                MAIN_MENUS = data.get("MAIN_MENUS", {})
                if not MAIN_MENUS:
                    MAIN_MENUS = {"ទូទៅ": list(MENUS)}
                
                await status_msg.edit_text("✅ បានស្តារការកំណត់ (Config) ឡើងវិញដោយជោគជ័យ! សូមចុច <code>/start</code> សារជាថ្មីដើម្បីមើលប៊ូតុង។", parse_mode="HTML")
            except Exception as e:
                logging.error(f"Error restoring config: {e}")
                if os.path.exists(temp_filepath):
                    os.remove(temp_filepath)
                await status_msg.edit_text("❌ មានបញ្ហាក្នុងការស្តារការកំណត់ឡើងវិញ។")
            return
        else:
            await update.message.reply_text("❌ ឯកសារមិនត្រឹមត្រូវ! សូមផ្ញើតែឯកសារ `.db` ឬ `.json` សម្រាប់ស្តារឡើងវិញ។")
            return

    if not file_name.endswith(".xlsx"):
        return

    status_msg = await update.message.reply_text("⏳ កំពុងទាញយក និងអានឯកសារ Excel...")
    
    try:
        file = await context.bot.get_file(document.file_id)
        temp_filepath = os.path.join(BASE_DIR, f"temp_{document.file_name}")
        await file.download_to_drive(temp_filepath)
        
        success, msg, imported_menus = import_accounts_from_excel(temp_filepath)
        
        if os.path.exists(temp_filepath):
            os.remove(temp_filepath)
            
        if success:
            # Check if there are new menus to add to config.json
            new_added_menus = []
            for menu in imported_menus:
                if menu not in MENUS:
                    MENUS.append(menu)
                    new_added_menus.append(menu)
            
            if new_added_menus:
                if not MAIN_MENUS:
                    MAIN_MENUS["ទូទៅ"] = []
                first_key = list(MAIN_MENUS.keys())[0]
                for menu in new_added_menus:
                    MAIN_MENUS[first_key].append(menu)
                    
                config_data["MENUS"] = MENUS
                config_data["MAIN_MENUS"] = MAIN_MENUS
                try:
                    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                        json.dump(config_data, f, indent=2)
                    menu_msg = f"\n\n⚙️ ព្រមទាំងបានបន្ថែម Menu ថ្មី៖ " + ", ".join([f"<b>{m}</b>" for m in new_added_menus])
                except Exception as e:
                    logging.error(f"Error saving config.json from excel import: {e}")
                    menu_msg = "\n\n⚠️ មានបញ្ហាក្នុងការបន្ថែម Menu ថ្មីទៅក្នុង config.json"
            else:
                menu_msg = ""
                
            await status_msg.edit_text(f"✅ {msg}{menu_msg} ទៅក្នុង Database រួចរាល់! សូមចុច <code>/start</code> សារជាថ្មីដើម្បីមើលប៊ូតុង។", parse_mode="HTML")
        else:
            await status_msg.edit_text(f"❌ ការបញ្ចូលបរាជ័យ៖ {msg}")
    except Exception as e:
        logging.error(f"Error importing excel: {e}")
        await status_msg.edit_text("❌ មានបញ្ហាក្នុងការអាន ឬរក្សាទុកឯកសារ Excel របស់អ្នក។")


async def send_auto_backup(bot, reason=""):
    """
    Sends an automatic backup copy of bot_data.db to all ADMIN_IDS via Telegram.
    """
    try:
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        if not os.path.exists(db_path):
            return
            
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT count(*) FROM accounts")
        acc_count = cursor.fetchone()[0]
        conn.close()

        timestamp = get_phnom_penh_time_str()
        caption = (
            f"📦 <b>[AUTO BACKUP] ទិន្នន័យបម្រុងទុកស្វ័យប្រវត្តិ</b>\n"
            f"🕒 ពេលវេលា: <code>{timestamp}</code>\n"
            f"📊 ចំនួនគណនីសរុប: <b>{acc_count}</b> Accounts\n"
        )
        if reason:
            caption += f"📝 មូលហេតុ: {reason}\n"
        caption += "\n💡 <i>រក្សាទុកឯកសារ .db នេះ! ប្រសិនបើ Server restart ហើយបាត់ទិន្នន័យ អ្នកគ្រាន់តែផ្ញើឯកសារនេះចូល Bot វិញ រួចចុច Restore។</i>"

        for admin_id in ADMIN_IDS:
            try:
                with open(db_path, 'rb') as db_file:
                    await bot.send_document(
                        chat_id=admin_id,
                        document=db_file,
                        filename=f"bot_data_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
                        caption=caption,
                        parse_mode="HTML"
                    )
            except Exception as e:
                logging.error(f"Error sending auto backup to admin {admin_id}: {e}")
    except Exception as e:
        logging.error(f"Error in send_auto_backup: {e}")

async def periodic_auto_backup_loop(bot):
    """
    Periodic task that runs on startup and every 6 hours to backup database and notify admin.
    """
    await asyncio.sleep(5)  # Short delay on startup
    
    # Boot notification & initial check
    try:
        db_path = os.path.join(BASE_DIR, 'bot_data.db')
        acc_count = 0
        if os.path.exists(db_path):
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT count(*) FROM accounts")
            acc_count = cursor.fetchone()[0]
            conn.close()
            
        if acc_count == 0:
            alert_msg = (
                "⚠️ <b>[ALERT] Bot ត្រូវបាន Restart លើ Server ប៉ុន្តែ Database ទទេ (0 Accounts)!</b>\n\n"
                "💡 <b>របៀប Restore យកទិន្នន័យមកវិញ៖</b>\n"
                "1. ផ្ញើឯកសារ <code>bot_data.db</code> (Backup ចាស់) ចូលមកក្នុង Bot នេះ\n"
                "2. ឬចូល <code>🛠️ Admin Panel</code> ➡️ <code>📤 ស្តារឡើងវិញ (Restore)</code>\n"
                "ទិន្នន័យទាំងអស់នឹងត្រឡប់មកវិញភ្លាមៗ!"
            )
            for admin_id in ADMIN_IDS:
                try:
                    await bot.send_message(chat_id=admin_id, text=alert_msg, parse_mode="HTML")
                except Exception as e:
                    logging.error(f"Error sending boot alert to admin {admin_id}: {e}")
        else:
            # Initial auto backup on boot if accounts exist
            await send_auto_backup(bot, reason="Bot ត្រូវបាន Start / Restart ឡើងវិញ")
    except Exception as e:
        logging.error(f"Error in initial boot check: {e}")

    # Loop every 6 hours
    while True:
        await asyncio.sleep(6 * 3600)
        try:
            await send_auto_backup(bot, reason="តាមកាលវិភាគ 6 ម៉ោងម្តង")
        except Exception as e:
            logging.error(f"Error in periodic_auto_backup_loop: {e}")

async def post_init(application: Application) -> None:
    try:
        await application.bot.set_my_commands([
            BotCommand("start", "ចាប់ផ្ដើមដំណើរការ Bot (Start Bot)")
        ])
        logging.info("Successfully set bot commands.")
    except Exception as e:
        logging.error(f"Error setting bot commands: {e}")
        
    # Start periodic auto-backup background task
    asyncio.create_task(periodic_auto_backup_loop(application.bot))

# ----------------- MAIN RUNNER -----------------
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain")
        self.end_headers()
        self.wfile.write(b"Bot is running OK")
    def log_message(self, format, *args):
        return

def start_health_check_server():
    port = int(os.environ.get("PORT", 8080))
    try:
        server = HTTPServer(("0.0.0.0", port), HealthCheckHandler)
        logging.info(f"Health check HTTP server listening on port {port}")
        server.serve_forever()
    except Exception as e:
        logging.error(f"Error starting health check server: {e}")

def main():
    if os.environ.get("PORT"):
        threading.Thread(target=start_health_check_server, daemon=True).start()
        
    init_db()
    
    # Use configuration-defined token
    application = Application.builder().token(TOKEN).post_init(post_init).build()
    
    # Commands & Messages Handlers
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("add", add_account_command))
    application.add_handler(CommandHandler("edit", edit_account_command))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("addmenu", add_menu_command))
    application.add_handler(CommandHandler("delmenu", del_menu_command))
    application.add_handler(CommandHandler("editmenu", edit_menu_command))
    application.add_handler(CommandHandler("addmainmenu", add_main_menu_command))
    application.add_handler(CommandHandler("delmainmenu", del_main_menu_command))
    application.add_handler(CommandHandler("editmainmenu", edit_main_menu_command))
    application.add_handler(CommandHandler("movemenu", move_menu_command))
    application.add_handler(CommandHandler("sample", sample_command))
    application.add_handler(CommandHandler("addadmin", add_admin_command))
    application.add_handler(CommandHandler("deladmin", del_admin_command))
    
    application.add_handler(MessageHandler(
        filters.Document.FileExtension("xlsx") | 
        filters.Document.FileExtension("db") | 
        filters.Document.FileExtension("json"), 
        handle_document_upload
    ))
    application.add_handler(MessageHandler(filters.Text("🛠️ Admin Panel"), admin_panel))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_menu_click))
    
    # Callback Query Handler សម្រាប់ប៊ូតុងចុច
    application.add_handler(CallbackQueryHandler(button_click))
    
    print("Bot ពេញលេញកំពុងដំណើរការ...")
    application.run_polling()

if __name__ == '__main__':
    main()