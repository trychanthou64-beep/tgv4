# 📖 មគ្គុទ្ទេសក៍ណែនាំសម្រាប់ការប្រើប្រាស់ Telegram Bot & Admin Panel

ផ្ទាំងគ្រប់គ្រង **🛠️ Admin Panel** មានមុខងារពេញលេញសម្រាប់គ្រប់គ្រង Telegram Bot៖

---

### 1. ➕ CREATE (បន្ថែមគណនី)
* **តាម Command:** `/add [MenuName] [Username] [Password]`
  *(ឧទាហរណ៍៖ `/add F88$ user123 pass456`)*
* **តាម Excel File:** ផ្ញើឯកសារ Excel (`.xlsx`) មកកាន់ Bot នោះ Bot នឹងបញ្ចូលទិន្នន័យដោយស្វ័យប្រវត្តិ។ ចុច `/sample` ឬប៊ូតុង **📥 ទាញយកគំរូ Excel** ដើម្បីទាញយកឯកសារគំរូ។

---

### 2. 📊 READ (មើលទិន្នន័យទាំងអស់)
* ទាញយកឯកសារ Excel តាមរយៈការផ្ញើសារ `/sample` ឬចុចប៊ូតុង **📥 ទាញយកគំរូ Excel**។ ឯកសារនេះរួមមាន 3 Sheets:
  - **`Data_Entry`**: ទម្រង់សម្រាប់បញ្ចូលគណនីថ្មី។
  - **`menu_name`**: បញ្ជី Menu និងចំនួនគណនីសរុប។
  - **`All Data`**: បង្ហាញរាល់ទិន្នន័យគណនីទាំងអស់ដែលមានក្នុងប្រព័ន្ធ។

---

### 3. 📝 UPDATE (កែប្រែទិន្នន័យ)
* **តាម Command:** `/edit [លេខID] [Userថ្មី] [Passថ្មី]`
  *(ឧទាហរណ៍៖ `/edit 1 newuser newpass`)*

---

### 4. 🗑️ DELETE (លុបទិន្នន័យ)
* ចុចប៊ូតុង **🗑️ លុបគណនីទាំងអស់ (Delete All)** ដើម្បីលុបទិន្នន័យគណនីទាំងអស់ចេញពីប្រព័ន្ធ។

---

### 5. 📁 MENUS MANAGEMENT (គ្រប់គ្រង Menu)

**ក. Main Menu (ផ្ទាំងធំ):**
* បន្ថែម៖ `/addmainmenu [ឈ្មោះMain]` (ឧទាហរណ៍៖ `/addmainmenu Casino`)
* កែប្រែ៖ `/editmainmenu [ឈ្មោះចាស់] [ឈ្មោះថ្មី]` (ឧទាហរណ៍៖ `/editmainmenu Sports Casino`)
* លុប៖ `/delmainmenu [ឈ្មោះMain]` (ឧទាហរណ៍៖ `/delmainmenu Casino`)

**ខ. Sub Menu (ប៊ូតុងខាងក្នុង):**
* បន្ថែម៖ `/addmenu [ឈ្មោះSub]` (ឧទាហរណ៍៖ `/addmenu Bet99`)
* កែប្រែ៖ `/editmenu [ឈ្មោះចាស់] [ឈ្មោះថ្មី]` (ឧទាហរណ៍៖ `/editmenu F88 F88_NEW`)
* លុប៖ `/delmenu [ឈ្មោះSub]` (ឧទាហរណ៍៖ `/delmenu Bet99`)
* បំលាស់ទី៖ `/movemenu [ឈ្មោះSub] [ឈ្មោះMain]` (ឧទាហរណ៍៖ `/movemenu Bet99 Casino`)

---

### 6. 👤 ADMIN MANAGEMENT (គ្រប់គ្រង Admin)
* បន្ថែម Admin ID ថ្មី៖ `/addadmin [លេខID]` (ឧទាហរណ៍៖ `/addadmin 123456789`)
* លុប Admin ID ចេញ៖ `/deladmin [លេខID]` (ឧទាហរណ៍៖ `/deladmin 123456789`)
*(ចំណាំ៖ ID `8558847170` គឺជា Default Admin ហើយមិនអាចលុបបានឡើយ)*

---

### 7. 📜 HISTORY MANAGEMENT (ប្រវត្តិទាញយក)
* ចុចប៊ូតុង **📜 ប្រវត្តិទាញយក (History)** ក្នុង Admin Panel ដើម្បីមើល ឬទាញយកប្រវត្តិការទាញយកគណនីរបស់ User ទាំងអស់ជា Excel។

---

### 8. 📦 BACKUP & RESTORE (បម្រុងទុក និងស្តារ)
* **បម្រុងទុក (Backup):** ចុចប៊ូតុង **📥 បម្រុងទុក (Backup)** ដើម្បីទាញយកឯកសារ Database (`bot_data.db`), Accounts (`backup_accounts.xlsx`), និង Configuration (`config.json`)។
* **ស្តារឡើងវិញ (Restore):** ចុច **📤 ស្តារឡើងវិញ (Restore)** រួចផ្ញើឯកសារ `.db` ឬ `.json` មកកាន់ Bot។
