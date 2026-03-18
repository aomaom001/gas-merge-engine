# GAS Promotion Merger — User Guide

> **Note:** This guide covers Tab 1 and Tab 2 only. Tab 3 is currently under development.

---

## Table of Contents

- [Deployment & Setup](#deployment--setup)
- [Login](#login)
- [Tab 1 — AI Chat Bot](#tab-1--ai-chat-bot)
- [Tab 2 — Manage / Merge Files](#tab-2--manage--merge-files)
- [Auto-filter Rules](#auto-filter-rules)

---

## Deployment & Setup

Follow these steps to deploy this project to a new Google Apps Script account.

### Step 1 — Copy the project files

Copy all files into the new Google Apps Script project:

| File | Type |
|---|---|
| `Code.gs` | Apps Script (server-side) |
| `Index.html` | HTML template (Web App UI) |

### Step 2 — Set Script Properties

Go to **Project Settings → Script Properties** and add the following keys:

| Key | Value |
|---|---|
| `USERNAME` | Login username (e.g. `admin`) |
| `PASSWORD` | Login password |
| `CLAUDE_API_KEY` | Your Anthropic Claude API key |
| `FOLDER_ID` | Google Drive folder ID where source files are stored |

> To get a folder ID: open the folder in Google Drive and copy the ID from the URL — the part after `/folders/`.

### Step 3 — Deploy as Web App

1. Click **Deploy → New deployment**
2. Select type: **Web App**
3. Set **Execute as:** `Me`
4. Set **Who has access:** `Anyone` (or restrict as needed)
5. Click **Deploy** and copy the Web App URL

### Step 4 — First run

Open the Web App URL in a browser, log in, then click **🔄 Sync** on Tab 1 or Tab 2 to load files from Drive.

---

## Login

Open the Web App URL and enter the username and password set in Script Properties.

> After 5 consecutive failed login attempts, the account is locked for 5 minutes.

---

## Tab 1 — 💬 AI Chat Bot

Use this tab to **ask questions about promotion data**. The AI reads the selected files and sheets, then answers in Thai (or the language of your question).

### Workflow

```
1. Sync  →  2. Select Files  →  3. Select Sheets  →  4. Type Question  →  5. Send
```

---

### Step 1 — Sync Files

Click **🔄 Sync** (bottom of the left sidebar) to fetch the latest file list from Google Drive.

> Re-sync whenever files are added or removed from Drive.

---

### Step 2 — Select Files

- Check ✅ one or more files for the AI to read
- Use the **Search files...** box to filter the list
- Click **☑️ Select / Deselect All** to toggle all at once

---

### Step 3 — Select Sheets

After selecting files, the sheet list loads automatically below.

- Check ✅ one or more sheets to include
- Use the **Search Sheet...** box to filter
- Click **☑️ Select / Deselect All Sheets** to toggle all at once

---

### Step 4 — Ask the AI

Type your question in the input bar at the bottom, then press **Send ➤** or hit `Enter`.

**Example questions:**
- `What iPhone promotions are available?`
- `What is the lowest price for Samsung?`
- `How many promotions have a 24-month contract?`
- `Summarize the maximum discount per brand`

---

### Additional Features

| Button / Feature | Description |
|---|---|
| 📋 Copy (top-right of each AI reply) | Copy the AI's answer to clipboard |
| 🗑️ Clear | Clear the entire chat history |
| 🌙 / ☀️ (top-right of nav bar) | Toggle Dark / Light Mode |

> **Note:** If the AI hits a rate limit, a countdown banner will appear and the request will retry automatically after 60 seconds.

---

## Tab 2 — 📦 Manage / Merge Files

Use this tab to **merge promotion data from multiple files** into a single Google Sheet, with strikethrough rows and empty rows filtered out automatically.

### Workflow

```
1. Select Files  →  2. Configure Output  →  3. Map Columns  →  4. Preview  →  5. Confirm Merge
```

---

### Step 1 — Select Files & Sync

**Sync Excel files (if .xlsx files are in Drive):**
Click **🔄 Sync Excel → Sheets** to automatically convert Excel files to Google Sheets.

**Select files:**
- Check ✅ one or more files to merge
- Use the **Search files...** box to filter
- Click 🔄 next to the section heading to refresh the file list

After selecting files, the system loads the available sheets for each file automatically.

---

### Step 2 — Configure Output

| Field | Description |
|---|---|
| **📁 Save to Folder** | Choose the destination folder in Drive |
| **＋** (next to folder dropdown) | Create a new folder directly |
| **✏️ Output File Name** | Name for the merged file to be created |

Then click **🚀 Start Advanced Merge** to proceed to column mapping.

---

### Step 3 — Map Columns

This screen shows **template columns** (left side) and **source columns from your files** (displayed as chips).

**How to map:**
- Click a source column chip and drag it into the desired template column slot
- One template column can accept multiple source columns — the system uses the first non-empty value

**Default Values:**
If a template column has no matching source column, a yellow warning banner appears — fill in a default value or leave it blank.

**Mapping Profiles (save your settings):**

| Button | Description |
|---|---|
| **-- Select Profile --** dropdown | Load a previously saved mapping profile |
| **✅ Apply** | Apply the selected profile |
| **🗑️** | Delete the selected profile |
| **Profile Name...** field + **💾 Save** | Save the current mapping as a new profile |
| **🔄 Reset** | Clear all mappings back to default |

> **Tip:** Save a Mapping Profile after the first setup — you won't need to reconfigure it next time.

---

### Step 4 — Preview

Click **🔍 Preview Before Merge** to see a sample of the output data.

- Shows up to **50 rows** (15 rows per file)
- Strikethrough rows are excluded and will not appear
- Rows with no **Brand & Model** value are filtered out
- Empty cells are highlighted with a light yellow background
- A yellow banner at the top shows the number of skipped rows

Click **⬅️ Back to Edit Mapping** to adjust mappings if needed.

---

### Step 5 — Confirm Merge

Click **🚀 Confirm Merge** to create the output file.

When complete, a summary is shown, for example:
```
Merged successfully: 320 rows  |  File: Promotion_March2026
Summary: 380 total rows | 45 strikethrough removed | 15 missing Brand/Model
```

A **🔗 Open in Google Sheets** link appears to open the result file immediately.

---

## Auto-filter Rules

The system automatically removes the following rows during merge:

| Type | Description |
|---|---|
| Strikethrough rows | Any row where text has strikethrough formatting is excluded entirely |
| Empty rows | Rows where every cell is blank |
| Missing Brand/Model | Rows where the "Brand & Model" column is empty |
