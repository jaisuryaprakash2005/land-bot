# Probridge Land Bot - Setup Guide

Paste or speak a mediator's land offer → AI fills your **Land Acquisition** columns → you check → **Save to Excel**.
Works inside **Excel on your laptop** and as an **app on your phone**. Cost: ₹0. All AI runs on your laptop.

```
Laptop:  Excel ── "Land Bot" panel ─┐
                                    ├─► Laptop server (start.ps1) ─► Whisper (voice→text)
Phone:   Land Bot app ── Tailscale ─┘                              └► Ollama Llama (text→columns)
Saving:  Excel panel → Office.js → open workbook
         Phone app   → Microsoft Graph → same workbook in OneDrive
```

## What's in this folder

| File | What it is |
|---|---|
| `config.json` | All settings: AI model, Excel table name, OneDrive path, columns, Status choices |
| `setup.ps1` | One-time setup (packages, AI models, certificate, Excel button) |
| `start.ps1` | Starts the bot. Keep its window open while you use it |
| `server/server.py` | Laptop server: voice → text, text → columns |
| `web/` | The chat screen (used by Excel and the phone) |
| `addin/manifest.xml` | Tells Excel about the "Land Bot" button |
| `scripts/autostart-on.ps1` | Optional: start the bot automatically at Windows login |

---

## Step 1 - Prepare the Excel sheet (5 min)

1. Open your Land Acquisition workbook.
2. **Rename the two "Status" headers** so they're different. Excel tables need unique names.
   For example, `Status` and `Status 2` (or `Deal Status` and `Follow-up Status`).
   The bot fills the **first** one from a dropdown (Prospect, Site visit…) and leaves the **second** one for you to type.
3. Click any cell in your data and press **Ctrl+T**, tick **My table has headers**, then click **OK**.
4. In **Table Design → Table Name** (top-left), type `Properties` and press Enter.
5. *(Optional)* Make "Prospect" turn green automatically: select the Status column, then go to
   **Home → Conditional Formatting → Highlight Cells Rules → Text that Contains** and type `Prospect` → choose the green fill.
6. Save the workbook in **OneDrive** as `Probridge/Land_Acquisition.xlsx`.
   If you use a different folder or name, change `onedrive_workbook_path` in `config.json`.

The bot matches columns by **header name**, so the column order doesn't matter.
S.No (next number) and Date (today) are filled automatically.

## Step 2 - One-time setup on the laptop (20-40 min, mostly downloads)

Open PowerShell **in this folder** and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

It will:
1. Install the Python packages.
2. Download the AI model `llama3.1:8b` (~4.9 GB).
3. Download the voice model Whisper `small` (~500 MB).
4. Install a **localhost HTTPS certificate**. Windows shows a security prompt to trust
   "Developer CA for Microsoft Office Add-ins". Click **Yes**, because Excel add-ins only load over https.
5. Add the **Land Bot** button to Excel.

## Step 3 - Start the bot

```powershell
powershell -ExecutionPolicy Bypass -File .\start.ps1
```

Keep that window open. To start the bot automatically every time you log in, run `scripts\autostart-on.ps1` once.

## Step 4 - Use it in Excel (laptop)

1. Close Excel completely, then open your workbook again.
2. Go to **Home → Land Bot** (far right of the ribbon).
   If the button isn't there, look under **Home → Add-ins → Developer Add-ins → Land Bot**.
3. The panel opens and says **Connected to table "Properties" (N rows)**.
4. Paste a mediator message (or 🎤 record, or 📎 attach a WhatsApp voice note), then click **Read message**.
5. Check the fields. Fields with an **orange** border need a look, and **AI** / **You** tags show who filled each one.
6. Click **Save to Excel**. A new row is added and selected.

The green dot at the top means the AI is ready, and a red dot means Ollama isn't running.

---

## Step 5 - Reach the laptop from your phone (Tailscale, free)

1. Install **Tailscale** on the laptop (tailscale.com/download) and on your phone (Play Store / App Store).
   Sign in with the **same account** on both.
2. In the Tailscale admin console (login.tailscale.com → **DNS**), turn on **MagicDNS** and **HTTPS Certificates**.
3. On the laptop, run this once (it keeps working after restarts):
   ```powershell
   tailscale serve --bg https+insecure://localhost:8443
   ```
4. It prints your private address, something like `https://jaisu-laptop.tail1234.ts.net/`.
   Your phone app address is that plus `app/`, for example `https://jaisu-laptop.tail1234.ts.net/app/`.
   Only devices signed in to *your* Tailscale can open it.

## Step 6 - Let the phone app save to Excel (Microsoft sign-in, free)

1. Go to **portal.azure.com → Microsoft Entra ID → App registrations → New registration**.
   - Name: `Probridge Land Bot`
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
   - Redirect URI: choose **Single-page application (SPA)**, then paste your phone app address from Step 5
     (e.g. `https://jaisu-laptop.tail1234.ts.net/app/`). Keep the `/app/` at the end.
2. Click **Register**. Copy the **Application (client) ID**.
3. Go to **API permissions → Add a permission → Microsoft Graph → Delegated** and add `Files.ReadWrite` and `User.Read`.
4. Open `config.json` and paste the ID:
   ```json
   "client_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
   ```
5. Restart `start.ps1`.

## Step 7 - Install the app on your phone

1. Make sure Tailscale is **on** on the phone.
2. Open your phone app address in **Chrome** (Android) or **Safari** (iPhone).
3. Tap **Sign in** and use the Microsoft account that owns the OneDrive workbook.
4. Add the app to your home screen. On Android, tap **⋮ → Add to Home screen / Install app**. On iPhone, tap **Share → Add to Home Screen**.
5. **Android bonus:** in WhatsApp, long-press a mediator's message or voice note, tap **Share → Land Bot**, and it opens already being read.

---

## Daily use

1. The laptop is on and `start.ps1` is running (or autostart is on).
2. A mediator sends details → paste, record, or share them into Land Bot.
3. Check the preview. Fix orange fields, set Status, add Previlling or Status 2 if needed.
4. **Save to Excel.**

The bot warns you about:
- **Possible duplicate:** the same Google pin, or the same area and mediator number, is already in the sheet.
- **TOTAL price:** the mediator gave a total price, or a per-cent price, instead of a per-acre price.
- **AI may have guessed:** a value isn't found in the message.
- **'Nil' negatives:** shown only when the mediator clearly said there are none.

The phone number and Google pin are copied **exactly** from the message by plain code, not by the AI.
If a mediator's number isn't in the message, the bot fills it from their earlier rows.

## Changing things later (`config.json`)

| Want to… | Change |
|---|---|
| Rename a column in Excel | Update its `"header"` in `columns` |
| Change Status dropdown choices | `status_options` |
| Faster but less accurate AI | `"ollama_model": "llama3.2:3b"`, then `ollama pull llama3.2:3b` |
| Better Tamil voice accuracy (slower) | `"whisper_model": "medium"` |
| Keep voice text in Tamil instead of English | `"whisper_task": "transcribe"` |
| Voice notes misspell a village or mediator name | Add the correct spelling to `whisper_hint` |

Restart `start.ps1` after any change.

## Troubleshooting

| Problem | Fix |
|---|---|
| Panel says "Cannot reach the laptop server" | Start `start.ps1` |
| Red dot: "Ollama is not running" | Open the Ollama app from the Start menu |
| "Model … not installed" | `ollama pull llama3.1:8b` |
| Excel panel is blank / "add-in error" | `start.ps1` must be running **before** you click Land Bot; re-run `setup.ps1` step 4 if the certificate is missing |
| "Excel table "Properties" not found" | Step 1, parts 3-4 |
| Phone: "Phone saving is not set up yet" | Step 6 (client ID) |
| Phone sign-in error about redirect URI | The Redirect URI in Azure must exactly match the app address, including `/app/` |
| Phone: 404 when saving | Check `onedrive_workbook_path` (folder and file name, case matters) |
| Reading takes 30-90 seconds | Normal on this laptop without a graphics card. The first message after start is slowest |
| Microphone doesn't work inside Excel | Use 📎 to attach the voice note file instead |
