# Outlook Email Cleanup Macro (First 150 Permanent Delete)

### ⚠️ Important: Classic Outlook Only
This macro **only works in Classic Outlook for Windows** (the version with the full ribbon and yellow envelope icon).  
It **will not work** in the “New Outlook” or in Outlook on the web.

---

## 💡 What this tool does

This small macro helps you safely clean up a busy Outlook folder by permanently deleting up to **150 of the oldest emails** received before a specific date.

It’s designed for quick clean-ups — for example, removing messages received before **30 October 2023** from your **Inbox** or another folder you select.

---

## 🧭 How it works

1. You select the folder you want to clean up (for example, *Inbox*, *Sent Items*, or a project folder).  
2. You press **Alt + F8** to run the macro.  
3. It asks you to enter a **date** (e.g. `30/10/2023` or `2023-10-30`).  
4. The macro looks through that folder and finds **emails received before** that date.  
5. It identifies the **oldest 150 emails** that meet the rule.  
   - If there are fewer than 150, it shows the exact number found.  
   - If there are more than 150, it says “first 150” and works on those.  
6. It shows you a confirmation message explaining exactly what will happen.  
7. If you click **Yes**, it **permanently deletes** those emails (they do **not** go to Deleted Items).  
8. If you click **No**, the macro stops and nothing is deleted.

---

## 🧩 Rules & Logic

- Works only on the **currently selected folder** (no subfolders).  
- Deletes **MailItems only** (not calendar invites, tasks, or reports).  
- Uses the email’s **Received Date** (not Sent Date).  
- Only emails **before** the date you enter are included.  
- Deletes up to **150 oldest matching emails** — never more.  
- All deletions are **permanent** (they bypass the Deleted Items folder).  
- Works the same in personal mailboxes and **shared mailboxes**, as long as you can open that folder in Outlook.  

---

## 🧰 How to install

1. Open **Classic Outlook for Windows**.  
2. Press **Alt + F11** to open the *Microsoft Visual Basic for Applications (VBA)* window.  
3. In the VBA editor:
   - Select **Insert → Module**.  
   - A new blank code window will appear.  
4. Open the file **`outlookMacro.vb`** from this GitHub repository.  
5. Copy everything in that file and paste it into the blank code window.  
6. Close the VBA editor (click the X in the corner or press **Alt + Q**).  

You’ve now installed the macro locally.

---

## ▶️ How to run the macro

1. In Outlook, go to the folder you want to process.  
2. Press **Alt + F8** on your keyboard.  
3. In the list that appears, select: **Purge_ByReceivedDate_First150**
4. Click **Run**.  
5. When prompted, type the date (e.g. `30/10/2023` or `2023-10-30`).  
- You can use either **day/month/year** or **year-month-day** format.  
- If your computer is set to U.S. format, it will also accept `10/30/2023`.  
6. Read the summary message carefully.  
7. Click **Yes** to confirm permanent deletion, or **No** to cancel.

---

## 📨 Using with shared mailboxes

You can use this macro in a shared mailbox as long as you have permission to open and view its folders.

To do this:
1. In Outlook’s folder pane, navigate to the shared mailbox.  
2. Click the folder you want to clean up (for example, *Inbox* or *Archive*).  
3. Run the macro using **Alt + F8** and follow the same steps.  

The script will only affect the folder you have open when you start it.

---

## ⚖️ Safety notes

- **Permanent means permanent.** Deleted emails do not go to the Deleted Items folder and can’t be recovered from there.  
- Always double-check the confirmation message before pressing **Yes**.  
- If you’re unsure, test the macro first in a low-risk folder.  
- The macro has no effect on other folders or accounts.  
- Outlook must remain open while it runs — don’t switch accounts or folders until it finishes.

---

## 🚧 Known limitations

| Limitation | Details |
|-------------|----------|
| **New Outlook** | The new Outlook app (with the toggle) doesn’t support macros. You must switch back to Classic Outlook. |
| **Date handling** | Outlook filters by local date/time. Machines set to U.S. format may interpret `10/11/2023` as 11 October instead of 10 November. Use `2023-10-11` to avoid confusion. |
| **Subfolders** | The macro only processes the selected folder. It doesn’t scan subfolders. |
| **Other item types** | It ignores calendar events, tasks, and non-mail items. |
| **Undo** | There is no “undo” after confirming deletion. |

---

## ✅ Example

If you select your **Inbox** and run the macro with a cutoff date of **30/10/2023**:

- The macro checks every email in the Inbox.  
- Finds all with a *Received Date* earlier than **30 October 2023**.  
- If it finds 92 matching emails, it will show “Found 92 items”.  
- If it finds 324, it will show “Found more than 150 — deleting the first 150 (oldest)”.  
- If you confirm “Yes”, those emails are permanently deleted.

---


