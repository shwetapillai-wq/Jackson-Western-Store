# JW Store: dashboard and weekly client email

Both tools read only the "Daily Data" tab of the JW Store tracker (Google Sheet version).

## Before you start (both tools)
1. Open the tracker in Google Drive. Use File > Save as Google Sheets so it is a real Google Sheet (title without .xlsx).
2. Keep the header names in Daily Data as they are: Date, Name, SKUs Worked, SKUs Completed, Hours, Comments.
3. Format the Date column as an unambiguous date, for example 24-Oct-2026 (Format > Number > Custom date and time).

## 1. Dashboard (dashboard folder)
1. Share the sheet: Share > General access > Anyone with the link > Viewer. This shares ALL tabs of the sheet with anyone who has the link.
2. Note the Sheet ID (the long code between /d/ and /edit) and the gid number of the Daily Data tab (after gid= in the address bar while that tab is open).
3. On GitHub create a NEW repository (for example jw-store-dashboard) and upload app.py and requirements.txt. Make sure the file is named exactly app.py, not app (1).py.
4. On share.streamlit.io click Create app, choose the new repository, main file app.py. In Advanced settings > Secrets paste:
       SHEET_ID = "your-sheet-id"
       GID = "your-daily-data-gid"
5. Deploy. The dashboard re-reads the sheet every 5 minutes.

## 2. Weekly email (weekly-email folder)
1. In the same Google Sheet open Extensions > Apps Script, paste in Code.gs and Save. In Project Settings set the time zone to Asia/Kolkata.
2. Run setupSheets (approve permissions). Fill the Client Config tab: client name, To emails, CC emails, Active = Yes, Show Daily Notes = No or Yes.
3. Run sendTestToMe. The email comes only to you, marked [TEST].
4. Change TEST_MODE to false near the top of the code, then run createWeeklyTrigger once. The email goes out every Monday around 9 AM for the last full Monday to Sunday week.
5. The Email Log tab records every send, skip and error.

## 3. Dev and DM teams: upload files, no typing (new)
Dev and DM keep sending the same files they do today. They just upload them to one shared Google Drive folder.

**One-time setup**
1. In Google Drive create a folder such as "JWS Weekly Reports" and share it with the Dev and DM teams (Editor). Copy the folder ID (the part of the folder's web address after /folders/).
2. In the Apps Script project (Extensions > Apps Script) click the + next to Files and add a script file named Files, then paste in Files.gs. Paste the folder ID at the top (FOLDER_ID).
3. Click the + next to Services, choose Drive API, version v3, and click Add.
4. Run previewFolderFiles. Check the Execution log: it shows the numbers the script reads from each file in the folder. Nothing is written yet. Compare with the real files.
5. Run importNewFiles once to add the numbers to the "Team Updates" and "SEO Keywords" tabs (created automatically). Run createWeeklyTrigger again so the daily import is scheduled (every day around 7 PM).

**What the teams do each week**
- Upload the Dev Excel and the DM PDF to the folder. Keep "Dev" in the Dev file's name. Any PDF in the folder is treated as the DM report.
- Files uploaded Thursday to Sunday count for that week. Files uploaded Monday to Wednesday count for the week before. This is how the script decides which week a file belongs to.

**What happens next**
- Each day the script reads new files and adds their numbers to the tabs. The dashboard's "Dev & SEO" tab shows them (add GID_UPDATES and GID_KEYWORDS in Streamlit Secrets, using the gid of those two tabs).
- Every Monday the client email includes the Dev and DM sections and attaches the original files for that week.
- The Email Log tab records every file read, with how many rows were found, or why a file could not be read.

**Good to know**
- Dev: reads the work descriptions and the "Total: N hours" line.
- DM: reads any bullet like "90 Backlinks Created" as a counted activity, so new activity names are picked up. It also reads links, short titles under a bullet, and the keyword ranking table. If the layout changes a lot, some numbers may be missed, but the original file is still attached.
- Always run previewFolderFiles on a new type of file before the first live email.
- You can still type rows into the Team Updates tab by hand. The script will not overwrite them.
