# Quick Sheet: from zip to Microsoft store

Quick Sheet adds a new worksheet. There's a **New Sheet** button on the Home tab, and a **Named Sheet** pane where you type the name first.

The folder contains two versions:

| File | What it is | Can it go in the store? |
|---|---|---|
| `NewSheet.bas` | VBA macro, for your own workbooks | **No.** The store doesn't accept VBA or Excel files. |
| `manifest.xml` + `*.html` + `quicksheet.js` + `assets/` | Office.js add-in, the same feature as a web add-in | **Yes.** This is what you submit. |

No Node.js or build step is needed. These are plain files that you put on a website.

---

## Part 1: Put the files online (GitHub Pages, free, about 10 min)

The store doesn't host your add-in. It just points at your HTTPS address.

1. Create a free account at **github.com**, or sign in. Note your **username**.
2. Click **+** (top right), then **New repository**.
   - Repository name: `quicksheet`
   - Select **Public**
   - Click **Create repository**
3. On the next page, click the **uploading an existing file** link.
4. Unzip `quicksheet.zip`. Open the `quicksheet` folder and drag **everything inside it** (including the `assets` folder) onto the GitHub page. Click **Commit changes**.
5. Go to **Settings**, then **Pages** (left menu).
   - Source: **Deploy from a branch**
   - Branch: **main**, folder **/(root)**, then **Save**
6. Wait 1–2 minutes, then open `https://YOUR-USERNAME.github.io/quicksheet/support.html` in a browser. If you see the support page, hosting works.

## Part 2: Point the manifest at your site

1. Open `manifest.xml` in **Notepad**. Right-click it and choose **Open with**, then **Notepad**.
2. Press **Ctrl+H**.
   - Find: `YOUR-GITHUB-USERNAME`
   - Replace: your GitHub username in **lower case**
   - Click **Replace All**, then save.
3. Upload the updated `manifest.xml` to the GitHub repo too, so your online copy matches: **Add file**, then **Upload files**, then **Commit**.

## Part 3: Test it in Excel (sideload)

The easiest place to test is Excel for the web.

1. Go to **excel.cloud.microsoft** (or office.com, then Excel) and open or create any workbook.
2. Go to **Home**, then **Add-ins**, then **More Add-ins**. Choose the **My Add-ins** tab, then **Upload My Add-in**. Pick your `manifest.xml` and click **Upload**.
3. On the Home tab you now have a **Quick Sheet** group:
   - **New Sheet** should add "New Sheet". Click it again to get "New Sheet (2)".
   - **Named Sheet** should open the pane. Type a name and click **Create sheet**.
4. Take your store screenshots now:
   - Use Snipping Tool (**Win+Shift+S**) and paste into **Paint**.
   - Go to **Resize**, choose **Pixels**, untick *Maintain aspect ratio*, and set **1366 × 768**.
   - Save as PNG. Take 1–3: the ribbon buttons, the pane, and a created sheet.

If you have Node.js installed, you can also run `npx office-addin-manifest validate manifest.xml` in that folder to check the manifest against Microsoft's rules.

## Part 4: Partner Center account (one-off, can take a few days)

1. Go to **partner.microsoft.com/dashboard** and sign in with your **work** account (mike@mantisnetworks.co).
2. Click the gear icon, then **Account settings**, then **Programs**. Find **Microsoft 365 and Copilot** and click **Get started**.
3. Fill in the company details exactly as registered (Mantis Networks (Pty) Ltd) and complete the email and business verification. Microsoft may ask for company documents. Wait for the "verified" status before submitting.

## Part 5: Create and submit the offer (about 30 min)

In Partner Center, go to **Marketplace offers**, open the **Microsoft 365 and Copilot** tab, click **+ New offer** and choose **Office add-in**.

1. **Offer name:** `Quick Sheet`. It must match `DisplayName` in the manifest.
2. **Packages:** upload `manifest.xml`.
3. **Properties:**
   - Category: *Productivity*
   - Privacy policy URL: `https://YOUR-USERNAME.github.io/quicksheet/privacy.html`
   - License terms: choose the **Standard Contract**
4. **Marketplace listing** (English):
   - Summary: *Add a new, uniquely named worksheet in one click.*
   - Description: what it does, how to use it (copy from `support.html`), and that it doesn't collect data.
   - Search keywords: *new sheet, worksheet, tab*
   - Support URL: `https://YOUR-USERNAME.github.io/quicksheet/support.html`
   - Logo: `assets/icon-300.png`
   - Screenshots: your 1366 × 768 PNGs from Part 3
5. **Availability:** all markets, release as soon as approved.
6. **Review and publish.** In **Notes for certification**, paste:
   > Free add-in, no sign-in or account needed. To test: open any workbook, Home tab > Quick Sheet group > "New Sheet" adds a sheet named "New Sheet" (numbered if taken). "Named Sheet" opens a task pane; type a name and click "Create sheet". The add-in does not read workbook content or send any data.
7. Click **Publish**. Microsoft tests it, which usually takes **3–5 business days**.
   - **Passed:** it goes live on AppSource and anyone can add it via Home, then Add-ins.
   - **Failed:** you get a report. Fix the listed items, then **Update** and resubmit.

---

## Be aware

- **This add-in may be rejected as "too simple".** Microsoft's certification rules expect an add-in to add meaningful value, and Excel already has a + button for new sheets. Treat this as a dry run: it gets your account verified, your hosting working, and teaches you every screen. Submitting costs nothing, and a rejection costs nothing either. Your real submission should be Mantis Updater.
- **Every change to the website files is live immediately.** Changes to `manifest.xml` (new buttons, new URLs) need a new submission with a higher `<Version>`.
- **Keep the GitHub repo online.** If the files disappear, the add-in breaks for everyone who installed it.
