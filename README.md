# Cellular Sales Outlook Signature Add-in — Phase 1 MVP
### Tailored for: Windows + Classic Outlook (older version)

## Files that matter for you today
```
setup-windows.bat   → double-click this to install & run everything
manifest.local.xml  → already configured to point at https://localhost:3000 (use this one for now)
manifest.xml        → the "real" version with placeholder hosting domain (for later, once IT sets up hosting)
```

## Step 1 — Run the setup script
1. Unzip the project folder if you haven't already.
2. Double-click **`setup-windows.bat`**.
3. A black terminal window will open and run 3 things automatically:
   - `npm install`
   - Certificate trust install (Windows will likely pop up a security prompt — click **Yes**)
   - Starts the server
4. When you see `Add-in dev server running at https://localhost:3000`, leave that window open and minimized. That's your server — closing it stops everything.

**If anything errors out**, the script pauses and tells you to copy the window — do that and send it to me exactly as shown.

## Step 2 — The classic Outlook risk check (do this first, before anything else)

Since you're on classic Outlook (older version), we need to confirm `setSignatureAsync` (the core API this whole add-in depends on) actually works on your build. This is the single biggest unknown from the scope doc — let's resolve it in the first 10 minutes rather than build everything around it and find out later.

**Sideload the add-in:**
1. In Outlook, go to the **Home** tab → **Get Add-ins** button (sometimes labeled "Store" or found under a "..." overflow menu depending on your ribbon layout).
2. Click **My add-ins** (left sidebar).
3. Scroll down to **Custom Addins** → **Add a custom add-in** → **Add from File...**
4. Browse to and select **`manifest.local.xml`** (not `manifest.xml`).
5. Click through any warning about it being an unverified/custom add-in — that's expected for sideloaded add-ins.

> **If you don't see "Add a custom add-in" as an option at all:** that's a signal your Outlook build is old enough that it may not support this path the normal way, and we'd need the alternate "shared folder catalog" sideload method instead (a registry-based approach). Tell me if this option is missing and I'll walk you through that alternative immediately.

**Test it:**
1. Compose a new email.
2. Look for an **"Insert Signature"** button in the ribbon (it may be in its own small group — look across the whole ribbon, it can land in an unexpected spot the first time).
3. Click it — a task pane should slide out on the right with the dropdown and fields.
4. Fill in a test name/title, click **Preview** (should show the signature rendering in the little frame — logos will look like plain red boxes right now, that's expected, it's the placeholder assets).
5. Click **Insert Signature**.

**This is the moment of truth:** if the signature appears in your email body, `setSignatureAsync` works on your Outlook build and the biggest risk is cleared. If you get an error message in the task pane, or nothing happens, tell me the exact wording and we'll troubleshoot from there — there are a couple of known fallback approaches if this specific API isn't available on your build.

## Step 3 — Swap in real branding (once Step 2 works)
1. Go to your SharePoint site collection, find the Email Signature Set Up page's asset library, and pull the actual logo/icon images as PNGs (not SVG).
2. Save them into the `assets/` folder in this project, using these exact filenames (or update `taskpane/templates.js` to match whatever names you use):
   - `cellularsales-logo.png`
   - `verizon-authorized-retailer.png`
   - `icon-instagram.png`, `icon-tiktok.png`, `icon-x.png`, `icon-facebook.png`, `icon-linkedin.png`
   - `greatest-workplaces-2026.png`
3. Restart the server (Ctrl+C in the terminal window, then run `npm start` again) to pick up the new files, then refresh the task pane preview.

## Step 4 — Show your boss
At this point you have a fully working local concept: real Outlook, real button, real inserted signature, real branding. That satisfies "working concept" — no hosting or IT dependency required for a demo.

## Later, once IT weighs in on hosting
Swap every `https://YOUR-HOSTING-DOMAIN` reference in `manifest.xml` for wherever they land, then it's ready for centralized deployment via the Microsoft 365 admin center.
