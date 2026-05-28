# DSR Dashboard — GitHub Pages Setup Guide

Do this **once** to get your permanent dashboard URL. Takes about 10 minutes.

---

## Step 1: Create a GitHub account (if you don't have one)

1. Go to https://github.com and sign up (free)
2. Choose a username — this will be part of your URL, e.g. `cpgi-amy`

---

## Step 2: Create a new repository

1. Click the **+** icon (top right) → **New repository**
2. Name it: `dsr-dashboard` (this becomes your URL)
3. Set to **Public** (required for free GitHub Pages)
4. Click **Create repository**

---

## Step 3: Install Git for Windows (if not already installed)

1. Go to https://git-scm.com/download/win
2. Download and install (all defaults are fine)
3. Open **Command Prompt** and verify: `git --version`

---

## Step 4: Initialize your local folder as a Git repo

Open Command Prompt and run these commands one by one:

```
cd "C:\Users\agnicolas\Documents\Claude\Projects\Daily Sales Report"
git init
git remote add origin https://github.com/YOUR-USERNAME/dsr-dashboard.git
copy DSR_Dashboard.html index.html
git add .
git commit -m "Initial DSR dashboard"
git branch -M main
git push -u origin main
```

Replace `YOUR-USERNAME` with your actual GitHub username.

---

## Step 5: Enable GitHub Pages

1. Go to your repo on GitHub: `https://github.com/YOUR-USERNAME/dsr-dashboard`
2. Click **Settings** tab
3. Scroll down to **Pages** (left sidebar)
4. Under **Source**, select: **Deploy from a branch**
5. Branch: **main** / folder: **/ (root)**
6. Click **Save**

Your URL will appear: `https://YOUR-USERNAME.github.io/dsr-dashboard/`

It takes ~1 minute for the first deploy. Refresh after a minute.

---

## Step 6: Daily update workflow

After processing each day's DSR Excel:

1. Open Command Prompt
2. Run: `deploy_github.bat`

That's it — the live site updates in ~30 seconds.

---

## Your live URL

```
https://YOUR-USERNAME.github.io/dsr-dashboard/
```

Share this URL with your team. It always shows the latest data.

---

## Security note

The dashboard HTML contains your DSR data baked in. Since the repo is public,
anyone with the URL can view it. If you need it private:
- Upgrade to GitHub Pro (~$4/month) for private repo Pages, OR
- Use Netlify (free, private sites available)
