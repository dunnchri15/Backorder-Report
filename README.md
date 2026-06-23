# Backorder Dashboard — Revolution Group

Live backorder reporting dashboard. Upload your daily backorder XLSX/CSV and view overdue items by coordinator, with persistent notes that survive each upload.

## Deploy to GitHub + Render (one-time setup)

### 1 — Push to GitHub

1. Go to [github.com](https://github.com) → **New repository**
2. Name it `backorder-dashboard`, set to **Private**, click **Create repository**
3. Upload all files from this folder (drag-and-drop onto the GitHub page, or use Git)

### 2 — Deploy on Render

1. Go to [render.com](https://render.com) → **New** → **Web Service**
2. Connect your GitHub account and select the `backorder-dashboard` repo
3. Render will auto-detect the `render.yaml` — confirm these settings:
   - **Runtime:** Python
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
4. Click **Create Web Service**

Render gives you a public URL like `https://backorder-dashboard.onrender.com` in ~2 minutes.

### 3 — (Optional) Persistent disk on Render

By default Render's free tier resets `/tmp` on each deploy, so uploaded data and notes won't survive a redeploy. To make notes truly persistent:

1. In Render → your service → **Disks** → **Add Disk**
2. Mount path: `/data`
3. Set the environment variable `DATA_DIR=/data`

With a disk, notes and feeder data survive indefinitely across deploys and restarts.

---

## Daily use

1. Go to your Render URL (e.g. `https://backorder-dashboard.onrender.com`)
2. Upload the backorder XLSX or CSV → click **Upload & Process**
3. View the dashboard — filtered to overdue by default
4. Click any project in the left panel to add/edit notes
5. Notes save to the server automatically — they persist the next day when you upload again

## Feeder CSVs (part source data)

Upload the 4 part-source CSV files once via the home page. They stay saved on the server. Only re-upload them when the source data changes. The app uses them to tag parts as **Purchased** or **Manufactured**.

Purchased tagging rules:
- Part Source = "Purchased" or "Purchased/Manufactured" → always Purchased
- Part Source = "Manufactured" + Planner in (Boes, Craig, DeVowe, Glynn, Salcedo, Slifer, Zhang, Tracy) → also Purchased

## File structure

```
backorder-dashboard/
├── app.py              # Flask server
├── requirements.txt    # Python dependencies
├── render.yaml         # Render deployment config
├── README.md
└── templates/
    ├── index.html      # Upload page
    └── dashboard.html  # Dashboard page
```
