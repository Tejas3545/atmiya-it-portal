# Atmiya University — B.Tech IT Attendance & Compensation Portal

Official web portal for Atmiya University B.Tech Information Technology students to check attendance records, weekly breakdown, points earned, and compensation status directly synced with the department Excel master file.

---

## 📁 Clean Directory Structure

```text
Dean LMS/
├── data/
│   └── IT.xlsx              # Master Attendance Spreadsheet for IT Department
├── static/
│   ├── index.html           # Portal Frontend UI
│   ├── style.css            # Professional Design Styles
│   ├── app.js               # Frontend Logic & API Integration
│   └── logo.png             # Atmiya University Seal
├── app.py                   # Flask Backend Engine & Excel Parser
├── requirements.txt         # Production Python Dependencies
├── Procfile                 # Production WSGI Config for Cloud Hosting
├── render.yaml              # Render.com Auto-Deployment Config
├── START SERVER.bat          # 1-Click Local Server Launcher for Windows
├── README.md                # Project & Hosting Documentation
└── .gitignore               # Git Exclusions
```

---

## 🚀 How to Run Locally

1. Make sure Python 3.9+ is installed.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the Flask application:
   ```bash
   python app.py
   ```
4. Open [http://localhost:5000](http://localhost:5000) in your browser.

---

## 🌐 How to Push to GitHub

1. Open Terminal / PowerShell in the project directory:
   ```bash
   git init
   git add .
   git commit -m "Initial commit - Atmiya IT Attendance Portal"
   ```
2. Create a new repository on GitHub (e.g. `atmiya-it-portal`).
3. Link and push:
   ```bash
   git branch -M main
   git remote add origin https://github.com/YOUR_GITHUB_USERNAME/atmiya-it-portal.git
   git push -u origin main
   ```

---

## ☁️ How to Host Live for FREE (Render.com)

1. Go to [Render.com](https://render.com) and create a free account.
2. Click **New +** -> **Web Service**.
3. Connect your GitHub repository `atmiya-it-portal`.
4. Render will automatically detect `Procfile` and `requirements.txt`:
   - **Environment**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
5. Click **Create Web Service**.
6. Within 2 minutes, your website will be live on a public URL like:
   `https://atmiya-it-portal.onrender.com`

---

## 🏛️ Future Expansion Plan (Multibranch & Main Dean LMS Portal)

This project is built with modular architecture:
1. **Branch Portals**: When you create portals for CSE, Civil, or Mechanical in the future:
   - Place spreadsheets in `data/CSE.xlsx`, `data/CIVIL.xlsx`, `data/MECH.xlsx`.
   - The backend routes `/api/it/student`, `/api/cse/student`, etc., isolate department logic.
2. **Central Dean LMS Main Portal**:
   - You can create a parent landing page with card buttons for each branch:
     - 💻 B.Tech Information Technology -> Links to `/it` or IT Portal
     - 🖥️ B.Tech Computer Science -> Links to `/cse`
     - 🏗️ B.Tech Civil Engineering -> Links to `/civil`
     - ⚙️ B.Tech Mechanical Engineering -> Links to `/mech`
