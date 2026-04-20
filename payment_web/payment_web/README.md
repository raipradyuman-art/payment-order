# Payment Order Web App — Deployment Guide

## ⚡ Quick Start (Run Locally First)

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
python app.py

# 3. Open browser
# http://localhost:5000
# Login: admin / admin123
```

---

## 🚀 Deploy FREE on Railway.app (Step by Step)

### Step 1 — GitHub pe upload karo

1. GitHub.com pe jaao → New repository banao: `payment-order`
2. Ye saari files upload karo:
   - `app.py`
   - `requirements.txt`
   - `Procfile`
   - `templates/` folder (saari .html files)
3. Commit karo

### Step 2 — Railway account banao

1. **railway.app** kholo
2. "Login with GitHub" click karo
3. Free account — credit card nahi chahiye

### Step 3 — Deploy karo

1. Railway dashboard pe "New Project" click karo
2. "Deploy from GitHub repo" select karo
3. Apna `payment-order` repo select karo
4. Railway automatically detect karega ki Flask app hai
5. **2-3 minute mein deploy ho jaayega!**

### Step 4 — Environment variables set karo

Railway dashboard → Your project → Variables tab:

```
SECRET_KEY = any-random-string-like-abc123xyz789
FLASK_ENV = production
```

### Step 5 — App open karo

Railway dashboard pe "Deployment URL" milega jaise:
`https://payment-order-production.up.railway.app`

**Ye URL kisi bhi browser pe, kisi bhi device pe open hoga!**

---

## 🔧 Production Checklist

- [ ] `SECRET_KEY` environment variable set karo (random string)
- [ ] Default admin password change karo (Users section mein)
- [ ] Company settings fill karo (letterhead upload)
- [ ] Bank accounts add karo

---

## 📱 Mobile Access

Railway pe deploy hone ke baad:
- **Office PC**: URL kholo
- **Mobile/Tablet**: Same URL browser mein kholo
- **Home laptop**: Same URL
- **Farm site**: Mobile data se bhi kaam karega

---

## 🗃️ Data Backup

Railway pe SQLite use ho raha hai. Data backup ke liye:
1. Railway dashboard → Files section
2. `payment.db` download karo

Ya automatic backup ke liye PostgreSQL upgrade karo:
```
DATABASE_URL = postgresql://user:pass@host:5432/dbname
```
(Railway PostgreSQL plugin se free mein milta hai)

---

## ⚠️ Important Notes

1. **Free tier limitations**: Railway free tier pe app 24/7 chalta hai, 500MB storage
2. **Uploads**: Documents/photos uploads ephemeral storage mein hain — Railway pe persist nahi honge restart pe. Production ke liye Cloudinary ya S3 use karo.
3. **HTTPS**: Railway automatically SSL/HTTPS provide karta hai — secure hai

---

## 🆘 Troubleshooting

**App start nahi ho raha?**
- Railway logs check karo (Deployments tab)
- `requirements.txt` mein sab packages hain?

**Database error?**
- `init_db()` automatically chalta hai pehli baar
- Ya manually: `python -c "from app import init_db; init_db()"`

**Login nahi ho raha?**
- Default: `admin` / `admin123`
- Agar bhool gaye: Railway console mein run karo:
  `python -c "from app import *; init_db()"`
