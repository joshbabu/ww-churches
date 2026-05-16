# Word & Worship — Church App Template

A complete, config-driven Christian songbook web app.
One codebase, infinite churches. Each church gets their own branded app with their own songs.

---

## How It Works

```
wordnworship.com/?church=bethel-assembly
         ↓
Loads: churches/bethel-assembly/config.json  (branding, colors, features)
       churches/bethel-assembly/songs.json   (their songs)
         ↓
Renders: Bethel Assembly branded app with their songs
```

Or via subdomain:
```
bethel.wordnworship.com  →  loads bethel-assembly config automatically
```

---

## Folder Structure

```
ww-churches/                    ← GitHub repo
  churches/
    telugu-brethren/
      config.json               ← branding + features
      songs.json                ← 1006 songs
    bethel-assembly/
      config.json
      songs.json
    your-church/
      config.json
      songs.json

app-template.html               ← single HTML file (deploy once)
pipeline.py                     ← song import tool
```

---

## Adding a New Church

### Step 1 — Create their config

Copy `churches/telugu-brethren/config.json` and edit:

```json
{
  "church": {
    "id": "your-church-id",
    "name": "Your Church Name",
    "community": "Your Community",
    "website": "https://your-church.wordnworship.com"
  },
  "theme": {
    "primary": "#065f46",
    "primaryLight": "#047857",
    "primaryDark": "#064e3b",
    "primaryFaint": "#ecfdf5",
    "primaryPale": "#d1fae5",
    "shadow": "0 4px 16px rgba(6,95,70,0.1)"
  },
  "features": {
    "songs": true,
    "bible": true,
    "bibleStudy": true,
    "directory": false,
    "aiTools": true
  }
}
```

### Step 2 — Import their songs

```bash
# From a Word file
python3 pipeline.py --input their-songs.docx --church your-church-id --push --token ghp_xxx

# From an Excel file
python3 pipeline.py --input their-songs.xlsx --church your-church-id --push --token ghp_xxx
```

### Step 3 — Share the URL

```
https://wordnworship.com/?church=your-church-id
```

Or set up a subdomain:
```
your-church.wordnworship.com
```

That's it — 3 steps, under 1 hour.

---

## Excel Format for Songs

If a church sends an Excel file, it should have these columns:

| no | te | en | author | catId | chorus_te | chorus_en | verse1_te | verse1_en | verse2_te | verse2_en |
|----|----|----|--------|-------|-----------|-----------|-----------|-----------|-----------|-----------|
| 1  | పాట పేరు | Song Title | రచయిత | 11 | పల్లవి text | Chorus text | వచనం 1 | Verse 1 | వచనం 2 | Verse 2 |

Minimum required columns: `no`, `te`

---

## Config Reference

```json
{
  "church": {
    "id": "unique-id",               // used in URL ?church=ID
    "name": "App Name",              // shown in topbar
    "nameLocal": "Local script name",
    "community": "Community Name",   // shown in copyright
    "tagline": "English tagline",
    "taglineLocal": "Local tagline",
    "copyright": "© 2026 Community",
    "website": "https://...",
    "email": "admin@...",
    "logo": "ww"                     // logo style: "ww" or custom
  },
  "theme": {
    "primary": "#3730a3",           // main brand color
    "primaryLight": "#4338ca",
    "primaryDark": "#312e81",
    "primaryFaint": "#eef2ff",      // very light tint for backgrounds
    "primaryPale": "#e0e7ff",       // light tint for highlights
    "accent": "#f59e0b",            // accent color (buttons, badges)
    "shadow": "0 4px 16px rgba(...)"
  },
  "language": {
    "default": "te",                // "te" or "en"
    "available": ["te", "en"]
  },
  "data": {
    "songsRepo": "joshbabu/ww-churches",
    "songsPath": "churches/ID/songs.json"
  },
  "features": {
    "songs": true,                  // A-Z song browser
    "bible": true,                  // Bible reader
    "bibleStudy": true,             // AI Bible study
    "directory": true,              // Assembly directory
    "aiTools": true,                // AI song assistant
    "inquire": true,                // Prayer request / contact form
    "songLibrary": true,            // Full songs library tab
    "favorites": true,              // Favorite songs
    "slideMode": true,              // Presentation mode
    "darkMode": true                // Dark mode toggle
  }
}
```

---

## Deployment

The `app-template.html` is a single file. Deploy it once:

```
wordnworship.com/index.html  ← one file serves all churches
```

Each church accesses it via:
```
wordnworship.com/?church=their-id
their-subdomain.wordnworship.com
```

No server-side code needed — pure static hosting.

---

## What Each Church Gets

- ✅ Their own branding (name, colors, logo)
- ✅ Their own songs
- ✅ Telugu + English Bible
- ✅ AI Bible study tools
- ✅ Prayer request / contact form
- ✅ Assembly directory (optional)
- ✅ Favorites, slide mode, dark mode
- ✅ Works offline (PWA)
- ✅ Free to use

---

## Built With

- Vanilla HTML/CSS/JS — no frameworks, no build tools
- GitHub as CDN for songs and config data
- Cloudflare Workers for auth and form submissions
- PWABuilder for Play Store packaging
