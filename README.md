# Adani Safety Performance Profile

Static **dashboard preview** (`dashboard-preview/`) plus **Power BI assets** (`PowerBI_Assets/`).

## Share with your team (GitHub + live URL)

### 1) Create a new repository on GitHub

1. Open [github.com/new](https://github.com/new).
2. Choose a repository name (example: `adani-safety-mis`).
3. Leave **Add a README** unchecked (this project already has files).
4. Click **Create repository**.

### 2) Push this folder from your PC

In PowerShell, from this folder:

```powershell
git init
git branch -M main
git add .
git commit -m "Initial commit: dashboard preview and Power BI assets"
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

Replace `YOUR_USERNAME` and `YOUR_REPO` with your GitHub username and repo name.

If GitHub asks for authentication, use a **Personal Access Token** (PAT) as the password, or sign in with GitHub Desktop / Git Credential Manager.

### 3) Turn on GitHub Pages (team review link)

1. On GitHub: **Settings → Pages**.
2. Under **Build and deployment → Source**, pick **Deploy from a branch**.
3. Branch: **main**, folder: **/ (root)**.
4. Save. After a minute or two, the site will be at:

**`https://YOUR_USERNAME.github.io/YOUR_REPO/`**

The root page redirects to the dashboard under `dashboard-preview/`.

### Optional: include Excel / PDF in the repo

By default, `*.xlsx` and `*.pdf` are listed in `.gitignore` to keep the repository smaller and avoid accidental publication of working files. To track them, remove those lines from `.gitignore` and run `git add -f "filename.xlsx"`.

### Local preview

Open `dashboard-preview/index.html` via a local static server (some browsers block module/script loading from `file://`):

```powershell
cd dashboard-preview
npx --yes serve -p 3000
```

Then visit `http://localhost:3000`.
