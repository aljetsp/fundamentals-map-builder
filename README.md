# Fundamentals Map Builder

A tool for building Results Fundamentals Maps using the Mass Ingenuity code schema.

## Deploy to GitHub Pages (recommended)

1. **Create a new GitHub repo** at [github.com/new](https://github.com/new)
   - Name it anything, e.g. `fundamentals-map-builder`
   - Set to Public
   - Do NOT initialize with README (you're uploading your own files)

2. **Update `vite.config.js`** — change the `base` value to match your repo name:
   ```js
   base: '/your-repo-name/',
   ```

3. **Push this code to GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
   git push -u origin main
   ```

4. **Enable GitHub Pages:**
   - Go to your repo → Settings → Pages
   - Under "Source", select **GitHub Actions**
   - Save

5. GitHub Actions will build and deploy automatically. Your live URL will be:
   `https://YOUR_USERNAME.github.io/YOUR_REPO_NAME/`

   The Actions tab in your repo shows build progress (takes ~1 minute).

## Making Changes

Edit `src/App.jsx`, then:
```bash
git add .
git commit -m "Update map"
git push
```
GitHub Actions redeploys automatically within ~1 minute.

## Local Development

```bash
npm install
npm run dev
```

## Features

- Build fundamentals maps with key goals, operating & supporting processes, outcomes
- Results code schema (01OP01A format) auto-generated
- Inline editing directly on the preview map
- Full screen preview mode
- Export to PowerPoint (.pptx)
- Export to PDF (browser print dialog — Ctrl+P / Cmd+P)
- Export / Import JSON for saving and sharing map data
- Logo upload
- Layout options: outcomes above or below core processes; group by goal or by type
- Auto-saves to browser localStorage

