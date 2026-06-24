# TPL SOW Generator

## How to deploy to Netlify

1. **Create a GitHub repo**
   - Go to github.com → New repository → name it `tpl-sow-generator` → Create
   - Upload the entire contents of this folder (drag and drop all files)

2. **Connect to Netlify**
   - Go to app.netlify.com → Add new site → Import an existing project
   - Choose GitHub → select your repo
   - Build settings are auto-detected from netlify.toml
   - Click **Deploy site**

3. **That's it**
   - Netlify installs dependencies and deploys automatically
   - The ⬇ Download PDF button calls a serverless function that uses 
     real Chrome (Puppeteer) to generate the PDF
   - Every subsequent push to GitHub auto-redeploys

## Local testing (optional)
```
npm install -g netlify-cli
npm install
netlify dev
```
Then open http://localhost:8888
