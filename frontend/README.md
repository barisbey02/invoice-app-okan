# Okan Invoice Generator

A web app for generating proforma invoices for Istanbul Okan University students.

## Project Structure

```
invoice-app/
├── backend/      → Express.js API (deploy to Render)
└── frontend/     → React app (deploy to GitHub Pages)
```

---

## Backend — Deploy to Render

1. Push the `backend/` folder to a GitHub repo (can be the same repo)
2. Go to [render.com](https://render.com) → New → Web Service
3. Connect your repo, set:
   - **Root directory**: `backend`
   - **Build command**: `npm install`
   - **Start command**: `node index.js`
   - **Environment**: Node
4. Deploy. Copy the URL (e.g. `https://okan-invoice.onrender.com`)

---

## Frontend — Deploy to GitHub Pages

1. In `frontend/`, create a `.env` file:
   ```
   REACT_APP_API_URL=https://your-render-url.onrender.com
   ```

2. In `frontend/package.json`, update the `homepage` field:
   ```json
   "homepage": "https://yourusername.github.io/your-repo-name"
   ```

3. Install and deploy:
   ```bash
   cd frontend
   npm install
   npm run deploy
   ```

This builds the React app and pushes it to the `gh-pages` branch automatically.

---

## Local Development

**Backend:**
```bash
cd backend
npm install
node index.js
# runs on http://localhost:3001
```

**Frontend:**
```bash
cd frontend
npm install
# create .env with REACT_APP_API_URL=http://localhost:3001
npm start
# runs on http://localhost:3000
```

---

## Customization

- To change bank info or signatory: edit `backend/index.js` (search for "BANK INFORMATION")
- To add/remove departments: edit `frontend/src/App.js` (the `DEPTS` array at the top)
- To change the tuition suggestion thresholds: edit the `getSuggestion()` function in `frontend/src/App.js`
