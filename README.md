# Attendance Automation Web App

Wraps the existing `attendance_automation.py` script in a single deployable web service: **FastAPI backend + React (Vite) frontend**, served from one process/port.

The core attendance logic is untouched — the web layer only handles upload, invocation, and download.

---

## Project Structure

```
.
├── attendance_automation.py       # original script (also copied to backend/app/)
├── backend/
│   ├── app/
│   │   ├── main.py                # FastAPI app + static file mount + CORS
│   │   ├── attendance_automation.py  # verbatim copy of the script
│   │   └── api/
│   │       └── routes.py          # POST /api/process, GET /api/health
│   └── requirements.txt
├── frontend/
│   ├── index.html
│   ├── package.json
│   ├── vite.config.js
│   └── src/
│       ├── main.jsx
│       ├── App.jsx
│       ├── App.css
│       ├── api.js
│       └── components/
│           ├── FileUploader.jsx
│           └── ProcessingStatus.jsx
├── Dockerfile                     # multi-stage: node build + python serve
└── README.md
```

---

## Local Development

### Backend

```bash
cd backend
pip install -r requirements.txt
# from backend/ directory:
uvicorn app.main:app --reload --port 8000
# or from repo root:
uvicorn backend.app.main:app --reload  # adjust PYTHONPATH as needed
```

The API is then at `http://localhost:8000/api/health` and `POST http://localhost:8000/api/process`.

### Frontend

```bash
cd frontend
npm install
npm run dev
```

Frontend dev server runs at `http://localhost:5173`.  
`vite.config.js` proxies `/api/*` to `http://localhost:8000`, so no CORS issues during dev.  
The FastAPI backend enables permissive CORS when `ENV != "production"` (default is `development`), allowing the Vite dev server to call it directly. In production (`ENV=production`, set in Dockerfile) CORS is not added.

### Running Both Together Locally (Production Mode)

Build the frontend and let FastAPI serve it:

```bash
cd frontend && npm run build
cd ../backend && uvicorn app.main:app --port 8000
# then open http://localhost:8000/
```

---

## Docker — Combined Single Service

Build and run the multi-stage image (builds frontend, copies `dist` into the Python image, serves everything on one port):

```bash
docker build -t attendance-automation .
docker run -p 8000:8000 attendance-automation
# or with a platform-injected PORT:
docker run -p 8000:8000 -e PORT=8000 attendance-automation
```

Then:
- UI at `http://localhost:8000/`
- API at `http://localhost:8000/api/health` and `http://localhost:8000/api/process`

---

## API

- **POST /api/process** — `multipart/form-data` with field `file` (`.xlsx`). Returns the processed workbook as `attachment` with headers `X-Employees-Processed` and `X-Warnings` (base64-encoded warning text). Errors return JSON `{"detail": "..."}` with 400/422/500.
- **GET /api/health** — `{"status":"ok"}`

Output filename follows the script's convention: `<input-stem>_output.xlsx`.

---

## Deployment

This repo is designed to deploy as **one service** on any platform that can run a Dockerfile and expose a port (Render, Railway, Fly.io, etc.):

1. Point the platform at the `Dockerfile` at the repo root.
2. It will build stage 1 (Node) then stage 2 (Python), expose `$PORT` (defaults to `8000`).
3. No env vars, databases, or volumes needed. Set `ENV=production` (already in Dockerfile) to disable dev CORS.

---

## Notes

- Fixed format: sheet `"MAIN"`, header row 2 — not configurable via UI.
- One file per request, stateless, temp files cleaned up after response via `BackgroundTasks`.
- Unrecognized day-status values that the script prints to stdout are captured via `contextlib.redirect_stdout` and surfaced in the `X-Warnings` header / yellow warning box in the UI.
