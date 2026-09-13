import os
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from app.api.routes import router as api_router

app = FastAPI(title="Attendance Automation")

# CORS for local development: permissive when not in production
# Set ENV=production in deployment to restrict
env = os.getenv("ENV", "development")
if env != "production":
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
        expose_headers=["X-Employees-Processed", "X-Warnings"],
    )

app.include_router(api_router, prefix="/api")

# Mount built frontend static files
# Search multiple possible locations for frontend/dist
possible_dist_paths = [
    Path(__file__).parent.parent.parent / "frontend" / "dist",  # repo root/frontend/dist when running from backend/app/main.py locally
    Path(__file__).parent / "static",  # alternative: backend/app/static
    Path(__file__).parent.parent / "frontend" / "dist",  # backend/frontend/dist in Docker
    Path("/app/frontend/dist"),  # Docker absolute
    Path(__file__).parent.parent / "static",
]

dist_path = None
for p in possible_dist_paths:
    if p.exists() and p.is_dir():
        dist_path = p
        break

if dist_path:
    app.mount("/", StaticFiles(directory=str(dist_path), html=True), name="static")
else:
    @app.get("/")
    async def root():
        return {"message": "Attendance Automation API. Frontend not built yet. Run 'npm run build' in frontend/."}
