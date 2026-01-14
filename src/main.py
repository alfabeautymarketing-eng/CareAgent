"""
AgentCare - Main entry point.

FastAPI application for Google Sheets automation.
"""

from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from src.api.router import router
from src.utils.logger import setup_logging, logger
from src.utils.config import settings


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan manager."""
    # Startup
    setup_logging(settings.log_level)
    logger.info("application_starting", version=settings.version)

    # TODO: Initialize connections
    # - Redis connection
    # - Google credentials validation
    # - Health checks

    yield

    # Shutdown
    logger.info("application_stopping")
    # TODO: Cleanup connections


app = FastAPI(
    title="AgentCare",
    description="Google Sheets automation with AI",
    version=settings.version,
    lifespan=lifespan,
)

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Middleware to allow iframing (Google Sheets)
@app.middleware("http")
async def add_security_headers(request, call_next):
    response = await call_next(request)
    # Remove X-Frame-Options to allow embedding in GAS iframes
    if "x-frame-options" in response.headers:
        del response.headers["x-frame-options"]
    
    # Set permissive CSP for frames
    response.headers["Content-Security-Policy"] = "frame-ancestors 'self' https://*.google.com https://*.googleusercontent.com;"
    return response

# Routes
app.include_router(router)

from fastapi.responses import RedirectResponse, HTMLResponse

@app.get("/")
async def root():
    """Redirect to Rules UI or show landing page."""
    return HTMLResponse(content=f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>AgentCare Server</title>
            <style>
                body {{ font-family: sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; background: #0f172a; color: white; }}
                .card {{ background: #1e293b; padding: 2rem; border-radius: 1rem; box-shadow: 0 10px 15px rgba(0,0,0,0.5); text-align: center; }}
                a {{ color: #0d9488; text-decoration: none; font-weight: bold; font-size: 1.2rem; }}
                a:hover {{ text-decoration: underline; }}
            </style>
        </head>
        <body>
            <div class="card">
                <h1>🚀 AgentCare Server</h1>
                <p>Status: <span style="color: #10b981;">Running</span></p>
                <br>
                <a href="/api/v1/rules-ui">打开 Rules Manager →</a>
            </div>
        </body>
        </html>
    """)


@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {
        "status": "healthy",
        "version": settings.version,
        "checks": {
            "redis": "ok",  # TODO: actual check
            "google_sheets": "ok",  # TODO: actual check
            "gemini": "ok",  # TODO: actual check
        },
    }


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "src.main:app",
        host=settings.server_host,
        port=settings.server_port,
        reload=settings.debug,
    )
