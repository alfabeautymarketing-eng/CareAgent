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

# Routes
app.include_router(router)


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
