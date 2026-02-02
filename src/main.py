"""
AgentCare - Main entry point.

FastAPI application for Google Sheets automation.
"""

from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from src.api.router import router
from src.api.tags_metadata import tags_metadata
from src.utils.logger import setup_logging, logger
from src.utils.config import settings


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan manager."""
    # Startup
    setup_logging(settings.log_level, log_file=settings.server_log_file)
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
    title="AgentCare Сервер",
    description="Автоматизация Google Таблиц с использованием ИИ",
    version=settings.version,
    lifespan=lifespan,
    openapi_tags=tags_metadata,
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

@app.get("/", include_in_schema=False)
async def root():
    """Перенаправление на интерфейсы управления или главную страницу."""
    return HTMLResponse(content=f"""
        <!DOCTYPE html>
        <html lang="ru">
        <head>
            <meta charset="UTF-8">
            <title>AgentCare Server</title>
            <style>
                body {{ 
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                    display: flex; 
                    justify-content: center; 
                    align-items: center; 
                    height: 100vh; 
                    margin: 0;
                    background: #0f172a; 
                    color: white; 
                }}
                .card {{ 
                    background: #1e293b; 
                    padding: 3rem; 
                    border-radius: 1.5rem; 
                    box-shadow: 0 20px 25px -5px rgba(0,0,0,0.5), 0 10px 10px -5px rgba(0,0,0,0.4); 
                    text-align: center;
                    max-width: 400px;
                    width: 90%;
                }}
                h1 {{ margin-top: 0; color: #f8fafc; font-size: 2rem; }}
                .status {{ color: #10b981; font-weight: bold; margin-bottom: 2rem; display: block; }}
                .links {{ display: flex; flex-direction: column; gap: 1rem; }}
                a {{ 
                    color: #fff; 
                    background: #0d9488;
                    text-decoration: none; 
                    font-weight: 600; 
                    padding: 1rem;
                    border-radius: 0.75rem;
                    transition: all 0.2s;
                }}
                a:hover {{ 
                    background: #0f766e;
                    transform: translateY(-2px);
                    box-shadow: 0 4px 12px rgba(13, 148, 136, 0.3);
                }}
                a.secondary {{
                    background: #334155;
                }}
                a.secondary:hover {{
                    background: #475569;
                }}
            </style>
        </head>
        <body>
            <div class="card">
                <h1>🚀 AgentCare</h1>
                <span class="status">● Сервер запущен</span>
                <div class="links">
                    <a href="/api/v1/rules-ui">Управление правилами</a>
                    <a href="/api/v1/logs-ui" class="secondary">Журнал синхронизации</a>
                </div>
            </div>
        </body>
        </html>
    """)


@app.get("/health", summary="Проверка здоровья сервера")
async def health_check():
    """Эндпоинт для проверки статуса всех соединений сервиса."""
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
