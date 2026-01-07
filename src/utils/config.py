"""
Configuration management using pydantic-settings.
"""

from typing import List, Optional
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
    )

    # App
    app_name: str = "AgentCare"
    version: str = "0.1.0"
    debug: bool = False
    log_level: str = "INFO"

    # Server
    server_host: str = "0.0.0.0"
    server_port: int = 8000
    server_workers: int = 4

    # CORS
    cors_origins: List[str] = ["*"]

    # Google
    google_credentials_file: str = "config/credentials.json"
    google_credentials_base64: Optional[str] = None

    # Gemini AI
    gemini_api_key: str = ""
    gemini_model: str = "gemini-2.0-flash"
    gemini_max_retries: int = 5

    # Redis
    redis_url: str = "redis://localhost:6379/0"
    cache_ttl: int = 3600

    # Webhook
    webhook_secret: str = ""

    # Notifications
    telegram_bot_token: Optional[str] = None
    telegram_chat_id: Optional[str] = None

    # Sync log storage
    sync_log_data_dir: str = "data/sync_logs"
    sync_log_retention_days: int = 90
    sync_log_max_entries: int = 200000


settings = Settings()
