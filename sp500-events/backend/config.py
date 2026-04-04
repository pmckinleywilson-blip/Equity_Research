from pathlib import Path
from pydantic_settings import BaseSettings
from functools import lru_cache

# Absolute path so all scripts (backend, scrapers) use the same DB
_DEFAULT_DB = f"sqlite:///{Path(__file__).parent.parent / 'sp500events.db'}"


class Settings(BaseSettings):
    database_url: str = _DEFAULT_DB
    secret_key: str = "change-me-in-production"
    resend_api_key: str = ""
    invite_from_email: str = "invites@sp500events.com"
    site_url: str = "http://localhost:8000"
    finnhub_api_key: str = ""
    sec_edgar_user_agent: str = "SP500Events admin@sp500events.com"
    rate_limit_default: str = "60/minute"
    rate_limit_api_key: str = "300/minute"

    class Config:
        env_file = ".env"


@lru_cache
def get_settings() -> Settings:
    return Settings()
