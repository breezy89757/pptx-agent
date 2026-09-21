from __future__ import annotations

import os
from dataclasses import dataclass
from functools import lru_cache

from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Settings:
    azure_openai_base_url: str
    azure_openai_api_key: str
    azure_openai_model: str


@lru_cache
def get_settings() -> Settings:
    missing = [
        name
        for name in ("AZURE_OPENAI_BASE_URL", "AZURE_OPENAI_API_KEY", "AZURE_OPENAI_MODEL")
        if not os.environ.get(name)
    ]
    if missing:
        raise RuntimeError(
            f"Missing required environment variables: {', '.join(missing)}. "
            "Copy .env.example to .env and fill them in."
        )
    return Settings(
        azure_openai_base_url=os.environ["AZURE_OPENAI_BASE_URL"],
        azure_openai_api_key=os.environ["AZURE_OPENAI_API_KEY"],
        azure_openai_model=os.environ["AZURE_OPENAI_MODEL"],
    )
