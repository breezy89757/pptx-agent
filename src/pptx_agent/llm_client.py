from __future__ import annotations

import base64
from functools import lru_cache
from pathlib import Path
from typing import TypeVar

from openai import OpenAI
from pydantic import BaseModel

from pptx_agent.config import get_settings

T = TypeVar("T", bound=BaseModel)


@lru_cache
def get_client() -> OpenAI:
    settings = get_settings()
    return OpenAI(base_url=settings.azure_openai_base_url, api_key=settings.azure_openai_api_key)


def complete_json(
    *,
    system_prompt: str,
    user_prompt: str,
    response_model: type[T],
) -> T:
    """Call the chat model and parse its response into `response_model`,
    using structured-output mode so every agent hands the next stage
    validated JSON rather than free text.

    Note: gpt-5.6-luna is a reasoning-tier model and only supports the
    default temperature (1), so we don't pass one."""
    settings = get_settings()
    client = get_client()
    completion = client.chat.completions.parse(
        model=settings.azure_openai_model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        response_format=response_model,
    )
    parsed = completion.choices[0].message.parsed
    if parsed is None:
        raise RuntimeError(f"Model refused or returned unparseable output: {completion.choices[0].message}")
    return parsed


def complete_json_with_image(
    *,
    system_prompt: str,
    user_prompt: str,
    image_path: Path,
    response_model: type[T],
) -> T:
    """Like complete_json, but attaches an image (the actual rendered slide)
    to the user turn -- used by the Visual QA Agent to review the real .pptx
    output instead of reasoning about markup."""
    settings = get_settings()
    client = get_client()
    b64 = base64.b64encode(image_path.read_bytes()).decode("ascii")
    completion = client.chat.completions.parse(
        model=settings.azure_openai_model,
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": user_prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ],
            },
        ],
        response_format=response_model,
    )
    parsed = completion.choices[0].message.parsed
    if parsed is None:
        raise RuntimeError(f"Model refused or returned unparseable output: {completion.choices[0].message}")
    return parsed
