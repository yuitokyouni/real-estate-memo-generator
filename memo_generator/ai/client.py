"""Anthropic Claude API client wrapper."""
import os
from pathlib import Path

import anthropic
from dotenv import load_dotenv

load_dotenv()

_PROMPTS_DIR = Path(__file__).parent / "prompts"


def _load_prompt(filename: str) -> str:
    return (_PROMPTS_DIR / filename).read_text(encoding="utf-8")


def get_client() -> anthropic.Anthropic:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError("ANTHROPIC_API_KEY is not set. Copy .env.example to .env and add your key.")
    return anthropic.Anthropic(api_key=api_key)


def generate_section(
    system_prompt: str,
    user_prompt: str,
    model: str | None = None,
    max_tokens: int = 1500,
) -> str:
    """Call Claude API and return the generated text for one memo section."""
    client = get_client()
    model = model or os.getenv("DEFAULT_MODEL", "claude-opus-4-5")

    message = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    )
    return message.content[0].text
