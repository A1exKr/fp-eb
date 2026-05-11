import json

from openai import OpenAI

from app.config import settings


class OpenAIService:
    def __init__(self) -> None:
        self.provider = settings.llm_provider.lower()
        self.base_url = None
        self.api_key = ""

        if self.provider == "litellm":
            self.base_url = settings.litellm_url.rstrip("/") if settings.litellm_url else ""
            self.api_key = settings.litellm_master_key
        else:
            self.api_key = settings.openai_api_key

        self.enabled = bool(self.api_key)
        self.client = None
        if self.enabled:
            client_kwargs = {"api_key": self.api_key}
            if self.base_url:
                client_kwargs["base_url"] = self.base_url
            self.client = OpenAI(**client_kwargs)

    def _missing_credentials_error(self) -> str:
        if self.provider == "litellm":
            return "LITELLM_URL and LITELLM_MASTER_KEY must be configured"
        return "OPENAI_API_KEY is not configured"

    def json_completion(self, system_prompt: str, user_prompt: str) -> dict:
        if not self.enabled:
            raise RuntimeError(self._missing_credentials_error())

        response = self.client.chat.completions.create(
            model=settings.openai_model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.2,
        )
        content = response.choices[0].message.content
        return json.loads(content)

    def text_completion(self, system_prompt: str, user_prompt: str) -> str:
        if not self.enabled:
            raise RuntimeError(self._missing_credentials_error())

        response = self.client.chat.completions.create(
            model=settings.openai_model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.4,
        )
        return response.choices[0].message.content.strip()
