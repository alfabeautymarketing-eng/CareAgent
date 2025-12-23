"""
Gemini API Client for AgentCare.
Direct integration with Google Gemini API.
"""

import time
import json
import httpx
from typing import Optional, Dict, Any, List

from src.utils.config import settings
from src.utils.logger import logger


class GeminiClient:
    """
    Client for Google Gemini API.
    Supports retry logic, model listing, and JSON responses.
    """

    BASE_URL = "https://generativelanguage.googleapis.com/v1beta"

    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        self.api_key = api_key or settings.gemini_api_key
        self.model = model or settings.gemini_model
        self.max_retries = settings.gemini_max_retries

        if not self.api_key:
            logger.warning("gemini_api_key_not_configured")

    def list_models(self) -> List[str]:
        """
        Get list of available Gemini models that support generateContent.

        Returns:
            List of model names (without 'models/' prefix)
        """
        try:
            url = f"{self.BASE_URL}/models?key={self.api_key}"

            with httpx.Client(timeout=30.0) as client:
                response = client.get(url)

            if response.status_code != 200:
                logger.error("list_models_failed",
                           status=response.status_code,
                           error=response.text)
                raise Exception(f"Failed to list models: {response.text}")

            data = response.json()
            models = data.get("models", [])

            # Filter models that support generateContent
            available = [
                m["name"].replace("models/", "")
                for m in models
                if "generateContent" in m.get("supportedGenerationMethods", [])
            ]

            logger.info("gemini_models_listed", count=len(available))
            return available

        except Exception as e:
            logger.error("list_models_error", error=str(e))
            raise

    def call_api(
        self,
        prompt: str,
        temperature: float = 0.1,
        max_output_tokens: int = 8192,
        json_response: bool = False,
        model: Optional[str] = None
    ) -> str:
        """
        Call Gemini API with text prompt.

        Args:
            prompt: Text prompt to send
            temperature: Creativity (0.0-1.0)
            max_output_tokens: Max response length
            json_response: If True, request JSON format
            model: Override default model

        Returns:
            Response text from Gemini
        """
        if not self.api_key:
            raise ValueError("Gemini API key not configured. Set GEMINI_API_KEY in .env")

        model_name = model or self.model
        url = f"{self.BASE_URL}/models/{model_name}:generateContent?key={self.api_key}"

        response_mime_type = "application/json" if json_response else "text/plain"

        payload = {
            "contents": [{
                "parts": [{
                    "text": prompt
                }]
            }],
            "generationConfig": {
                "temperature": temperature,
                "topK": 40,
                "topP": 0.95,
                "maxOutputTokens": max_output_tokens,
                "responseMimeType": response_mime_type
            }
        }

        logger.info("gemini_api_call_start",
                   model=model_name,
                   prompt_length=len(prompt),
                   json_response=json_response)

        # Retry logic with exponential backoff
        wait_times = [2, 5, 10, 20, 30]  # seconds
        last_error = None

        for attempt in range(1, self.max_retries + 1):
            try:
                start_time = time.time()

                with httpx.Client(timeout=120.0) as client:
                    response = client.post(
                        url,
                        json=payload,
                        headers={"Content-Type": "application/json"}
                    )

                duration = time.time() - start_time
                status = response.status_code

                if status == 200:
                    logger.info("gemini_api_call_success",
                               attempt=attempt,
                               duration=f"{duration:.2f}s")

                    result = response.json()
                    return self._extract_text(result)

                error_text = response.text
                logger.warning("gemini_api_error",
                             attempt=attempt,
                             status=status,
                             error=error_text[:500])

                # Retry on server errors or rate limit
                if status in (429, 500, 502, 503, 504) and attempt < self.max_retries:
                    wait = wait_times[min(attempt - 1, len(wait_times) - 1)]
                    logger.info("gemini_api_retry",
                               wait_seconds=wait,
                               attempt=attempt)
                    time.sleep(wait)
                    continue

                # Handle specific errors
                if status == 404 and "not found" in error_text.lower():
                    raise ValueError(
                        f'Model "{model_name}" not found. '
                        f'Try: gemini-2.5-flash, gemini-2.0-flash, or gemini-1.5-pro. '
                        f'Use list_models() to see available models.'
                    )

                last_error = Exception(f"Gemini API error ({status}): {error_text}")

                # Don't retry on client errors (4xx except 429)
                if 400 <= status < 500 and status != 429:
                    break

            except httpx.TimeoutException as e:
                last_error = e
                logger.warning("gemini_api_timeout", attempt=attempt)
                if attempt < self.max_retries:
                    wait = wait_times[min(attempt - 1, len(wait_times) - 1)]
                    time.sleep(wait)
                    continue

            except Exception as e:
                last_error = e
                logger.error("gemini_api_exception",
                           attempt=attempt,
                           error=str(e))
                break

        logger.error("gemini_api_all_retries_failed")
        raise last_error or Exception("Gemini API call failed after all retries")

    def _extract_text(self, result: Dict[str, Any]) -> str:
        """Extract text from Gemini API response."""
        if not result.get("candidates"):
            logger.warning("gemini_empty_response", response=str(result)[:500])
            raise ValueError("Empty response from Gemini (no candidates)")

        candidate = result["candidates"][0]

        # Try different response formats
        if candidate.get("content", {}).get("parts"):
            text = candidate["content"]["parts"][0].get("text", "")
        elif candidate.get("text"):
            text = candidate["text"]
        elif candidate.get("output"):
            text = candidate["output"]
        else:
            logger.warning("gemini_unknown_format",
                         candidate_keys=list(candidate.keys()))
            raise ValueError(f"Could not extract text from response: {list(candidate.keys())}")

        logger.info("gemini_response_received", response_length=len(text))
        return text

    def analyze_inci(
        self,
        inci_text: str,
        purpose: Optional[str] = None,
        application: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Analyze INCI composition and classify product.

        Args:
            inci_text: INCI composition text (from PDF)
            purpose: Product purpose (назначение)
            application: Product application (применение)

        Returns:
            Analysis result dict with category, type, compositions
        """
        prompt = self._build_analysis_prompt(inci_text, purpose, application)

        response = self.call_api(
            prompt=prompt,
            temperature=0.1,
            max_output_tokens=4096,
            json_response=True
        )

        return self._parse_analysis_response(response)

    def _build_analysis_prompt(
        self,
        inci_text: str,
        purpose: Optional[str] = None,
        application: Optional[str] = None
    ) -> str:
        """Build the analysis prompt for INCI classification."""
        context_parts = []
        if purpose:
            context_parts.append(f"Назначение: {purpose}")
        if application:
            context_parts.append(f"Применение: {application}")
        context = "\n".join(context_parts)

        return f"""Ты эксперт по косметической продукции и классификации товаров по кодам ТН ВЭД.

ЗАДАЧА: Проанализируй состав косметического продукта и определи:
1. Код ТН ВЭД (3304, 3305 или 3307)
2. Категорию продукта
3. Вид продукции
4. Разрешительный документ (СГР или Декларация)
5. Составы на русском и английском языках

ИСХОДНЫЕ ДАННЫЕ:

Состав INCI:
{inci_text}

{context}

КАТЕГОРИИ ТН ВЭД:
- 3304: Средства косметические или для макияжа и средства для ухода за кожей
- 3305: Средства для ухода за волосами
- 3307: Средства для бритья, дезодоранты, соли для ванн и прочие косметические средства

ОТВЕТ должен быть в формате JSON:

{{
  "category_code": "3304",
  "category": "Средства косметические или для макияжа...",
  "product_type": "Крем для лица",
  "product_reason": "Обоснование почему это крем для лица",
  "reg_doc": "СГР" или "Точно декларация",
  "reg_reason": "Обоснование выбора документа",
  "active_ru_no_pct": "Активные ингредиенты без % (рус)",
  "active_en_no_pct": "Active ingredients without % (eng)",
  "active_ru_pct": "Активные с % (рус)",
  "active_en_pct": "Active with % (eng)",
  "full_ru_no_pct": "Полный состав без % (рус)",
  "full_en_no_pct": "Full composition without % (eng)",
  "full_ru_pct": "Полный состав с % (рус)",
  "full_en_pct": "Full composition with % (eng)"
}}

Верни ТОЛЬКО валидный JSON, без дополнительных комментариев."""

    def _parse_analysis_response(self, response: str) -> Dict[str, Any]:
        """Parse and validate the analysis response."""
        try:
            # Clean markdown code blocks if present
            cleaned = response.strip()

            if cleaned.startswith("```json"):
                cleaned = cleaned[7:]
            if cleaned.startswith("```"):
                cleaned = cleaned[3:]
            if cleaned.endswith("```"):
                cleaned = cleaned[:-3]

            cleaned = cleaned.strip()
            result = json.loads(cleaned)

            # Validate required fields
            required = [
                "category_code",
                "category",
                "product_type",
                "reg_doc",
                "active_ru_no_pct",
                "full_ru_no_pct"
            ]

            missing = [f for f in required if not result.get(f)]
            if missing:
                raise ValueError(f"Missing required fields: {missing}")

            logger.info("inci_analysis_parsed",
                       category=result.get("category_code"),
                       product_type=result.get("product_type"))

            return result

        except json.JSONDecodeError as e:
            logger.error("gemini_response_parse_error",
                        error=str(e),
                        response=response[:500])
            raise ValueError(f"Failed to parse Gemini response as JSON: {e}")

    def test_connection(self) -> Dict[str, Any]:
        """
        Test Gemini API connection with a simple request.

        Returns:
            Dict with status, model, and available models
        """
        try:
            # Simple test prompt
            response = self.call_api(
                prompt="Ответь одним словом: OK",
                temperature=0.0,
                max_output_tokens=10,
                json_response=False
            )

            # Get available models
            try:
                models = self.list_models()
            except Exception:
                models = []

            return {
                "status": "online",
                "response": response.strip(),
                "model": self.model,
                "available_models": models[:10]  # Limit to first 10
            }

        except Exception as e:
            logger.error("gemini_connection_test_failed", error=str(e))
            return {
                "status": "error",
                "error": str(e),
                "model": self.model
            }


# Runtime settings (can be updated without restart)
_runtime_api_key: Optional[str] = None
_runtime_model: Optional[str] = None


def set_runtime_config(api_key: Optional[str] = None, model: Optional[str] = None):
    """
    Set runtime configuration for Gemini API.
    This allows updating API key and model without restarting the server.
    """
    global _runtime_api_key, _runtime_model, _gemini_client

    if api_key is not None:
        _runtime_api_key = api_key
        logger.info("gemini_runtime_api_key_updated")

    if model is not None:
        _runtime_model = model
        logger.info("gemini_runtime_model_updated", model=model)

    # Reset singleton to use new config
    _gemini_client = None


def get_runtime_config() -> Dict[str, Any]:
    """Get current runtime configuration."""
    return {
        "api_key_set": bool(_runtime_api_key or settings.gemini_api_key),
        "model": _runtime_model or settings.gemini_model
    }


# Singleton instance
_gemini_client: Optional[GeminiClient] = None


def get_gemini_client() -> GeminiClient:
    """Get or create the Gemini client singleton."""
    global _gemini_client
    if _gemini_client is None:
        # Use runtime config if available, otherwise fall back to settings
        api_key = _runtime_api_key or settings.gemini_api_key
        model = _runtime_model or settings.gemini_model
        _gemini_client = GeminiClient(api_key=api_key, model=model)
    return _gemini_client
