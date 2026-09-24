"""Ollama-native tool-calling adapter for the 3GPP Assistant.

This module translates Ollama REST response dictionaries into small,
provider-independent objects.  It does not implement the research loop.
"""

from __future__ import annotations

from typing import Any, Dict, List, Mapping, Optional, Sequence

from core.ai.ollama_client import OllamaClient, load_ollama_config
from modules.assistant.core.tool_registry import ToolDefinition


from modules.assistant.core.model_types import ModelResponse, ModelToolCall


class OllamaAgentError(RuntimeError):
    """Base error raised by the Ollama agent adapter."""


class OllamaModelNotConfiguredError(OllamaAgentError):
    """Raised when no model has been selected for Ollama."""


class OllamaProtocolError(OllamaAgentError):
    """Raised when Ollama returns a response that cannot be interpreted safely."""


class OllamaAgentAdapter:
    """Thin translation layer between Assistant domain objects and Ollama."""

    def __init__(
        self,
        client: Optional[OllamaClient] = None,
        model: Optional[str] = None,
        options: Optional[Dict[str, Any]] = None,
    ):
        cfg = load_ollama_config()
        self.client = client or OllamaClient(
            host=cfg.get("host"),
            proxy_mode=cfg.get("proxy_mode", "direct"),
        )
        self.model = str(model or cfg.get("selected_model", "")).strip()
        self.options = dict(options or {})

    def complete(
        self,
        *,
        messages: Sequence[Mapping[str, Any]],
        tools: Sequence[ToolDefinition] = (),
        options: Optional[Dict[str, Any]] = None,
    ) -> ModelResponse:
        if not self.model:
            raise OllamaModelNotConfiguredError(
                "No Ollama model is configured. Select an agent-capable local model first."
            )

        tool_schemas = [tool.ollama_schema() for tool in tools]
        merged_options = dict(self.options)
        if options:
            merged_options.update(options)

        response = self.client.chat(
            model=self.model,
            messages=[dict(message) for message in messages],
            tools=tool_schemas or None,
            options=merged_options or None,
        )
        return self._parse_response(response)

    def _parse_response(self, response: Mapping[str, Any]) -> ModelResponse:
        if not isinstance(response, Mapping):
            raise OllamaProtocolError("Ollama returned a non-object chat response.")

        message = response.get("message")
        if not isinstance(message, Mapping):
            raise OllamaProtocolError("Ollama chat response did not contain a valid message object.")

        content = str(message.get("content") or "")
        parsed_calls: List[ModelToolCall] = []

        raw_calls = message.get("tool_calls") or []
        if raw_calls and not isinstance(raw_calls, list):
            raise OllamaProtocolError("Ollama message.tool_calls was not a list.")

        for raw_call in raw_calls:
            if not isinstance(raw_call, Mapping):
                raise OllamaProtocolError("Ollama returned an invalid tool call object.")

            function = raw_call.get("function")
            if not isinstance(function, Mapping):
                raise OllamaProtocolError("Ollama tool call did not contain a function object.")

            name = str(function.get("name") or "").strip()
            if not name:
                raise OllamaProtocolError("Ollama tool call did not contain a function name.")

            arguments = function.get("arguments")
            if arguments is None:
                arguments = {}
            if not isinstance(arguments, Mapping):
                raise OllamaProtocolError(
                    f"Ollama tool call '{name}' arguments were not a JSON object."
                )

            parsed_calls.append(
                ModelToolCall(
                    name=name,
                    arguments=dict(arguments),
                    call_id=str(raw_call.get("id") or ""),
                )
            )

        metrics = {
            key: response.get(key)
            for key in (
                "total_duration",
                "load_duration",
                "prompt_eval_count",
                "prompt_eval_duration",
                "eval_count",
                "eval_duration",
                "done_reason",
            )
            if key in response
        }

        return ModelResponse(
            content=content,
            tool_calls=tuple(parsed_calls),
            raw_message=dict(message),
            metrics=metrics,
            model=str(response.get("model") or self.model),
        )

    @staticmethod
    def assistant_message(response: ModelResponse) -> Dict[str, Any]:
        """Create the assistant-history message required before tool results."""
        message: Dict[str, Any] = {
            "role": "assistant",
            "content": response.content,
        }
        if response.tool_calls:
            message["tool_calls"] = [
                {
                    **({"id": call.call_id} if call.call_id else {}),
                    "function": {
                        "name": call.name,
                        "arguments": dict(call.arguments),
                    },
                }
                for call in response.tool_calls
            ]
        return message

    @staticmethod
    def tool_result_message(
        tool_name: str,
        content: str,
        *,
        call_id: str = "",
    ) -> Dict[str, Any]:
        """Create an Ollama tool-role message for a completed tool request."""
        message: Dict[str, Any] = {
            "role": "tool",
            "content": str(content),
        }
        # Ollama versions/models differ in whether call IDs are emitted/needed.
        # Preserve one when supplied, but do not invent it.
        if call_id:
            message["tool_call_id"] = call_id
        # The name is useful for local diagnostics and accepted by current
        # Ollama chat message handling; it also makes histories readable.
        if tool_name:
            message["name"] = str(tool_name)
        return message
