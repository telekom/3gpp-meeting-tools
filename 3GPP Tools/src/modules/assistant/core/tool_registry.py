"""Validated semantic tool registry for the 3GPP Assistant."""
from __future__ import annotations
import json
from dataclasses import dataclass
from enum import Enum
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Tuple
from modules.assistant.core.agent_models import ToolResult, ToolStatus

class ToolPermission(str, Enum):
    KNOWLEDGE_READ = "knowledge_read"

class ArgumentType(str, Enum):
    STRING = "string"
    INTEGER = "integer"
    BOOLEAN = "boolean"

@dataclass(frozen=True)
class ToolArgument:
    name: str
    type: ArgumentType
    description: str
    required: bool = False
    default: Any = None
    minimum: Optional[int] = None
    maximum: Optional[int] = None
    max_length: Optional[int] = None
    enum: Tuple[Any, ...] = ()

@dataclass(frozen=True)
class ToolDefinition:
    name: str
    description: str
    arguments: Tuple[ToolArgument, ...]
    handler: Callable[..., ToolResult]
    permission: ToolPermission = ToolPermission.KNOWLEDGE_READ

    def ollama_schema(self) -> Dict[str, Any]:
        props, required = {}, []
        for arg in self.arguments:
            s = {"type": arg.type.value, "description": arg.description}
            if arg.enum: s["enum"] = list(arg.enum)
            if arg.minimum is not None: s["minimum"] = arg.minimum
            if arg.maximum is not None: s["maximum"] = arg.maximum
            if arg.max_length is not None and arg.type == ArgumentType.STRING: s["maxLength"] = arg.max_length
            if arg.default is not None: s["default"] = arg.default
            props[arg.name] = s
            if arg.required: required.append(arg.name)
        return {"type":"function","function":{"name":self.name,"description":self.description,
            "parameters":{"type":"object","properties":props,"required":required,"additionalProperties":False}}}

class ToolRegistry:
    def __init__(self, allowed_permissions: Optional[Iterable[ToolPermission]] = None):
        self._definitions = {}
        self._allowed_permissions = set(allowed_permissions or {ToolPermission.KNOWLEDGE_READ})

    def register(self, definition: ToolDefinition) -> None:
        if definition.name in self._definitions:
            raise ValueError(f"Tool already registered: {definition.name}")
        self._definitions[definition.name] = definition

    def get(self, name: str):
        return self._definitions.get(str(name or "").strip())

    def definitions(self):
        return tuple(self._definitions.values())

    def ollama_tools(self):
        return [d.ollama_schema() for d in self.definitions()]

    def execute(self, name: str, arguments: Optional[Mapping[str, Any]]) -> ToolResult:
        d = self.get(name)
        if d is None:
            return ToolResult(ToolStatus.INVALID_REQUEST, message=f"Unknown tool '{name}'.",
                              metadata={"reason":"unknown_tool","available_tools":list(self._definitions)})
        if d.permission not in self._allowed_permissions:
            return ToolResult(ToolStatus.INVALID_REQUEST, message=f"Tool '{name}' is not authorized.",
                              metadata={"reason":"unauthorized_tool"})
        normalized, error = self._validate(d, arguments or {})
        if error:
            return ToolResult(ToolStatus.INVALID_REQUEST, message=error,
                              metadata={"reason":"invalid_arguments","tool":name})
        try:
            result = d.handler(**normalized)
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Tool '{name}' failed unexpectedly: {exc}",
                              metadata={"reason":"tool_exception","tool":name})
        if not isinstance(result, ToolResult):
            return ToolResult(ToolStatus.ERROR, message=f"Tool '{name}' returned an invalid result.")
        return result

    @staticmethod
    def canonical_call_key(name: str, arguments: Optional[Mapping[str, Any]]) -> str:
        return json.dumps({"tool":str(name or "").strip(),"arguments":dict(arguments or {})},
                          sort_keys=True, ensure_ascii=False, separators=(",",":"), default=str)

    @staticmethod
    def _validate(d: ToolDefinition, arguments: Mapping[str, Any]):
        if not isinstance(arguments, Mapping):
            return {}, "Tool arguments must be an object."
        known = {a.name for a in d.arguments}
        extras = sorted(set(arguments)-known)
        if extras: return {}, f"Unexpected argument(s): {', '.join(extras)}."
        out = {}
        for a in d.arguments:
            if a.name not in arguments or arguments[a.name] is None:
                if a.required: return {}, f"Missing required argument '{a.name}'."
                if a.default is not None: out[a.name] = a.default
                continue
            v = arguments[a.name]
            if a.type == ArgumentType.STRING:
                if not isinstance(v,str): return {}, f"Argument '{a.name}' must be a string."
                v=v.strip()
                if a.required and not v: return {}, f"Argument '{a.name}' must not be empty."
                if a.max_length is not None and len(v)>a.max_length:
                    return {}, f"Argument '{a.name}' exceeds maximum length {a.max_length}."
            elif a.type == ArgumentType.INTEGER:
                if isinstance(v,bool) or not isinstance(v,int): return {}, f"Argument '{a.name}' must be an integer."
                if a.minimum is not None and v<a.minimum: return {}, f"Argument '{a.name}' must be >= {a.minimum}."
                if a.maximum is not None and v>a.maximum: return {}, f"Argument '{a.name}' must be <= {a.maximum}."
            elif a.type == ArgumentType.BOOLEAN and not isinstance(v,bool):
                return {}, f"Argument '{a.name}' must be a boolean."
            if a.enum and v not in a.enum: return {}, f"Argument '{a.name}' has an unsupported value."
            out[a.name]=v
        return out, None
