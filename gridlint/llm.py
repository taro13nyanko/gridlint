"""Provider-agnostic LLM client with record/replay.

Why: judges must be able to run the product with ZERO API keys (replay mode),
and the builder must not be locked to one vendor. Every call goes through
`complete()`; responses are cached under fixtures/ keyed by a hash of the
request, so a demo can be replayed byte-for-byte offline.

Providers (env LLM_PROVIDER):
  replay      - only serve cached fixtures; error if missing (default when no key)
  anthropic   - ANTHROPIC_API_KEY, model default claude-sonnet-5
  openai      - OPENAI_API_KEY
  groq        - GROQ_API_KEY (OpenAI-compatible)
  gemini      - GEMINI_API_KEY (OpenAI-compatible endpoint)
  ollama      - local, OLLAMA_HOST (default http://localhost:11434)
  claude-cli  - `claude -p` (Claude Code subscription; dev only)

Set LLM_RECORD=1 to write fixtures while using a live provider.
"""
from __future__ import annotations

import hashlib
import json
import os
import subprocess
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import httpx

FIXTURE_DIR = Path(os.environ.get("LLM_FIXTURE_DIR", Path(__file__).resolve().parent.parent / "fixtures" / "llm"))


class LLMError(RuntimeError):
    pass


class FixtureMissing(LLMError):
    pass


@dataclass
class LLMResult:
    text: str
    provider: str
    model: str
    cached: bool
    latency_ms: int
    usage: dict[str, Any] = field(default_factory=dict)

    def json(self) -> Any:
        return parse_json(self.text)


def parse_json(text: str) -> Any:
    """Tolerant JSON extraction: strips code fences and leading prose."""
    t = text.strip()
    if t.startswith("```"):
        t = t.split("\n", 1)[1] if "\n" in t else t[3:]
        if t.rstrip().endswith("```"):
            t = t.rstrip()[:-3]
    t = t.strip()
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        pass
    for opener, closer in (("{", "}"), ("[", "]")):
        i = t.find(opener)
        j = t.rfind(closer)
        if i != -1 and j > i:
            try:
                return json.loads(t[i : j + 1])
            except json.JSONDecodeError:
                continue
    raise LLMError("model did not return JSON: " + t[:200])


def _detect_provider() -> str:
    p = os.environ.get("LLM_PROVIDER", "").strip().lower()
    if p:
        return p
    if os.environ.get("ANTHROPIC_API_KEY"):
        return "anthropic"
    if os.environ.get("OPENAI_API_KEY"):
        return "openai"
    if os.environ.get("GROQ_API_KEY"):
        return "groq"
    if os.environ.get("GEMINI_API_KEY"):
        return "gemini"
    return "replay"


DEFAULT_MODELS = {
    "anthropic": "claude-sonnet-5",
    "openai": "gpt-5-mini",
    "groq": "llama-3.3-70b-versatile",
    "gemini": "gemini-2.5-flash",
    "ollama": "qwen3:8b",
    "claude-cli": "claude-cli",
    "replay": "replay",
}

OPENAI_COMPAT = {
    "openai": ("https://api.openai.com/v1", "OPENAI_API_KEY"),
    "groq": ("https://api.groq.com/openai/v1", "GROQ_API_KEY"),
    "gemini": ("https://generativelanguage.googleapis.com/v1beta/openai", "GEMINI_API_KEY"),
    "ollama": (os.environ.get("OLLAMA_HOST", "http://localhost:11434").rstrip("/") + "/v1", None),
}


def provider_info() -> dict[str, str]:
    p = _detect_provider()
    return {
        "provider": p,
        "model": os.environ.get("LLM_MODEL") or DEFAULT_MODELS.get(p, "?"),
        "record": "1" if os.environ.get("LLM_RECORD") == "1" else "0",
    }


def _key(system: str, messages: list[dict[str, str]], json_mode: bool, tag: str) -> str:
    h = hashlib.sha256()
    h.update(json.dumps({"s": system, "m": messages, "j": json_mode, "t": tag}, ensure_ascii=False, sort_keys=True).encode())
    return h.hexdigest()[:24]


def complete(
    messages: list[dict[str, str]],
    *,
    system: str = "",
    json_mode: bool = False,
    max_tokens: int = 2048,
    temperature: float = 0.0,
    tag: str = "",
) -> LLMResult:
    """Run one completion. `tag` namespaces fixtures (e.g. 'extract_v2')."""
    provider = _detect_provider()
    model = os.environ.get("LLM_MODEL") or DEFAULT_MODELS.get(provider, "")
    key = _key(system, messages, json_mode, tag)
    fx = FIXTURE_DIR / f"{key}.json"

    if fx.exists():
        data = json.loads(fx.read_text(encoding="utf-8"))
        return LLMResult(
            text=data["text"], provider="replay", model=data.get("model", "?"), cached=True, latency_ms=0,
            usage=data.get("usage", {}),
        )
    if provider == "replay":
        raise FixtureMissing(
            f"No API key set and no fixture for this request (tag={tag!r}). "
            "Set ANTHROPIC_API_KEY / OPENAI_API_KEY / GROQ_API_KEY / GEMINI_API_KEY, or use the demo data."
        )

    t0 = time.time()
    if provider == "anthropic":
        text, usage = _anthropic(model, system, messages, max_tokens, temperature, json_mode)
    elif provider in OPENAI_COMPAT:
        text, usage = _openai_compat(provider, model, system, messages, max_tokens, temperature, json_mode)
    elif provider == "claude-cli":
        text, usage = _claude_cli(system, messages, json_mode)
    else:
        raise LLMError(f"unknown LLM_PROVIDER {provider!r}")
    ms = int((time.time() - t0) * 1000)

    if os.environ.get("LLM_RECORD") == "1":
        FIXTURE_DIR.mkdir(parents=True, exist_ok=True)
        fx.write_text(
            json.dumps(
                {"tag": tag, "provider": provider, "model": model, "text": text, "usage": usage,
                 "system": system, "messages": messages},
                ensure_ascii=False, indent=1,
            ),
            encoding="utf-8",
        )
    return LLMResult(text=text, provider=provider, model=model, cached=False, latency_ms=ms, usage=usage)


def _anthropic(model, system, messages, max_tokens, temperature, json_mode):
    import anthropic  # lazy: optional dependency

    client = anthropic.Anthropic()
    kw: dict[str, Any] = {"model": model, "max_tokens": max_tokens, "messages": messages, "temperature": temperature}
    if system or json_mode:
        kw["system"] = (system + ("\nRespond with JSON only." if json_mode else "")).strip()
    for attempt in range(3):
        try:
            r = client.messages.create(**kw)
            text = "".join(getattr(b, "text", "") for b in r.content)
            return text, {"input": r.usage.input_tokens, "output": r.usage.output_tokens}
        except anthropic.RateLimitError:
            time.sleep(2 * (attempt + 1))
        except anthropic.APIStatusError as e:  # pragma: no cover
            if e.status_code >= 500 and attempt < 2:
                time.sleep(2)
                continue
            raise LLMError(str(e)) from e
    raise LLMError("anthropic: rate limited")


def _openai_compat(provider, model, system, messages, max_tokens, temperature, json_mode):
    base, key_env = OPENAI_COMPAT[provider]
    key = os.environ.get(key_env, "ollama") if key_env else "ollama"
    msgs = ([{"role": "system", "content": system}] if system else []) + messages
    body: dict[str, Any] = {"model": model, "messages": msgs, "max_tokens": max_tokens, "temperature": temperature}
    if json_mode and provider != "ollama":
        body["response_format"] = {"type": "json_object"}
    if json_mode and provider == "ollama":
        body["format"] = "json"
    for attempt in range(3):
        r = httpx.post(
            f"{base}/chat/completions", json=body, timeout=120,
            headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
        )
        if r.status_code == 429 or r.status_code >= 500:
            time.sleep(2 * (attempt + 1))
            continue
        if r.status_code >= 400:
            raise LLMError(f"{provider} {r.status_code}: {r.text[:300]}")
        d = r.json()
        return d["choices"][0]["message"]["content"], d.get("usage", {})
    raise LLMError(f"{provider}: retries exhausted")


def _claude_cli(system, messages, json_mode):
    """Development-only provider: the Claude Code CLI, used to record fixtures.

    Runs in a scratch directory outside any project, because the CLI loads the
    CLAUDE.md of wherever it starts and applies those standing instructions to
    its answer -- which an API call never does.
    """
    import tempfile

    prompt = "\n\n".join(f"[{m['role']}]\n{m['content']}" for m in messages)
    if system:
        prompt = f"[system]\n{system}\n\n" + prompt
    if json_mode:
        prompt += "\n\nRespond with JSON only, no prose, no code fences."
    with tempfile.TemporaryDirectory(prefix="gridlint-llm-") as cwd:
        r = subprocess.run(
            ["claude", "-p", prompt, "--output-format", "text", "--tools", ""],
            capture_output=True, text=True, encoding="utf-8", timeout=300, cwd=cwd,
        )
    if r.returncode != 0:
        raise LLMError("claude -p failed: " + (r.stderr or "")[:300])
    return clean_cli_output(r.stdout), {}


#: Section headings a chat assistant adds after an answer that an API never returns.
_CLI_CUTS = ("more natural way to say it", "a few notes on the choices",
             "notes on the choices", "note on choices", "notes on choices",
             "a note on the choices")


def clean_cli_output(text: str) -> str:
    """Strip the conversational scaffolding the CLI wraps around an answer.

    The CLI applies the operator's own standing instructions from CLAUDE.md, so
    a recorded fixture would otherwise carry a trailing writing-coach section
    and a bold "Draft for the file owner" label into the product's UI.
    """
    kept: list[str] = []
    for line in (text or "").splitlines():
        stripped = line.strip().lower().strip("*_# >")
        if any(stripped.startswith(cut) for cut in _CLI_CUTS):
            break
        if line.strip() in ("---", "***", "___") and kept:
            break
        kept.append(line)

    while kept and (not kept[0].strip() or _is_label(kept[0])):
        kept.pop(0)
    while kept and not kept[-1].strip():
        kept.pop()
    body = "\n".join(ln.lstrip("> ").rstrip() for ln in kept)
    return body.strip().replace("**", "")


def _is_label(line: str) -> bool:
    """True for a heading line such as "**Draft for the file owner**" or "Draft:"."""
    t = line.strip().rstrip(":")
    if not t:
        return False
    bold_only = t.startswith("**") and t.endswith("**") and t.count("**") == 2
    short_heading = line.strip().endswith(":") and len(t) < 60 and t.count(" ") < 6
    return bold_only or short_heading
