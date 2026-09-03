"""The provider-agnostic model layer, and the hygiene of the recorded fixtures."""
from __future__ import annotations

import glob
import importlib
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "fixtures" / "llm"


@pytest.fixture(autouse=True)
def _restore_llm_module():
    """Tests here reload gridlint.llm with a temporary fixture directory.
    Put the real module back afterwards so later tests still see the shipped fixtures."""
    yield
    from gridlint import explain, llm
    importlib.reload(llm)
    importlib.reload(explain)


def _load(tmp_path, monkeypatch, **env):
    for k in ("LLM_PROVIDER", "ANTHROPIC_API_KEY", "OPENAI_API_KEY", "GROQ_API_KEY",
              "GEMINI_API_KEY", "LLM_RECORD", "LLM_MODEL"):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setenv("LLM_FIXTURE_DIR", str(tmp_path))
    for k, v in env.items():
        monkeypatch.setenv(k, v)
    from gridlint import llm
    importlib.reload(llm)
    return llm


def test_with_no_key_the_provider_is_replay_and_says_so_clearly(tmp_path, monkeypatch):
    llm = _load(tmp_path, monkeypatch)
    assert llm.provider_info()["provider"] == "replay"
    with pytest.raises(llm.FixtureMissing) as e:
        llm.complete([{"role": "user", "content": "hi"}], tag="t")
    assert "API key" in str(e.value)


def test_a_fixture_is_served_before_any_network_call(tmp_path, monkeypatch):
    llm = _load(tmp_path, monkeypatch, LLM_PROVIDER="anthropic")   # no key: a call would fail
    msgs = [{"role": "user", "content": "hello"}]
    key = llm._key("sys", msgs, True, "tag1")
    (tmp_path / f"{key}.json").write_text(json.dumps({"text": '{"a":1}', "model": "m"}), encoding="utf-8")
    r = llm.complete(msgs, system="sys", json_mode=True, tag="tag1")
    assert r.cached and r.json() == {"a": 1}


def test_the_cache_key_separates_different_prompts(tmp_path, monkeypatch):
    llm = _load(tmp_path, monkeypatch)
    m = [{"role": "user", "content": "a"}]
    assert llm._key("", m, False, "v1") != llm._key("", m, False, "v2")
    assert llm._key("", m, False, "v1") != llm._key("", [{"role": "user", "content": "b"}], False, "v1")
    assert llm._key("", m, False, "v1") == llm._key("", m, False, "v1")


def test_json_extraction_survives_fences_and_prose(tmp_path, monkeypatch):
    llm = _load(tmp_path, monkeypatch)
    assert llm.parse_json('```json\n{"x": [1,2]}\n```') == {"x": [1, 2]}
    assert llm.parse_json('Sure, here it is: {"x": 1} — hope that helps') == {"x": 1}
    assert llm.parse_json("[1,2,3]") == [1, 2, 3]
    with pytest.raises(llm.LLMError):
        llm.parse_json("there is no json in this sentence")


def test_cli_scaffolding_is_stripped(tmp_path, monkeypatch):
    llm = _load(tmp_path, monkeypatch)
    raw = ("**Draft for the file owner**\n\n"
           "> The total leaves out a row.\n"
           "> It matters because the runway is wrong.\n\n"
           "---\n\n"
           "**More natural way to say it**\n\n"
           "Some coaching that belongs to the operator, not the product.")
    out = llm.clean_cli_output(raw)
    assert out == "The total leaves out a row.\nIt matters because the runway is wrong."
    assert "More natural" not in out and "**" not in out


def test_cleaner_leaves_an_ordinary_answer_alone(tmp_path, monkeypatch):
    llm = _load(tmp_path, monkeypatch)
    text = "The total in C16 leaves out Contractors, so runway reads 38.6 instead of 5.2."
    assert llm.clean_cli_output(text) == text


@pytest.mark.parametrize("path", sorted(glob.glob(str(FIXTURES / "*.json"))))
def test_recorded_fixtures_carry_no_chat_scaffolding(path):
    """Fixtures ship in the repo and are shown to users, so they must be clean prose."""
    d = json.loads(Path(path).read_text(encoding="utf-8"))
    text = d["text"]
    assert text.strip(), f"{path} is empty"
    for banned in ("More natural way", "Draft for the file owner", "Notes on choices",
                   "notes on the choices"):
        assert banned.lower() not in text.lower(), f"{path} still contains {banned!r}"
    assert not text.lstrip().startswith(">"), f"{path} starts with a blockquote"
    assert "**" not in text, f"{path} contains markdown emphasis"


def test_every_fixture_is_reachable_from_the_bundled_samples():
    """A fixture nobody can replay is dead weight; every one should belong to a sample."""
    from gridlint.audit import audit
    from gridlint import explain

    used: set[str] = set()
    from gridlint import llm

    original = llm.complete

    def spy(messages, *, system="", json_mode=False, max_tokens=2048, temperature=0.0, tag=""):
        used.add(llm._key(system, messages, json_mode, tag))
        return original(messages, system=system, json_mode=json_mode,
                        max_tokens=max_tokens, temperature=temperature, tag=tag)

    explain.complete = spy
    try:
        for sample in ("board-model.xlsx", "runway.xlsx"):
            d = audit(ROOT / "samples" / sample).to_dict()
            for fd in d["findings"]:
                f = explain._finding_from_dict(fd)
                try:
                    explain.explain_finding(f, sheet_context=explain.sheet_context(d["path"], f.sheet))
                except Exception:
                    pass
            try:
                explain.review_note(d)
            except Exception:
                pass
    finally:
        explain.complete = original

    on_disk = {Path(p).stem for p in glob.glob(str(FIXTURES / "*.json"))}
    orphans = on_disk - used
    assert not orphans, f"fixtures nothing replays: {sorted(orphans)}"
