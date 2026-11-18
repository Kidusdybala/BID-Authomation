"""One client for every provider that speaks the OpenAI chat API
(Gemini, Groq, OpenRouter, Ollama, OpenAI ...).

Token-saving features:
  * every answer is cached on disk (same prompt -> 0 tokens on re-run)
  * per-agent fallback chain (big model first, another model if rate-limited)
  * usage.jsonl log so you can see what each agent costs
"""
import hashlib
import json
import os
import re
import time
import urllib.error
import urllib.request
from pathlib import Path


class LLMUnavailable(Exception):
    pass


def load_env(path):
    p = Path(path)
    if not p.exists():
        return
    for line in p.read_text("utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def parse_json(text):
    t = re.sub(r"<think>.*?</think>", "", text or "", flags=re.S).strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t, flags=re.M).strip()
    try:
        return json.loads(t)
    except ValueError:
        pass
    starts = [i for i in (t.find("{"), t.find("[")) if i != -1]
    if starts:
        s = min(starts)
        e = max(t.rfind("}"), t.rfind("]"))
        if e > s:
            try:
                return json.loads(t[s:e + 1])
            except ValueError:
                pass
    raise ValueError("not valid JSON")


def _json_ok(text):
    try:
        parse_json(text)
        return True
    except ValueError:
        return False


class LLM:
    def __init__(self, cfg, workdir, enabled=True):
        self.cfg = cfg
        self.dir = Path(workdir)
        self.enabled = enabled
        self.cache = self.dir / ".cache" / "llm"
        self.cache.mkdir(parents=True, exist_ok=True)
        self.usage_log = self.dir / "usage.jsonl"

    def chain(self, agent):
        a = self.cfg.get("agents", {})
        return a.get(agent) or a.get("default") or []

    def available(self, agent):
        if not self.enabled:
            return False
        for s in self.chain(agent):
            prov = self.cfg.get("providers", {}).get(s["provider"], {})
            env = prov.get("api_key_env")
            if not env or os.environ.get(env):
                return True
        return False

    def complete(self, agent, system, user, json_mode=False, temperature=0.2,
                 max_tokens=4096, cache=True):
        if not self.enabled:
            raise LLMUnavailable("LLM disabled (--no-llm)")
        chain = self.chain(agent)
        if not chain:
            raise LLMUnavailable(f'no model configured for agent "{agent}"')
        key = hashlib.sha256(json.dumps(
            [agent, system, user, json_mode,
             [(s["provider"], s["model"]) for s in chain]]).encode()).hexdigest()[:40]
        cf = self.cache / f"{key}.json"
        if cache and cf.exists():
            d = json.loads(cf.read_text("utf-8"))
            self._log(agent, d["provider"], d["model"], 0, 0, True)
            return d["text"]

        errors = []
        retry_config = self.cfg.get("retry_attempts", {})
        retry_attempts = max(1, int(retry_config.get(agent, retry_config.get("default", 3))))
        agent_timeout = self.cfg.get("agent_timeouts", {}).get(agent)
        for step in chain:
            prov = self.cfg.get("providers", {}).get(step["provider"])
            if not prov:
                errors.append(f'{step["provider"]}: provider not defined')
                continue
            env = prov.get("api_key_env")
            api_key = os.environ.get(env, "") if env else ""
            if env and not api_key:
                errors.append(f"{step['provider']}: {env} not set")
                continue
            use_json = json_mode
            request_provider = dict(prov)
            if agent_timeout:
                request_provider["timeout"] = agent_timeout
            for attempt in range(retry_attempts):
                try:
                    text, usage = self._call(request_provider, api_key, step["model"], system, user,
                                             use_json, temperature, max_tokens)
                    if cache and (not json_mode or _json_ok(text)):
                        cf.write_text(json.dumps({"provider": step["provider"],
                                                  "model": step["model"], "text": text}),
                                      "utf-8")
                    self._log(agent, step["provider"], step["model"],
                              usage.get("prompt_tokens", 0),
                              usage.get("completion_tokens", 0), False)
                    return text
                except urllib.error.HTTPError as e:
                    body = e.read()[:300].decode("utf-8", "ignore")
                    if e.code == 400 and use_json:
                        use_json = False      # model has no native JSON mode
                        continue
                    if e.code in (429, 500, 502, 503, 504):
                        errors.append(f'{step["provider"]}/{step["model"]}: HTTP {e.code}')
                        if attempt + 1 < retry_attempts:
                            time.sleep(min(3 * 2 ** attempt, 30))
                        continue
                    errors.append(f'{step["provider"]}/{step["model"]}: HTTP {e.code} {body}')
                    break
                except (urllib.error.URLError, TimeoutError, OSError) as e:
                    errors.append(f'{step["provider"]}/{step["model"]}: {e}')
                    if attempt + 1 < retry_attempts:
                        time.sleep(2)
        raise LLMUnavailable("all models failed: " + " | ".join(errors[-4:]))

    def complete_json(self, agent, system, user, **kw):
        text = self.complete(agent, system, user, json_mode=True, **kw)
        try:
            return parse_json(text)
        except ValueError:
            fixed = self.complete(agent, "Return ONLY valid JSON. No commentary.",
                                  "Repair this into valid JSON:\n" + text[:12000],
                                  json_mode=True, cache=False)
            return parse_json(fixed)

    def _call(self, prov, api_key, model, system, user, json_mode, temperature, max_tokens):
        body = {"model": model, "temperature": temperature, "max_tokens": max_tokens,
                "messages": [{"role": "system", "content": system},
                             {"role": "user", "content": user}]}
        if json_mode:
            body["response_format"] = {"type": "json_object"}
        headers = {"Content-Type": "application/json"}
        if api_key:
            headers["Authorization"] = "Bearer " + api_key
        req = urllib.request.Request(prov["base_url"].rstrip("/") + "/chat/completions",
                                     data=json.dumps(body).encode(), headers=headers)
        with urllib.request.urlopen(req, timeout=prov.get("timeout", 240)) as r:
            data = json.loads(r.read().decode("utf-8"))
        text = data["choices"][0]["message"]["content"] or ""
        text = re.sub(r"<think>.*?</think>", "", text, flags=re.S).strip()
        return text, data.get("usage", {}) or {}

    def _log(self, agent, provider, model, pt, ct, cached):
        with self.usage_log.open("a", encoding="utf-8") as f:
            f.write(json.dumps({"ts": int(time.time()), "agent": agent, "provider": provider,
                                "model": model, "prompt_tokens": pt,
                                "completion_tokens": ct, "cached": cached}) + "\n")


def usage_summary(workdir):
    p = Path(workdir) / "usage.jsonl"
    rows = [json.loads(l) for l in p.read_text("utf-8").splitlines() if l.strip()] if p.exists() else []
    agg = {}
    for r in rows:
        a = agg.setdefault(r["agent"], {"calls": 0, "cached": 0, "prompt": 0, "completion": 0})
        a["calls"] += 0 if r["cached"] else 1
        a["cached"] += 1 if r["cached"] else 0
        a["prompt"] += r["prompt_tokens"]
        a["completion"] += r["completion_tokens"]
    return agg
