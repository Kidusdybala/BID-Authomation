"""Research BOM products with Gemini's Google Search grounding."""
import json
import urllib.error
import urllib.parse
import urllib.request

from .library import Library
from .llm import parse_json
from .util import collapse


SYSTEM = """Research the requested product using the live Google Search results available to you.
Prefer the manufacturer's own product page and official datasheets; use reputable distributors only
when the manufacturer does not publish the information. Do not infer specifications. Return one JSON
object: {\"name\": string, \"description\": [string], \"features\": [{\"label\": string,
\"value\": string, \"quote\": string, \"source_url\": string}]}. Each feature must be supported by
one of the cited search sources. Keep quotes short and exact when possible. If you cannot find a
reliable source for the exact model, return empty name, description, and features."""


def research_product(api_key, model, model_name, item_hint="", requirements=(), timeout=180):
    """Return a draft library product and warnings, or (None, warnings) if research fails."""
    prompt = (f"BOM item: {item_hint}\nExact requested model/part number: {model_name}\n"
              f"Tender requirements for this item: {json.dumps(list(requirements), ensure_ascii=False)}\n\n"
              "Search Google for this exact model and extract only specifications relevant to the item "
              "and its tender requirements. Include a source_url for every feature.")
    url = "https://generativelanguage.googleapis.com/v1beta/models/" + urllib.parse.quote(model, safe="") + ":generateContent"
    body = {
        "contents": [{"parts": [{"text": SYSTEM + "\n\n" + prompt}]}],
        "tools": [{"google_search": {}}],
        "generationConfig": {"temperature": 0.2, "maxOutputTokens": 4096},
    }
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json", "x-goog-api-key": api_key})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            result = json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError, ValueError) as exc:
        return None, [f"Google Search research failed for '{model_name}': {exc}"]

    try:
        candidate = result["candidates"][0]
        text = "".join(p.get("text", "") for p in candidate["content"]["parts"])
        data = parse_json(text)
    except (KeyError, IndexError, TypeError, ValueError) as exc:
        return None, [f"Google Search returned no usable product data for '{model_name}': {exc}"]
    if not isinstance(data, dict):
        return None, [f"Google Search returned an unexpected response for '{model_name}'."]

    chunks = (candidate.get("groundingMetadata") or {}).get("groundingChunks") or []
    sources = {}
    for chunk in chunks:
        web = chunk.get("web") or {}
        uri = web.get("uri", "")
        if uri.startswith("https://") or uri.startswith("http://"):
            sources[uri] = collapse(web.get("title", "")) or uri
    features = []
    for feature in data.get("features") or []:
        if not isinstance(feature, dict) or not collapse(feature.get("value")):
            continue
        uri = collapse(feature.get("source_url", ""))
        if uri not in sources:
            continue
        features.append({"label": collapse(feature.get("label", "")),
                         "value": collapse(feature["value"]),
                         "source": f"{sources[uri]} — {uri}",
                         "quote": collapse(feature.get("quote", "")), "verified": True})
    if not features:
        return None, [f"Google Search found no source-backed specifications for '{model_name}'."]

    product = Library.new_product(model_name, name=collapse(data.get("name", "")) or item_hint,
                                  status="draft")
    product["source"] = {"web_research": list(sources.values())}
    heading = collapse(data.get("name", "")) or item_hint or model_name
    product["blocks"] = [{"heading": heading, "description": [collapse(s) for s in (data.get("description") or []) if collapse(s)],
                          "features": features, "image": None, "image_width_cm": None,
                          "caption": heading}]
    return product, [f"{model_name}: web-researched product saved as a draft. Verify each linked source and specification before approval."]
