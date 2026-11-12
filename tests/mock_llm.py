"""A tiny fake OpenAI-style server so the whole pipeline can be tested with no keys.
It answers each agent in a believable way (and makes two deliberate mistakes to prove the
safety checks work: an invented spec, and a wrong 'complies')."""
import json
import re
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

CALLS = []


def words(s):
    return set(re.findall(r"[a-z0-9]{3,}", s.lower()))


def answer(system, user):
    agent = re.match(r"\[agent:(\w+)\]", system).group(1)
    if agent == "tender":
        g = lambda k: (re.search(rf"^{k}:\s*(.+)$", user, re.M | re.I) or [None, ""])[1].strip()
        reqs = [{"item_no": m[0], "text": m[1].strip()} for m in re.findall(r"^REQ ([\d.]+):\s*(.+)$", user, re.M)]
        return {"project_name": g("Project"), "client": g("Client"), "tender_no": g("Tender No"),
                "location": g("Location"), "deadline": "", "summary": "The university wants to improve campus safety.",
                "purpose": ["Improve campus safety."], "objectives": ["Better coverage"], "scope": [],
                "requirements": reqs, "general_requirements": []}
    if agent == "datasheet" and "DATASHEET PAGES" in user:
        feats = []
        for pg in re.finditer(r"\[#(\d+)\][^\n]*\n(.*?)(?=\n\n\[#|\Z)", user, re.S):
            for ln in pg.group(2).splitlines():
                m = re.match(r"([A-Za-z ]+): (.+)", ln)
                if m:
                    feats.append({"label": m[1], "value": m[2], "quote": ln, "page": int(pg.group(1))})
        feats.append({"label": "Waterproof rating", "value": "IP99", "quote": "Fully waterproof to IP99", "page": 1})
        return {"name": "Thermal Camera", "description": ["Zeta TH900X is a thermal imaging camera for perimeter detection."],
                "features": feats}
    if agent == "datasheet":                      # second pass
        d = json.loads(user.split("\n\nPAGES:")[0])
        res = []
        for r in d["requirements"]:
            if "speaker" in r["text"].lower():
                res.append({"i": r["i"], "found": False})
        return {"results": res}
    if agent == "compliance":
        d = json.loads(user)
        res = []
        for r in d["requirements"]:
            rw, best, sc = words(r["text"]), None, 0
            for f in d["features"]:
                s = len(rw & words(f["text"]))
                if s > sc:
                    best, sc = f, s
            if "12 MP" in r["text"] and best:
                res.append({"i": r["i"], "features": [best["i"]], "status": "complies", "note": "wrongly says yes"})
            elif best and sc >= 1:
                res.append({"i": r["i"], "features": [best["i"]], "status": "complies", "note": ""})
            else:
                res.append({"i": r["i"], "features": [], "status": "not_found", "note": ""})
        return {"results": res}
    if agent == "writer":
        d = json.loads(re.search(r"WRITE THESE SECTIONS:\n(.*)$", user, re.S).group(1))
        return {"sections": {s["title"]: ["Riverside University needs a safer campus. We will supply the systems in the BOM.",
                                          "Our design targets 98% coverage."] for s in d}}
    if agent == "reviewer":
        return {"findings": [{"where": "Item 3", "severity": "low", "issue": "Rack quantity looks low for a 42U rack project.", "fix": "Check"}]}
    return {}


class H(BaseHTTPRequestHandler):
    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        system, user = body["messages"][0]["content"], body["messages"][1]["content"]
        try:
            text = json.dumps(answer(system, user))
        except Exception as e:                     # surface mock bugs loudly
            text = json.dumps({"mock_error": str(e)})
        CALLS.append(re.match(r"\[agent:(\w+)\]", system).group(1))
        out = {"choices": [{"message": {"content": text}}], "usage": {"prompt_tokens": len(user) // 4, "completion_tokens": len(text) // 4}}
        b = json.dumps(out).encode()
        self.send_response(200); self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b))); self.end_headers(); self.wfile.write(b)

    def log_message(self, *a):
        pass


def start(port=0):
    srv = HTTPServer(("127.0.0.1", port), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, srv.server_address[1]
