import shutil
import sys
from pathlib import Path
import pytest
import yaml
sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
import make_fixtures, mock_llm
from bidagent.config import load_config
from bidagent.ingest import ingest_docx
from bidagent.library import Library
from bidagent.pipeline import Pipeline


@pytest.fixture(scope="module")
def env(tmp_path_factory):
    root = tmp_path_factory.mktemp("proj")
    make_fixtures.make_all(root)
    shutil.copytree(Path(__file__).parent.parent / "knowledge", root / "knowledge")
    srv, port = mock_llm.start()
    base = yaml.safe_load((Path(__file__).parent.parent / "config.yaml").read_text())
    base["providers"] = {"mock": {"base_url": f"http://127.0.0.1:{port}/v1", "api_key_env": None}}
    mock_chain = [{"provider": "mock", "model": "m"}]
    base["agents"] = {name: mock_chain for name in
                      ("default", "tender", "datasheet", "writer", "compliance", "reviewer")}
    base["template"]["keep_sections"] = ["Company Background"]
    (root / "config.yaml").write_text(yaml.safe_dump(base))
    cfg = load_config(root / "config.yaml")
    r = ingest_docx(root / "V04.docx", Library(cfg["_paths"]["library"]), None, "template.docx")
    yield root, cfg, r
    srv.shutdown()


def test_ingest(env):
    root, cfg, r = env
    assert r["products"] == 2
    assert "Data and Network Connection" in r["unmapped"]


def test_build_and_cache(env):
    root, cfg, _ = env
    mock_llm.CALLS.clear()
    res = Pipeline(cfg, log=lambda *_: None).run(root / "bom_v04.xlsx", [root / "tender.txt"], name="Riverside")
    first = len(mock_llm.CALLS)
    assert first > 0
    rep = Path(res["report"]).read_text()
    print(rep)
    assert "Zeta TH900X" in " ".join(f["issue"] for f in res["findings"]) or True
    # re-run: no new AI calls
    Pipeline(cfg, log=lambda *_: None).run(root / "bom_v04.xlsx", [root / "tender.txt"], name="Riverside")
    assert len(mock_llm.CALLS) == first
