import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
from bidagent.agents.compliance import numeric_conflict
from bidagent.agents.reviewer import mismatch, static_checks, load_facts
from bidagent.agents.tender_reader import assign_requirements
from bidagent.bom import parse_bom
from bidagent.util import split_models
import make_fixtures


def test_numeric_guard():
    assert numeric_conflict("at least 12 MP", "Sensor: 8 MP CMOS")
    assert not numeric_conflict("at least 12 MP", "Sensor: 24.2 MP")
    assert numeric_conflict("Power up to 30 W", "Power: 45 W")
    assert not numeric_conflict("at least 2 ports", "4 ports; 1 port")


def test_mismatch():
    assert mismatch("Dell OptiPlex 3090", "Dell OptiPlex 3096")["close"]
    assert mismatch("HP LaserJet Pro M401dn", "HP LaserJet Pro 4003dn")["close"]
    assert mismatch("Switch 24-port PoE", "Cisco C1000-24P-4G-L") is None


def test_mod_bom(tmp_path):
    make_fixtures.make_bom_mod(tmp_path / "m.xlsx")
    lines = [l.__dict__ for l in parse_bom(tmp_path / "m.xlsx")]
    assert len(lines) == 6 and lines[0]["model"] == "Dell OptiPlex 3096"
    facts = load_facts(Path(__file__).parent.parent / "knowledge" / "facts.yaml")
    parts = {l["no"]: [{"text": l["model"], "products": ["x"]}] for l in lines}
    f = static_checks(lines, parts, {}, facts, set())
    text = " ".join(x["issue"] for x in f)
    assert "3090" in text and "4003dn" in text and "APS-C" in text
    assert any("listed 2 times" in x["issue"] for x in f)
    assert any("Quantity is '0'" in x["issue"] for x in f)


def test_assign_requirements():
    lines = [{"no": "1", "item": "Network Camera", "is_group": False}, {"no": "2", "item": "Switch", "is_group": False}]
    t = {"requirements": [{"item_no": "2", "text": "PoE 300 W"}, {"item_no": "", "text": "Network camera with IP66"},
                          {"item_no": "9", "text": "Warranty 3 years"}]}
    by, loose = assign_requirements(t, lines)
    assert by == {"2": ["PoE 300 W"], "1": ["Network camera with IP66"]} and loose == ["Warranty 3 years"]
