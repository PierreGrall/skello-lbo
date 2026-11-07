"""Tests for the engine and for the workbook's formula wiring. Run with: pytest tests -q"""
from __future__ import annotations

import copy
import json
import shutil
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from engine import curve_multiple, load_inputs, run, shift_driver, xirr  # noqa: E402
from recalc import recalc  # noqa: E402
from datetime import date  # noqa: E402

INP = load_inputs()


@pytest.fixture(params=INP["scenarios"])
def scenario(request):
    return request.param


def test_sources_equal_uses():
    e = run(INP, "Base").entry
    d = INP["deal"]
    assert e["total_sources"] == pytest.approx(d["primary_cash"] + e["deal_fees"] + e["oid"] + e["secondary"])


def test_pre_deal_holders_receive_ev_plus_cash_less_costs():
    e = run(INP, "Base").entry
    assert e["old_holder_value"] == pytest.approx(e["ev_entry"] + INP["deal"]["existing_cash"] - e["deal_fees"] - e["oid"])


def test_arr_rolls_forward(scenario):
    r = run(INP, scenario).rows
    for t in range(6):
        assert r["arr_fr"][t] == pytest.approx(r["arr_fr_open"][t] - r["arr_fr_churn"][t] + r["arr_fr_upsell"][t] + r["arr_fr_new"][t])
        assert r["arr_total"][t] == pytest.approx(r["arr_fr"][t] + r["arr_in"][t] + r["arr_acq"][t])


def test_cash_rolls_forward_and_respects_minimum(scenario):
    res = run(INP, scenario)
    r = res.rows
    cash = res.entry["cash_entry"]
    for t in range(6):
        cash += r["pre_fin_cf"][t] + r["draw"][t] + r["gap"][t] - r["repay"][t]
        assert r["cash"][t] == pytest.approx(cash)
        assert r["cash"][t] >= INP["deal"]["min_cash"] - 1e-9
        assert r["line"][t] <= INP["deal"]["line_commitment"] + 1e-9


def test_bridge_sums_to_exit_equity(scenario):
    b = run(INP, scenario).bridge
    parts = b["entry_equity"] + b["organic"] + b["acquired"] + b["multiple"] + b["net_debt"] + b["costs"]
    assert parts == pytest.approx(b["exit_equity"])


def test_xirr_hand_case():
    # 100 becomes 200 over exactly 730 days: 2^(365/730) - 1.
    assert xirr([-100, 200], [date(2026, 1, 1), date(2028, 1, 1)]) == pytest.approx(2 ** (365 / 730) - 1, abs=1e-8)


def test_moic_hand_case_base():
    """Rebuild the Base Dec-30 exit by hand from the balance rows."""
    res = run(INP, "Base")
    d, ex = INP["deal"], INP["exit"]
    ev = res.rows["revenue"][4] * res.exits["Dec-30"]["multiple"]
    equity = ev * (1 - ex["cost_pct"]) - res.rows["net_debt"][4]
    mip = ex["mip_share"] * max(equity - ex["mip_hurdle_moic"] * d["post_money"], 0)
    assert (equity - mip) * d["bp_equity"] / d["post_money"] / d["bp_equity"] == pytest.approx(res.exits["Dec-30"]["moic"])


def test_curve_interpolates_and_is_flat_beyond_its_ends():
    assert curve_multiple(INP, 40) == pytest.approx(6.0)
    assert curve_multiple(INP, 5) == 2.5
    assert curve_multiple(INP, 90) == 12.0


def test_preference_protects_the_downside_and_costs_base_nothing():
    down, base = run(INP, "Downside").exits["Dec-30"], run(INP, "Base").exits["Dec-30"]
    assert down["moic_pref"] > down["moic"]
    assert base["moic_pref"] == pytest.approx(base["moic"])


def test_financial_buyer_pays_less_than_the_curve_in_base():
    res = run(INP, "Base")
    assert res.next_buyer["ev"] < res.exits["Dec-30"]["ev"]


def test_retention_flex_lowers_new_logos_when_retention_rises():
    from analysis import retention_flex
    up = retention_flex(INP, "Base", 0.03)
    assert up["drivers"]["grr_france"]["Base"][1] == pytest.approx(0.89)
    assert up["drivers"]["new_france"]["Base"][1] < INP["drivers"]["new_france"]["Base"][1]


def test_returns_rise_with_retention_and_fall_with_price():
    base = run(INP, "Base").exits["Dec-30"]["moic"]
    better = run(shift_driver(INP, "Base", "grr_france", delta=0.03), "Base").exits["Dec-30"]["moic"]
    dearer = copy.deepcopy(INP)
    dearer["deal"]["post_money"] = 450.0
    assert better > base
    assert run(dearer, "Base").exits["Dec-30"]["moic"] < base


def test_interest_is_on_opening_balance_so_no_circularity():
    r = run(INP, "Base").rows
    rate = INP["financing"]["euribor"] + INP["financing"]["margin"]
    assert r["interest"][0] == pytest.approx(40 * rate * 0.5)


def test_scenarios_are_ordered():
    m = {s: run(INP, s).exits["Dec-30"]["moic"] for s in INP["scenarios"]}
    assert m["Downside"] < m["Base"] < m["Plan"]


def test_plan_matches_public_targets():
    r = run(INP, "Plan").rows
    assert r["arr_total"][4] >= 200  # More than EUR 200M by 2030 (Evidence E14).
    assert r["arr_total"][2] >= 100  # EUR 100M by 2028 (Evidence E15).


def test_ltm_calibration_ties_to_public_facts():
    wb = load_workbook(ROOT / "model" / "Skello_LBO_Model.xlsx", data_only=True)
    t = wb["Triangulation"]
    assert t["D13"].value > 0  # Profitable.
    assert abs(t["D15"].value / 400 - 1) < 0.05  # About 400 staff.


def _set_scalar(ws, label, value):
    for row in ws.iter_rows(min_col=2, max_col=2):
        if row[0].value == label:
            ws[f"D{row[0].row}"] = value
            return
    raise KeyError(label)


def _shift_driver_cells(ws, label, scenario_row, delta):
    for row in ws.iter_rows(min_col=2, max_col=2):
        if row[0].value == label:
            r = row[0].row + 1 + scenario_row
            for col in "FGHIJK":
                ws[f"{col}{r}"] = ws[f"{col}{r}"].value + delta
            return
    raise KeyError(label)


def test_workbook_follows_flexed_inputs(tmp_path):
    """Stress the workbook's own input cells and require the formulas to land on the engine's answer.

    This is what catches a hard-coded number hiding inside a formula, which a
    reconciliation at the delivered inputs alone would never reveal.
    """
    if shutil.which("soffice") is None:
        pytest.skip("LibreOffice not available")
    path = tmp_path / "flexed.xlsx"
    shutil.copy(ROOT / "model" / "Skello_LBO_Model.xlsx", path)
    wb = load_workbook(path)
    ws = wb["Inputs"]
    _set_scalar(ws, "Primary cash to the balance sheet", 25.0)
    _set_scalar(ws, "Unitranche", 70.0)
    _set_scalar(ws, "3M Euribor (flat)", 0.05)
    _set_scalar(ws, "Post-money equity value", 410.0)
    _set_scalar(ws, "Tax losses carried forward at closing", 5.0)
    _shift_driver_cells(ws, "Gross revenue retention, France", 0, -0.06)
    _shift_driver_cells(ws, "Acquired ARR (EUR M)", 0, 8.0)
    wb.save(path)
    recalc(path)

    flexed = copy.deepcopy(INP)
    flexed["deal"].update(primary_cash=25.0, unitranche=70.0, post_money=410.0)
    flexed["financing"]["euribor"] = 0.05
    flexed["tax"]["nol_opening"] = 5.0
    flexed["drivers"]["grr_france"]["Base"] = [v - 0.06 for v in flexed["drivers"]["grr_france"]["Base"]]
    flexed["drivers"]["acquired_arr"]["Base"] = [v + 8.0 for v in flexed["drivers"]["acquired_arr"]["Base"]]
    res = run(flexed, "Base")

    maps = json.loads((ROOT / "outputs" / "cell_map.json").read_text())["Base"]
    values = load_workbook(path, data_only=True)["Base"]
    for key in ("arr_total", "ebitda", "tax", "cash", "line", "gap", "net_debt", "nol"):
        for i, col in enumerate("FGHIJK"):
            assert values[f"{col}{maps['rows'][key]}"].value == pytest.approx(res.rows[key][i], abs=1e-6), (key, col)
    assert values[f"J{maps['rows']['x_moic']}"].value == pytest.approx(res.exits["Dec-30"]["moic"], abs=1e-6)
    assert values[maps["irr"]["Dec-30"]].value == pytest.approx(res.exits["Dec-30"]["irr"], abs=1e-6)
