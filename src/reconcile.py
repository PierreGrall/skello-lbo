"""Cell-by-cell reconciliation of the recalculated workbook against the Python engine.

The workbook and engine.py are two implementations of the same model. A formula
slip in either surfaces here as a break. This proves consistency, not correctness:
an error of reasoning shared by both would pass.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from openpyxl import load_workbook

from analysis import firepower, required_arr
from engine import ROOT, curve_multiple, load_inputs, run

XLSX = ROOT / "model" / "Skello_LBO_Model.xlsx"
CELL_MAP = ROOT / "outputs" / "cell_map.json"
TOLERANCE = 1e-6
COLS = ["F", "G", "H", "I", "J", "K"]
EXIT_COL = {"Dec-29": "I", "Dec-30": "J", "Dec-31": "K"}
ROW_KEYS = (
    "arr_fr_open arr_fr_churn arr_fr_upsell arr_fr_new arr_fr arr_in_open arr_in_churn arr_in_upsell arr_in_new arr_in "
    "arr_acq_open arr_acq_add arr_acq_growth arr_acq arr_org arr_total arr_growth org_growth rev_org rev_acq revenue cogs sm_var sm_cs "
    "rd ga ebitda_org ebitda_acq ebitda ebitda_margin r40 capex nwc cir interest commitment interest_income fin_charges "
    "deductible taxable nol_open nol_use nol tax ma_spend integration pre_fin_cf need draw gap repay line cash unitranche "
    "net_debt yield_per_dau fr_dau fr_penetration interest_cover debt_to_arr debt_to_ebitda fte_implied".split()
)
EXIT_KEYS = {"x_mult": "multiple", "x_ev": "ev", "x_ev_arr": "ev_arr", "x_ev_ebitda": "ev_ebitda", "x_costs": "costs", "x_equity": "equity",
             "x_follow_on": "follow_on", "x_mip": "mip", "x_dist": "distributable", "x_bp": "bp_proceeds", "x_bp_inv": "bp_invested",
             "x_moic": "moic", "x_bp_pref": "bp_pref", "x_moic_pref": "moic_pref", "x_moic_pref_bp": "moic_pref_bp_only"}


def reconcile(path: Path = XLSX) -> tuple[int, list[str]]:
    inp = load_inputs()
    wb = load_workbook(path, data_only=True)
    maps = json.loads(CELL_MAP.read_text())
    compared, breaks = 0, []

    def compare(where, excel, python):
        nonlocal compared
        compared += 1
        if excel is None:
            excel = 0.0
        if abs(float(excel) - float(python)) > TOLERANCE:
            breaks.append(f"{where}: excel {excel} vs python {python}")

    for scenario in inp["scenarios"]:
        ws = wb[scenario]
        m = maps[scenario]
        res = run(inp, scenario)
        for key in ROW_KEYS:
            row = m["rows"][key]
            for i, col in enumerate(COLS):
                compare(f"{scenario}!{col}{row} {key}", ws[f"{col}{row}"].value, res.rows[key][i])
        for label, col in EXIT_COL.items():
            for xkey, ekey in EXIT_KEYS.items():
                compare(f"{scenario} {label} {xkey}", ws[f"{col}{m['rows'][xkey]}"].value, res.exits[label][ekey])
            compare(f"{scenario} {label} irr", ws[m["irr"][label]].value, res.exits[label]["irr"])
        for key, ref in m["bridge"].items():
            if key != "check":
                compare(f"{scenario} bridge {key}", ws[ref].value, res.bridge[key])

    deal = wb["Deal"]
    res = run(inp, "Base")
    for ref, key in {"D8": "total_sources", "D12": "deal_fees", "D13": "oid", "D14": "secondary", "D20": "cash_entry",
                     "D21": "ev_entry", "D22": "m_entry", "D23": "net_debt_entry", "E29": "bp_share", "D32": "rollover",
                     "D33": "old_holder_value"}.items():
        compare(f"Deal!{ref} {key}", deal[ref].value, res.entry[key])

    tri = wb["Triangulation"]
    d_ltm = tri["D16"].value
    compare("Triangulation curve at entry", tri["D17"].value, curve_multiple(inp, d_ltm))

    nb_tab = wb["NextBuyer"]
    for scenario in inp["scenarios"]:
        nb = run(inp, scenario).next_buyer
        for key, ref in maps["_next_buyer"][scenario].items():
            if key != "premium":
                compare(f"NextBuyer {scenario} {key}", nb_tab[ref].value, nb[key])

    analytics = wb["Analytics"]
    fp = firepower(inp)
    cells = maps["_analytics"]["cells"]
    for key, fkey in {"near_term": "near_term", "cum_fcf": "cum_fcf_pre_ma", "available": "available", "organic": "organic_dec30",
                      "needed": "needed_at_base", "cost": "cost_at_base", "gap": "gap_at_base", "bp_gap": "bp_share_gap",
                      "debt_low": "debt_capacity_low", "debt_high": "debt_capacity_high"}.items():
        compare(f"Analytics {key}", analytics[cells[key]].value, fp[fkey])
    req = required_arr(inp)
    for target, row in maps["_analytics"]["required"].items():
        for col, m in zip("DEFGHIJ", [4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0]):
            compare(f"Analytics required {target} {m}", analytics[f"{col}{row}"].value, req["targets"][target][req["multiples"].index(m)])

    checks = wb["Checks"]
    failures = [r[0].value for r in checks.iter_rows(min_row=5, max_col=5, min_col=2) if r[3].value == "FAIL"]
    breaks += [f"Checks tab: {name}" for name in failures]
    return compared, breaks


if __name__ == "__main__":
    n, found = reconcile()
    if found:
        print(f"RECONCILIATION FAILED: {len(found)} breaks out of {n} comparisons")
        for line in found[:40]:
            print("  ", line)
        sys.exit(1)
    print(f"Reconciled {n} values across Excel and Python: 0 breaks (tolerance {TOLERANCE})")
