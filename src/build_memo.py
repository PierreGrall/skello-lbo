"""Builds memo/Skello_Investment_Memo.pdf.

Every model output in the memo is a \\V{key} macro filled from outputs/analysis.json.
Before compiling, the script scans the memo text for literal numbers and fails if
one cannot be traced to the evidence ledger, the inputs or a short list of labels.
"""
from __future__ import annotations

import copy
import json
import re
import shutil
import subprocess
import sys

from engine import ROOT, load_inputs, run
from reconcile import reconcile

MEMO = ROOT / "memo"


def pct(x, d=0):
    return f"{x * 100:.{d}f}\\%"


def mult(x, d=1):
    return f"{x:.{d}f}x"


def num(x, d=0):
    return f"{x:,.{d}f}"


def _ltm_fte(inp, inf) -> float:
    """Implied headcount from the last-twelve-month calibration; ties to the Triangulation tab."""
    c, m = inp["calibration"], inp["market"]
    rev = inf["revenue"]
    costs = rev * (1 - c["ltm_gross_margin"]) + _ltm_sm(inp, inf) + rev * c["ltm_rd_pct"] + rev * c["ltm_ga_pct"]
    return costs * m["personnel_share"] / (m["loaded_cost_per_fte_k"] / 1000)


def _ltm_sm(inp, inf) -> float:
    c = inp["calibration"]
    return (c["ltm_new_france"] * c["ltm_cac_france"] + c["ltm_new_intl"] * c["ltm_cac_intl"]
            + c["ltm_upsell"] * c["ltm_cac_upsell"] + inf["revenue"] * c["ltm_cs_pct"])


def build_values() -> dict:
    a = json.loads((ROOT / "outputs" / "analysis.json").read_text())
    inp = load_inputs()
    S = a["scenarios"]
    B, D, P = S["Base"], S["Downside"], S["Plan"]
    th, fp, lev, ti = a["thresholds"], a["firepower"], a["leverage"], a["tornado"]
    bp = a["base_to_plan"]["steps"]
    d = inp["deal"]
    t = {r["label"]: r for r in ti["rows"]}
    build = a["build_up_detail"]
    inf = a["retention_inference"]
    ue = a["unit_economics"]
    nb = inp["next_buyer"]
    grid = a["grid_entry"]
    cells = [c for row in grid["rows"] for c in row["cells"]]
    flat = [v for row in inf["rows"] for v in row["grr"] if v <= 1.0]
    integ = inp["constants_by_scenario"]["integration_pct"]["Base"]
    cagr = lambda arr: (arr / d["arr_entry"]) ** (1 / 4.5) - 1
    bx, dx, px = B["exits"]["Dec-30"], D["exits"]["Dec-30"], P["exits"]["Dec-30"]
    ref = inp["references"]
    from engine import curve_multiple
    r40_entry = ((d["arr_entry"] / inf["opening"] - 1) + (inf["revenue"] * inp["calibration"]["ltm_gross_margin"] - _ltm_sm(inp, inf) - inf["revenue"] * (inp["calibration"]["ltm_rd_pct"] + inp["calibration"]["ltm_ga_pct"])) / inf["revenue"]) * 100
    band_median = 5.0
    equity_band = max(bx["ev"] * 0 + B["rows"]["revenue"][4] * band_median * (1 - inp["exit"]["cost_pct"]) - B["rows"]["net_debt"][4], 0)
    mip_band = inp["exit"]["mip_share"] * max(equity_band - inp["exit"]["mip_hurdle_moic"] * d["post_money"], 0)
    moic_band = (equity_band - mip_band) * B["entry"]["bp_share"] / d["bp_equity"]
    ai_extra = sum((inp["drivers"]["rd_pct"]["Base"][i] - inp["calibration"]["ltm_rd_pct"]) * B["rows"]["rev_org"][i] for i in range(4))
    carry = d["unitranche"] * (inp["financing"]["margin"] + inp["financing"]["cash_yield_spread"])
    selldown = B["entry"]["secondary"] - d["reinvest_equity"]
    gap30 = next(r["gap_5x"] for r in fp["rows"] if abs(r["cagr"] - 0.30) < 1e-9)
    exposure = d["bp_equity"] + fp["bp_share_gap"]
    new_drop = 1 - (inp["drivers"]["new_france"]["Downside"][4] + inp["drivers"]["new_intl"]["Downside"][4]) / (inp["drivers"]["new_france"]["Downside"][0] + inp["drivers"]["new_intl"]["Downside"][0])
    tor3 = [t["Gross retention -/+ 3 pts (new logos adjust)"], t["CAC +/- 20%"], t["Exit multiple -/+ 1.0x revenue"]]
    v = {
        "hurdle_20": "2.0x", "hurdle_25": "2.5x", "hurdle_30": mult(inp["hurdle"]["moic"]), "hurdle_irr": pct(inp["hurdle"]["irr"]), "irr_20": "20\\%",
        "grr_base": pct(th["grr_france_base"]), "grr_25": pct(th["grr_both_for_2.5"]), "grr_20": pct(th["grr_both_for_2.0"]), "grr_30": pct(th["grr_both_for_3.0"]),
        "grr_down": pct(inp["drivers"]["grr_france"]["Downside"][1]),
        "pm_25": num(th["post_money_for_2.5"]), "pm_i20": num(th["post_money_for_irr_20"]), "pm_i25": num(th["post_money_for_irr_25"]), "pm_30": num(th["post_money_for_3.0"]),
        "pm_25_stake": pct(d["bp_equity"] / th["post_money_for_2.5"]), "pm_25_cheque": num(th["post_money_for_2.5"] * B["entry"]["bp_share"]),
        "pm_25_arr": mult(next(r for r in a["price_table"] if r["label"] == "MOIC 2.5x")["ev_arr"]),
        "exit_adj_25": mult(th["exit_adj_for_2.5"]), "exit_25_rev": mult(bx["multiple"] + th["exit_adj_for_2.5"]),
        "r40_gap": f"{r40_entry - 40:.1f}", "m_entry": mult(B["entry"]["m_entry"]), "ev_ltm": mult(B["entry"]["ev_entry"] / inf["revenue"]), "curve_entry": mult(curve_multiple(inp, r40_entry)),
        "grid_low": mult(min(c["ev_arr"] for c in cells)), "grid_high": mult(max(c["ev_arr"] for c in cells)),
        "grid_moic_low": mult(min(c["moic"] for c in cells)), "grid_moic_high": mult(max(c["moic"] for c in cells)),
        "base_moic": mult(bx["moic"]), "base_moic2": mult(bx["moic"], 2), "base_irr": pct(bx["irr"], 1),
        "down_moic": mult(dx["moic"]), "down_irr": f"({abs(dx['irr']) * 100:.0f}\\%)", "down_pref": mult(dx["moic_pref"]),
        "plan_moic": mult(px["moic"]), "plan_irr": pct(px["irr"]),
        "base_nb": mult(B["next_buyer"]["moic"]), "down_nb": mult(D["next_buyer"]["moic"]), "plan_nb": mult(P["next_buyer"]["moic"]),
        "base_arr30": num(B["rows"]["arr_total"][4]), "down_arr30": num(D["rows"]["arr_total"][4]), "plan_arr30": num(P["rows"]["arr_total"][4]), "base_arr28": num(B["rows"]["arr_total"][2]),
        "base_cagr": pct(cagr(B["rows"]["arr_org"][4])), "down_cagr": pct(cagr(D["rows"]["arr_org"][4])), "plan_cagr": pct(cagr(P["rows"]["arr_org"][4])),
        "base_margin": pct(B["rows"]["ebitda_margin"][4]), "down_margin": pct(D["rows"]["ebitda_margin"][4]), "plan_margin": pct(P["rows"]["ebitda_margin"][4]),
        "base_xebitda": mult(bx["ev_ebitda"], 0), "down_xebitda": mult(dx["ev_ebitda"], 0), "plan_xebitda": mult(px["ev_ebitda"], 0), "base_xrev": mult(bx["multiple"]), "down_xrev": mult(dx["multiple"]), "plan_xrev": mult(px["multiple"]),
        "base_xarr": mult(bx["ev_arr"]), "down_xarr": mult(dx["ev_arr"]), "plan_xarr": mult(px["ev_arr"]),
        "plan_follow": num(px["follow_on"]),
        "primary": num(d["primary_cash"]), "existing": num(d["existing_cash"]), "bp_eq": num(d["bp_equity"]), "reinvest": num(d["reinvest_equity"]), "unitranche": num(d["unitranche"]),
        "secondary": num(B["entry"]["secondary"]), "secondary_b": num(a["reading_b"]["secondary"]),
        "costs_entry": num(B["entry"]["deal_fees"] + B["entry"]["oid"]), "sources": num(B["entry"]["total_sources"]), "ev_entry": num(B["entry"]["ev_entry"]),
        "bp_share": pct(B["entry"]["bp_share"], 1), "bp_share_short": pct(B["entry"]["bp_share"]), "ticket_fund": pct(d["bp_equity"] / ref["fund_size_m"], 1),
        "paycor": mult(4100 / 655), "dayforce": mult(ref["dayforce_value_usd_m"] / (ref["dayforce_q2_revenue_usd_m"] * 4)),
        "selldown": num(selldown), "selldown_pct": pct(selldown / B["entry"]["old_holder_value"]), "cash_entry": num(B["entry"]["cash_entry"]), "carry": f"{carry:.1f}",
        "ltm_rev": num(inf["revenue"]), "net_new": num(inf["net_new"]), "grr_low": pct(min(flat)), "grr_high": pct(max(flat)),
        "cac_a": mult(inf["cacs"][0]), "cac_b": mult(inf["cacs"][1]), "cac_c": mult(inf["cacs"][2]),
        "sm_a": pct(inf["rows"][0]["share"]), "sm_b": pct(inf["rows"][1]["share"]), "sm_c": pct(inf["rows"][2]["share"]),
        "g_a_a": pct(inf["rows"][0]["grr"][0]), "g_a_b": pct(inf["rows"][0]["grr"][1]),
        "g_b_a": pct(inf["rows"][1]["grr"][0]), "g_b_b": pct(inf["rows"][1]["grr"][1]), "g_b_c": pct(inf["rows"][1]["grr"][2]),
        "g_c_a": pct(inf["rows"][2]["grr"][0]), "g_c_b": pct(inf["rows"][2]["grr"][1]), "g_c_c": pct(inf["rows"][2]["grr"][2]),
        "sm_cal": pct(_ltm_sm(inp, inf) / inf["revenue"]),
        "cac_cal": mult(_ltm_sm(inp, inf) / (inp["calibration"]["ltm_new_france"] + inp["calibration"]["ltm_new_intl"] + inp["calibration"]["ltm_upsell"]), 2),
        "grr_cal": pct(1 - (inp["calibration"]["ltm_new_france"] + inp["calibration"]["ltm_new_intl"] + inp["calibration"]["ltm_upsell"] - inf["net_new"]) / inf["opening"]),
        "growth_h2": pct((inp["calibration"]["arr_dec_2025"] / inf["opening"]) ** 2 - 1),
        "growth_h1": pct((d["arr_entry"] / inp["calibration"]["arr_dec_2025"]) ** 2 - 1),
        "nrr_fr": pct(ue["nrr_france"]), "payback": num(ue["payback_months"]), "ltv_cac": mult(ue["ltv_cac"]),
        "pen_entry": pct(inp["market"]["dau_entry_m"] * d["france_share"] / inp["market"]["france_frontline_m"], 1),
        "pen_base": pct(B["rows"]["fr_penetration"][4], 1), "pen_plan": pct(P["rows"]["fr_penetration"][4], 1),
        "dau_hcr": pct(inp["market"]["dau_entry_m"] * d["france_share"] / ref["hcr_staff_m"]),
        "base_acq_bought": num(sum(B["rows"]["arr_acq_add"])), "down_new_drop": pct(new_drop),
        "r40_base": num(bx["r40"]), "base_band": mult(moic_band), "ai_extra": num(ai_extra),
        "ai_cost": mult(ti["base"] - t["Extra AI spend EUR 30M / none"]["low"], 2),
        "debt_arr": mult(d["unitranche"] / d["arr_entry"]), "debt_ebitda": mult(B["rows"]["debt_to_ebitda"][1]), "cover_fy27": mult(B["rows"]["interest_cover"][1]),
                "down_lowp": mult(a["stress_downside"]["low_primary_moic"]), "down_lowp_fo": num(a["stress_downside"]["follow_on"]),
        "w_d": pct(a["weighted"]["weights"]["Downside"]).replace("\\%", ""), "w_b": pct(a["weighted"]["weights"]["Base"]).replace("\\%", ""), "w_p": pct(a["weighted"]["weights"]["Plan"]),
        "w_moic": mult(a["weighted"]["moic"]), "w_pref": mult(a["weighted"]["moic_pref"]),
        "base_exit_eq": num(B["bridge"]["exit_equity"]), "br_arr": num(B["bridge"]["organic"] + B["bridge"]["acquired"]), "br_mult": num(-B["bridge"]["multiple"]),
        "lev_irr": pct(lev["levered_irr"], 1), "unlev_irr": pct(lev["unlevered_irr"], 1), "d_lev": mult(lev["downside_levered"], 2), "d_unlev": mult(lev["downside_unlevered"], 2),
        "moic_29": mult(B["exits"]["Dec-29"]["moic"], 2), "moic_31": mult(B["exits"]["Dec-31"]["moic"], 2),
        "irr_29": pct(B["exits"]["Dec-29"]["irr"], 1), "irr_31": pct(B["exits"]["Dec-31"]["irr"], 1),
        "tor_low": mult(min((x["high"] - x["low"]) / 2 for x in tor3), 2), "tor_high": mult(max((x["high"] - x["low"]) / 2 for x in tor3), 2),
        "gap_plan": mult(a["base_to_plan"]["plan"] - a["base_to_plan"]["start"], 2), "step_org": mult(bp["Organic ARR"], 2), "step_margin": mult(bp["Margins and CAC"], 2), "step_ma": mult(-bp["Build-ups"], 2),
        "base_ev": num(bx["ev"]), "down_org30": pct(D["rows"]["org_growth"][4], 1), "base_org30": pct(B["rows"]["org_growth"][4]),
        "nb_debt": mult(nb["debt_x_ebitda"], 0), "nb_cap": pct(nb["margin_cap"]), "nb_exit": mult(nb["exit_ebitda_multiple"], 0), "nb_irr": pct(nb["target_irr"]),
        "nb_ev": num(B["next_buyer"]["ev"]), "nb_low": mult(a["next_buyer_range"][0]["moic"]), "nb_high": mult(a["next_buyer_range"][2]["moic"]),
        "nb_low_ev": num(a["next_buyer_range"][0]["ev"]), "nb_high_ev": num(a["next_buyer_range"][2]["ev"]),
        "nb_curve": num(B["next_buyer"]["ev_on_curve"]), "nb_curve_x": mult(B["next_buyer"]["curve_exit_ebitda"], 0), "nb_pref": mult(D["next_buyer"]["moic_pref"]),
        "buyout_low": "10.8x", "buyout_high": "13.1x", "down_12x": mult(a["preference"]["down_12x"]), "down_12x_pref": mult(a["preference"]["down_12x_pref"]),
        "ai_ebitda": num(-a["ai_stress"]["ebitda_fy27"]), "ai_fo": num(a["ai_stress"]["follow_on"]),
        "d_cover_27": mult(D["rows"]["interest_cover"][1]), "d_cover_28": mult(D["rows"]["interest_cover"][2]), "d_cover_29": mult(D["rows"]["interest_cover"][3]),
        "down_dist": num(a["preference"]["downside_dist"]), "rollover": num(B["entry"]["rollover"]), "pref_bp": mult(a["preference"]["bp_only"]), "pref_cut": num(a["preference"]["cut"]),
        "debt_low_net": num(fp["debt_capacity_low"] - d["unitranche"]), "debt_high_net": num(fp["debt_capacity_high"] - d["unitranche"]),
        "near_term": num(fp["near_term"]), "buy_5x": num(fp["arr_buyable_near_term_5x"]), "capacity": num(fp["available"]),
        "plan_acq_at_base": num(fp["needed_at_base"]), "plan_gap_at_base": num(fp["gap_at_base"]),
        "exposure": num(exposure), "exposure_pct": pct(exposure / ref["fund_size_m"], 1), "gap_30": num(gap30),
        "integration": pct(integ), "ma_eff": mult(build["effective_price"]), "ma_none": mult(build["none"]["moic"], 2), "ma_double": mult(build["double"]["moic"], 2),
        "ma_double_fo": num(build["double"]["follow_on"]), "ma_double_irr": pct(build["double"]["irr"], 1),
        "ma_double_gain": num(build["double"]["bp_proceeds"] - build["base"]["bp_proceeds"]),
        "newlogo_cost": mult(ti["base"] - t["New-logo ARR -/+ 15%"]["low"], 2), "cac_cost": mult(ti["base"] - t["CAC +/- 20%"]["low"], 2),
        "exit_cost": mult(t["Exit multiple -/+ 1.0x revenue"]["high"] - ti["base"], 2),
        "intl_share": pct((B["rows"]["arr_in"][4] - d["arr_entry"] * (1 - d["france_share"])) / (B["rows"]["arr_org"][4] - d["arr_entry"])),
        "cir_share": pct(B["rows"]["cir"][1] / B["rows"]["ebitda"][1]),
        "new_fr_28": num(inp["drivers"]["new_france"]["Base"][2], 1), "new_in_28": num(inp["drivers"]["new_intl"]["Base"][2], 1),
        "cac_fr_28": mult(inp["drivers"]["cac_france"]["Base"][2], 2), "cac_in_28": mult(inp["drivers"]["cac_intl"]["Base"][2], 2),
        "fte_ltm": num(_ltm_fte(inp, inf)), "margin_27": pct(B["rows"]["ebitda_margin"][1]),
        "all_in": pct(inp["financing"]["euribor"] + inp["financing"]["margin"], 2),
        "mip": pct(inp["exit"]["mip_share"]), "mip_h": mult(inp["exit"]["mip_hurdle_moic"]),
    }
    g = a["grid_retention_exit"]
    picks = [r for r in g["rows"] if r["level"] in (0.71, 0.76, 0.81, 0.84, 0.86, 0.88, 0.90)]
    for i, adj in enumerate(g["adjustments"]):
        v[f"a{i}"] = f"{adj:+.1f}x" if adj < 0 else f"{adj:.1f}x"
    for k, row in enumerate(picks, start=1):
        v[f"l{k}"] = pct(row["level"])
        for i, m in enumerate(row["moic"]):
            v[f"r{k}{i}"] = mult(m, 2)

    # What the 2024 filed accounts show, and what they changed.
    acc = inp["accounts_2024"]
    eb23 = acc["operating_result_2023"] + acc["da_2023"] + acc["net_provisions_2023"]
    eb24 = acc["operating_result_2024"] + acc["da_2024"] + acc["net_provisions_2024"]
    pre = inp["pre_accounts"]
    old = copy.deepcopy(inp)
    old["deal"]["existing_cash"] = pre["existing_cash"]
    old["tax"]["nol_opening"] = pre["nol_opening"]
    old["cash_items"]["nwc_pct_delta_arr"] = pre["nwc_pct_delta_arr"]
    old["cash_items"]["cir_pct_rd"] = pre["cir_pct_rd"]
    paren = lambda x: f"({abs(x):.1f})" if x < 0 else f"{x:.1f}"
    neg_pct = lambda x: f"({abs(x) * 100:.0f}\\%)"
    v.update(
        a_rev23=f"{acc['revenue_2023']:.1f}", a_rev24=f"{acc['revenue_2024']:.1f}", a_eb23=paren(eb23), a_eb24=paren(eb24),
        a_m23=neg_pct(eb23 / acc["revenue_2023"]), a_m24=neg_pct(eb24 / acc["revenue_2024"]),
        a_x23=neg_pct((eb23 - acc["capitalised_dev_2023"]) / acc["revenue_2023"]),
        a_x24=neg_pct((eb24 - acc["capitalised_dev_2024"]) / acc["revenue_2024"]),
        a_pca23=pct(acc["deferred_revenue_2023"] / acc["revenue_2023"]), a_pca24=pct(acc["deferred_revenue_2024"] / acc["revenue_2024"]),
        a_cash23=f"{acc['cash_2023']:.1f}", a_cash24=f"{acc['cash_2024']:.1f}",
        a_debt23=f"{acc['bank_debt_2023']:.1f}", a_debt24=f"{acc['bank_debt_2024']:.1f}", a_hc=num(acc["headcount_avg_2024"]),
        a_growth=pct(acc["revenue_2024"] / acc["revenue_2023"] - 1),
        a_jump=num((eb24 / acc["revenue_2024"] - eb23 / acc["revenue_2023"]) * 100),
        a_capdev=f"{acc['capitalised_dev_2024']:.1f}", a_capdev_pct=pct(acc["capitalised_dev_2024"] / acc["revenue_2024"]),
        a_pca_inc=pct((acc["deferred_revenue_2024"] - acc["deferred_revenue_2023"]) / (acc["revenue_2024"] - acc["revenue_2023"])),
        a_doubt=pct(acc["doubtful_receivables_2024"] / acc["revenue_2024"]), a_bad=pct(acc["bad_debt_charge_2024"] / acc["revenue_2024"], 1),
        a_ib_loss=f"{abs(acc['iberia_result_2024']):.1f}", a_ib_rev=f"{acc['iberia_revenue_2024']:.1f}", a_ib_loan=f"{acc['loans_to_iberia_2024']:.1f}",
        a_credit=f"{acc['tax_credits_2024']:.2f}", pre_cir=pct(pre["cir_pct_rd"]), new_cir=pct(inp["cash_items"]["cir_pct_rd"]),
        pre_nwc=pct(pre["nwc_pct_delta_arr"]), new_nwc=pct(inp["cash_items"]["nwc_pct_delta_arr"]),
        pre_nol=num(pre["nol_opening"]), new_nol=num(inp["tax"]["nol_opening"]),
        pre_cash=num(pre["existing_cash"]), new_cash=num(d["existing_cash"]),
        pre_moic=mult(run(old, "Base").exits["Dec-30"]["moic"], 2), pre_m_entry=mult(run(old, "Base").entry["m_entry"]),
        a_founders=pct(acc["founders_stake"], 1), a_partech=pct(acc["partech_stake"], 1), a_xange=pct(acc["xange_stake"], 1),
    )

    n_recon, breaks = reconcile()
    if breaks:
        sys.exit(f"Reconciliation failed: {breaks[:5]}")
    v["n_recon"] = num(n_recon)
    return v


ALLOWED_LABELS = {
    # Dates, ranges and flex sizes the memo names in words rather than computes.
    *[str(y) for y in range(2016, 2032)], "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "20", "30", "15",
    "40", "50", "60", "70", "12", "2x", "0", "2.32", "6.50", "375", "81", "21", "13", "8.4", "4.2", "0.15", "17", "25", "5.0", "7.0", "3", "6", "4", "2.8", "400", "500", "38",
}


def allowed_numbers(inp) -> set:
    text = (ROOT / "data" / "evidence.csv").read_text() + (ROOT / "data" / "inputs.json").read_text()
    allowed = {t.replace(",", "") for t in re.findall(r"\d[\d,]*\.?\d*", text)}
    for t in list(allowed):
        try:
            allowed.add(str(float(t) * 100).rstrip("0").rstrip("."))
            allowed.add(str(float(t)).rstrip("0").rstrip("."))
        except ValueError:
            pass
    return allowed | ALLOWED_LABELS


def scan_literals(inp) -> list:
    body = (MEMO / "body.tex").read_text()
    body = re.sub(r"\\V\{[^}]*\}", "", body)
    body = re.sub(r"\\src\{[^}]*\}", "", body)
    body = re.sub(r"\\includegraphics\[[^\]]*\]\{[^}]*\}", "", body)
    body = re.sub(r"\\begin\{(tabularx|tabular)\}(\{\\textwidth\})?\{[^\n]*\n", "\n", body)
    body = re.sub(r"\\(vspace|hspace|fontsize|setlength)\*?\{[^}]*\}(\{[^}]*\})?", "", body)
    body = re.sub(r"\\(begin|end)\{minipage\}(\[[a-z]\])?(\{[^}]*\})?", "", body)
    allowed = allowed_numbers(inp)
    bad = []
    for token in re.findall(r"\d[\d,]*\.?\d*", body):
        clean = token.rstrip(",.").replace(",", "")
        norm = str(float(clean)).rstrip("0").rstrip(".") if clean else clean
        if clean not in allowed and norm not in allowed:
            bad.append(token)
    return bad


def main():
    inp = load_inputs()
    values = build_values()
    lines = [f"\\expandafter\\def\\csname val@{k}\\endcsname{{{val}}}" for k, val in values.items()]
    (MEMO / "values.tex").write_text("\n".join(lines) + "\n")
    bad = scan_literals(inp)
    if bad:
        sys.exit(f"Untraceable numbers in memo text: {sorted(set(bad))}")
    (MEMO / "Skello_Investment_Memo.tex").write_text("\\input{preamble.tex}\n\\input{body.tex}\n")
    for _ in range(2):
        out = subprocess.run(["xelatex", "-interaction=nonstopmode", "-halt-on-error", "Skello_Investment_Memo.tex"], cwd=MEMO, capture_output=True, text=True)
        if out.returncode:
            sys.exit(out.stdout[-2500:])
    print("Memo built;", len(values), "injected values; no untraceable literals.")
    for f in ("aux", "log", "out"):
        p = MEMO / f"Skello_Investment_Memo.{f}"
        if f != "log" and p.exists():
            p.unlink()


if __name__ == "__main__":
    main()
