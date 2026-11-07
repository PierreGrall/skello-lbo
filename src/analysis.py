"""Sensitivities, break-even solves and attribution, all from the Python engine.

Results go to outputs/analysis.json, which feeds the workbook's static grids, the
charts and the memo. Exit value is linked to the Rule of 40, so operating flexes
also move the exit multiple, unless a flex says otherwise.
"""
from __future__ import annotations

import itertools
import json
from datetime import date

from scipy.optimize import brentq

from engine import ROOT, flexed, load_inputs, run, shift_driver, xirr

OUT = ROOT / "outputs"
GRR_KEYS = ("grr_france", "grr_intl")
NEW_KEYS = ("new_france", "new_intl")
CAC_KEYS = ("cac_france", "cac_intl")


def shift_many(inp, scenario, keys, delta=0.0, factor=1.0):
    out = inp
    for key in keys:
        out = shift_driver(out, scenario, key, delta=delta, factor=factor)
    return out


def set_exit_adj(inp, scenario, value):
    return flexed(inp, lambda x: x["constants_by_scenario"]["exit_adj"].__setitem__(scenario, value))


def set_deal(inp, **kw):
    return flexed(inp, lambda x: x["deal"].update(kw))


def moic(inp, s="Base", label="Dec-30"):
    return run(inp, s).exits[label]["moic"]


def irr(inp, s="Base", label="Dec-30"):
    return run(inp, s).exits[label]["irr"]


def solve(fn, target, low, high):
    """Root of fn(x) = target on [low, high]; None when the target is out of range."""
    f_low, f_high = fn(low) - target, fn(high) - target
    if f_low * f_high > 0:
        return None
    return brentq(lambda x: fn(x) - target, low, high, xtol=1e-6)


def retention_flex(inp, scenario, delta):
    """Shifts retention in both geographies and cuts new logos to stay consistent with public LTM growth.

    Higher retention means less churn was replaced last year, so less gross new ARR was
    sold. Holding new logos at Base while raising retention would count the same growth twice.
    """
    c = inp["calibration"]
    opening = inp["deal"]["arr_entry"] / (1 + c["growth_ltm"])
    factor = 1 - delta * opening / (c["ltm_new_france"] + c["ltm_new_intl"])
    return shift_many(shift_many(inp, scenario, GRR_KEYS, delta=delta), scenario, NEW_KEYS, factor=factor)


def add_rd(inp, scenario, per_year, periods=(1, 2)):
    """Adds an absolute amount of R&D, expressed through the percentage driver."""
    res = run(inp, scenario)

    def apply(x):
        for t in periods:
            x["drivers"]["rd_pct"][scenario][t] += per_year / res.rows["rev_org"][t]
    return flexed(inp, apply)


def grid_exit_vs_new_logo(inp):
    adjustments = [-1.5, -1.0, -0.5, 0.0, 0.5, 1.0, 1.5]
    factors = [0.8, 0.9, 1.0, 1.1, 1.2]
    rows = []
    for f in factors:
        base = shift_many(inp, "Base", NEW_KEYS, factor=f)
        rows.append(dict(
            factor=f, arr_dec30=run(base, "Base").rows["arr_total"][4],
            moic=[moic(set_exit_adj(base, "Base", a)) for a in adjustments],
            irr=[irr(set_exit_adj(base, "Base", a)) for a in adjustments],
        ))
    return dict(adjustments=adjustments, rows=rows)


def grid_retention_vs_exit(inp):
    """Both-geography retention shift against the exit-multiple adjustment, with the French level shown."""
    grr0 = inp["drivers"]["grr_france"]["Base"][1]
    levels = [0.71, 0.76, 0.81, 0.84, 0.86, 0.88, 0.90, 0.92]
    adjustments = [-1.0, -0.5, 0.0, 0.5, 1.0]
    rows = []
    for level in levels:
        flex = retention_flex(inp, "Base", level - grr0)
        r = run(flex, "Base")
        rows.append(dict(level=level, arr_dec30=r.rows["arr_total"][4], multiple=r.exits["Dec-30"]["ev_arr"],
                         moic=[moic(set_exit_adj(flex, "Base", a)) for a in adjustments]))
    return dict(adjustments=adjustments, rows=rows)


def grid_entry(inp):
    """Reported stake of 40-50% on EUR 150M means EUR 300-375M post-money; primary cash is undisclosed."""
    post = [300.0, 325.0, 350.0, 375.0]
    primary = [20.0, 40.0, 60.0, 80.0, 100.0]
    rows = []
    for p in post:
        row = []
        for c in primary:
            r = run(set_deal(inp, post_money=p, primary_cash=c), "Base")
            row.append(dict(ev_arr=r.entry["m_entry"], moic=r.exits["Dec-30"]["moic"], irr=r.exits["Dec-30"]["irr"]))
        rows.append(dict(post=p, stake=inp["deal"]["bp_equity"] / p, cells=row))
    return dict(primary=primary, rows=rows)


def grid_exit_year(inp):
    adjustments = [-1.0, -0.5, 0.0, 0.5, 1.0]
    rows = []
    for label in ("Dec-29", "Dec-30", "Dec-31"):
        rows.append(dict(label=label, moic=[moic(set_exit_adj(inp, "Base", a), label=label) for a in adjustments],
                         irr=[irr(set_exit_adj(inp, "Base", a), label=label) for a in adjustments]))
    return dict(adjustments=adjustments, rows=rows)


def tornado(inp):
    none_ma = shift_driver(inp, "Base", "acquired_arr", factor=0.0)
    cases = [
        ("Exit multiple -/+ 1.0x revenue", set_exit_adj(inp, "Base", -1.0), set_exit_adj(inp, "Base", 1.0)),
        ("Gross retention -/+ 3 pts (new logos adjust)", retention_flex(inp, "Base", -0.03), retention_flex(inp, "Base", 0.03)),
        ("New-logo ARR -/+ 15%", shift_many(inp, "Base", NEW_KEYS, factor=0.85), shift_many(inp, "Base", NEW_KEYS, factor=1.15)),
        ("CAC +/- 20%", shift_many(inp, "Base", CAC_KEYS, factor=1.2), shift_many(inp, "Base", CAC_KEYS, factor=0.8)),
        ("Post-money EUR 375M / EUR 300M", set_deal(inp, post_money=375.0), set_deal(inp, post_money=300.0)),
        ("Primary cash EUR 20M / EUR 100M", set_deal(inp, primary_cash=20.0), set_deal(inp, primary_cash=100.0)),
        ("Price and upsell -/+ 2 pts", shift_many(inp, "Base", ("upsell_france", "upsell_intl"), delta=-0.02), shift_many(inp, "Base", ("upsell_france", "upsell_intl"), delta=0.02)),
        ("Extra AI spend EUR 30M / none", add_rd(inp, "Base", 15.0), inp),
        ("Unitranche EUR 0M / EUR 80M, same EV", set_deal(inp, unitranche=0.0, post_money=390.0), set_deal(inp, unitranche=80.0, post_money=310.0)),
        ("Build-up price +/- 1.0x ARR", shift_driver(inp, "Base", "ma_multiple", delta=1.0), shift_driver(inp, "Base", "ma_multiple", delta=-1.0)),
        ("Build-ups: none / double", none_ma, shift_driver(inp, "Base", "acquired_arr", factor=2.0)),
        ("Euribor +/- 100 bps", flexed(inp, lambda x: x["financing"].__setitem__("euribor", 0.0332)), flexed(inp, lambda x: x["financing"].__setitem__("euribor", 0.0132))),
    ]
    rows = [dict(label=label, low=moic(low), high=moic(high)) for label, low, high in cases]
    rows.sort(key=lambda r: abs(r["high"] - r["low"]), reverse=True)
    return dict(base=moic(inp), rows=rows)


def thresholds(inp):
    grr0 = inp["drivers"]["grr_france"]["Base"][1]
    out = dict(grr_france_base=grr0)
    for target in (1.5, 2.0, 2.5, 3.0):
        shift = solve(lambda d: moic(retention_flex(inp, "Base", d)), target, -0.25, 0.13)
        out[f"grr_both_for_{target}"] = None if shift is None else grr0 + shift
        out[f"new_logo_factor_for_{target}"] = solve(lambda f: moic(shift_many(inp, "Base", NEW_KEYS, factor=f)), target, 0.1, 3.0)
        out[f"exit_adj_for_{target}"] = solve(lambda a: moic(set_exit_adj(inp, "Base", a)), target, -4.0, 8.0)
        out[f"post_money_for_{target}"] = solve(lambda p: moic(set_deal(inp, post_money=p)), target, 150.5, 800.0)
    out["post_money_for_irr_25"] = solve(lambda p: irr(set_deal(inp, post_money=p)), 0.25, 150.5, 800.0)
    out["post_money_for_irr_20"] = solve(lambda p: irr(set_deal(inp, post_money=p)), 0.20, 150.5, 800.0)
    return out


def price_table(inp):
    """What each solved post-money means, keeping the EUR 150M cheque."""
    th = thresholds(inp)
    rows = []
    for label, p in (("Reported", inp["deal"]["post_money"]), ("IRR 20%", th["post_money_for_irr_20"]), ("MOIC 2.5x", th["post_money_for_2.5"]),
                     ("IRR 25%", th["post_money_for_irr_25"]), ("MOIC 3.0x", th["post_money_for_3.0"])):
        r = run(set_deal(inp, post_money=p), "Base")
        d = run(set_deal(inp, post_money=p), "Downside")
        rows.append(dict(label=label, post=p, ev=r.entry["ev_entry"], ev_arr=r.entry["m_entry"], stake=r.entry["bp_share"],
                         moic=r.exits["Dec-30"]["moic"], irr=r.exits["Dec-30"]["irr"], downside=d.exits["Dec-30"]["moic"]))
    return rows


def required_arr(inp):
    """Dec-30 ARR needed for a target MOIC at a given EV/ARR, at Base net debt and with no new equity."""
    d, ex = inp["deal"], inp["exit"]
    res = run(inp, "Base")
    nd = res.rows["net_debt"][4]
    share = res.entry["bp_share"]
    multiples = [3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0]
    table = {}
    for target in (2.0, 2.5, 3.0):
        dist = target * d["bp_equity"] / share
        hurdle = ex["mip_hurdle_moic"] * d["post_money"]
        equity = dist if dist <= hurdle else (dist - ex["mip_share"] * hurdle) / (1 - ex["mip_share"])
        table[str(target)] = [(equity + nd) / (m * (1 - ex["cost_pct"])) for m in multiples]
    return dict(multiples=multiples, targets=table, net_debt=nd)


def leverage_effect(inp):
    """Returns on 100% of the equity with and without the debt package, at the same enterprise value."""
    def total_equity(x, scenario="Base"):
        res = run(x, scenario)
        closing = date.fromisoformat(x["closing_date"])
        ends = [date.fromisoformat(e) for e in x["period_end"]]
        e = res.exits["Dec-30"]
        flows = [-x["deal"]["post_money"]] + [-res.rows["gap"][t] for t in range(5)]
        flows[-1] += e["equity"]
        return e["equity"] / (x["deal"]["post_money"] + e["follow_on"]), xirr(flows, [closing] + ends[:5])

    unlev = set_deal(inp, post_money=inp["deal"]["post_money"] + inp["deal"]["unitranche"], unitranche=0.0, line_commitment=0.0)
    lev_m, lev_i = total_equity(inp)
    un_m, un_i = total_equity(unlev)
    d_lev, _ = total_equity(inp, "Downside")
    d_un, _ = total_equity(unlev, "Downside")
    return dict(levered_moic=lev_m, levered_irr=lev_i, unlevered_moic=un_m, unlevered_irr=un_i, irr_points=(lev_i - un_i) * 100,
                downside_levered=d_lev, downside_unlevered=d_un)


def base_to_plan(inp):
    """Attribution of the MOIC gap between Base and Plan, averaged over every order of the steps."""
    groups = {
        "Organic ARR": ["grr_france", "grr_intl", "upsell_france", "upsell_intl", "new_france", "new_intl"],
        "Build-ups": ["acquired_arr", "ma_multiple", "acquired_growth", "acquired_margin"],
        "Margins and CAC": ["cac_france", "cac_intl", "cac_upsell", "gross_margin", "cs_fixed_pct", "rd_pct", "ga_pct"],
    }
    totals = {g: 0.0 for g in groups}
    orders = list(itertools.permutations(groups))
    for order in orders:
        current = flexed(inp, lambda x: None)
        previous = moic(current)
        for g in order:
            for key in groups[g]:
                current["drivers"][key]["Base"] = list(inp["drivers"][key]["Plan"])
            now = moic(current)
            totals[g] += now - previous
            previous = now
    return dict(start=moic(inp), plan=moic(inp, "Plan"), steps={g: v / len(orders) for g, v in totals.items()})


def firepower(inp):
    """Build-up capacity against the EUR 200M ARR target. Live on the workbook's Analytics tab."""
    d = inp["deal"]
    r = run(inp, "Base")
    rows = r.rows
    fcf = sum(rows["pre_fin_cf"][t] + rows["ma_spend"][t] + rows["integration"][t] for t in range(5))
    cash_entry = r.entry["cash_entry"]
    near_term = cash_entry - d["min_cash"] + d["line_commitment"]
    available = near_term + fcf
    integration = inp["constants_by_scenario"]["integration_pct"]["Base"]
    organic = rows["arr_org"][4]
    out = dict(cum_fcf_pre_ma=fcf, near_term=near_term, available=available, organic_dec30=organic,
               base_organic_cagr=(organic / d["arr_entry"]) ** (1 / 4.5) - 1,
               arr_buyable_near_term_5x=near_term / (5.0 * (1 + integration)),
               debt_capacity_low=rows["ebitda"][4] * 4.0, debt_capacity_high=rows["ebitda"][4] * 5.0,
               needed_at_base=max(200.0 - organic, 0.0))
    out["cost_at_base"] = out["needed_at_base"] * 5.0 * (1 + integration)
    out["gap_at_base"] = max(out["cost_at_base"] - available, 0.0)
    out["bp_share_gap"] = r.entry["bp_share"] * out["gap_at_base"]
    rows_out = []
    for cagr in (0.20, 0.25, 0.30, 0.35):
        o = d["arr_entry"] * (1 + cagr) ** 4.5
        needed = max(200.0 - o, 0.0)
        c5, c6 = needed * 5.0 * (1 + integration), needed * 6.0 * (1 + integration)
        rows_out.append(dict(cagr=cagr, organic=o, acquired=needed, cost_5x=c5, cost_6x=c6, gap_5x=max(c5 - available, 0.0), gap_6x=max(c6 - available, 0.0)))
    out["rows"] = rows_out
    return out


def build_up_detail(inp):
    out = {}
    for label, factor in (("none", 0.0), ("base", 1.0), ("double", 2.0)):
        r = run(shift_driver(inp, "Base", "acquired_arr", factor=factor), "Base")
        e = r.exits["Dec-30"]
        out[label] = dict(moic=e["moic"], irr=e["irr"], arr=r.rows["arr_total"][4], follow_on=e["follow_on"], bp_proceeds=e["bp_proceeds"], ev_arr=e["ev_arr"])
    out["effective_price"] = inp["drivers"]["ma_multiple"]["Base"][2] * (1 + inp["constants_by_scenario"]["integration_pct"]["Base"])
    return out


def retention_inference(inp):
    """Gross retention implied by LTM growth, restricted to S&M levels that leave the business profitable."""
    c, cash = inp["calibration"], inp["cash_items"]
    closing = inp["deal"]["arr_entry"]
    opening = closing / (1 + c["growth_ltm"])
    revenue = (opening + closing) / 2 * (1 + cash["services_pct"])
    net_new = closing - opening
    max_sm = c["ltm_gross_margin"] - c["ltm_rd_pct"] - c["ltm_ga_pct"]
    shares, cacs = (0.35, 0.40, max_sm), (0.8, 1.0, 1.2)
    rows = []
    for share in shares:
        rows.append(dict(share=share, grr=[1 - (share * revenue / cac - net_new) / opening for cac in cacs]))
    return dict(cacs=list(cacs), rows=rows, revenue=revenue, net_new=net_new, opening=opening, max_sm=max_sm)


def unit_economics(inp):
    c = inp["calibration"]
    d = inp["drivers"]
    grr = d["grr_france"]["Base"][1]
    gm = c["ltm_gross_margin"]
    cac = c["ltm_cac_france"]
    arr = inp["deal"]["arr_entry"]
    ref = inp["references"]
    return dict(
        nrr_france=grr + d["upsell_france"]["Base"][1], nrr_intl=d["grr_intl"]["Base"][1] + d["upsell_intl"]["Base"][1],
        payback_months=cac / gm * 12, ltv_cac=(gm / (1 - grr)) / cac,
        arr_per_customer=arr * 1e6 / ref["customers"], arr_per_customer_month=arr * 1e6 / ref["customers"] / 12,
        dau_per_customer=inp["market"]["dau_entry_m"] * 1e6 / ref["customers"], yield_month=arr / inp["market"]["dau_entry_m"] / 12,
    )


def stress_downside(inp):
    r = run(set_deal(inp, primary_cash=20.0), "Downside")
    return dict(low_primary_moic=r.exits["Dec-30"]["moic"], low_primary_pref=r.exits["Dec-30"]["moic_pref"], follow_on=r.exits["Dec-30"]["follow_on"])


def weighted(inp):
    w = inp["weights"]
    res = {s: run(inp, s) for s in w}
    return dict(weights=w, moic=sum(w[s] * res[s].exits["Dec-30"]["moic"] for s in w),
                moic_pref=sum(w[s] * res[s].exits["Dec-30"]["moic_pref"] for s in w),
                moic_next_buyer=sum(w[s] * res[s].next_buyer["moic"] for s in w))


def preference_value(inp):
    """The post-money cut that is worth as much as the preference, at the illustrative weights."""
    w = inp["weights"]
    with_pref = weighted(inp)["moic_pref"]

    def wm(p):
        x = set_deal(inp, post_money=p)
        return sum(w[s] * moic(x, s) for s in w)
    p = solve(wm, with_pref, 200.0, 350.0)
    down = run(inp, "Downside")
    e = down.exits["Dec-30"]
    # Downside priced on EBITDA instead of the revenue floor.
    from engine import distribution
    ebitda_ev = down.rows["ebitda"][4] * 12.0
    w12 = distribution(inp, down.entry, ebitda_ev, down.rows["net_debt"][4], e["follow_on"])
    rollover_paid = max(e["distributable"] - (inp["deal"]["bp_equity"] + inp["deal"]["reinvest_equity"]), 0.0)
    return dict(equivalent_post_money=p, cut=inp["deal"]["post_money"] - p, downside_dist=e["distributable"], rollover_paid=rollover_paid,
                down_12x=w12["moic"], down_12x_pref=w12["moic_pref"], bp_only=e["moic_pref_bp_only"],
                nb_pref=down.next_buyer["moic_pref"])


def ai_stress(inp):
    x = add_rd(inp, "Base", 15.0)
    r = run(x, "Base")
    return dict(ebitda_fy27=r.rows["ebitda"][1], follow_on=r.exits["Dec-30"]["follow_on"], moic=r.exits["Dec-30"]["moic"])


def next_buyer_range(inp):
    rows = []
    for mult in (13.0, 15.0, 17.0):
        x = flexed(inp, lambda y: y["next_buyer"].__setitem__("exit_ebitda_multiple", mult))
        nb = run(x, "Base").next_buyer
        rows.append(dict(multiple=mult, ev=nb["ev"], ev_arr=nb["ev_arr"], moic=nb["moic"], ev_ebitda=nb["ev_ebitda"]))
    return rows


def main():
    inp = load_inputs()
    results = {s: run(inp, s) for s in inp["scenarios"]}
    out = dict(
        scenarios={s: dict(rows=r.rows, entry=r.entry, exits=r.exits, bridge=r.bridge, next_buyer=r.next_buyer) for s, r in results.items()},
        grid_exit_new_logo=grid_exit_vs_new_logo(inp),
        grid_retention_exit=grid_retention_vs_exit(inp),
        grid_entry=grid_entry(inp),
        grid_exit_year=grid_exit_year(inp),
        tornado=tornado(inp),
        thresholds=thresholds(inp),
        price_table=price_table(inp),
        required_arr=required_arr(inp),
        leverage=leverage_effect(inp),
        base_to_plan=base_to_plan(inp),
        firepower=firepower(inp),
        build_up_detail=build_up_detail(inp),
        retention_inference=retention_inference(inp),
        unit_economics=unit_economics(inp),
        stress_downside=stress_downside(inp),
        weighted=weighted(inp),
        next_buyer_range=next_buyer_range(inp),
        preference=preference_value(inp),
        ai_stress=ai_stress(inp),
        reading_b=dict(secondary=run(set_deal(inp, reinvest_equity=10.0), "Base").entry["secondary"],
                       moic=run(set_deal(inp, reinvest_equity=10.0), "Base").exits["Dec-30"]["moic"]),
    )
    OUT.mkdir(exist_ok=True)
    (OUT / "analysis.json").write_text(json.dumps(out, indent=1))
    return out


if __name__ == "__main__":
    o = main()
    round2 = lambda v: round(v, 3) if isinstance(v, float) else v
    print("THRESHOLDS", {k: round2(v) for k, v in o["thresholds"].items()})
    print("PRICE", [{k: round2(v) for k, v in r.items()} for r in o["price_table"]])
    print("LEV", {k: round2(v) for k, v in o["leverage"].items()})
    print("B2P", o["base_to_plan"])
    print("FIRE", {k: round2(v) for k, v in o["firepower"].items() if k != "rows"})
    for r in o["firepower"]["rows"]:
        print({k: round(v, 1) for k, v in r.items()})
    print("BUILD", o["build_up_detail"])
    print("INFER", o["retention_inference"])
    print("UNIT", {k: round2(v) for k, v in o["unit_economics"].items()})
    print("STRESS", o["stress_downside"], "WEIGHTED", o["weighted"])
    print("NB", o["next_buyer_range"], o["scenarios"]["Base"]["next_buyer"])
    print("PREF", o["preference"], "AI", o["ai_stress"])
    print("TORNADO", round(o["tornado"]["base"], 3))
    for r in o["tornado"]["rows"]:
        print(f"  {r['label']:38} {r['low']:.2f} {r['high']:.2f}")
    print("RET x EXIT", o["grid_retention_exit"]["adjustments"])
    for r in o["grid_retention_exit"]["rows"]:
        print(r["level"], round(r["arr_dec30"]), round(r["multiple"], 2), [round(x, 2) for x in r["moic"]])
    print("ENTRY", o["grid_entry"]["primary"])
    for r in o["grid_entry"]["rows"]:
        print(r["post"], round(r["stake"], 3), [(round(c["ev_arr"], 1), round(c["moic"], 2)) for c in r["cells"]])
