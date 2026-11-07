"""Stand-alone Python implementation of the Skello LBO model.

A second implementation of the workbook's logic: it catches formula slips in
either version, though not an error of reasoning that both share. All amounts
are EUR millions. Interest is charged on opening balances, so there is no
circular reference. Exit value is LTM revenue times a multiple read from the
Rule of 40 curve, so margin, CAC and retention all move the exit price.
"""
from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from scipy.optimize import brentq

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUTS = ROOT / "data" / "inputs.json"
EXIT_INDEX = {"Dec-29": 3, "Dec-30": 4, "Dec-31": 5}


def load_inputs(path: Path = DEFAULT_INPUTS) -> dict:
    return json.loads(Path(path).read_text())


def xirr(flows: list[float], dates: list[date]) -> float:
    """Annualised return on actual/365 day count, matching Excel's XIRR."""
    if max(flows) <= 0:
        return -1.0  # Total loss: no positive cash flow to solve against.
    t0 = dates[0]

    def npv(rate: float) -> float:
        return sum(f / (1 + rate) ** ((d - t0).days / 365) for f, d in zip(flows, dates))

    return brentq(npv, -0.99, 10.0, xtol=1e-10)


def curve_multiple(inp: dict, r40: float) -> float:
    """EV / Revenue from the Rule of 40 curve; flat beyond the end points."""
    xs, ys = inp["r40_curve"]["x"], inp["r40_curve"]["y"]
    if r40 <= xs[0]:
        return ys[0]
    if r40 >= xs[-1]:
        return ys[-1]
    for i in range(len(xs) - 1):
        if xs[i] <= r40 < xs[i + 1]:
            return ys[i] + (r40 - xs[i]) * (ys[i + 1] - ys[i]) / (xs[i + 1] - xs[i])
    raise ValueError(r40)


@dataclass
class Result:
    rows: dict = field(default_factory=dict)
    entry: dict = field(default_factory=dict)
    exits: dict = field(default_factory=dict)
    bridge: dict = field(default_factory=dict)
    next_buyer: dict = field(default_factory=dict)


def distribution(inp: dict, res_entry: dict, ev: float, net_debt: float, follow_on: float) -> dict:
    """Exit waterfall for a given enterprise value."""
    d, ex = inp["deal"], inp["exit"]
    costs = ev * ex["cost_pct"]
    equity = max(ev - net_debt - costs, 0.0)  # Limited liability.
    invested_all = d["post_money"] + follow_on
    mip = ex["mip_share"] * max(equity - ex["mip_hurdle_moic"] * invested_all, 0)
    dist = equity - mip
    share = res_entry["bp_share"]
    bp = dist * share
    bp_invested = d["bp_equity"] + share * follow_on
    pref_amount = d["bp_equity"] + d["reinvest_equity"]
    bp_pref = max(bp, min(dist, pref_amount) * d["bp_equity"] / pref_amount)
    bp_pref_only = max(bp, min(dist, d["bp_equity"]))
    return dict(costs=costs, equity=equity, mip=mip, distributable=dist, bp_proceeds=bp, bp_invested=bp_invested,
                moic=bp / bp_invested, bp_pref=bp_pref, moic_pref=bp_pref / bp_invested, moic_pref_bp_only=bp_pref_only / bp_invested)


def next_buyer_value(inp: dict, revenue: float, margin: float, growth: float) -> dict:
    """Most a financial buyer could pay at exit and still earn its target IRR."""
    nb = inp["next_buyer"]
    ebitda0 = revenue * margin
    debt = nb["debt_x_ebitda"] * max(ebitda0, 0)
    rev, m, g = revenue, margin, growth
    floor = min(nb["growth_floor"], growth)  # A slow grower does not recover to the floor.
    cum_fcf = 0.0
    ebitda = ebitda0
    for _ in range(nb["years"]):
        g = max(g - nb["growth_fade"], floor)
        rev *= 1 + g
        m = min(m + nb["margin_step"], max(nb["margin_cap"], margin))
        ebitda = rev * m
        cum_fcf += ebitda * nb["fcf_conversion"] - debt * nb["interest_rate"]
    exit_ev = ebitda * nb["exit_ebitda_multiple"]
    exit_equity = exit_ev - debt + cum_fcf
    equity_in = exit_equity / (1 + nb["target_irr"]) ** nb["years"]
    # The same buyer if its own exit were priced on the Rule of 40 curve instead of an EBITDA multiple.
    curve_ev = rev * curve_multiple(inp, (g + m) * 100)
    curve_in = (curve_ev - debt + cum_fcf) / (1 + nb["target_irr"]) ** nb["years"]
    return dict(ev=equity_in + debt, debt=debt, exit_ev=exit_ev, cum_fcf=cum_fcf, exit_equity=exit_equity,
                equity_in=equity_in, final_revenue=rev, final_ebitda=ebitda, ev_on_curve=curve_in + debt,
                curve_exit_ebitda=curve_ev / ebitda if ebitda > 0 else 0.0)


def run(inp: dict, scenario: str) -> Result:
    d, fin, tax = inp["deal"], inp["financing"], inp["tax"]
    cash_items, mkt = inp["cash_items"], inp["market"]
    drv = {k: v[scenario] for k, v in inp["drivers"].items()}
    lengths = inp["period_length"]
    n = len(lengths)
    closing = date.fromisoformat(inp["closing_date"])
    ends = [date.fromisoformat(x) for x in inp["period_end"]]

    deal_fees = d["deal_fee_pct"] * (d["post_money"] + d["unitranche"])
    oid = d["oid_pct"] * d["unitranche"]
    total_sources = d["bp_equity"] + d["reinvest_equity"] + d["unitranche"]
    secondary = total_sources - d["primary_cash"] - deal_fees - oid
    cash_entry = d["existing_cash"] + d["primary_cash"]
    ev_entry = d["post_money"] + d["unitranche"] - cash_entry
    rollover = d["post_money"] - d["bp_equity"] - d["reinvest_equity"]
    bp_share = d["bp_equity"] / d["post_money"]
    m_entry = ev_entry / d["arr_entry"]

    r = {k: [0.0] * n for k in (
        "arr_fr_open arr_fr_churn arr_fr_upsell arr_fr_new arr_fr "
        "arr_in_open arr_in_churn arr_in_upsell arr_in_new arr_in "
        "arr_acq_open arr_acq_add arr_acq_growth arr_acq arr_org arr_total arr_growth org_growth "
        "rev_org rev_acq revenue cogs sm_var sm_cs rd ga ebitda_org ebitda_acq ebitda ebitda_margin r40 "
        "capex nwc cir interest commitment interest_income tax ma_spend integration pre_fin_cf "
        "need draw gap repay line cash unitranche net_debt "
        "fin_charges deductible taxable nol_open nol_use nol "
        "yield_per_dau fr_dau fr_penetration interest_cover debt_to_arr debt_to_ebitda fte_implied".split()
    )}

    arr_fr_prev = d["arr_entry"] * d["france_share"]
    arr_in_prev = d["arr_entry"] * (1 - d["france_share"])
    arr_acq_prev = 0.0
    arr_total_prev = d["arr_entry"]
    cash_prev, line_prev, nol_prev = cash_entry, 0.0, tax["nol_opening"]
    elapsed = 0.0
    yield_entry = d["arr_entry"] / mkt["dau_entry_m"]
    dec25 = inp["calibration"]["arr_dec_2025"]

    for t in range(n):
        L = lengths[t]
        elapsed += L
        fr_churn = arr_fr_prev * (1 - drv["grr_france"][t]) * L
        fr_upsell = arr_fr_prev * drv["upsell_france"][t] * L
        fr_new = drv["new_france"][t] * L
        arr_fr = arr_fr_prev - fr_churn + fr_upsell + fr_new

        in_churn = arr_in_prev * (1 - drv["grr_intl"][t]) * L
        in_upsell = arr_in_prev * drv["upsell_intl"][t] * L
        in_new = drv["new_intl"][t] * L
        arr_in = arr_in_prev - in_churn + in_upsell + in_new

        acq_add = drv["acquired_arr"][t]
        acq_growth = arr_acq_prev * drv["acquired_growth"][t] * L
        arr_acq = arr_acq_prev + acq_add + acq_growth

        arr_org = arr_fr + arr_in
        arr_total = arr_org + arr_acq

        rev_org = (arr_fr_prev + arr_in_prev + arr_org) / 2 * L * (1 + cash_items["services_pct"])
        rev_acq = (arr_acq_prev + arr_acq) / 2 * L
        revenue = rev_org + rev_acq
        cogs = rev_org * (1 - drv["gross_margin"][t])
        sm_var = (fr_new * drv["cac_france"][t] + in_new * drv["cac_intl"][t]
                  + (fr_upsell + in_upsell) * drv["cac_upsell"][t])
        sm_cs = rev_org * drv["cs_fixed_pct"][t]
        rd = rev_org * drv["rd_pct"][t]
        ga = rev_org * drv["ga_pct"][t]
        ebitda_org = rev_org - cogs - sm_var - sm_cs - rd - ga
        ebitda_acq = rev_acq * drv["acquired_margin"][t]
        ebitda = ebitda_org + ebitda_acq

        all_in = fin["euribor"] + fin["margin"]
        interest = (d["unitranche"] + line_prev) * all_in * L
        commitment = (d["line_commitment"] - line_prev) * fin["margin"] * fin["commitment_fee_share"] * L
        interest_income = cash_prev * max(fin["euribor"] - fin["cash_yield_spread"], 0) * L

        capex = revenue * cash_items["capex_pct"]
        fin_charges = interest + commitment - interest_income
        cap = max(tax["interest_floor"] * L, tax["interest_cap_pct"] * max(ebitda, 0))
        deductible = min(fin_charges, cap) if fin_charges > 0 else fin_charges
        taxable = ebitda - capex - deductible
        if taxable > 0:
            allowance = tax["nol_free"] * L + tax["nol_excess_share"] * max(taxable - tax["nol_free"] * L, 0)
            nol_use = min(nol_prev, allowance, taxable)
            nol = nol_prev - nol_use
        else:
            nol_use = 0.0
            nol = nol_prev - taxable
        tax_paid = max(taxable - nol_use, 0) * tax["rate"]

        nwc = cash_items["nwc_pct_delta_arr"] * (arr_total - arr_total_prev - acq_add)
        cir = rd * cash_items["cir_pct_rd"]
        ma_spend = acq_add * drv["ma_multiple"][t]
        integration = ma_spend * inp["constants_by_scenario"]["integration_pct"][scenario]
        pre_fin = (ebitda - capex + nwc + cir - interest - commitment + interest_income
                   - tax_paid - ma_spend - integration)

        position = cash_prev + pre_fin
        need = max(d["min_cash"] - position, 0)
        draw = min(need, d["line_commitment"] - line_prev)
        gap = need - draw
        repay = min(line_prev, max(position - d["min_cash"], 0))
        line = line_prev + draw - repay
        cash = position + draw + gap - repay

        for key, value in dict(
            arr_fr_open=arr_fr_prev, arr_fr_churn=fr_churn, arr_fr_upsell=fr_upsell, arr_fr_new=fr_new, arr_fr=arr_fr,
            arr_in_open=arr_in_prev, arr_in_churn=in_churn, arr_in_upsell=in_upsell, arr_in_new=in_new, arr_in=arr_in,
            arr_acq_open=arr_acq_prev, arr_acq_add=acq_add, arr_acq_growth=acq_growth, arr_acq=arr_acq,
            arr_org=arr_org, arr_total=arr_total,
            rev_org=rev_org, rev_acq=rev_acq, revenue=revenue, cogs=cogs, sm_var=sm_var, sm_cs=sm_cs,
            rd=rd, ga=ga, ebitda_org=ebitda_org, ebitda_acq=ebitda_acq, ebitda=ebitda,
            ebitda_margin=ebitda / revenue, capex=capex, nwc=nwc, cir=cir, interest=interest, commitment=commitment,
            interest_income=interest_income, tax=tax_paid, ma_spend=ma_spend, integration=integration,
            pre_fin_cf=pre_fin, need=need, draw=draw, gap=gap, repay=repay, line=line, cash=cash,
            unitranche=d["unitranche"], net_debt=d["unitranche"] + line - cash,
            fin_charges=fin_charges, deductible=deductible, taxable=taxable,
            nol_open=nol_prev, nol_use=nol_use, nol=nol,
        ).items():
            r[key][t] = value

        year_ago = dec25 if t == 0 else r["arr_total"][t - 1]
        org_year_ago = dec25 if t == 0 else r["arr_org"][t - 1]
        r["arr_growth"][t] = arr_total / year_ago - 1
        r["org_growth"][t] = arr_org / org_year_ago - 1
        r["r40"][t] = (r["org_growth"][t] + r["ebitda_margin"][t]) * 100

        y = yield_entry * (1 + mkt["yield_growth"]) ** elapsed
        r["yield_per_dau"][t] = y
        r["fr_dau"][t] = arr_fr / y
        r["fr_penetration"][t] = r["fr_dau"][t] / mkt["france_frontline_m"]
        fixed = interest + commitment
        r["interest_cover"][t] = ebitda / fixed if fixed else 0.0
        gross = d["unitranche"] + line
        r["debt_to_arr"][t] = gross / arr_total
        r["debt_to_ebitda"][t] = gross / (ebitda / L) if ebitda > 0 else 0.0
        r["fte_implied"][t] = (revenue - ebitda) * mkt["personnel_share"] / L / (mkt["loaded_cost_per_fte_k"] / 1000)

        arr_fr_prev, arr_in_prev, arr_acq_prev = arr_fr, arr_in, arr_acq
        arr_total_prev = arr_total
        cash_prev, line_prev, nol_prev = cash, line, nol

    entry = dict(
        deal_fees=deal_fees, oid=oid, total_sources=total_sources, secondary=secondary, cash_entry=cash_entry,
        ev_entry=ev_entry, m_entry=m_entry, rollover=rollover, bp_share=bp_share,
        net_debt_entry=d["unitranche"] - cash_entry, old_holder_value=secondary + rollover,
    )
    adj = inp["constants_by_scenario"]["exit_adj"][scenario]

    exits = {}
    for label, idx in EXIT_INDEX.items():
        multiple = curve_multiple(inp, r["r40"][idx]) + adj
        ev = r["revenue"][idx] * multiple
        follow_on = sum(r["gap"][: idx + 1])
        w = distribution(inp, entry, ev, r["net_debt"][idx], follow_on)
        flows = [-d["bp_equity"]] + [-bp_share * r["gap"][t] for t in range(idx + 1)]
        flows[-1] += w["bp_proceeds"]
        exits[label] = dict(
            multiple=multiple, ev=ev, follow_on=follow_on, irr=xirr(flows, [closing] + ends[: idx + 1]),
            ev_arr=ev / r["arr_total"][idx], ev_ebitda=ev / r["ebitda"][idx] if r["ebitda"][idx] > 0 else 0.0,
            r40=r["r40"][idx], **w,
        )

    i = EXIT_INDEX["Dec-30"]
    x = exits["Dec-30"]
    bridge = dict(
        entry_equity=d["post_money"],
        organic=(r["arr_org"][i] - d["arr_entry"]) * m_entry,
        acquired=r["arr_acq"][i] * m_entry,
        multiple=(x["ev_arr"] - m_entry) * r["arr_total"][i],
        net_debt=-(r["net_debt"][i] - entry["net_debt_entry"]),
        costs=-x["costs"],
        exit_equity=x["equity"],
    )
    nb = next_buyer_value(inp, r["revenue"][i], r["ebitda_margin"][i], r["org_growth"][i])
    nb_w = distribution(inp, entry, nb["ev"], r["net_debt"][i], x["follow_on"])
    nb.update(moic=nb_w["moic"], moic_pref=nb_w["moic_pref"], ev_arr=nb["ev"] / r["arr_total"][i],
              ev_ebitda=nb["ev"] / r["ebitda"][i] if r["ebitda"][i] > 0 else 0.0, ev_revenue=nb["ev"] / r["revenue"][i])
    return Result(rows=r, entry=entry, exits=exits, bridge=bridge, next_buyer=nb)


def flexed(inp: dict, fn) -> dict:
    """Return a deep copy of the inputs with `fn` applied, leaving the original untouched."""
    out = copy.deepcopy(inp)
    fn(out)
    return out


def shift_driver(inp: dict, scenario: str, key: str, delta: float = 0.0, factor: float = 1.0) -> dict:
    def apply(x):
        x["drivers"][key][scenario] = [v * factor + delta for v in x["drivers"][key][scenario]]
    return flexed(inp, apply)


if __name__ == "__main__":
    data = load_inputs()
    for s in data["scenarios"]:
        res = run(data, s)
        print(f"\n=== {s} ===  EV {res.entry['ev_entry']:.1f}  EV/ARR {res.entry['m_entry']:.2f}x")
        for key in ["arr_org", "arr_acq", "arr_total", "org_growth", "revenue", "ebitda", "ebitda_margin",
                    "r40", "cash", "line", "gap", "fr_penetration", "interest_cover", "debt_to_ebitda", "fte_implied"]:
            vals = res.rows[key]
            fmt = "{:8.1%}" if key in ("org_growth", "ebitda_margin", "fr_penetration") else "{:8.2f}"
            print(f"{key:16}", " ".join(fmt.format(v) for v in vals))
        for label, e in res.exits.items():
            print(label, {k: round(e[k], 3) for k in ("multiple", "ev", "ev_arr", "ev_ebitda", "equity", "moic", "moic_pref", "irr", "follow_on")})
        print("next buyer", {k: round(v, 2) for k, v in res.next_buyer.items()})
