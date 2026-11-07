"""Builds model/Skello_LBO_Model.xlsx with live formulas from data/inputs.json.

Three model tabs (Base, Downside, Plan) are generated from one row specification,
so they are formula-identical; only the scenario index in cell D3 differs.
Sensitivity grids come from the Python engine and are labelled as static values;
everything else recalculates when an input changes.
"""
from __future__ import annotations

import csv
import json
import re
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from engine import ROOT, load_inputs

OUT_XLSX = ROOT / "model" / "Skello_LBO_Model.xlsx"
CELL_MAP = ROOT / "outputs" / "cell_map.json"
ANALYSIS = ROOT / "outputs" / "analysis.json"

FONT = "Arial"
NAVY = "1F2A44"
GREY = "F2F3F5"
MID = "D9DCE3"
BLUE_INPUT = "0000FF"
GREEN_LINK = "006100"
KEY_FILL = "FFF7CC"

NUM = '#,##0.0_);(#,##0.0);"-"_)'
NUM2 = '#,##0.00_);(#,##0.00);"-"_)'
INT = '#,##0_);(#,##0);"-"_)'
PCT = '0.0%_);(0.0%);"-"_)'
PCT0 = '0%_);(0%);"-"_)'
MULT = '0.00"x"_)'
MULT1 = '0.0"x"_)'
DATE = 'mmm-yy'
PERIOD_COLS = ["F", "G", "H", "I", "J", "K"]
ALL_COLS = ["E"] + PERIOD_COLS
THIN = Side(style="thin", color="7A7F8C")
HAIR = Side(style="hair", color="B5B9C4")


def font(bold=False, color="000000", italic=False, size=9):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def put(ws, ref, value, fmt=None, bold=False, color="000000", italic=False, fill=None, size=9, align=None, wrap=False):
    cell = ws[ref]
    cell.value = value
    cell.font = font(bold, color, italic, size)
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = PatternFill("solid", start_color=fill, end_color=fill)
    if align or wrap:
        cell.alignment = Alignment(horizontal=align, vertical="top" if wrap else None, wrap_text=wrap)
    return cell


def band(ws, row, text, first="B", last="L", fill=NAVY, color="FFFFFF"):
    a, b = ord(first) - 64, ord(last) - 64
    for col in range(a, b + 1):
        ws.cell(row=row, column=col).fill = PatternFill("solid", start_color=fill, end_color=fill)
    put(ws, f"{first}{row}", text, bold=True, color=color, fill=fill)


def finish(ws, widths, freeze=None, landscape=True, tab=None):
    ws.sheet_view.showGridLines = False
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    if freeze:
        ws.freeze_panes = freeze
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    if tab:
        ws.sheet_properties.tabColor = tab


# ---------------------------------------------------------------- Inputs
class InputsSheet:
    def __init__(self, wb, inp):
        self.ws = wb.create_sheet("Inputs")
        self.inp = inp
        self.scalar = {}
        self.driver = {}
        self.constant = {}
        self.row = 1

    def title(self):
        ws = self.ws
        put(ws, "B1", "Inputs: every assumption in the model", bold=True, size=12)
        put(ws, "B2", "EUR millions unless stated. Blue = hard-coded input. Shaded = assumption with no public evidence, and the first things diligence would replace. "
                      "Rates are annualised; the H2-26 stub takes half a year. References such as E11 point to the Evidence tab.", italic=True, color="555555")
        self.row = 4

    def section(self, text):
        band(self.ws, self.row, text)
        self.row += 1

    def add(self, key, label, value, unit, fmt, note, key_assumption=False):
        ws, r = self.ws, self.row
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", unit, color="555555")
        put(ws, f"D{r}", value, fmt=fmt, color=BLUE_INPUT, fill=KEY_FILL if key_assumption else None)
        put(ws, f"F{r}", note, color="555555", italic=True)
        self.scalar[key] = f"Inputs!$D${r}"
        self.row += 1

    def drivers(self, key, label, fmt, note, key_assumption=False):
        ws = self.ws
        r0 = self.row
        put(ws, f"B{r0}", label, bold=True)
        put(ws, f"L{r0}", note, color="555555", italic=True)
        for col, p in zip(PERIOD_COLS, self.inp["periods"]):
            put(ws, f"{col}{r0}", p, bold=True, align="right", color="555555")
        for i, s in enumerate(self.inp["scenarios"]):
            r = r0 + 1 + i
            put(ws, f"C{r}", s, color="555555")
            for col, v in zip(PERIOD_COLS, self.inp["drivers"][key][s]):
                put(ws, f"{col}{r}", v, fmt=fmt, color=BLUE_INPUT, fill=KEY_FILL if key_assumption else None)
        self.driver[key] = (r0 + 1, r0 + 3)
        self.row = r0 + 4

    def scenario_constant(self, key, label, fmt, note):
        ws, r0 = self.ws, self.row
        put(ws, f"B{r0}", label, bold=True)
        put(ws, f"F{r0}", note, color="555555", italic=True)
        for i, s in enumerate(self.inp["scenarios"]):
            r = r0 + 1 + i
            put(ws, f"C{r}", s, color="555555")
            put(ws, f"D{r}", self.inp["constants_by_scenario"][key][s], fmt=fmt, color=BLUE_INPUT)
        self.constant[key] = (r0 + 1, r0 + 3)
        self.row = r0 + 4


def build_inputs(wb, inp):
    sheet = InputsSheet(wb, inp)
    sheet.title()
    d, f, t, x, m, c, cal = inp["deal"], inp["financing"], inp["tax"], inp["exit"], inp["market"], inp["cash_items"], inp["calibration"]

    sheet.section("Timing")
    r = sheet.row
    put(sheet.ws, f"B{r}", "Closing date")
    put(sheet.ws, f"D{r}", date.fromisoformat(inp["closing_date"]), fmt="dd-mmm-yy", color=BLUE_INPUT)
    put(sheet.ws, f"F{r}", "Announced early July 2026; 30 June used as a clean period boundary.", italic=True, color="555555")
    sheet.scalar["closing_date"] = f"Inputs!$D${r}"
    sheet.row += 1
    r = sheet.row
    put(sheet.ws, f"B{r}", "Period end")
    put(sheet.ws, f"B{r+1}", "Period length (years)")
    for col, e, l in zip(PERIOD_COLS, inp["period_end"], inp["period_length"]):
        put(sheet.ws, f"{col}{r}", date.fromisoformat(e), fmt=DATE, color=BLUE_INPUT)
        put(sheet.ws, f"{col}{r+1}", l, fmt="0.0", color=BLUE_INPUT)
    sheet.period_end_row, sheet.length_row = r, r + 1
    sheet.row += 3

    sheet.section("Transaction (see Evidence E08-E13)")
    sheet.add("arr_entry", "ARR at closing", d["arr_entry"], "EUR M", NUM, "E01: reported as more than EUR 50M.")
    sheet.add("france_share", "France share of ARR", d["france_share"], "%", PCT, "E07: international is 20% of revenue, applied to ARR.")
    sheet.add("post_money", "Post-money equity value", d["post_money"], "EUR M", NUM, "E11: CFNEWS estimate, not a company figure.")
    sheet.add("bp_equity", "Bridgepoint equity", d["bp_equity"], "EUR M", NUM, "E08.")
    sheet.add("reinvest_equity", "New money from founders and VCs", d["reinvest_equity"], "EUR M", NUM, "E09. Set to 10 for the reading in E10; returns do not change.")
    sheet.add("unitranche", "Unitranche", d["unitranche"], "EUR M", NUM, "E12.")
    sheet.add("line_commitment", "Capex line commitment (undrawn at closing)", d["line_commitment"], "EUR M", NUM, "E12.")
    sheet.add("primary_cash", "Primary cash to the balance sheet", d["primary_cash"], "EUR M", NUM, "ASSUMPTION. The press gives no primary / secondary split; this one input moves entry EV/ARR between 5.5x and 7.1x.", True)
    sheet.add("existing_cash", "Cash on balance sheet before the deal", d["existing_cash"], "EUR M", NUM, "ASSUMPTION, net of legacy bank debt. End-2024: EUR 8.4M cash, EUR 4.2M bank debt (E52); near breakeven in 2025 (E03).", True)
    sheet.add("deal_fee_pct", "Transaction costs, % of post-money plus debt", d["deal_fee_pct"], "%", PCT, "ASSUMPTION. Advisers, diligence, legal.", True)
    sheet.add("oid_pct", "Original issue discount", d["oid_pct"], "%", PCT, "E22: issue price 97.5-98.5, midpoint 98.")
    sheet.add("min_cash", "Minimum operating cash", d["min_cash"], "EUR M", NUM, "ASSUMPTION. About two months of Dec-27 cash costs.", True)
    sheet.row += 1

    sheet.section("Financing")
    sheet.add("euribor", "3M Euribor (flat)", f["euribor"], "%", '0.00%', "E21.")
    sheet.add("margin", "Margin on unitranche and line", f["margin"], "%", '0.00%', "E22: top of the 5.00-6.00% range plus part of the software premium. Actual terms undisclosed.")
    sheet.add("commitment_fee_share", "Commitment fee, share of margin", f["commitment_fee_share"], "%", PCT0, "ASSUMPTION. Market convention, 35% of margin.", True)
    sheet.add("cash_yield_spread", "Deposit rate below Euribor", f["cash_yield_spread"], "%", '0.00%', "ASSUMPTION.", True)
    sheet.row += 1

    sheet.section("Tax (France)")
    sheet.add("tax_rate", "Corporate income tax, with social surcharge", t["rate"], "%", '0.00%', "French statutory rate: 25% plus 3.3% surcharge.")
    sheet.add("nol_opening", "Tax losses carried forward at closing", t["nol_opening"], "EUR M", NUM, "Accumulated accounting losses were EUR 31.8M at end-2024 (E56); tax losses assumed equal, little change in 2025.", True)
    sheet.add("nol_free", "Loss relief: fully usable tranche, per year", t["nol_free"], "EUR M", NUM, "French rule: EUR 1M, then 50% of the excess.")
    sheet.add("nol_excess_share", "Loss relief: share of the excess", t["nol_excess_share"], "%", PCT0, "French rule.")
    sheet.add("interest_floor", "Interest deductibility floor, per year", t["interest_floor"], "EUR M", NUM, "E34.")
    sheet.add("interest_cap_pct", "Interest deductibility cap, % of tax EBITDA", t["interest_cap_pct"], "%", PCT0, "E34. EBITDA is used as a proxy for tax EBITDA.")
    sheet.row += 1

    sheet.section("Exit")
    sheet.add("exit_cost_pct", "Exit costs, % of EV", x["cost_pct"], "%", PCT, "ASSUMPTION.", True)
    sheet.add("mip_share", "Management incentive, share of equity above hurdle", x["mip_share"], "%", PCT, "ASSUMPTION. Low because management already holds about 40% (E13).", True)
    sheet.add("mip_hurdle", "Management incentive hurdle, multiple of equity invested", x["mip_hurdle_moic"], "x", MULT, "ASSUMPTION.", True)
    sheet.add("hurdle_moic", "Reference gross MOIC", inp["hurdle"]["moic"], "x", MULT, "MY ASSUMPTION, read from E19. Not a Bridgepoint figure.", True)
    sheet.add("hurdle_irr", "Reference gross IRR", inp["hurdle"]["irr"], "%", PCT, "MY ASSUMPTION. E19 reports 30% on older, smaller funds.", True)
    sheet.row += 1

    nb = inp["next_buyer"]
    sheet.section("Next financial buyer at exit (NextBuyer tab)")
    sheet.add("nb_irr", "Buyer's target IRR", nb["target_irr"], "%", PCT, "ASSUMPTION.", True)
    sheet.add("nb_exit", "Buyer's exit EV / EBITDA", nb["exit_ebitda_multiple"], "x", MULT1, "ASSUMPTION. Inside the 12-18x range for mature businesses (E24); all-sector buyout median is 10.8-13.1x (E45).", True)
    sheet.add("nb_debt", "Buyer's opening debt / EBITDA", nb["debt_x_ebitda"], "x", MULT1, "ASSUMPTION.", True)
    sheet.add("nb_rate", "Buyer's all-in interest rate", nb["interest_rate"], "%", '0.00%', "Equals Euribor plus margin above.", True)
    sheet.add("nb_conv", "Cash conversion before interest, % of EBITDA", nb["fcf_conversion"], "%", PCT0, "ASSUMPTION. After capex and tax.", True)
    sheet.add("nb_fade", "Growth fade per year", nb["growth_fade"], "pts", PCT, "ASSUMPTION.", True)
    sheet.add("nb_floor", "Growth floor", nb["growth_floor"], "%", PCT, "ASSUMPTION.", True)
    sheet.add("nb_step", "Margin gain per year", nb["margin_step"], "pts", PCT, "ASSUMPTION.", True)
    sheet.add("nb_cap", "Margin cap", nb["margin_cap"], "%", PCT, "ASSUMPTION. Mature SMB SaaS.", True)
    sheet.row += 1

    sheet.section("Cash items and market")
    sheet.add("services_pct", "Non-recurring revenue, % of recurring", c["services_pct"], "%", PCT, "ASSUMPTION. Self-serve SMB onboarding.", True)
    sheet.add("capex_pct", "Capex, % of revenue (also tax depreciation)", c["capex_pct"], "%", PCT, "ASSUMPTION. Development cost treated as expensed, so EBITDA is close to cash EBITDA.", True)
    sheet.add("nwc_pct", "Working capital inflow, % of organic ARR added", c["nwc_pct_delta_arr"], "%", PCT, "Deferred revenue rose by 46% of the revenue increase in 2024 (E53); 25% allows for receivables.", True)
    sheet.add("cir_pct", "Research tax credit, % of R&D cost", c["cir_pct_rd"], "%", PCT, "Tax credits were EUR 0.15M in 2024 (E56), about 3% of an estimated R&D cost.", True)
    sheet.add("dau_entry", "Daily users at closing", m["dau_entry_m"], "M", NUM2, "E04.")
    sheet.add("yield_growth", "Growth in ARR per daily user, per year", m["yield_growth"], "%", PCT, "ASSUMPTION. Price and module mix.", True)
    sheet.add("frontline", "French frontline workers", m["france_frontline_m"], "M", NUM, "E28: broad company definition, includes industry.")
    sheet.add("fte_cost", "Loaded cost per employee", m["loaded_cost_per_fte_k"], "EUR k", INT, "ASSUMPTION. Paris / Lille / Barcelona blend.", True)
    sheet.add("personnel_share", "Personnel share of cash costs", m["personnel_share"], "%", PCT0, "ASSUMPTION. Typical for SaaS at this scale.", True)
    sheet.row += 1

    sheet.section("Last-twelve-month calibration (Triangulation tab)")
    sheet.add("arr_dec25", "ARR at end-2025", cal["arr_dec_2025"], "EUR M", NUM, "E01.")
    sheet.add("growth_ltm", "Reported growth, read as June to June", cal["growth_ltm"], "%", PCT0, "E02. Period not stated in the source, so this reading is an assumption.", True)
    sheet.add("arr_mid25", "ARR at mid-2025", f"={sheet.scalar['arr_entry']}/(1+{sheet.scalar['growth_ltm']})", "EUR M", NUM, "Derived.")
    sheet.ws[f"D{sheet.row-1}"].font = font(color="000000")
    sheet.add("ltm_new_france", "LTM new-logo ARR, France", cal["ltm_new_france"], "EUR M", NUM, "ASSUMPTION, chosen with the lines below to reproduce a small profit and 400 staff.", True)
    sheet.add("ltm_new_intl", "LTM new-logo ARR, international", cal["ltm_new_intl"], "EUR M", NUM, "ASSUMPTION.", True)
    sheet.add("ltm_upsell", "LTM upsell ARR", cal["ltm_upsell"], "EUR M", NUM, "ASSUMPTION. 10% of opening ARR.", True)
    sheet.add("ltm_gm", "LTM gross margin", cal["ltm_gross_margin"], "%", PCT, "ASSUMPTION.", True)
    sheet.add("ltm_cs", "LTM customer success, % of revenue", cal["ltm_cs_pct"], "%", PCT, "ASSUMPTION.", True)
    sheet.add("ltm_rd", "LTM R&D, % of revenue", cal["ltm_rd_pct"], "%", PCT, "ASSUMPTION.", True)
    sheet.add("ltm_ga", "LTM G&A, % of revenue", cal["ltm_ga_pct"], "%", PCT, "ASSUMPTION.", True)
    sheet.add("ltm_cac_fr", "LTM CAC ratio, France", cal["ltm_cac_france"], "x", MULT, "ASSUMPTION. S&M per EUR 1 of new ARR.", True)
    sheet.add("ltm_cac_in", "LTM CAC ratio, international", cal["ltm_cac_intl"], "x", MULT, "ASSUMPTION.", True)
    sheet.add("ltm_cac_up", "LTM CAC ratio, upsell", cal["ltm_cac_upsell"], "x", MULT, "ASSUMPTION.", True)
    sheet.add("headcount", "Reported headcount", cal["headcount"], "people", INT, "E06.")
    sheet.row += 1

    ref = inp["references"]
    sheet.section("Reference figures quoted in the Triangulation tab")
    sheet.add("customers", "Customers", ref["customers"], "no.", INT, "E04.")
    sheet.add("hcr_staff", "French hospitality salaried staff", ref["hcr_staff_m"], "M", NUM2, "E29.")
    sheet.add("planday_price", "Planday total price", ref["planday_price_m"], "EUR M", NUM, "E36, including earn-out.")
    sheet.add("planday_users", "Planday employee users", ref["planday_users_m"], "M", NUM2, "E36.")
    sheet.add("dayforce_value", "Dayforce deal value", ref["dayforce_value_usd_m"], "USD M", INT, "E25.")
    sheet.add("dayforce_q2", "Dayforce Q2-25 revenue", ref["dayforce_q2_revenue_usd_m"], "USD M", NUM, "E25.")
    sheet.add("paycor_ev", "Paycor enterprise value", 4100, "USD M", INT, "E44.")
    sheet.add("paycor_rev", "Paycor revenue", 655, "USD M", INT, "E44.")
    sheet.add("fund_size", "BDC V fund size", ref["fund_size_m"], "EUR M", INT, "E18.")
    sheet.row += 1

    acc = inp["accounts_2024"]
    sheet.section("2024 filed accounts, Skello SAS (Evidence E50-E57)")
    for key, label, fmt, note in [
        ("revenue_2023", "Revenue FY2023", NUM, "E50."), ("revenue_2024", "Revenue FY2024", NUM, "E50."),
        ("operating_result_2023", "Operating result FY2023", NUM, "E51."), ("operating_result_2024", "Operating result FY2024", NUM, "E51."),
        ("da_2023", "Depreciation and amortisation FY2023", NUM, "E51."), ("da_2024", "Depreciation and amortisation FY2024", NUM, "E51."),
        ("net_provisions_2023", "Net provision charges FY2023", NUM, "E51, E54."), ("net_provisions_2024", "Net provision charges FY2024", NUM, "E51, E54."),
        ("capitalised_dev_2023", "Capitalised development FY2023", NUM, "E51."), ("capitalised_dev_2024", "Capitalised development FY2024", NUM, "E51."),
        ("cash_2024", "Cash at end-2024", NUM, "E52."), ("bank_debt_2024", "Bank borrowings at end-2024", NUM, "E52."),
        ("deferred_revenue_2023", "Deferred revenue at end-2023", NUM, "E53."), ("deferred_revenue_2024", "Deferred revenue at end-2024", NUM, "E53."),
        ("doubtful_receivables_2024", "Doubtful receivables, gross", NUM, "E54."), ("bad_debt_charge_2024", "Bad-debt provision charge FY2024", NUM, "E54."),
        ("headcount_avg_2024", "Average headcount, French entity", INT, "E56."), ("personnel_cost_2024", "Personnel cost FY2024", NUM, "E56."),
        ("iberia_revenue_2024", "Skello Iberia revenue FY2024", NUM, "E55."), ("iberia_result_2024", "Skello Iberia result FY2024", NUM, "E55."),
    ]:
        sheet.add(key, label, acc[key], "EUR M" if fmt != INT else "people", fmt, note)
    sheet.row += 1

    sheet.section("Scenario constants")
    sheet.scenario_constant("exit_adj", "Exit multiple adjustment to the Rule of 40 curve (turns of revenue)", MULT1, "Zero in all cases: exit EV is LTM revenue times the curve value. Change it to test a premium or a discount.")
    sheet.scenario_constant("integration_pct", "Integration cost, % of build-up price", PCT, "ASSUMPTION.")

    sheet.section("Operating drivers by scenario")
    sheet.drivers("grr_france", "Gross revenue retention, France", PCT, "ASSUMPTION: the deal's central unknown. Benchmarks in E30 are 71-81%; the Triangulation tab bounds it.", True)
    sheet.drivers("grr_intl", "Gross revenue retention, international", PCT, "ASSUMPTION: younger cohorts retain worse.", True)
    sheet.drivers("upsell_france", "Upsell, % of opening ARR, France", PCT, "ASSUMPTION. Supported by E33 module adoption.", True)
    sheet.drivers("upsell_intl", "Upsell, % of opening ARR, international", PCT, "ASSUMPTION.", True)
    sheet.drivers("new_france", "New-logo ARR, France (EUR M, annual rate)", NUM, "ASSUMPTION. Implied penetration is shown in each model tab.", True)
    sheet.drivers("new_intl", "New-logo ARR, international (EUR M, annual rate)", NUM, "ASSUMPTION. Southern Europe first, then Northern Europe (Maddyness, E03).", True)
    sheet.drivers("cac_france", "CAC ratio, France", MULT, "ASSUMPTION. S&M per EUR 1 of new-logo ARR.", True)
    sheet.drivers("cac_intl", "CAC ratio, international", MULT, "ASSUMPTION. Higher: no brand, smaller reference base.", True)
    sheet.drivers("cac_upsell", "CAC ratio, upsell", MULT, "ASSUMPTION.", True)
    sheet.drivers("gross_margin", "Gross margin", PCT, "ASSUMPTION. Support-heavy SMB base.", True)
    sheet.drivers("cs_fixed_pct", "Customer success, % of revenue", PCT, "ASSUMPTION.", True)
    sheet.drivers("rd_pct", "R&D, % of revenue", PCT, "ASSUMPTION. Elevated near term: a 'significant part' of proceeds goes to AI (E02).", True)
    sheet.drivers("ga_pct", "G&A, % of revenue", PCT, "ASSUMPTION.", True)
    sheet.drivers("acquired_arr", "Acquired ARR (EUR M)", NUM, "ASSUMPTION. Targets of EUR 5-20M ARR (E16).", True)
    sheet.drivers("ma_multiple", "Build-up price, EV / ARR", MULT1, "ASSUMPTION. Inside the 3-6x time-and-attendance range (E24).", True)
    sheet.drivers("acquired_growth", "Growth of acquired ARR", PCT, "ASSUMPTION. No cross-sell synergy assumed.", True)
    sheet.drivers("acquired_margin", "EBITDA margin on acquired revenue", PCT, "ASSUMPTION.", True)
    finish(sheet.ws, {"A": 2, "B": 50, "C": 10, "D": 11, "E": 2, "F": 9, "G": 9, "H": 9, "I": 9, "J": 9, "K": 9, "L": 70}, freeze="C4", tab="0000FF")
    return sheet


# ---------------------------------------------------------------- Deal
def build_deal(wb, sheet: InputsSheet):
    ws = wb.create_sheet("Deal")
    s = sheet.scalar
    put(ws, "B1", "Transaction reconstruction", bold=True, size=12)
    put(ws, "B2", "Outside-in view from press reports. EUR millions. Only post-money, cheque sizes and debt are reported; the rest follows from the primary-cash assumption.", italic=True, color="555555")
    refs = {}

    def row(r, key, label, formula, fmt=NUM, bold=False, link=False):
        put(ws, f"B{r}", label, bold=bold)
        put(ws, f"D{r}", formula, fmt=fmt, bold=bold, color=GREEN_LINK if link else "000000")
        if bold:
            for col in "BCD":
                ws[f"{col}{r}"].border = Border(top=THIN)
        if key:
            refs[key] = f"Deal!$D${r}"

    band(ws, 4, "Sources", last="D")
    row(5, "src_bp", "Bridgepoint (BDC V)", f"={s['bp_equity']}", link=True)
    row(6, "src_re", "Founders and VCs, new money", f"={s['reinvest_equity']}", link=True)
    row(7, "src_debt", "Unitranche (Eurazeo)", f"={s['unitranche']}", link=True)
    row(8, "total_sources", "Total sources", "=SUM(D5:D7)", bold=True)

    band(ws, 10, "Uses", last="D")
    row(11, "primary", "Primary cash to balance sheet", f"={s['primary_cash']}", link=True)
    row(12, "deal_fees", "Transaction costs", f"={s['deal_fee_pct']}*({s['post_money']}+{s['unitranche']})")
    row(13, "oid", "Original issue discount", f"={s['oid_pct']}*{s['unitranche']}")
    row(14, "secondary", "Purchase of existing shares (balancing item)", "=D8-D11-D12-D13")
    row(15, "total_uses", "Total uses", "=SUM(D11:D14)", bold=True)

    band(ws, 17, "Entry valuation", last="D")
    row(18, "post_money", "Post-money equity value", f"={s['post_money']}", link=True)
    row(19, None, "Unitranche", f"={s['unitranche']}", link=True)
    row(20, "cash_entry", "Less: cash after closing", f"={s['existing_cash']}+{s['primary_cash']}")
    row(21, "ev_entry", "Enterprise value", "=D18+D19-D20", bold=True)
    row(22, "m_entry", "EV / ARR", f"=D21/{s['arr_entry']}", MULT)
    row(23, "net_debt_entry", "Net debt / (net cash) at closing", "=D19-D20")
    row(24, "ev_ltm", "EV / LTM revenue", "=D21/Triangulation!$D$8", MULT)
    row(25, "debt_arr", "Unitranche / ARR", f"={s['unitranche']}/{s['arr_entry']}", MULT)

    band(ws, 27, "Cap table after closing", last="F")
    put(ws, "D28", "EUR M", bold=True, align="right")
    put(ws, "E28", "%", bold=True, align="right")
    put(ws, "F28", "Press check", bold=True, align="right")
    put(ws, "B29", "Bridgepoint")
    put(ws, "D29", f"={s['bp_equity']}", fmt=NUM, color=GREEN_LINK)
    put(ws, "E29", "=D29/$D$31", fmt=PCT)
    put(ws, "F29", "40-50%", color="555555", align="right")
    put(ws, "B30", "Founders, management, VCs and others (rollover plus new money)")
    put(ws, "D30", "=D31-D29", fmt=NUM)
    put(ws, "E30", "=D30/$D$31", fmt=PCT)
    put(ws, "F30", "Mgmt ~40%", color="555555", align="right")
    put(ws, "B31", "Total", bold=True)
    put(ws, "D31", f"={s['post_money']}", fmt=NUM, bold=True, color=GREEN_LINK)
    put(ws, "E31", "=SUM(E29:E30)", fmt=PCT, bold=True)
    refs["bp_share"] = "Deal!$E$29"
    row(32, "rollover", "Of which rolled over without new money", f"=D31-{s['bp_equity']}-{s['reinvest_equity']}")
    row(33, "old_holder_value", "Value received by pre-deal holders (cash plus rollover)", "=D14+D32")
    row(34, "ev_check", "Check: EV + existing cash - costs", f"=D21+{s['existing_cash']}-D12-D13")

    band(ws, 36, "What the undisclosed primary cash does to entry price", last="H")
    for col, h in zip("DEFGH", ["Primary", "Cash after", "EV", "EV / ARR", "Secondary"]):
        put(ws, f"{col}37", h, bold=True, align="right")
    for i, p in enumerate((20, 40, 60, 80, 100)):
        r = 38 + i
        put(ws, f"D{r}", p, fmt=NUM, color=BLUE_INPUT)
        put(ws, f"E{r}", f"={s['existing_cash']}+D{r}", fmt=NUM)
        put(ws, f"F{r}", f"={s['post_money']}+{s['unitranche']}-E{r}", fmt=NUM)
        put(ws, f"G{r}", f"=F{r}/{s['arr_entry']}", fmt=MULT)
        put(ws, f"H{r}", f"=$D$8-D{r}-$D$12-$D$13", fmt=NUM)

    band(ws, 44, "Reading the EUR 200M headline (Evidence E08-E10)", last="H")
    put(ws, "B45", "CFNEWS: EUR 150M + EUR 50M of equity, plus EUR 40M debt outside the headline.")
    put(ws, "B46", "Maddyness: the EUR 200M mixes equity and debt, so new equity is about EUR 160M.")
    put(ws, "B47", "Post-money, Bridgepoint's cheque and debt are the same in both readings, so returns are the same. Only the secondary amount changes.", bold=True)

    band(ws, 49, "Alignment and carry", last="H")
    row(50, "net_selldown", "Net sell-down by pre-deal holders (purchase of shares less new money)", "=D14-D6")
    row(51, "selldown_pct", "As a share of the value they held", "=D50/D33", PCT)
    row(52, "carry", "Annual negative carry: cash after closing at deposit rate against the unitranche", f"=D19*({s['euribor']}+{s['margin']})-D20*MAX({s['euribor']}-{s['cash_yield_spread']},0)")
    row(53, "cheque_fund", "Bridgepoint cheque as a share of BDC V", f"={s['bp_equity']}/{s['fund_size']}", PCT)
    finish(ws, {"A": 2, "B": 62, "C": 2, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11}, tab=NAVY)
    return refs


# ---------------------------------------------------------------- Triangulation
def build_triangulation(wb, sheet: InputsSheet):
    ws = wb.create_sheet("Triangulation")
    s = sheet.scalar
    put(ws, "B1", "Triangulation: what the public facts imply", bold=True, size=12)
    put(ws, "B2", "Live formulas, driven from the Inputs tab.", italic=True, color="555555")

    def row(r, label, formula, fmt=NUM, note=None, bold=False):
        put(ws, f"B{r}", label, bold=bold)
        put(ws, f"D{r}", formula, fmt=fmt, bold=bold)
        if note:
            put(ws, f"F{r}", note, italic=True, color="555555")

    band(ws, 4, "1. Reconstructed last twelve months", last="L")
    row(5, "Opening ARR, mid-2025", f"={s['arr_mid25']}")
    row(6, "Closing ARR, mid-2026", f"={s['arr_entry']}")
    row(7, "Net new ARR", "=D6-D5")
    row(8, "Revenue", f"=(D5+D6)/2*(1+{s['services_pct']})", note="Average ARR plus non-recurring.")
    row(9, "Gross profit", f"=D8*{s['ltm_gm']}")
    row(10, "Sales and marketing", f"={s['ltm_new_france']}*{s['ltm_cac_fr']}+{s['ltm_new_intl']}*{s['ltm_cac_in']}+{s['ltm_upsell']}*{s['ltm_cac_up']}+D8*{s['ltm_cs']}")
    row(11, "R&D", f"=D8*{s['ltm_rd']}")
    row(12, "G&A", f"=D8*{s['ltm_ga']}")
    row(13, "EBITDA", "=D9-D10-D11-D12", bold=True)
    row(14, "EBITDA margin", "=D13/D8", PCT, "Consistent with 'profitable since 2025' (E03).")
    row(15, "Implied headcount", f"=(D8-D13)*{s['personnel_share']}/({s['fte_cost']}/1000)", INT, "Compare with about 400 reported (E06).")
    row(16, "Rule of 40 at entry (points)", "=((D6/D5-1)+D14)*100", NUM, "EV / LTM revenue at entry is on the Deal tab; curve value is in the next row.")
    row(17, "Curve EV / Revenue at that score", f"=FORECAST(D16,Inputs!$H${sheet.curve_y}:$I${sheet.curve_y},Inputs!$H${sheet.curve_x}:$I${sheet.curve_x})", MULT, "Line between the midpoints of the 30-40 and 40-50 bands (E23). Valid for scores of 35-45.")

    band(ws, 19, "2. The three public ARR points", last="L")
    row(20, "ARR mid-2025 (from reported growth)", f"={s['arr_mid25']}")
    row(21, "ARR end-2025", f"={s['arr_dec25']}")
    row(22, "ARR mid-2026", f"={s['arr_entry']}")
    row(23, "Annualised growth H2-25", "=(D21/D20)^2-1", PCT)
    row(24, "Annualised growth H1-26", "=(D22/D21)^2-1", PCT, "Could be seasonality, price, definition or rounding. A EUR 42M end-2025 figure would remove the gap.")

    band(ws, 26, "3. Gross retention implied by LTM growth, at profitable levels of S&M", last="L")
    put(ws, "B27", "Gross new ARR = S&M / CAC ratio. Churn = gross new ARR - net new ARR. GRR = 1 - churn / opening ARR.", italic=True, color="555555")
    put(ws, "B28", "S&M share of revenue (rows); CAC ratio (columns)", bold=True)
    for col, cac in zip("DEF", (0.8, 1.0, 1.2)):
        put(ws, f"{col}28", cac, fmt=MULT, color=BLUE_INPUT)
    put(ws, "C29", 0.35, fmt=PCT0, color=BLUE_INPUT)
    put(ws, "C30", 0.40, fmt=PCT0, color=BLUE_INPUT)
    put(ws, "C31", f"={s['ltm_gm']}-{s['ltm_rd']}-{s['ltm_ga']}", fmt=PCT0)
    for r in (29, 30, 31):
        for col in "DEF":
            put(ws, f"{col}{r}", f'=IF(1-($C{r}*$D$8/{col}$28-$D$7)/$D$5>1,"n/a",1-($C{r}*$D$8/{col}$28-$D$7)/$D$5)', fmt=PCT, align="right")
    put(ws, "B32", "Bottom row is the break-even S&M level on the cost assumptions above; beyond it the business would not be profitable.", italic=True, color="555555")
    row(33, "Calibrated: S&M share / blended CAC / retention", "=D10/D8", PCT0)
    put(ws, "E33", f"=D10/({s['ltm_new_france']}+{s['ltm_new_intl']}+{s['ltm_upsell']})", fmt=MULT)
    put(ws, "F33", f"=1-({s['ltm_new_france']}+{s['ltm_new_intl']}+{s['ltm_upsell']}-D7)/D5", fmt=PCT)

    band(ws, 35, "4. Unit economics", last="L")
    row(36, "ARR per customer (EUR)", f"={s['arr_entry']}*1000000/{s['customers']}", INT)
    row(37, "ARR per customer per month (EUR)", "=D36/12", INT, "List price from EUR 79-89 per site (E26).")
    row(38, "Daily users per customer", f"={s['dau_entry']}*1000000/{s['customers']}", NUM)
    row(39, "ARR per daily user per month (EUR)", f"={s['arr_entry']}/{s['dau_entry']}/12", NUM2, "Combo lists EUR 3.50-6.00 per employee (E27, from Skello's own blog). Not like for like.")
    row(40, "ARR per employee (EUR k)", f"={s['arr_entry']}*1000/{s['headcount']}", INT)
    row(41, "Payback on French new logos (months)", f"={s['ltm_cac_fr']}/{s['ltm_gm']}*12", NUM, "On my calibrated CAC, not reported.")
    row(42, "Gross-margin LTV / CAC, France", f"=({s['ltm_gm']}/(1-INDEX(Inputs!$G${sheet.driver['grr_france'][0]}:$G${sheet.driver['grr_france'][1]},1)))/{s['ltm_cac_fr']}", MULT1, "Generous: either a very good business or a CAC set too low.")

    band(ws, 44, "5. Market room, France", last="L")
    row(45, "French daily users at entry (M)", f"={s['dau_entry']}*{s['france_share']}", NUM2)
    row(46, "Share of French frontline workers", f"=D45/{s['frontline']}", PCT, "Company framing of 12M, including industry (E28).")
    row(47, "French daily users / hospitality staff", f"=D45/{s['hcr_staff']}", PCT, "Growth already has to come from outside the original hospitality core.")

    band(ws, 49, "6. Price reference points", last="L")
    row(50, "Skello EV per daily user (EUR)", f"=Deal!$D$21/{s['dau_entry']}", INT)
    row(51, "Planday price per employee user (EUR)", f"={s['planday_price']}/{s['planday_users']}", INT, "E36. 2021 market; users are not daily users. Direction only.")
    row(52, "Dayforce value / annualised revenue", f"={s['dayforce_value']}/({s['dayforce_q2']}*4)", MULT, "E25. Growing 9.8%; far larger.")
    row(53, "Paycor EV / revenue", f"={s['paycor_ev']}/{s['paycor_rev']}", MULT, "E44.")

    band(ws, 55, "7. What the 2024 filed accounts show", last="L")
    put(ws, "D56", "FY2023", bold=True, align="right")
    put(ws, "E56", "FY2024", bold=True, align="right")
    def two(r, label, f23, f24, fmt=NUM, note=None):
        put(ws, f"B{r}", label)
        if f23 is not None:
            put(ws, f"D{r}", f23, fmt=fmt)
        put(ws, f"E{r}", f24, fmt=fmt)
        if note:
            put(ws, f"F{r}", note, italic=True, color="555555")
    two(57, "Revenue growth", None, f"={s['revenue_2024']}/{s['revenue_2023']}-1", PCT, "Consistent with the reported 40% (E02).")
    two(58, "Reported EBITDA", f"={s['operating_result_2023']}+{s['da_2023']}+{s['net_provisions_2023']}", f"={s['operating_result_2024']}+{s['da_2024']}+{s['net_provisions_2024']}")
    two(59, "Reported EBITDA margin", f"=D58/{s['revenue_2023']}", f"=E58/{s['revenue_2024']}", PCT)
    two(60, "EBITDA with development expensed", f"=D58-{s['capitalised_dev_2023']}", f"=E58-{s['capitalised_dev_2024']}")
    two(61, "Margin with development expensed", f"=D60/{s['revenue_2023']}", f"=E60/{s['revenue_2024']}", PCT, "The model's EBITDA is on this basis.")
    two(62, "Deferred revenue / revenue", f"={s['deferred_revenue_2023']}/{s['revenue_2023']}", f"={s['deferred_revenue_2024']}/{s['revenue_2024']}", PCT, "Customers pay ahead: growth releases cash.")
    two(63, "Increase in deferred revenue / increase in revenue", None, f"=({s['deferred_revenue_2024']}-{s['deferred_revenue_2023']})/({s['revenue_2024']}-{s['revenue_2023']})", PCT)
    two(64, "Doubtful receivables / revenue", None, f"={s['doubtful_receivables_2024']}/{s['revenue_2024']}", PCT, "Customer failures show up here before they show in churn.")
    two(65, "Bad-debt charge / revenue", None, f"={s['bad_debt_charge_2024']}/{s['revenue_2024']}", PCT)
    two(66, "Skello Iberia margin", None, f"={s['iberia_result_2024']}/{s['iberia_revenue_2024']}", PCT, "International growth is funded by the French entity.")
    two(67, "Personnel cost per head (EUR k)", None, f"={s['personnel_cost_2024']}/{s['headcount_avg_2024']}*1000", INT, "Compare with the EUR 80k blended assumption, which includes Spain.")
    two(68, "Net cash at end-2024", None, f"={s['cash_2024']}-{s['bank_debt_2024']}")
    finish(ws, {"A": 2, "B": 52, "C": 7, "D": 11, "E": 9, "F": 9, "G": 9, "H": 9, "I": 9, "J": 9, "K": 9, "L": 9}, tab="7A7F8C")


# ---------------------------------------------------------------- Model tabs
def model_spec():
    """Row specification shared by the three model tabs.

    Tokens: [key] same column; [key<] previous column; {in:x} Inputs scalar;
    {deal:x} Deal cell; {drv:x} scenario row of a driver; {cs:x} scenario constant.
    """
    drivers = [
        ("grr_france", "Gross retention, France", PCT), ("grr_intl", "Gross retention, international", PCT),
        ("upsell_france", "Upsell, France", PCT), ("upsell_intl", "Upsell, international", PCT),
        ("new_france", "New-logo ARR, France (annual rate)", NUM), ("new_intl", "New-logo ARR, international (annual rate)", NUM),
        ("cac_france", "CAC ratio, France", MULT), ("cac_intl", "CAC ratio, international", MULT), ("cac_upsell", "CAC ratio, upsell", MULT),
        ("gross_margin", "Gross margin", PCT), ("cs_fixed_pct", "Customer success, % revenue", PCT), ("rd_pct", "R&D, % revenue", PCT),
        ("ga_pct", "G&A, % revenue", PCT), ("acquired_arr", "Acquired ARR", NUM), ("ma_multiple", "Build-up price, EV / ARR", MULT1),
        ("acquired_growth", "Growth of acquired ARR", PCT), ("acquired_margin", "Margin on acquired revenue", PCT),
    ]
    spec = [("section", "Timeline")]
    spec += [
        dict(key="date", label="Period end", fmt=DATE, formula="={period_end}", open="={in:closing_date}", link=True),
        dict(key="L", label="Period length (years)", fmt="0.0", formula="={length}", link=True),
        dict(key="elapsed", label="Years since closing", fmt="0.0", formula="=[elapsed<]+[L]", open=0),
        ("section", "Drivers for this scenario (linked from Inputs)"),
    ]
    spec += [dict(key=f"a_{k}", label=l, fmt=f, formula=f"={{drv:{k}}}", link=True) for k, l, f in drivers]
    spec += [
        ("section", "ARR: France"),
        dict(key="arr_fr_open", label="Opening", formula="=[arr_fr<]"),
        dict(key="arr_fr_churn", label="Less: churn and contraction", formula="=[arr_fr_open]*(1-[a_grr_france])*[L]"),
        dict(key="arr_fr_upsell", label="Plus: upsell", formula="=[arr_fr_open]*[a_upsell_france]*[L]"),
        dict(key="arr_fr_new", label="Plus: new logos", formula="=[a_new_france]*[L]"),
        dict(key="arr_fr", label="Closing ARR, France", formula="=[arr_fr_open]-[arr_fr_churn]+[arr_fr_upsell]+[arr_fr_new]", open="={in:arr_entry}*{in:france_share}", bold=True),
        ("section", "ARR: international"),
        dict(key="arr_in_open", label="Opening", formula="=[arr_in<]"),
        dict(key="arr_in_churn", label="Less: churn and contraction", formula="=[arr_in_open]*(1-[a_grr_intl])*[L]"),
        dict(key="arr_in_upsell", label="Plus: upsell", formula="=[arr_in_open]*[a_upsell_intl]*[L]"),
        dict(key="arr_in_new", label="Plus: new logos", formula="=[a_new_intl]*[L]"),
        dict(key="arr_in", label="Closing ARR, international", formula="=[arr_in_open]-[arr_in_churn]+[arr_in_upsell]+[arr_in_new]", open="={in:arr_entry}*(1-{in:france_share})", bold=True),
        ("section", "ARR: acquired"),
        dict(key="arr_acq_open", label="Opening", formula="=[arr_acq<]"),
        dict(key="arr_acq_add", label="Plus: acquired in period", formula="=[a_acquired_arr]"),
        dict(key="arr_acq_growth", label="Plus: growth on acquired base", formula="=[arr_acq_open]*[a_acquired_growth]*[L]"),
        dict(key="arr_acq", label="Closing ARR, acquired", formula="=[arr_acq_open]+[arr_acq_add]+[arr_acq_growth]", open=0, bold=True),
        ("section", "ARR: total"),
        dict(key="arr_org", label="Organic ARR", formula="=[arr_fr]+[arr_in]", open="=[arr_fr]+[arr_in]"),
        dict(key="arr_total", label="Total ARR", formula="=[arr_org]+[arr_acq]", open="=[arr_org]+[arr_acq]", bold=True),
        dict(key="arr_year_ago", label="Total ARR twelve months earlier", formula="=[arr_total<]", first="={in:arr_dec25}"),
        dict(key="arr_growth", label="Total ARR growth, twelve months (includes acquired)", fmt=PCT, formula="=[arr_total]/[arr_year_ago]-1"),
        dict(key="org_year_ago", label="Organic ARR twelve months earlier", formula="=[arr_org<]", first="={in:arr_dec25}"),
        dict(key="org_growth", label="Organic ARR growth, twelve months", fmt=PCT, formula="=[arr_org]/[org_year_ago]-1"),
        ("section", "P&L"),
        dict(key="rev_org", label="Organic revenue", formula="=([arr_fr_open]+[arr_in_open]+[arr_org])/2*[L]*(1+{in:services_pct})"),
        dict(key="rev_acq", label="Acquired revenue", formula="=([arr_acq_open]+[arr_acq])/2*[L]"),
        dict(key="revenue", label="Revenue", formula="=[rev_org]+[rev_acq]", bold=True),
        dict(key="cogs", label="Less: cost of revenue (organic)", formula="=[rev_org]*(1-[a_gross_margin])"),
        dict(key="sm_var", label="Less: sales and marketing, acquisition", formula="=[arr_fr_new]*[a_cac_france]+[arr_in_new]*[a_cac_intl]+([arr_fr_upsell]+[arr_in_upsell])*[a_cac_upsell]"),
        dict(key="sm_cs", label="Less: customer success", formula="=[rev_org]*[a_cs_fixed_pct]"),
        dict(key="rd", label="Less: R&D", formula="=[rev_org]*[a_rd_pct]"),
        dict(key="ga", label="Less: G&A", formula="=[rev_org]*[a_ga_pct]"),
        dict(key="ebitda_org", label="Organic EBITDA", formula="=[rev_org]-[cogs]-[sm_var]-[sm_cs]-[rd]-[ga]"),
        dict(key="ebitda_acq", label="Acquired EBITDA", formula="=[rev_acq]*[a_acquired_margin]"),
        dict(key="ebitda", label="EBITDA", formula="=[ebitda_org]+[ebitda_acq]", bold=True),
        dict(key="ebitda_margin", label="EBITDA margin", fmt=PCT, formula="=[ebitda]/[revenue]"),
        dict(key="r40", label="Rule of 40, organic growth plus margin (points)", fmt=NUM, formula="=([org_growth]+[ebitda_margin])*100"),
        ("section", "Financing costs (on opening balances, so no circularity)"),
        dict(key="interest", label="Interest on unitranche and line", formula="=({in:unitranche}+[line<])*({in:euribor}+{in:margin})*[L]"),
        dict(key="commitment", label="Commitment fee on undrawn line", formula="=({in:line_commitment}-[line<])*{in:margin}*{in:commitment_fee_share}*[L]"),
        dict(key="interest_income", label="Interest income", formula="=[cash<]*MAX({in:euribor}-{in:cash_yield_spread},0)*[L]"),
        ("section", "Tax"),
        dict(key="capex", label="Capex (also tax depreciation)", formula="=[revenue]*{in:capex_pct}"),
        dict(key="fin_charges", label="Net financial charges", formula="=[interest]+[commitment]-[interest_income]"),
        dict(key="deductible", label="Deductible charges (higher of floor and 30% of EBITDA)", formula="=IF([fin_charges]>0,MIN([fin_charges],MAX({in:interest_floor}*[L],{in:interest_cap_pct}*MAX([ebitda],0))),[fin_charges])"),
        dict(key="taxable", label="Taxable result before losses", formula="=[ebitda]-[capex]-[deductible]"),
        dict(key="nol_open", label="Tax losses, opening", formula="=[nol<]"),
        dict(key="nol_use", label="Losses used", formula="=IF([taxable]>0,MIN([nol_open],{in:nol_free}*[L]+{in:nol_excess_share}*MAX([taxable]-{in:nol_free}*[L],0),[taxable]),0)"),
        dict(key="nol", label="Tax losses, closing", formula="=IF([taxable]>0,[nol_open]-[nol_use],[nol_open]-[taxable])", open="={in:nol_opening}"),
        dict(key="tax", label="Tax paid", formula="=MAX([taxable]-[nol_use],0)*{in:tax_rate}", bold=True),
        ("section", "Build-ups"),
        dict(key="ma_spend", label="Build-up spend", formula="=[arr_acq_add]*[a_ma_multiple]"),
        dict(key="integration", label="Integration cost", formula="=[ma_spend]*{cs:integration_pct}"),
        ("section", "Cash flow before external funding (signed)"),
        dict(key="cf_ebitda", label="EBITDA", formula="=[ebitda]"),
        dict(key="cf_capex", label="Capex", formula="=-[capex]"),
        dict(key="nwc", label="Working capital", formula="={in:nwc_pct}*([arr_total]-[arr_total<]-[arr_acq_add])"),
        dict(key="cir", label="Research tax credit (counted when earned)", formula="=[rd]*{in:cir_pct}"),
        dict(key="cf_fin", label="Net financing costs", formula="=-([interest]+[commitment]-[interest_income])"),
        dict(key="cf_tax", label="Tax", formula="=-[tax]"),
        dict(key="cf_ma", label="Build-ups and integration", formula="=-([ma_spend]+[integration])"),
        dict(key="pre_fin_cf", label="Cash flow before external funding", formula="=[cf_ebitda]+[cf_capex]+[nwc]+[cir]+[cf_fin]+[cf_tax]+[cf_ma]", bold=True),
        ("section", "Funding waterfall: cash, then capex line, then new equity"),
        dict(key="position", label="Cash before funding", formula="=[cash<]+[pre_fin_cf]"),
        dict(key="need", label="Shortfall against minimum cash", formula="=MAX({in:min_cash}-[position],0)"),
        dict(key="draw", label="Capex line drawn", formula="=MIN([need],{in:line_commitment}-[line<])"),
        dict(key="gap", label="New equity required (funding gap)", formula="=[need]-[draw]", bold=True),
        dict(key="repay", label="Capex line repaid", formula="=MIN([line<],MAX([position]-{in:min_cash},0))"),
        dict(key="line", label="Capex line, closing", formula="=[line<]+[draw]-[repay]", open=0),
        dict(key="cash", label="Cash, closing", formula="=[position]+[draw]+[gap]-[repay]", open="={deal:cash_entry}", bold=True),
        dict(key="unitranche", label="Unitranche, closing", formula="={in:unitranche}", open="={in:unitranche}"),
        dict(key="net_debt", label="Net debt / (net cash)", formula="=[unitranche]+[line]-[cash]", open="=[unitranche]+[line]-[cash]", bold=True),
        ("section", "Credit and market read-across (no covenant is modelled; terms are undisclosed)"),
        dict(key="interest_cover", label="EBITDA / cash interest and fees", fmt=MULT, formula="=IF([interest]+[commitment]=0,0,[ebitda]/([interest]+[commitment]))"),
        dict(key="debt_to_arr", label="Gross debt / ARR", fmt=MULT, formula="=({in:unitranche}+[line])/[arr_total]"),
        dict(key="debt_to_ebitda", label="Gross debt / annualised EBITDA", fmt=MULT1, formula="=IF([ebitda]>0,({in:unitranche}+[line])/([ebitda]/[L]),0)"),
        dict(key="yield_per_dau", label="ARR per daily user (EUR)", fmt=NUM, formula="={in:arr_entry}/{in:dau_entry}*(1+{in:yield_growth})^[elapsed]"),
        dict(key="fr_dau", label="Implied French daily users (M)", fmt=NUM2, formula="=[arr_fr]/[yield_per_dau]"),
        dict(key="fr_penetration", label="Share of French frontline workers", fmt=PCT, formula="=[fr_dau]/{in:frontline}"),
        dict(key="fte_implied", label="Implied headcount (period-average cost base)", fmt=INT, formula="=([revenue]-[ebitda])*{in:personnel_share}/[L]/({in:fte_cost}/1000)"),
        ("section", "Exit waterfall, for an exit at each period end"),
        dict(key="x_r", label="Rule of 40, limited to the curve's range", fmt=NUM, formula="=MAX(MIN([r40],{cx_last}),{cx_first})"),
        dict(key="x_m", label="Curve segment", fmt="0", formula="=MIN(MATCH([x_r],{cx},1),5)"),
        dict(key="x_curve", label="EV / LTM revenue from the curve", fmt=MULT, formula="=INDEX({cy},[x_m])+([x_r]-INDEX({cx},[x_m]))*(INDEX({cy},[x_m]+1)-INDEX({cy},[x_m]))/(INDEX({cx},[x_m]+1)-INDEX({cx},[x_m]))"),
        dict(key="x_mult", label="EV / LTM revenue applied (curve plus adjustment)", fmt=MULT, formula="=[x_curve]+{cs:exit_adj}", bold=True),
        dict(key="x_ev", label="Enterprise value", formula="=[revenue]/[L]*[x_mult]", bold=True),
        dict(key="x_ev_arr", label="Implied EV / ARR", fmt=MULT, formula="=[x_ev]/[arr_total]"),
        dict(key="x_ev_ebitda", label="Implied EV / EBITDA", fmt=MULT1, formula="=IF([ebitda]>0,[x_ev]/[ebitda]*[L],0)"),
        dict(key="x_costs", label="Less: exit costs", formula="=[x_ev]*{in:exit_cost_pct}"),
        dict(key="x_equity", label="Equity value", formula="=MAX([x_ev]-[net_debt]-[x_costs],0)", bold=True),
        dict(key="x_follow_on", label="Cumulative new equity", formula="=[x_follow_on<]+[gap]", open=0),
        dict(key="x_mip", label="Less: management incentive", formula="={in:mip_share}*MAX([x_equity]-{in:mip_hurdle}*({in:post_money}+[x_follow_on]),0)"),
        dict(key="x_dist", label="Distributable to shareholders", formula="=[x_equity]-[x_mip]"),
        dict(key="x_bp", label="Bridgepoint proceeds, ordinary shares", formula="=[x_dist]*{deal:bp_share}", bold=True),
        dict(key="x_bp_inv", label="Bridgepoint capital invested", formula="={in:bp_equity}+{deal:bp_share}*[x_follow_on]"),
        dict(key="x_moic", label="Bridgepoint MOIC, ordinary shares", fmt=MULT, formula="=[x_bp]/[x_bp_inv]", bold=True),
        dict(key="x_bp_pref", label="Bridgepoint proceeds with a 1x senior preference on new money", formula="=MAX([x_bp],MIN([x_dist],{in:bp_equity}+{in:reinvest_equity})*{in:bp_equity}/({in:bp_equity}+{in:reinvest_equity}))"),
        dict(key="x_moic_pref", label="Bridgepoint MOIC with the preference", fmt=MULT, formula="=[x_bp_pref]/[x_bp_inv]"),
        dict(key="x_moic_pref_bp", label="Bridgepoint MOIC with a 1x senior preference on its own cheque only", fmt=MULT, formula="=MAX([x_bp],MIN([x_dist],{in:bp_equity}))/[x_bp_inv]"),
        ("section", "Bridgepoint cash flows and IRR by exit date (ordinary shares)"),
    ]
    for label in ("Dec-29", "Dec-30", "Dec-31"):
        spec.append(dict(key=f"flow_{label}", label=f"Exit {label}: cash flows", fmt=NUM, exit_label=label))
    spec.append(("irr_block",))
    spec.append(("section", "Equity value bridge to a Dec-30 exit, 100% of equity"))
    spec.append(("bridge",))
    return spec


EXIT_COL = {"Dec-29": "I", "Dec-30": "J", "Dec-31": "K"}


def build_model_tab(wb, name, index, sheet: InputsSheet, deal_refs, inp):
    ws = wb.create_sheet(name)
    put(ws, "B1", f"Model: {name} case", bold=True, size=12)
    put(ws, "B2", "EUR millions. Formula-identical across the Base, Downside and Plan tabs; only D3 differs.", italic=True, color="555555")
    put(ws, "B3", "Scenario index (1 Base, 2 Downside, 3 Plan)")
    put(ws, "D3", index, fmt="0", color=BLUE_INPUT)
    put(ws, "E4", "Closing", bold=True, align="right")
    for col, p in zip(PERIOD_COLS, inp["periods"]):
        put(ws, f"{col}4", p, bold=True, align="right")

    spec = model_spec()
    rows, r = {}, 5
    for item in spec:
        if isinstance(item, tuple):
            if item[0] == "section":
                r += 2
            elif item[0] == "irr_block":
                r += 2
            elif item[0] == "bridge":
                r += 9
            continue
        rows[item["key"]] = r
        r += 1

    cx = f"Inputs!$F${sheet.curve_x}:$K${sheet.curve_x}"
    cy = f"Inputs!$F${sheet.curve_y}:$K${sheet.curve_y}"

    def resolve(tmpl, col):
        prev = chr(ord(col) - 1)

        def token(m):
            key = m.group(1)
            if key.endswith("<"):
                return f"{prev}{rows[key[:-1]]}"
            return f"{col}{rows[key]}"

        out = re.sub(r"\[([A-Za-z0-9_\-<]+)\]", token, tmpl)
        out = out.replace("{period_end}", f"Inputs!{col}${sheet.period_end_row}")
        out = out.replace("{length}", f"Inputs!{col}${sheet.length_row}")
        out = out.replace("{cx_last}", f"Inputs!$K${sheet.curve_x}").replace("{cx_first}", f"Inputs!$F${sheet.curve_x}")
        out = out.replace("{cx}", cx).replace("{cy}", cy)
        out = re.sub(r"\{in:(\w+)\}", lambda m: sheet.scalar[m.group(1)], out)
        out = re.sub(r"\{deal:(\w+)\}", lambda m: deal_refs[m.group(1)], out)

        def drv(m):
            a, b = sheet.driver[m.group(1)]
            return f"INDEX(Inputs!{col}${a}:{col}${b},$D$3)"

        def cs(m):
            a, b = sheet.constant[m.group(1)]
            return f"INDEX(Inputs!$D${a}:$D${b},$D$3)"

        out = re.sub(r"\{drv:(\w+)\}", drv, out)
        out = re.sub(r"\{cs:(\w+)\}", cs, out)
        return out

    r = 5
    bridge_rows = {}
    irr_rows = {}
    for item in spec:
        if isinstance(item, tuple):
            if item[0] == "section":
                r += 1
                band(ws, r, item[1], last="K")
                r += 1
            elif item[0] == "irr_block":
                put(ws, f"B{r}", "Bridgepoint IRR (XIRR, actual/365; -100% if no proceeds)", bold=True)
                for label, col in EXIT_COL.items():
                    fr = rows[f"flow_{label}"]
                    put(ws, f"{col}{r}", f"=IFERROR(XIRR($E${fr}:$K${fr},$E${rows['date']}:$K${rows['date']}),-1)", fmt=PCT, bold=True)
                    irr_rows[label] = f"{col}{r}"
                r += 2
            elif item[0] == "bridge":
                J = "J"
                entries = [
                    ("entry_equity", "Post-money equity at closing", f"={sheet.scalar['post_money']}", GREEN_LINK),
                    ("organic", "Organic ARR growth at entry multiple", f"=({J}{rows['arr_org']}-{sheet.scalar['arr_entry']})*{deal_refs['m_entry']}", "000000"),
                    ("acquired", "Acquired ARR at entry multiple", f"={J}{rows['arr_acq']}*{deal_refs['m_entry']}", "000000"),
                    ("multiple", "Change in EV / ARR on exit ARR", f"=({J}{rows['x_ev_arr']}-{deal_refs['m_entry']})*{J}{rows['arr_total']}", "000000"),
                    ("net_debt", "Change in net debt (includes build-up spend)", f"=-({J}{rows['net_debt']}-E{rows['net_debt']})", "000000"),
                    ("costs", "Exit costs", f"=-{J}{rows['x_costs']}", "000000"),
                ]
                start = r
                for key, label, formula, colour in entries:
                    put(ws, f"B{r}", label)
                    put(ws, f"E{r}", formula, fmt=NUM, color=colour)
                    bridge_rows[key] = f"E{r}"
                    r += 1
                put(ws, f"B{r}", "Equity value at Dec-30 (before incentive)", bold=True)
                put(ws, f"E{r}", f"=SUM(E{start}:E{r-1})", fmt=NUM, bold=True)
                ws[f"E{r}"].border = Border(top=THIN)
                bridge_rows["exit_equity"] = f"E{r}"
                r += 1
                put(ws, f"B{r}", "Check against the exit waterfall")
                put(ws, f"E{r}", f"=E{r-1}-J{rows['x_equity']}", fmt=NUM2)
                bridge_rows["check"] = f"E{r}"
                r += 1
            continue

        key = item["key"]
        fmt = item.get("fmt", NUM)
        bold = item.get("bold", False)
        put(ws, f"B{r}", item["label"], bold=bold)
        color = GREEN_LINK if item.get("link") else "000000"
        if "exit_label" in item:
            label = item["exit_label"]
            exit_col = EXIT_COL[label]
            put(ws, f"E{r}", f"=-{sheet.scalar['bp_equity']}", fmt=fmt, color=GREEN_LINK)
            for col in PERIOD_COLS:
                if col > exit_col:
                    formula = "=0"
                elif col == exit_col:
                    formula = f"=-{deal_refs['bp_share']}*{col}{rows['gap']}+{col}{rows['x_bp']}"
                else:
                    formula = f"=-{deal_refs['bp_share']}*{col}{rows['gap']}"
                put(ws, f"{col}{r}", formula, fmt=fmt)
        else:
            if "open" in item:
                value = item["open"]
                value = resolve(value, "E") if isinstance(value, str) else value
                is_link = isinstance(value, str) and re.fullmatch(r"=(Inputs|Deal)![$A-Z0-9]+", value) is not None
                put(ws, f"E{r}", value, fmt=fmt, bold=bold, color=GREEN_LINK if is_link else "000000")
            for i, col in enumerate(PERIOD_COLS):
                tmpl = item["first"] if (i == 0 and "first" in item) else item["formula"]
                cell_color = GREEN_LINK if (i == 0 and "first" in item) else color
                put(ws, f"{col}{r}", resolve(tmpl, col), fmt=fmt, bold=bold, color=cell_color)
        if bold:
            for col in ["B", "C", "D"] + ALL_COLS:
                ws[f"{col}{r}"].border = Border(top=HAIR)
        r += 1
    finish(ws, {"A": 2, "B": 56, "C": 2, "D": 7, "E": 10, "F": 10, "G": 10, "H": 10, "I": 10, "J": 10, "K": 10}, freeze="E5",
           tab={"Base": "2E5EAA", "Downside": "B3261E", "Plan": "2E7D32"}[name])
    return dict(rows=rows, irr=irr_rows, bridge=bridge_rows)


# ---------------------------------------------------------------- Next financial buyer
def build_next_buyer(wb, maps, sheet: InputsSheet, inp):
    ws = wb.create_sheet("NextBuyer")
    s = sheet.scalar
    put(ws, "B1", "What a financial buyer could pay at a Dec-30 exit", bold=True, size=12)
    put(ws, "B2", "EUR millions. Live. The buyer levers, grows, expands margin and exits, earning its target IRR on equity. It is a second lens on the Rule of 40 curve, not a forecast.", italic=True, color="555555")
    refs = {}
    r = 4
    for name in inp["scenarios"]:
        m = maps[name]["rows"]
        band(ws, r, f"{name} case", last="J")
        r += 1
        for col, lab in zip("EFGHIJ", ["Dec-30", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"]):
            put(ws, f"{col}{r}", lab, bold=True, align="right")
        r += 1
        g, rev, mg, eb, fcf = r, r + 1, r + 2, r + 3, r + 4
        put(ws, f"B{g}", "Organic growth")
        put(ws, f"B{rev}", "Revenue")
        put(ws, f"B{mg}", "EBITDA margin")
        put(ws, f"B{eb}", "EBITDA")
        put(ws, f"B{fcf}", "Free cash flow after interest")
        put(ws, f"E{g}", f"={name}!J{m['org_growth']}", fmt=PCT, color=GREEN_LINK)
        put(ws, f"E{rev}", f"={name}!J{m['revenue']}", fmt=NUM, color=GREEN_LINK)
        put(ws, f"E{mg}", f"={name}!J{m['ebitda_margin']}", fmt=PCT, color=GREEN_LINK)
        put(ws, f"E{eb}", f"=E{rev}*E{mg}", fmt=NUM)
        debt_row = r + 6
        for i, col in enumerate("FGHIJ"):
            prev = "EFGHI"[i]
            put(ws, f"{col}{g}", f"=MAX({prev}{g}-{s['nb_fade']},MIN({s['nb_floor']},$E{g}))", fmt=PCT)
            put(ws, f"{col}{rev}", f"={prev}{rev}*(1+{col}{g})", fmt=NUM)
            put(ws, f"{col}{mg}", f"=MIN({prev}{mg}+{s['nb_step']},MAX({s['nb_cap']},$E{mg}))", fmt=PCT)
            put(ws, f"{col}{eb}", f"={col}{rev}*{col}{mg}", fmt=NUM)
            put(ws, f"{col}{fcf}", f"={col}{eb}*{s['nb_conv']}-$E${debt_row}*{s['nb_rate']}", fmt=NUM)
        r += 6
        lines = [
            ("debt", "Opening debt", f"={s['nb_debt']}*MAX(E{eb},0)", NUM),
            ("exit_ev", "Buyer's exit enterprise value", f"=J{eb}*{s['nb_exit']}", NUM),
            ("exit_equity", "Buyer's exit equity", f"=E{r+1}-E{r}+SUM(F{fcf}:J{fcf})", NUM),
            ("equity_in", "Equity the buyer can invest", f"=E{r+2}/(1+{s['nb_irr']})^5", NUM),
            ("ev", "Most the buyer can pay (EV)", f"=E{r+3}+E{r}", NUM),
            ("ev_arr", "EV / ARR", f"=E{r+4}/{name}!J{m['arr_total']}", MULT),
            ("ev_ebitda", "EV / EBITDA", f"=IF(E{eb}>0,E{r+4}/E{eb},0)", MULT1),
        ]
        out = {}
        for key, label, formula, fmt in lines:
            put(ws, f"B{r}", label, bold=key == "ev")
            put(ws, f"E{r}", formula, fmt=fmt, bold=key == "ev")
            out[key] = f"E{r}"
            r += 1
        ev = out["ev"]
        eq = f"MAX({ev}*(1-{s['exit_cost_pct']})-{name}!J{m['net_debt']},0)"
        mip = f"{s['mip_share']}*MAX({eq}-{s['mip_hurdle']}*({s['post_money']}+{name}!J{m['x_follow_on']}),0)"
        put(ws, f"B{r}", "Bridgepoint MOIC at that price, ordinary shares", bold=True)
        put(ws, f"E{r}", f"=({eq}-{mip})*Deal!$E$29/{name}!J{m['x_bp_inv']}", fmt=MULT, bold=True)
        out["moic"] = f"E{r}"
        r += 1
        put(ws, f"B{r}", "Curve exit EV, premium to that price")
        put(ws, f"E{r}", f"={name}!J{m['x_ev']}/{ev}-1", fmt=PCT)
        out["premium"] = f"E{r}"
        r += 2
        refs[name] = out
    finish(ws, {"A": 2, "B": 46, "C": 2, "D": 2, "E": 10, "F": 10, "G": 10, "H": 10, "I": 10, "J": 10}, tab="7A7F8C")
    return refs


# ---------------------------------------------------------------- Analytics
def build_analytics(wb, maps, analysis, deal_refs, sheet: InputsSheet):
    ws = wb.create_sheet("Analytics")
    s = sheet.scalar
    rows = maps["Base"]["rows"]
    put(ws, "B1", "Analytics behind the memo", bold=True, size=12)
    put(ws, "B2", "Sections 1-2 are live. Section 3 holds solved values from src/analysis.py, each with the Goal Seek that reproduces it.", italic=True, color="555555")

    band(ws, 4, "1. Live: Dec-30 ARR needed for a target MOIC (Base net debt, no new equity)", last="K")
    multiples = [4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0]
    put(ws, "B5", "Target MOIC (rows); exit EV / ARR (columns)", bold=True)
    cols = "DEFGHIJ"
    for col, m in zip(cols, multiples):
        put(ws, f"{col}5", m, fmt=MULT1, color=BLUE_INPUT)
    req = {}
    for i, t in enumerate((2.0, 2.5, 3.0)):
        r = 6 + i
        put(ws, f"C{r}", t, fmt=MULT1, color=BLUE_INPUT)
        for col in cols:
            dist = f"$C{r}*{s['bp_equity']}/{deal_refs['bp_share']}"
            hurdle = f"{s['mip_hurdle']}*{s['post_money']}"
            eq = f"IF({dist}<={hurdle},{dist},({dist}-{s['mip_share']}*{hurdle})/(1-{s['mip_share']}))"
            put(ws, f"{col}{r}", f"=({eq}+Base!$J${rows['net_debt']})/({col}$5*(1-{s['exit_cost_pct']}))", fmt=INT)
        req[t] = r

    band(ws, 10, "2. Live: build-up capacity against the EUR 200M ARR target (before growth on acquired ARR)", last="K")
    cf = "+".join(f"Base!{c}{rows['pre_fin_cf']}+Base!{c}{rows['ma_spend']}+Base!{c}{rows['integration']}" for c in "FGHIJ")
    lines = [
        ("near_term", "Cash above minimum at closing, plus capex line", f"={deal_refs['cash_entry']}-{s['min_cash']}+{s['line_commitment']}", NUM),
        ("cum_fcf", "Cash generated to Dec-30 before build-ups (Base)", f"={cf}", NUM),
        ("available", "Capacity to Dec-30", "=D11+D12", NUM),
        ("organic", "Base organic ARR, Dec-30", f"=Base!J{rows['arr_org']}", NUM),
        ("needed", "Acquired ARR needed for EUR 200M", "=MAX(200-D14,0)", NUM),
        ("cost", "All-in cost at the multiple in column C, plus integration", f"=D15*C16*(1+INDEX(Inputs!$D${sheet.constant['integration_pct'][0]}:$D${sheet.constant['integration_pct'][1]},1))", NUM),
        ("gap", "Capital beyond capacity", "=MAX(D16-D13,0)", NUM),
        ("bp_gap", "Bridgepoint pro rata share", f"=D17*{deal_refs['bp_share']}", NUM),
        ("exposure", "Bridgepoint exposure, % of BDC V", f"=({s['bp_equity']}+D18)/{s['fund_size']}", PCT),
        ("debt_low", "Gross debt capacity at the multiple in column C, FY30 EBITDA", f"=Base!J{rows['ebitda']}*C20", NUM),
        ("debt_high", "Gross debt capacity at the multiple in column C, FY30 EBITDA", f"=Base!J{rows['ebitda']}*C21", NUM),
    ]
    out = {}
    for i, (key, label, formula, fmt) in enumerate(lines):
        r = 11 + i
        put(ws, f"B{r}", label)
        put(ws, f"D{r}", formula, fmt=fmt)
        out[key] = f"D{r}"

    put(ws, "C16", 5.0, fmt=MULT1, color=BLUE_INPUT)
    put(ws, "C20", 4.0, fmt=MULT1, color=BLUE_INPUT)
    put(ws, "C21", 5.0, fmt=MULT1, color=BLUE_INPUT)

    band(ws, 23, "3. Solved values (static; rerun src/analysis.py after changing inputs)", last="K")
    put(ws, "B24", "Item", bold=True)
    put(ws, "D24", "Value", bold=True, align="right")
    put(ws, "E24", "How to reproduce in Excel", bold=True)
    th = analysis["thresholds"]
    gs_grr = "Goal Seek Base MOIC by shifting both retention rows and cutting new logos by shift x opening ARR / LTM new logos"
    gs_pm = "Goal Seek Base MOIC or IRR by changing post-money on Inputs"
    solved = [
        ("French retention for 2.0x (international and new logos move with it)", th["grr_both_for_2.0"], PCT, gs_grr),
        ("French retention for 2.5x (international and new logos move with it)", th["grr_both_for_2.5"], PCT, gs_grr),
        ("French retention for 3.0x (international and new logos move with it)", th["grr_both_for_3.0"], PCT, gs_grr),
        ("Post-money for 20% IRR", th["post_money_for_irr_20"], NUM, gs_pm),
        ("Post-money for 2.5x", th["post_money_for_2.5"], NUM, gs_pm),
        ("Post-money for 25% IRR", th["post_money_for_irr_25"], NUM, gs_pm),
        ("Post-money for 3.0x", th["post_money_for_3.0"], NUM, gs_pm),
        ("New-logo ARR factor for 3.0x", th["new_logo_factor_for_3.0"], MULT, "Goal Seek by scaling both new-logo rows on Inputs"),
        ("Total-equity IRR with the debt package", analysis["leverage"]["levered_irr"], PCT, "Base tab flows on 100% of equity"),
        ("Total-equity IRR with no debt, same EV", analysis["leverage"]["unlevered_irr"], PCT, "Set unitranche and line to 0; raise post-money by 40"),
        ("Base to Plan: organic ARR (average of all orders)", analysis["base_to_plan"]["steps"]["Organic ARR"], MULT, "Copy Plan driver rows into Base rows, group by group"),
        ("Base to Plan: margins and CAC (average of all orders)", analysis["base_to_plan"]["steps"]["Margins and CAC"], MULT, "As above"),
        ("Base to Plan: build-ups (average of all orders)", analysis["base_to_plan"]["steps"]["Build-ups"], MULT, "As above"),
    ]
    for i, (label, value, fmt, how) in enumerate(solved):
        r = 25 + i
        put(ws, f"B{r}", label)
        put(ws, f"D{r}", value, fmt=fmt, color="444444")
        put(ws, f"E{r}", how, italic=True, color="555555")
    finish(ws, {"A": 2, "B": 52, "C": 7, "D": 9, "E": 9, "F": 9, "G": 9, "H": 9, "I": 9, "J": 9, "K": 9}, tab="7A7F8C")
    return dict(required=req, cells=out)


# ---------------------------------------------------------------- Sensitivities
def static(ws, ref, value, fmt):
    put(ws, ref, value, fmt=fmt, color="444444")


def build_sensitivities(wb, analysis, maps, sheet: InputsSheet):
    ws = wb.create_sheet("Sensitivities")
    put(ws, "B1", "Sensitivities on the Base case, exit Dec-30, ordinary shares", bold=True, size=12)
    put(ws, "B2", "Exit value follows the Rule of 40 curve, so operating flexes also move the multiple. Grey values are static output of src/analysis.py; "
                  "the Checks tab flags them as STALE if inputs change. Grid E is live.", italic=True, color="555555")
    base = maps["Base"]["rows"]
    s = sheet.scalar
    g = analysis["grid_exit_new_logo"]
    band(ws, 4, "A. MOIC: exit multiple against the curve, turns of revenue (columns); new-logo ARR relative to Base (rows)", last="L")
    put(ws, "B5", "New-logo ARR", bold=True)
    put(ws, "C5", "ARR Dec-30", bold=True, align="right")
    cols = "DEFGHIJ"
    plus = '+0.0"x";-0.0"x";"Curve"'
    for col, m in zip(cols, g["adjustments"]):
        put(ws, f"{col}5", m, fmt=plus, bold=True, align="right")
    for i, row in enumerate(g["rows"]):
        r = 6 + i
        put(ws, f"B{r}", row["factor"], fmt=PCT0, bold=True, align="left")
        static(ws, f"C{r}", row["arr_dec30"], NUM)
        for col, v in zip(cols, row["moic"]):
            static(ws, f"{col}{r}", v, MULT)
    center_a = "G8"
    band(ws, 12, "A2. IRR on the same grid", last="L")
    for col, m in zip(cols, g["adjustments"]):
        put(ws, f"{col}13", m, fmt=plus, bold=True, align="right")
    for i, row in enumerate(g["rows"]):
        r = 14 + i
        put(ws, f"B{r}", row["factor"], fmt=PCT0, bold=True, align="left")
        for col, v in zip(cols, row["irr"]):
            static(ws, f"{col}{r}", v, PCT)
    g = analysis["grid_retention_exit"]
    band(ws, 20, "B. MOIC: exit multiple against the curve (columns); French gross retention, with international and new logos moving with it (rows)", last="L")
    put(ws, "B21", "French retention", bold=True)
    put(ws, "C21", "ARR Dec-30", bold=True, align="right")
    for col, a in zip("DEFGH", g["adjustments"]):
        put(ws, f"{col}21", a, fmt=plus, bold=True, align="right")
    put(ws, "I21", "EV/ARR", bold=True, align="right")
    for i, row in enumerate(g["rows"]):
        r = 22 + i
        put(ws, f"B{r}", row["level"], fmt=PCT0, bold=True, align="left")
        static(ws, f"C{r}", row["arr_dec30"], NUM)
        for col, v in zip("DEFGH", row["moic"]):
            static(ws, f"{col}{r}", v, MULT)
        static(ws, f"I{r}", row["multiple"], MULT)
    center_b = "F26"
    g = analysis["grid_entry"]
    band(ws, 30, "C. Entry: post-money (rows); primary cash (columns). EV / ARR, then Base MOIC", last="L")
    for col, p in zip("DEFGH", g["primary"]):
        put(ws, f"{col}31", p, fmt=INT, bold=True, align="right")
    put(ws, "B31", "Post-money and implied stake", bold=True)
    r = 32
    for row in g["rows"]:
        put(ws, f"B{r}", row["post"], fmt=INT, bold=True, align="left")
        put(ws, f"C{r}", row["stake"], fmt=PCT, color="444444")
        for col, c in zip("DEFGH", row["cells"]):
            static(ws, f"{col}{r}", c["ev_arr"], MULT1)
            static(ws, f"{col}{r+4}", c["moic"], MULT)
        put(ws, f"B{r+4}", row["post"], fmt=INT, bold=True, align="left")
        r += 1
    t = analysis["tornado"]
    band(ws, 41, "D. One-at-a-time flexes on MOIC", last="L")
    put(ws, "D42", "Low case", bold=True, align="right")
    put(ws, "E42", "High case", bold=True, align="right")
    put(ws, "F42", "Spread", bold=True, align="right")
    for i, row in enumerate(t["rows"]):
        r = 43 + i
        put(ws, f"B{r}", row["label"])
        static(ws, f"D{r}", row["low"], MULT)
        static(ws, f"E{r}", row["high"], MULT)
        put(ws, f"F{r}", f"=ABS(E{r}-D{r})", fmt=MULT)
    r0 = 43 + len(t["rows"]) + 1
    band(ws, r0, "E. Live: MOIC by exit date and adjustment to the Base exit multiple", last="L")
    adjustments = [-1.0, -0.5, 0.0, 0.5, 1.0]
    for col, a in zip("DEFGH", adjustments):
        put(ws, f"{col}{r0+1}", a, fmt=plus, color=BLUE_INPUT, bold=True, align="right")
    for i, (label, col_x) in enumerate(EXIT_COL.items()):
        r = r0 + 2 + i
        put(ws, f"B{r}", label, bold=True)
        for col in "DEFGH":
            ev = f"(Base!{col_x}{base['x_mult']}+{col}${r0+1})*Base!{col_x}{base['revenue']}/Base!{col_x}{base['L']}"
            eq = f"MAX({ev}*(1-{s['exit_cost_pct']})-Base!{col_x}{base['net_debt']},0)"
            mip = f"{s['mip_share']}*MAX({eq}-{s['mip_hurdle']}*({s['post_money']}+Base!{col_x}{base['x_follow_on']}),0)"
            put(ws, f"{col}{r}", f"=({eq}-{mip})*Deal!$E$29/Base!{col_x}{base['x_bp_inv']}", fmt=MULT)
    center_d = f"F{r0+3}"
    finish(ws, {"A": 2, "B": 40, "C": 11, "D": 9, "E": 9, "F": 9, "G": 9, "H": 9, "I": 9, "J": 9, "K": 2, "L": 2}, tab="7A7F8C")
    return dict(center_a=center_a, center_b=center_b, center_d=center_d)


# ---------------------------------------------------------------- Checks
def build_checks(wb, maps, sens, deal_refs, sheet: InputsSheet):
    ws = wb.create_sheet("Checks")
    put(ws, "B1", "Integrity checks", bold=True, size=12)
    put(ws, "B2", "Mechanical checks only: they show the file is consistent, not that its assumptions are right. STALE means a static grid needs src/analysis.py to be rerun; REVIEW means a calibration target is no longer met. Neither is a formula error.", italic=True, color="555555")
    for col, h in zip("BCDE", ["Check", "Value", "Tolerance", "Result"]):
        put(ws, f"{col}4", h, bold=True, fill=MID, align=None if col == "B" else "right")
    checks = [
        ("Sources equal uses", f"={deal_refs['total_sources']}-{deal_refs['total_uses']}", 1e-6, "FAIL"),
        ("Pre-deal holder value equals EV plus existing cash less costs", "=Deal!$D$33-Deal!$D$34", 1e-6, "FAIL"),
        ("Cap table sums to 100%", "=Deal!$E$31-1", 1e-9, "FAIL"),
        ("LTM calibration is profitable, as reported", "=IF(Triangulation!$D$13>0,0,1)", 0, "FAIL"),
        ("LTM calibration headcount within 5% of reported", f"=IF(ABS(Triangulation!$D$15/{sheet.scalar['headcount']}-1)<=0.05,0,1)", 0, "FAIL"),
    ]
    for name, m in maps.items():
        rows = m["rows"]
        checks += [
            (f"{name}: France ARR roll-forward", f"={name}!K{rows['arr_fr']}-({name}!E{rows['arr_fr']}-SUM({name}!F{rows['arr_fr_churn']}:K{rows['arr_fr_churn']})+SUM({name}!F{rows['arr_fr_upsell']}:K{rows['arr_fr_upsell']})+SUM({name}!F{rows['arr_fr_new']}:K{rows['arr_fr_new']}))", 1e-6, "FAIL"),
            (f"{name}: cash roll-forward", f"={name}!K{rows['cash']}-({name}!E{rows['cash']}+SUM({name}!F{rows['pre_fin_cf']}:K{rows['pre_fin_cf']})+SUM({name}!F{rows['draw']}:K{rows['draw']})+SUM({name}!F{rows['gap']}:K{rows['gap']})-SUM({name}!F{rows['repay']}:K{rows['repay']}))", 1e-6, "FAIL"),
            (f"{name}: cash never below minimum", f"=IF(MIN({name}!F{rows['cash']}:K{rows['cash']})>={sheet.scalar['min_cash']}-0.000001,0,1)", 0, "FAIL"),
            (f"{name}: line never above commitment", f"=IF(MAX({name}!F{rows['line']}:K{rows['line']})<={sheet.scalar['line_commitment']}+0.000001,0,1)", 0, "FAIL"),
            (f"{name}: tax losses never negative", f"=IF(MIN({name}!F{rows['nol']}:K{rows['nol']})>=-0.000001,0,1)", 0, "FAIL"),
            (f"{name}: equity bridge ties to exit waterfall", f"={name}!{m['bridge']['check']}", 1e-6, "FAIL"),
        ]
    plan = maps["Plan"]["rows"]
    b_moic = f"Base!J{maps['Base']['rows']['x_moic']}"
    checks += [
        ("Plan reaches the reported EUR 200M ARR by Dec-30", f"=IF(Plan!J{plan['arr_total']}>=200,0,1)", 0, "REVIEW"),
        ("Live grid E centre equals live Base MOIC", f"=Sensitivities!{sens['center_d']}-{b_moic}", 1e-6, "FAIL"),
        ("Static grid A centre equals live Base MOIC", f"=Sensitivities!{sens['center_a']}-{b_moic}", 1e-6, "STALE"),
        ("Static grid B centre equals live Base MOIC", f"=Sensitivities!{sens['center_b']}-{b_moic}", 1e-6, "STALE"),
    ]
    r = 5
    for label, formula, tol, flag in checks:
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", formula, fmt='0.000000;-0.000000;0')
        put(ws, f"D{r}", tol, fmt="0.0E+00", color=BLUE_INPUT)
        put(ws, f"E{r}", f'=IF(ABS(C{r})<=D{r},"OK","{flag}")', align="right")
        r += 1
    put(ws, f"B{r+1}", "Failures", bold=True)
    put(ws, f"E{r+1}", f'=COUNTIF(E5:E{r-1},"FAIL")', fmt="0", bold=True)
    put(ws, f"B{r+2}", "Static grids out of date", bold=True)
    put(ws, f"E{r+2}", f'=COUNTIF(E5:E{r-1},"STALE")', fmt="0", bold=True)
    finish(ws, {"A": 2, "B": 66, "C": 13, "D": 11, "E": 9}, tab="2E7D32")
    return f"Checks!$E${r+1}", f"Checks!$E${r+2}", f"COUNTA(Checks!$B$5:$B${r-1})"


# ---------------------------------------------------------------- Evidence
def build_evidence(wb):
    ws = wb.create_sheet("Evidence")
    put(ws, "B1", "Evidence ledger", bold=True, size=12)
    put(ws, "B2", "Every external figure, its status and where it comes from. 'Conflict' marks points where public sources disagree.", italic=True, color="555555")
    headers = ["ID", "Item", "Value", "Status", "Source", "Date", "Note", "URL"]
    for col, h in zip("BCDEFGHI", headers):
        put(ws, f"{col}4", h, bold=True, fill=MID)
    with open(ROOT / "data" / "evidence.csv", newline="") as f:
        for i, row in enumerate(csv.DictReader(f)):
            r = 5 + i
            for col, key in zip("BCDEFGH", ["id", "item", "value", "status", "source", "date", "note"]):
                put(ws, f"{col}{r}", row[key], wrap=True)
            c = put(ws, f"I{r}", row["url"], color="1F4E9E")
            c.hyperlink = row["url"]
    finish(ws, {"A": 2, "B": 5, "C": 28, "D": 48, "E": 15, "F": 22, "G": 10, "H": 42, "I": 60}, freeze="C5", tab="7A7F8C")


# ---------------------------------------------------------------- Summary
def build_summary(ws, maps, deal_refs, sheet: InputsSheet, nb, fail_ref, stale_ref, count_expr):
    s = sheet.scalar
    put(ws, "B1", "Skello: what BDC V has to believe", bold=True, size=14)
    put(ws, "B2", "Outside-in LBO model. Public sources only; no access to company or Bridgepoint data. A student case study by Pierre Grall, not affiliated with Skello or Bridgepoint.", italic=True, color="555555")
    put(ws, "B3", f'=IF({fail_ref}=0,{count_expr}&" mechanical checks run, none failed"&IF({stale_ref}>0,"; static grids out of date: "&{stale_ref},""),"CHECK FAILURES: "&{fail_ref})', bold=True, color="006100")

    band(ws, 5, "Entry", last="G")
    entry = [
        ("Post-money equity value (CFNEWS estimate)", deal_refs["post_money"], NUM), ("Enterprise value (needs my primary and existing cash assumptions)", deal_refs["ev_entry"], NUM),
        ("EV / ARR", deal_refs["m_entry"], MULT), ("EV / LTM revenue", deal_refs["ev_ltm"], MULT),
        ("Unitranche / ARR", deal_refs["debt_arr"], MULT), ("Bridgepoint stake", deal_refs["bp_share"], PCT),
    ]
    for i, (label, ref, fmt) in enumerate(entry):
        put(ws, f"B{6+i}", label)
        put(ws, f"E{6+i}", f"={ref}", fmt=fmt, color=GREEN_LINK)

    r = 13
    band(ws, r, "Scenarios, exit Dec-30 unless stated", last="G")
    r += 1
    names = ["Downside", "Base", "Plan"]
    for col, n in zip("EFG", names):
        put(ws, f"{col}{r}", n, bold=True, align="right")
    r += 1
    lines = [
        ("ARR Dec-28", "H", "arr_total", NUM), ("ARR Dec-30", "J", "arr_total", NUM), ("  of which acquired", "J", "arr_acq", NUM),
        ("Organic ARR growth FY30", "J", "org_growth", PCT), ("Revenue FY30", "J", "revenue", NUM), ("EBITDA FY30", "J", "ebitda", NUM),
        ("EBITDA margin FY30", "J", "ebitda_margin", PCT), ("Rule of 40 FY30 (points, organic)", "J", "r40", NUM),
        ("Share of French frontline workers", "J", "fr_penetration", PCT),
        ("New equity required", "J", "x_follow_on", NUM), ("Exit EV / LTM revenue (curve)", "J", "x_mult", MULT), ("Exit enterprise value", "J", "x_ev", NUM),
        ("Implied EV / ARR", "J", "x_ev_arr", MULT), ("Implied EV / EBITDA", "J", "x_ev_ebitda", MULT1),
        ("Net debt / (net cash)", "J", "net_debt", NUM), ("Equity value", "J", "x_equity", NUM),
        ("Bridgepoint proceeds", "J", "x_bp", NUM), ("Bridgepoint capital invested", "J", "x_bp_inv", NUM),
    ]
    for label, cx, key, fmt in lines:
        put(ws, f"B{r}", label)
        for col, n in zip("EFG", names):
            put(ws, f"{col}{r}", f"={n}!{cx}{maps[n]['rows'][key]}", fmt=fmt, color=GREEN_LINK)
        r += 1
    put(ws, f"B{r}", "MOIC, ordinary shares", bold=True)
    for col, n in zip("EFG", names):
        put(ws, f"{col}{r}", f"={n}!J{maps[n]['rows']['x_moic']}", fmt=MULT, bold=True, color=GREEN_LINK, fill=GREY)
    r += 1
    put(ws, f"B{r}", "IRR, ordinary shares", bold=True)
    for col, n in zip("EFG", names):
        put(ws, f"{col}{r}", f"={n}!{maps[n]['irr']['Dec-30']}", fmt=PCT, bold=True, color=GREEN_LINK, fill=GREY)
    r += 1
    put(ws, f"B{r}", "MOIC with a 1x senior preference on new money")
    for col, n in zip("EFG", names):
        put(ws, f"{col}{r}", f"={n}!J{maps[n]['rows']['x_moic_pref']}", fmt=MULT, color=GREEN_LINK)
    r += 1
    put(ws, f"B{r}", "MOIC with a 1x senior preference on Bridgepoint's cheque only")
    for col, n in zip("EFG", names):
        put(ws, f"{col}{r}", f"={n}!J{maps[n]['rows']['x_moic_pref_bp']}", fmt=MULT, color=GREEN_LINK)
    r += 1
    put(ws, f"B{r}", "Most a financial buyer could pay (EV)")
    for col, n in zip("EFG", names):
        put(ws, f"{col}{r}", f"=NextBuyer!{nb[n]['ev']}", fmt=NUM, color=GREEN_LINK)
    r += 1
    put(ws, f"B{r}", "MOIC at that price")
    for col, n in zip("EFG", names):
        put(ws, f"{col}{r}", f"=NextBuyer!{nb[n]['moic']}", fmt=MULT, color=GREEN_LINK)
    r += 1
    for lab, key_col in (("Dec-29", "I"), ("Dec-31", "K")):
        put(ws, f"B{r}", f"MOIC, exit {lab}")
        for col, n in zip("EFG", names):
            put(ws, f"{col}{r}", f"={n}!{key_col}{maps[n]['rows']['x_moic']}", fmt=MULT, color=GREEN_LINK)
        r += 1
        put(ws, f"B{r}", f"IRR, exit {lab}")
        for col, n in zip("EFG", names):
            put(ws, f"{col}{r}", f"={n}!{maps[n]['irr'][lab]}", fmt=PCT, color=GREEN_LINK)
        r += 1
    put(ws, f"B{r}", "My reference: gross MOIC / IRR (assumption, Evidence E19)")
    put(ws, f"F{r}", f"={s['hurdle_moic']}", fmt=MULT, color=GREEN_LINK)
    put(ws, f"G{r}", f"={s['hurdle_irr']}", fmt=PCT, color=GREEN_LINK)
    r += 2

    band(ws, r, "Base case: where Dec-30 equity value comes from (100% of equity)", last="G")
    r += 1
    labels = [("entry_equity", "Post-money equity at closing"), ("organic", "Organic ARR growth at entry multiple"), ("acquired", "Acquired ARR at entry multiple"),
              ("multiple", "Change in EV / ARR on exit ARR"), ("net_debt", "Change in net debt (build-up spend, net of cash generated)"), ("costs", "Exit costs"), ("exit_equity", "Equity value at Dec-30")]
    for key, label in labels:
        put(ws, f"B{r}", label, bold=key == "exit_equity")
        put(ws, f"F{r}", f"=Base!{maps['Base']['bridge'][key]}", fmt=NUM, color=GREEN_LINK, bold=key == "exit_equity")
        r += 1
    r += 1
    band(ws, r, "How to read this workbook", last="G")
    r += 1
    guide = [
        "Inputs holds every assumption. Blue cells can be changed; shaded cells have no public evidence and are what diligence would replace.",
        "Deal rebuilds sources and uses, the cap table and the entry multiple. Triangulation tests the assumptions against public facts.",
        "Base, Downside and Plan are formula-identical calculation tabs. NextBuyer tests the exit. Analytics supports the memo's numbers.",
        "Colours: blue = input, black = formula, green = link to another tab, grey = static output of a script. EUR millions. No macros, no circular references.",
        "No balance sheet or covenant test: terms are undisclosed and cash, debt and tax losses are rolled forward directly.",
    ]
    for line in guide:
        put(ws, f"B{r}", line)
        r += 1
    finish(ws, {"A": 2, "B": 54, "C": 2, "D": 2, "E": 11, "F": 11, "G": 11, "H": 2}, landscape=False, tab="F2C94C")


def main():
    inp = load_inputs()
    analysis = json.loads(ANALYSIS.read_text())
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    sheet = build_inputs(wb, inp)

    # The Rule of 40 curve sits at the foot of Inputs.
    ws = sheet.ws
    r = sheet.row
    band(ws, r, "Rule of 40 curve: EV / Revenue (Evidence E23 band medians at band midpoints, joined by straight lines)")
    r += 1
    t = inp["r40_curve"]
    put(ws, f"B{r}", "Rule of 40 score (band midpoint)")
    put(ws, f"B{r+1}", "Median EV / Revenue")
    for col, a, b in zip(PERIOD_COLS, t["x"], t["y"]):
        put(ws, f"{col}{r}", a, fmt="0", color=BLUE_INPUT)
        put(ws, f"{col}{r+1}", b, fmt=MULT1, color=BLUE_INPUT)
    sheet.curve_x, sheet.curve_y = r, r + 1

    deal_refs = build_deal(wb, sheet)
    build_triangulation(wb, sheet)
    maps = {}
    for i, name in enumerate(inp["scenarios"]):
        maps[name] = build_model_tab(wb, name, i + 1, sheet, deal_refs, inp)
    nb = build_next_buyer(wb, maps, sheet, inp)
    analytics = build_analytics(wb, maps, analysis, deal_refs, sheet)
    sens = build_sensitivities(wb, analysis, maps, sheet)
    fail_ref, stale_ref, count_expr = build_checks(wb, maps, sens, deal_refs, sheet)
    build_evidence(wb)
    build_summary(summary, maps, deal_refs, sheet, nb, fail_ref, stale_ref, count_expr)

    OUT_XLSX.parent.mkdir(exist_ok=True)
    wb.save(OUT_XLSX)
    maps["_next_buyer"] = nb
    maps["_analytics"] = analytics
    CELL_MAP.write_text(json.dumps(maps, indent=1))
    print("saved", OUT_XLSX)


if __name__ == "__main__":
    main()
