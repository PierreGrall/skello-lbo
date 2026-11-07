# Skello: what BDC V has to believe

An outside-in look at Bridgepoint Development Capital's July 2026 investment in Skello, a French workforce-management SaaS. It asks one question: on public information, what would an investor have to believe for the deal to work at the reported price?

**Student case study from public sources only, including Skello's filed 2024 accounts.** No access to Skello or Bridgepoint data, and no affiliation with either. Reported figures are sourced in the evidence ledger. Everything else is my estimate.

## The short version

- From public sources alone, the deal does not reach a 3.0x reference at about €350M post-money on ordinary shares. My Base case returns about 2.1x and 18% at a December 2030 exit, and my Downside about 0.5x.
- That gap measures what an outsider cannot see. Two answers would matter most in diligence: French gross retention and the instrument. Retention near 90% with a 1x senior preference on the new money gives about 2.5x.
- 2.5x on price alone needs about €294M post-money, below the range the press allows.
- The exit is the weak point: Base's €804M is about 26x EBITDA, while a sponsor underwriting a 15x exit could pay about €658M.
- Debt adds about one point of IRR. Returns come from growth and the exit price.
- Checked against Skello's 2024 filed accounts: revenue up 41% to €25.4M; reported EBITDA margin -10% (-18% with development expensed); deferred revenue at 35% of revenue; doubtful receivables at 13% of revenue; the Spanish subsidiary lost €2.8M on €3.5M of revenue.

BDC V signed with diligence I do not have. The memo shows how far each unknown would have to move.

## What is here

| Path | What it is |
|---|---|
| `memo/Skello_Investment_Memo.pdf` | The memo. Start here. |
| `model/Skello_LBO_Model.xlsx` | The model, in live formulas: inputs, deal, three scenarios, next-buyer test, analytics, evidence |
| `data/inputs.json`, `data/evidence.csv` | Every assumption, and every external figure with its source |
| `src/engine.py` | A second implementation of the model in Python, reconciled to the workbook |
| `src/analysis.py` | Sensitivities, break-even solves and attribution |
| `src/build_model.py`, `src/build_memo.py`, `src/charts.py` | Generate the workbook, memo and charts from the inputs |
| `tests/test_model.py` | Mechanical tests, including a stress test of the workbook's own input cells |

## Rebuilding

Needs Python 3.11+, LibreOffice and XeLaTeX with TeX Gyre Pagella and Inter.

```
pip install -r requirements.txt
python run_all.py
```

The build stops if the workbook and the Python engine disagree, or if the memo contains a number with no source.

## What the checks do and do not prove

They show the file is internally consistent. They do not show the assumptions are right, and an error of reasoning shared by both implementations would pass. The main limits are stated in the memo: retention, primary cash, the exit benchmark and the instrument are all unknown from outside.

## Tooling

Built with AI assistance for code and drafting. All assumptions and conclusions are mine to defend.
