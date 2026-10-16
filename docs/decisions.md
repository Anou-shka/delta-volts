# Analyst decision log

I make the investment calls; the code implements them. This log records each decision that shapes the model, when it was made, and why. Assumptions that move the valuation by more than ~5% are decided here before they're built.

| # | Date | Decision | Why | Effect on the model |
|---|---|---|---|---|
| D1 | 16 Oct 2026 | **Anchor the AI data-centre revenue line only on Delta's own disclosures.** | The widely quoted split (AI power supplies >20% and AI cooling ~9% of FY2025 revenue) comes from press coverage of the 4Q25 call. It isn't in Delta's slides, and Delta published no 4Q25 transcript. A thesis built on an unverifiable figure won't hold up under questioning. | AI-line anchors (all from the 2Q26 meeting transcript):<br>• AI-related products **>25% of FY2026 revenue** (p.12);<br>• liquid cooling **~10% of FY2025 revenue** and **>12% of FY2026** (p.12);<br>• data-centre business **>50% of 1H26 revenue** (p.2).<br>The press figures stay in `sources.csv` as `TODO-SOURCE` and aren't used. |
| D2 | 16 Oct 2026 | **Wait for Q3 2026 results before locking the base period.** | Q3 results are expected around 29–30 Oct 2026. They will be the latest reviewed statements, and they add a quarter of AI-driven growth to the base. Locking now would mean revising every forecast two weeks later. | The base period moves from 1H26 to 9M26 when the Q3 2026 report is published. Until then, Jul–Sep monthly revenue (NT$197.1bn, unaudited) is the Q3 cross-check. The historical layer (M1) is final for FY2021–FY2025 and 1H26. |
| D3 | 16 Oct 2026 | Operating EBITDA (EBIT + D&A) for valuation, not Delta's presented EBITDA. | Delta's figure (PBT + finance costs + D&A) includes non-operating gains, which don't belong in an EV multiple. | FY2025 EBITDA NT$111.8bn (model) vs NT$117.9bn (Delta's slides). Both are shown. |
| D4 | 16 Oct 2026 | Capex = cash paid for PPE + intangibles (cash flow statement). | It's audited and consistent across years. Management (NT$46.6bn) and press (NT$46.1bn) figures use undisclosed definitions. | FY2025 capex NT$47.2bn. |
| D5 | 16 Oct 2026 | Segment history on the latest restated basis from FY2023; FY2021–FY2022 on the old 3-segment basis. | Delta carved Mobility out of Power Electronics in the FY2024 report and reclassified products again in FY2025. Restating earlier years myself would mean inventing an allocation. | Segment growth rates before FY2023 are shown for Power Electronics including Mobility. |

## Pending

- Review of the 10-figure sample in [`m1-review-sample.md`](m1-review-sample.md).
- Q3 2026 results (D2): add the Q3 report to `fetch.py`, extend the extractor's period map, and rerun.
