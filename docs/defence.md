# Interview defence sheet

33 questions an interviewer could ask about this initiation, grouped by topic. **The answers are mine to write**; the pointer under each question says where the evidence is in the repo.

Tip: answer with a number, a source and a judgement, in that order.

## Business

**1. Walk me through what Delta Electronics actually sells, and to whom.**

<sub>Evidence: Company overview; FY2025 segment note</sub>

> *My answer:*
>
>
>

**2. Why did Infrastructure profit go from NT$5.5bn to NT$28.4bn in one year?**

<sub>Evidence: Segments sheet; D5, D7</sub>

> *My answer:*
>
>
>

**3. Delta sells 36% of its revenue in the US but has almost no assets there. How do you think about that?**

<sub>Evidence: Geography chart; Risks</sub>

> *My answer:*
>
>
>

**4. What is Mobility doing in this company, and would you keep it?**

<sub>Evidence: Segments sheet (margins); M3 notes</sub>

> *My answer:*
>
>
>

**5. No customer is above 10% of revenue. Does that mean customer concentration isn't a risk?**

<sub>Evidence: Risks table; 3Q26 transcript p.7</sub>

> *My answer:*
>
>
>

## AI thesis

**6. How much of Delta is an AI bet, and how do you know?**

<sub>Evidence: D1, D7; Segments sheet disclosures table</sub>

> *My answer:*
>
>
>

**7. Why didn't you use the widely quoted '20% AI power, 9% cooling' split?**

<sub>Evidence: D1; sources.csv TODO-SOURCE rows</sub>

> *My answer:*
>
>
>

**8. Why does a 1 MW rack matter more to Delta than to a chip designer?**

<sub>Evidence: Industry section; NVIDIA citations</sub>

> *My answer:*
>
>
>

**9. What would make AI data-centre demand fall in 2028, and what would you watch to see it coming?**

<sub>Evidence: Key debate; Bear case</sub>

> *My answer:*
>
>
>

**10. Management talks about HVDC and solid-state transformers. How much of that is in your numbers?**

<sub>Evidence: assumptions.csv content gains; 2Q26 transcript</sub>

> *My answer:*
>
>
>

## Forecasts

**11. Your FY2027 EPS is well below consensus. Defend the gap, or tell me why you'd close it.**

<sub>Evidence: M4 notes; Comps sheet consensus</sub>

> *My answer:*
>
>
>

**12. Why build the AI line from hyperscaler capex times a content gain rather than from rack shipments?**

<sub>Evidence: D8</sub>

> *My answer:*
>
>
>

**13. How did you get from nine months of actuals to a full-year FY2026 number?**

<sub>Evidence: D6; Segments FY2026E formula</sub>

> *My answer:*
>
>
>

**14. Why is SG&A a balancing item? Doesn't that hide cost assumptions?**

<sub>Evidence: D9; Checks (SG&A range)</sub>

> *My answer:*
>
>
>

**15. Where does the shift to 3-year depreciation for AI equipment show up in your model?**

<sub>Evidence: D9; Schedules (PP&E)</sub>

> *My answer:*
>
>
>

**16. Payable days went from 69 to 95. What did you assume and why?**

<sub>Evidence: Schedules; assumptions.csv dpo</sub>

> *My answer:*
>
>
>

## Valuation

**17. Your DCF says NT$675 and the stock is NT$990 after halving. Why isn't that a Buy?**

<sub>Evidence: Reverse DCF; football field</sub>

> *My answer:*
>
>
>

**18. Why cap terminal growth at the NT$ risk-free rate?**

<sub>Evidence: D14; valuation_inputs.csv</sub>

> *My answer:*
>
>
>

**19. The Gordon value implies 7.3x EBITDA and peers trade at 13x. Which do you believe?**

<sub>Evidence: D14; DCF sheet cross-check</sub>

> *My answer:*
>
>
>

**20. Why a bottom-up beta of 1.48 instead of the 1.73 regression beta?**

<sub>Evidence: D12</sub>

> *My answer:*
>
>
>

**21. How do you treat Delta Thailand's minority interest in the equity bridge, and why not at book?**

<sub>Evidence: D13</sub>

> *My answer:*
>
>
>

**22. Which peers are fair comparisons and which aren't?**

<sub>Evidence: Comps sheet notes; D15</sub>

> *My answer:*
>
>
>

**23. What does the reverse DCF tell you, and what are its limits?**

<sub>Evidence: Sensitivity sheet section 3; D17</sub>

> *My answer:*
>
>
>

## Risks

**24. Which single risk would change your rating, and what's the indicator?**

<sub>Evidence: Risks table</sub>

> *My answer:*
>
>
>

**25. How exposed is Delta to US tariffs on Thai and Taiwanese electronics?**

<sub>Evidence: Geography; Risks</sub>

> *My answer:*
>
>
>

**26. Q3 gross margin fell to 34.1%. Is that the start of price competition or a one-off?**

<sub>Evidence: 3Q26 transcript p.4-5</sub>

> *My answer:*
>
>
>

**27. What happens to your numbers if the NT dollar strengthens 10% against the US dollar?**

<sub>Evidence: Risks (FX); 88% of revenue outside Taiwan</sub>

> *My answer:*
>
>
>

## Model mechanics

**28. How do you know your historical numbers are right?**

<sub>Evidence: M1 notes; 440 checks; review sample</sub>

> *My answer:*
>
>
>

**29. Why does your balance-sheet check prove anything if you built the balance sheet from reported numbers?**

<sub>Evidence: M2 notes</sub>

> *My answer:*
>
>
>

**30. How do you make sure the Excel and the Python never disagree?**

<sub>Evidence: tests/test_model_matches_excel.py</sub>

> *My answer:*
>
>
>

**31. How did you handle Delta's segment restatements?**

<sub>Evidence: D5</sub>

> *My answer:*
>
>
>

**32. If I change one assumption, what has to happen for the report to update?**

<sub>Evidence: README 'Run it'; D18</sub>

> *My answer:*
>
>
>

**33. What did you use AI tools for in this project, and what did you decide yourself?**

<sub>Evidence: Decision log D1-D18</sub>

> *My answer:*
>
>
>
