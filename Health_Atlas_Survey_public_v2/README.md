# Health Atlas of Bangladesh: patient cost survey (v2: waves 1 + 2)

This branch holds v2: 59 NIKDU inpatients + 40 NIKDU outpatients. The `Survey-V1` branch holds wave 1 only.

Student research on what illness costs households in Bangladesh: out-of-pocket medical spending, catastrophic spending, how families pay, and how the path to care affects cost.

**This repository holds the method, code and aggregate results only.** The patient-level data are confidential and are not included. Respondents were told no identifying information would be recorded or shared.

## Survey

| Wave | Setting | Forms | Period |
|---|---|---|---|
| 1 | NIKDU (National Institute of Kidney Diseases & Urology, Dhaka), inpatients: paediatric, male and female urology and nephrology wards | 59 | Sep 2026 |
| 2 | NIKDU, urology and nephrology outpatient clinics | 40 | Oct 2026 |
| 3 (planned) | Private cardiology practice (see disclosure below) | 50 | — |

These are convenience samples at a public tertiary referral hospital. They describe the high-cost end of health spending in Bangladesh, not the average household. For national or district rates, see HIES 2022.

The questionnaire is in [`questionnaire.md`](questionnaire.md).

## Headline results (waves 1 + 2, n = 99)

Medical out-of-pocket cost covers fees, tests, medicines and bed/procedure charges for one illness, from first care to the interview. It includes care before NIKDU. Transport and food are excluded, as in SDG 3.8.2.

| | All | Inpatients | Outpatients |
|---|---|---|---|
| Median medical cost (BDT) | 65,500 | 88,400 | 26,700 |
| Diagnostics share of medical cost (itemised forms) | 63% | 61% | 67% |
| Catastrophic: over 10% of a year's household spending within 12 months | 75% | 93% | 50% |
| Catastrophic at the 25% threshold | 41% | 54% | 23% |
| Borrowed or sold assets to pay | 61% | 78% | 38% |
| Delayed care because of cost | 42% | 44% | 38% |

Full tables are in [`summary_tables_v2.csv`](summary_tables_v2.csv). Any figure based on fewer than 5 patients is suppressed. The written results are in [`results_v2.md`](results_v2.md).

## Method

- **Money:** BDT. "k" = 1,000; "lakh" = 100,000. Ranges are coded at the midpoint. Rates (per day/week/month/trip) are multiplied out; where a duration was missing, the illness length was used and the assumption logged.
- **Medical cost** = the sum of the four medical items, or the respondent's stated all-inclusive medical total if that is larger or is the only figure given.
- **Catastrophic spending:** medical cost falling in one 12-month window, `oop_medical × min(1, 12 / months of care)`, divided by 12 × monthly household spending. Monthly spending is the exact figure if given, otherwise the bracket midpoint (<15k → 12,000; 15–30k → 22,500; 30–60k → 45,000; 60–100k → 80,000; >100k → 120,000). Thresholds: 10% and 25%.
- **Cost shares** use fully itemised forms only.
- **Regressions:** OLS on ln(medical cost) with robust (HC1) standard errors. Controls: outpatient, dialysis, under 18, urology vs nephrology, ln(months of care), ln(nights admitted), spending bracket.

## Files

| File | What it is |
|---|---|
| `build_workbook.py` | Builds the coded dataset, Excel workbook (codebook, cost audit trail, data-issue log, summaries, sensitivity checks, regressions) and Stata do-file from the hand-coded records. |
| `districts.py` | The 64 districts mapped to the 8 divisions. |
| `survey_records_TEMPLATE.py` | The format of the hand-coded records, with one fictional example. |
| `survey_analysis_v2.do` | Stata analysis: recomputes derived variables, then descriptives and regressions. |
| `questionnaire.md` | The blank survey instrument. |
| `summary_tables_v2.csv` | Aggregate results, all / inpatients / outpatients, with n for every cell. |
| `results_v2.md` | Written results. |
| `survey_public_v2.csv` | **De-identified data: 99 rows, one per patient** (see below). |
| `public_codebook_v2.md` | What every column in the public file means, and how it was de-identified. |
| `survey_public_analysis_v2.do` | Stata analysis that runs directly on the public file. |
| `make_public_dataset.py` | The code that made the public file from the confidential data. |
| `public_check_v2.txt` | Privacy check result and a comparison of public vs exact results. |

`build_workbook.py` and `make_public_dataset.py` need the confidential files (`survey_records.py`, `survey_records_wave2.py`, `care_months_wave1.py`) to run. They are excluded by `.gitignore`. `survey_public_analysis_v2.do` runs on the public file alone.

## Data availability

`survey_public_v2.csv` is a de-identified version of the data:
- Serial numbers, all free text, exact age, household size and occupation are removed.
- Household district, division and region are kept.
- Age, residence and spending use the survey form's own categories (18 or older; city/town/village; the five spending brackets).
- Costs are exact, as coded from the forms.
- District is kept for all 99 patients. On the other traits (sex, 18+, region, residence), every patient matches at least 3 others; a few residence, sex or 18+ values are blanked to achieve this.

Results from the public file are identical to the published figures. The original survey forms are not public.

## Disclosure

The physician at the private cardiology practice planned for wave 3 is related to the author. This will be stated in every report that uses wave 3 data.
