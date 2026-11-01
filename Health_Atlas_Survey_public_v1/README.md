# Health Atlas of Bangladesh: patient cost survey (v1: wave 1)

This branch holds v1: 59 inpatients at NIKDU (National Institute of Kidney Diseases & Urology, Dhaka), interviewed in September 2026 on the paediatric, male and female urology and nephrology wards. The `Survey-V2` branch adds 40 outpatients.

Student research on what illness costs households in Bangladesh: out-of-pocket medical spending, catastrophic spending, how families pay, and how the path to care affects cost.

**This repository holds a de-identified dataset, the method and aggregate results.** The original survey forms are confidential and are not included. Respondents were told no identifying information would be recorded or shared.

This is a convenience sample at a public tertiary referral hospital. It describes the high-cost end of health spending in Bangladesh, not the average household.

## Headline results (n = 59)

Medical out-of-pocket cost covers fees, tests, medicines and bed/procedure charges for one illness, from first care to the interview. It includes care before NIKDU. Transport and food are excluded, as in SDG 3.8.2.

| Figure | Value |
|---|---|
| Median medical cost | ৳88,350 (58 usable forms) |
| Diagnostics share of medical cost (itemised forms) | 61% (25 forms) |
| Catastrophic: over 10% of a year's household spending within 12 months | 93% (57 forms with budget data) |
| Catastrophic at the 25% threshold | 54% |
| Borrowed or sold assets to pay | 78% (58) |
| Delayed care because of cost | 44% (59) |

More indicators are in [`summary_tables_v1.csv`](summary_tables_v1.csv). Any figure based on fewer than 5 patients is suppressed.

## Files

| File | What it is |
|---|---|
| `survey_public_v1.csv` | **De-identified data: 59 rows, one per patient.** |
| `public_codebook_v1.md` | What every column means, and how it was de-identified. |
| `survey_public_analysis_v1.do` | Stata analysis that runs directly on the public file. |
| `summary_tables_v1.csv` | Aggregate results with n for every figure. |
| `public_check_v1.txt` | Privacy check result and a comparison of public vs exact results. |
| `make_public_dataset.py` | The code that made the public file from the confidential data (needs the confidential files to run). |
| `questionnaire.md` | The blank survey instrument. |

## Method

- **Money:** BDT. "k" = 1,000; "lakh" = 100,000. Ranges are coded at the midpoint. Rates (per day/week/month/trip) are multiplied out.
- **Medical cost** = the sum of the four medical items, or the respondent's stated all-inclusive medical total if that is larger or is the only figure given.
- **Catastrophic spending:** medical cost falling in one 12-month window, divided by 12 × monthly household spending (bracket midpoint). Thresholds: 10% and 25%.
- **Cost shares** use fully itemised forms only.

## Data availability

`survey_public_v1.csv` is a de-identified version of the data:
- Serial numbers, all free text, exact age, household size and occupation are removed.
- Household district, division and region are kept.
- Age, residence and spending use the survey form's own categories (18 or older; city/town/village; the five spending brackets).
- Costs are exact, as coded from the forms.
- District is kept for all 59 patients. On the other traits (sex, 18+, region, residence), every patient matches at least 3 others; a few residence, sex or 18+ values are blanked to achieve this.

Results from the public file are identical to the published figures.

## Disclosure

The physician at the private cardiology practice planned for a later wave is related to the author. This will be stated in every report that uses those data.
