# Codebook: survey_public_v2.csv

De-identified data from the Health Atlas patient cost survey at NIKDU, Dhaka: 59 NIKDU inpatients (wave 1, Sep 2026) + 40 NIKDU outpatients (wave 2, Oct 2026). 99 rows, one per patient. Blank = missing.

## What was done to protect respondents

- Removed: serial numbers, interview notes and all free text (condition description, care path, occupation, attendant relation, answers as written), exact age, household size, occupation, exact spending, interview completion.
- Kept: household district, division and region.
- Bands from the survey form: patient 18 or older; city / town / village; the five monthly spending brackets; the form's answer lists.
- Bands set for this file (the form has none): nights admitted; months of care; work days lost; providers seen capped at 4+.
- Costs are exact, as coded from the forms (rules in the README). `oop_medical` can exceed the item sum when the respondent gave a larger all-inclusive total.
- District is kept on every row. Some districts have only one or two patients; see public_check for counts.
- On the other traits, every row matches at least 3 patients on visit type, sex, 18+, region and residence (a blank counts as matching anything). Where a patient would otherwise stand out, residence, 18+ or sex was blanked on that row: residence 12, adult 4, sex 2.
- Rows are in random order with new ids.

| Variable | Meaning | Codes / units |
|---|---|---|
| `id` | Public row id (random order; not the form serial number) | number |
| `wave` | Survey wave | 1 NIKDU inpatients (Sep 2026); 2 NIKDU outpatients (Oct 2026) |
| `visit_type` | Visit type | 1 inpatient; 2 outpatient |
| `specialty` | Department | 1 urology; 2 nephrology |
| `sex` | Patient sex | 1 male; 2 female; blank = suppressed for privacy |
| `adult` | Patient is 18 or older (form screening question) | 1 yes; 0 no (child); blank = missing or suppressed |
| `resp_patient` | Respondent is the patient | 1 patient; 0 attendant |
| `region` | Household region | 1 Dhaka division; 2 outside Dhaka division |
| `division` | Household division | name (8 divisions) |
| `district` | Household district (BBS spelling) | name (all rows) |
| `residence` | Residence type (form) | 1 city; 2 town; 3 village; blank = not answered or suppressed for privacy |
| `spend_bracket` | Monthly household spending (form brackets) | 1 under 15,000; 2 15,000-30,000; 3 30,000-60,000; 4 60,000-100,000; 5 above 100,000 BDT |
| `cond_cat` | Condition category | 1 kidney failure/dialysis; 2 CKD/kidney damage; 3 glomerular; 4 stone; 5 obstruction/hydronephrosis; 6 tumour/cancer; 7 infection; 8 other/undiagnosed |
| `dialysis` | Dialysis during the illness | 0/1 |
| `nights_band` | Nights in current admission | 0 none (outpatient); 1 1-7; 2 8-14; 3 15-30; 4 31+ |
| `chronic` | Symptoms for 90+ days | 0/1 |
| `care_months_band` | Months over which the reported costs were spent | 1 up to 1; 2 1-3; 3 3-6; 4 6-12; 5 12-24; 6 over 24 |
| `cost_scope` | Length of the costed illness | 1 within 12 months; 2 longer than 12 months |
| `first_care` | First point of care (form) | 1 pharmacy/drug seller; 2 village doctor/informal; 3 community clinic; 4 govt hospital; 5 private clinic/hospital; 6 this hospital directly; 7 nothing at first |
| `n_providers` | Providers seen before this hospital | 0-3; 4 = four or more |
| `reason` | Why this hospital (form) | 1 referred; 2 reputation; 3 closest; 4 cost; 5 doctor known; 6 own choice/other |
| `referred` | Referred | 0/1 |
| `premed` | Took medicine before a prescription | 0/1 |
| `premed_antibiotic` | ... antibiotic | 0/1 |
| `premed_painkiller` | ... painkiller | 0/1 |
| `premed_other` | ... other | 0/1 |
| `cost_fees` | Doctor/consultation fees | BDT, exact as coded |
| `cost_diag` | Diagnostic tests and imaging | BDT, exact as coded |
| `cost_med` | Medicines | BDT, exact as coded |
| `cost_bed` | Bed, procedure, operation, dialysis | BDT, exact as coded |
| `oop_medical` | Medical out-of-pocket cost (analysis variable) | BDT, exact as coded |
| `cost_transport` | Transport, all trips | BDT, exact as coded |
| `cost_food` | Food and lodging for attendants | BDT, exact as coded |
| `oop_total` | Medical + transport + food | BDT, exact as coded |
| `informal_pay` | Unbilled payment reported | 0/1 |
| `scheme_paid` | Insurance/employer/NGO/govt scheme paid | 0/1 |
| `fin_savings` | Paid from savings | 0/1 |
| `fin_borrow_rel` | Borrowed from relatives/friends | 0/1 |
| `fin_borrow_lender` | Borrowed from lender/NGO/bank | 0/1 |
| `fin_borrow_unspec` | Borrowed, source not recorded | 0/1 |
| `fin_sold_asset` | Sold/pawned an asset | 0/1 |
| `fin_abroad` | Help from relatives abroad | 0/1 |
| `fin_income` | Paid from current income | 0/1 |
| `fin_distress` | Any borrowing or asset sale | 0/1 |
| `days_lost_band` | Work days lost (patient + attendants) | 0 none; 1 1-30; 2 31-90; 3 91-365; 4 over 365 |
| `delay` | Delayed coming to hospital | 0/1 |
| `delay_reason` | Reason for delay | 1 funds; 2 transport; 3 other (non-cost); 4 not specified |
| `delay_cost` | Delay because of cost | 0/1 |
| `fewer_meds` | Bought fewer medicines than prescribed | 0/1 |
| `skipped_test` | Skipped a test | 0/1 |
| `left_early` | Left earlier than advised | 0/1 |
| `any_cost_coping` | Any of the four because of cost | 0/1 |
| `next_choice` | Where next time | 1 same hospital; 2 other facility; 3 depends/doctor; 4 unsure; 9 missing |
| `next_reason` | Why | 1 quality; 2 familiar; 3 cheap/govt; 4 doctor advice; 5 close; 6 other |
| `cost_quality` | Cost data quality | 1 fully itemised; 2 partly itemised/estimated; 3 total only |
| `usable_cost` | Usable for cost analysis | 0/1 |
| `oop_share` | 12-month medical cost / annual household spending | fraction |
| `cat10` | Catastrophic: share > 10% | 0/1 |
| `cat25` | Catastrophic: share > 25% | 0/1 |
| `cat25_illness` | Whole-illness cost > 25% of one year of spending | 0/1 |
