* Health Atlas of Bangladesh - patient cost survey analysis
* v2 (waves 1 + 2, n = 99). Generated 1 Nov 2026. Put this file next to survey_clean_v2.csv and run: do survey_analysis_v2.do
* CONFIDENTIAL row-level data: never commit the CSV or this log output with individual rows to a public place.
clear all
set more off
import delimited using "survey_clean_v2.csv", clear varnames(1) encoding("utf-8") stringcols(2)

* ---- variable labels
label variable id "Row id"
label variable serial "Serial number on paper form"
label variable wave "Survey wave"
label variable hospital "Hospital"
label variable ward "Ward / department"
label variable specialty "Specialty"
label variable paediatric "Paediatric patient (<18: paediatric ward 1-2, or age <18 for other wards)"
label variable visit_type "Visit type"
label variable interview_date "Interview date"
label variable sex "Patient sex"
label variable age "Patient age"
label variable resp_patient "Respondent is the patient"
label variable resp_relation "Attendant relation to patient"
label variable district "Household district"
label variable division "Household division"
label variable residence "Residence type"
label variable hh_size "Household size"
label variable occupation "Main earner occupation (as said)"
label variable occ_cat "Main earner occupation category"
label variable spend_bracket "Monthly household spending bracket"
label variable spend_exact "Monthly spending if an exact figure was given"
label variable monthly_cons "Monthly consumption used for shares"
label variable cond_text "Condition (as recorded)"
label variable cond_cat "Condition category"
label variable dialysis "Dialysis during episode"
label variable nights_current "Nights in current admission"
label variable nights_costed "Inpatient nights covered by the reported costs"
label variable onset_days "Days since symptoms began"
label variable chronic "Symptoms for 90+ days"
label variable care_months "Months over which the reported costs were spent (first care for this disease -> "
label variable first_care "First point of care (C1)"
label variable first_informal "First care at pharmacy or informal provider (H3)"
label variable first_private "First care at private clinic/hospital"
label variable first_direct "Came to this hospital first"
label variable care_path "Care path (as recorded)"
label variable n_providers "Providers seen before this hospital (C2)"
label variable reason "Why this hospital (C3)"
label variable referred "Referred"
label variable premed "Took medicine before a prescription (C4)"
label variable premed_antibiotic "Pre-prescription antibiotic"
label variable premed_painkiller "Pre-prescription painkiller"
label variable premed_other "Pre-prescription other medicine"
label variable cost_fees "D1 Doctor/consultation fees"
label variable cost_diag "D2 Diagnostics & imaging"
label variable cost_med "D3 Medicines"
label variable cost_bed "D4 Bed, procedure, operation (incl. dialysis sessions, blood, ICU)"
label variable cost_items_sum "Sum of reported medical items D1-D4"
label variable cost_reported_total "All-inclusive medical total stated by respondent"
label variable oop_medical "Medical out-of-pocket spending (analysis variable)"
label variable cost_transport "D5 Transport, all trips"
label variable cost_food "D6 Food & lodging for attendants"
label variable oop_nonmedical "Transport + food"
label variable oop_total "Total economic cost (medical + non-medical)"
label variable informal_pay "D7 Unbilled payment reported (tips/staff/brokers)"
label variable scheme_paid "D8 Insurance/employer/NGO/govt scheme paid"
label variable fin_savings "E1 Paid from savings"
label variable fin_borrow_rel "E1 Borrowed from relatives/friends"
label variable fin_borrow_lender "E1 Borrowed from lender/NGO/bank"
label variable fin_borrow_unspec "E1 Borrowed, source not recorded"
label variable fin_sold_asset "E1 Sold/pawned/mortgaged asset (incl. land)"
label variable fin_abroad "E1 Help from relatives abroad"
label variable fin_income "E1 Paid from current income"
label variable fin_distress "Any borrowing or asset sale"
label variable days_lost "E2 Work days lost (patient + attendants)"
label variable job_loss "Someone lost their job during the illness"
label variable delay "E3 Delayed coming to hospital"
label variable delay_reason "Reason for delay"
label variable delay_cost "Delay attributed to cost (funds or unspecified)"
label variable fewer_meds "E3 Bought fewer medicines than prescribed"
label variable skipped_test "E3 Skipped a test"
label variable left_early "E3 Left earlier than advised"
label variable any_cost_coping "Any of E3 due to cost"
label variable next_choice "E4 Where next time"
label variable next_reason "E4 Why"
label variable cost_scope "Length of the costed illness"
label variable cost_quality "Cost data quality"
label variable usable_cost "Usable for cost analysis"
label variable form_complete "Completion as ticked on form"
label variable oop_per_month "Medical OOP per month of care"
label variable oop_12m "Medical OOP falling in one 12-month window"
label variable oop_share "12-month medical OOP / annual household consumption (SDG 3.8.2 style)"
label variable cat10 "Catastrophic: 12-month OOP > 10% of annual consumption"
label variable cat25 "Catastrophic: 12-month OOP > 25% of annual consumption"
label variable share_illness "Whole-illness medical OOP / one year of consumption"
label variable cat25_illness "Whole-illness OOP > 25% of one year of consumption"
label variable ln_oop_medical "ln(medical OOP)"
label variable ln_oop_total "ln(total economic cost)"
label variable n_issues "Number of data issues logged"

* ---- value labels
label define wave_lbl 1 "W1 NIKDU inpatients" 2 "W2 NIKDU outpatients" 3 "W3 private cardiology", replace
label values wave wave_lbl
label define hospital_lbl 1 "NIKDU" 2 "NICVD" 3 "Private cardiology practice", replace
label values hospital hospital_lbl
label define ward_lbl 1 "Paediatric Urology" 2 "Paediatric Nephrology" 3 "Male Urology" 4 "Male Nephrology" 5 "Female Urology" 6 "Female Nephrology" 7 "NIKDU Outpatient" 8 "NICVD Cardiology (inpatient)" 9 "NICVD Cardiac Surgery (inpatient)" 10 "NICVD Outpatient" 11 "Private cardiology practice", replace
label values ward ward_lbl
label define visit_type_lbl 1 "Inpatient" 2 "Outpatient", replace
label values visit_type visit_type_lbl
label define sex_lbl 1 "Male" 2 "Female", replace
label values sex sex_lbl
label define residence_lbl 1 "City" 2 "Town" 3 "Village", replace
label values residence residence_lbl
label define occ_cat_lbl 1 "Agriculture/fishing" 2 "Wage/manual" 3 "Small business" 4 "Salaried/skilled" 5 "Driver/transport" 6 "Migrant abroad" 7 "Pension/govt" 8 "Other", replace
label values occ_cat occ_cat_lbl
label define spend_bracket_lbl 1 "<15k" 2 "15-30k" 3 "30-60k" 4 "60-100k" 5 ">100k", replace
label values spend_bracket spend_bracket_lbl
label define cond_cat_lbl 1 "Kidney failure-dialysis" 2 "CKD/kidney damage" 3 "Glomerular" 4 "Stone" 5 "Obstruction/hydronephrosis" 6 "Tumour/cancer" 7 "Infection" 8 "Other/undiagnosed" 9 "Heart (cardiology)" 10 "Cardiac surgery", replace
label values cond_cat cond_cat_lbl
label define first_care_lbl 1 "Pharmacy" 2 "Village doctor/informal" 3 "Community clinic" 4 "Govt hospital" 5 "Private clinic/hospital" 6 "This hospital directly" 7 "Nothing", replace
label values first_care first_care_lbl
label define reason_lbl 1 "Referred" 2 "Reputation" 3 "Closest" 4 "Cost" 5 "Doctor known" 6 "Own choice/other", replace
label values reason reason_lbl
label define delay_reason_lbl 1 "Funds" 2 "Transport" 3 "Other (non-cost)" 4 "Not specified", replace
label values delay_reason delay_reason_lbl
label define next_choice_lbl 1 "Same hospital" 2 "Other facility" 3 "Depends/doctor" 4 "Unsure" 9 "Missing", replace
label values next_choice next_choice_lbl
label define next_reason_lbl 1 "Quality" 2 "Familiar" 3 "Cheap/govt" 4 "Doctor advice" 5 "Close" 6 "Other", replace
label values next_reason next_reason_lbl
label define cost_scope_lbl 1 "Within 12 months" 2 "Longer than 12 months", replace
label values cost_scope cost_scope_lbl
label define cost_quality_lbl 1 "Itemised" 2 "Partial" 3 "Total only", replace
label values cost_quality cost_quality_lbl
label define specialty_lbl 1 "Urology" 2 "Nephrology" 3 "Cardiology", replace
label values specialty specialty_lbl

* ---- recompute derived variables (so newly entered rows only need the "entered" columns)
replace specialty = cond(inlist(ward,1,3,5),1, cond(inlist(ward,2,4,6),2, cond(ward>=8 & ward<=11,3, specialty)))
replace paediatric = inlist(ward,1,2) if !missing(ward) & ward<=6
replace paediatric = (age<18) if !missing(age) & ward>6
replace monthly_cons = cond(!missing(spend_exact), spend_exact, cond(spend_bracket==1,12000, cond(spend_bracket==2,22500, cond(spend_bracket==3,45000, cond(spend_bracket==4,80000, cond(spend_bracket==5,120000,.))))))
replace chronic = onset_days>=90 if !missing(onset_days)
replace first_informal = inlist(first_care,1,2) if !missing(first_care)
replace first_informal = . if missing(first_care)
replace first_private = cond(missing(first_care), ., first_care==5)
replace first_direct = cond(missing(first_care), ., first_care==6)
replace referred = reason==1 if !missing(reason)
egen _items = rowtotal(cost_fees cost_diag cost_med cost_bed), missing
replace cost_items_sum = _items
drop _items
egen _nonmed = rowtotal(cost_transport cost_food)
replace oop_nonmedical = _nonmed
drop _nonmed
replace oop_total = oop_medical + oop_nonmedical
replace fin_distress = (fin_borrow_rel==1 | fin_borrow_lender==1 | fin_borrow_unspec==1 | fin_sold_asset==1) if !missing(fin_savings)
replace delay_cost = cond(missing(delay), ., (delay==1 & inlist(delay_reason,1,4)))
replace any_cost_coping = cond(missing(delay), ., (delay_cost==1 | fewer_meds==1 | skipped_test==1 | left_early==1))
replace cost_scope = cond(care_months>12,2,1) if !missing(care_months)
replace oop_per_month = oop_medical/max(care_months,1)
replace oop_12m = oop_medical*min(1,12/care_months)
replace oop_share = oop_12m/(12*monthly_cons) if usable_cost==1 & !missing(monthly_cons)
replace share_illness = oop_medical/(12*monthly_cons) if usable_cost==1 & !missing(monthly_cons)
replace cat25_illness = share_illness>0.25 if !missing(share_illness)
replace cat10 = oop_share>0.10 if !missing(oop_share)
replace cat25 = oop_share>0.25 if !missing(oop_share)
replace ln_oop_medical = ln(oop_medical) if oop_medical>0 & !missing(oop_medical)
replace ln_oop_total = ln(oop_total) if oop_total>0 & !missing(oop_total)
gen byte uro = specialty==1
gen ln_care = ln(max(care_months,0.5))
gen ln_nights = ln(max(nights_current,1))
gen byte outpatient = visit_type==2

* ---- Phase 1 descriptives
tab ward visit_type
tab wave
tabstat oop_medical oop_total cost_transport if usable_cost==1, by(cost_scope) stat(n mean p50 p25 p75) format(%12.0fc)
tabstat oop_medical if usable_cost==1, by(ward) stat(n mean p50) format(%12.0fc)
* Inpatients vs outpatients
tabstat oop_medical oop_total if usable_cost==1, by(visit_type) stat(n mean p25 p50 p75) format(%12.0fc)
tabstat oop_share if usable_cost==1, by(visit_type) stat(n p25 p50 p75) format(%5.3f)
tabstat cat10 cat25 fin_distress delay_cost any_cost_coping, by(visit_type) stat(n mean) format(%5.3f)
ranksum oop_medical if usable_cost==1, by(visit_type)
tabstat oop_medical if usable_cost==1, by(cond_cat) stat(n p50) format(%12.0fc)
summarize cat10 cat25 if usable_cost==1
summarize cat10 cat25 cat25_illness if usable_cost==1 & cost_scope==1
tab first_care
summarize fin_* delay_cost any_cost_coping
* H2: shares of medical OOP (fully itemised forms)
preserve
keep if usable_cost==1 & cost_quality==1
collapse (sum) cost_fees cost_diag cost_med cost_bed cost_items_sum
foreach v in cost_fees cost_diag cost_med cost_bed {
    display "`v' share = " %5.3f `v'/cost_items_sum
}
restore

* ---- Phase 2: survey regressions (H3). Robust SEs.
local ctrl outpatient dialysis paediatric uro ln_care ln_nights spend_bracket
reg ln_oop_medical first_informal `ctrl' if usable_cost==1, vce(robust)
reg ln_oop_medical n_providers `ctrl' if usable_cost==1, vce(robust)
reg ln_oop_medical first_private `ctrl' if usable_cost==1, vce(robust)
reg ln_oop_medical premed `ctrl' if usable_cost==1, vce(robust)
* Outpatients only (wave 2)
reg ln_oop_medical n_providers dialysis paediatric uro ln_care spend_bracket if usable_cost==1 & outpatient==1, vce(robust)
* Sensitivity: drop the two extreme-cost forms (outliers A and B; serials listed in the private data notes)
reg ln_oop_medical n_providers `ctrl' if usable_cost==1 & !inlist(serial,"85","88"), vce(robust)
* After the private cardiology wave: add hospital dummy and condition fixed effects
* reg ln_oop_medical first_informal n_providers i.hospital i.visit_type i.cond_cat ln_nights i.spend_bracket ln_care if usable_cost==1, vce(robust)
* Robustness: single episodes only; total cost incl. transport; drop top 5% of OOP
* reg ln_oop_medical n_providers `ctrl' if usable_cost==1 & cost_scope==1, vce(robust)
* reg ln_oop_total n_providers `ctrl' if usable_cost==1, vce(robust)
* Catastrophic spending (linear probability model)
* reg cat25 n_providers first_informal `ctrl' if usable_cost==1, vce(robust)
