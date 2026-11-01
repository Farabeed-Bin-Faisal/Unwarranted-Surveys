* Health Atlas of Bangladesh - patient cost survey: analysis of the PUBLIC de-identified file
* Put this file next to survey_public_v2.csv and run: do survey_public_analysis_v2.do
* Codebook: public_codebook_v2.md. Costs are exact, so results match the published tables.
clear all
set more off
import delimited using "survey_public_v2.csv", clear varnames(1) encoding("utf-8")

* ---- value labels
label define wave_lbl 1 "W1 NIKDU inpatients" 2 "W2 NIKDU outpatients", replace
label values wave wave_lbl
label define visit_lbl 1 "Inpatient" 2 "Outpatient", replace
label values visit_type visit_lbl
label define spec_lbl 1 "Urology" 2 "Nephrology", replace
label values specialty spec_lbl
label define sex_lbl 1 "Male" 2 "Female", replace
label values sex sex_lbl
label define region_lbl 1 "Dhaka division" 2 "Outside Dhaka division", replace
label values region region_lbl
label define res_lbl 1 "City" 2 "Town" 3 "Village", replace
label values residence res_lbl
label define br_lbl 1 "<15k" 2 "15-30k" 3 "30-60k" 4 "60-100k" 5 ">100k", replace
label values spend_bracket br_lbl
label define cond_lbl 1 "Kidney failure-dialysis" 2 "CKD/kidney damage" 3 "Glomerular" 4 "Stone" 5 "Obstruction/hydronephrosis" 6 "Tumour/cancer" 7 "Infection" 8 "Other/undiagnosed", replace
label values cond_cat cond_lbl
label define fc_lbl 1 "Pharmacy" 2 "Village doctor/informal" 3 "Community clinic" 4 "Govt hospital" 5 "Private clinic/hospital" 6 "This hospital directly" 7 "Nothing", replace
label values first_care fc_lbl
label define nights_lbl 0 "None" 1 "1-7" 2 "8-14" 3 "15-30" 4 "31+", replace
label values nights_band nights_lbl
label define care_lbl 1 "<=1 month" 2 "1-3" 3 "3-6" 4 "6-12" 5 "12-24" 6 ">24 months", replace
label values care_months_band care_lbl
label define cq_lbl 1 "Itemised" 2 "Partial" 3 "Total only", replace
label values cost_quality cq_lbl

* ---- analysis variables
gen ln_oop_medical = ln(oop_medical) if oop_medical > 0 & !missing(oop_medical)
gen byte outpatient = visit_type == 2
gen byte uro = specialty == 1
gen byte child = 1 - adult
gen byte first_private = first_care == 5 if !missing(first_care)
gen byte first_informal = inlist(first_care, 1, 2) if !missing(first_care)

* ---- descriptives
tab wave
tab division
tab district, sort
tabstat oop_medical oop_total if usable_cost == 1, by(visit_type) stat(n p25 p50 p75) format(%12.0fc)
tabstat cat10 cat25 fin_distress delay_cost, by(visit_type) stat(n mean) format(%5.3f)
ranksum oop_medical if usable_cost == 1, by(visit_type)
tab first_care visit_type, col
* Cost shares (fully itemised forms)
preserve
keep if usable_cost == 1 & cost_quality == 1
egen items = rowtotal(cost_fees cost_diag cost_med cost_bed)
collapse (sum) cost_fees cost_diag cost_med cost_bed items
foreach v in cost_fees cost_diag cost_med cost_bed {
    display "`v' share = " %5.3f `v'/items
}
restore

* ---- regressions (robust SEs). Bands replace exact months of care and nights.
local ctrl outpatient dialysis child uro i.care_months_band nights_band spend_bracket
reg ln_oop_medical n_providers `ctrl' if usable_cost == 1, vce(robust)
reg ln_oop_medical first_private `ctrl' if usable_cost == 1, vce(robust)
reg ln_oop_medical premed `ctrl' if usable_cost == 1, vce(robust)
reg ln_oop_medical first_informal `ctrl' if usable_cost == 1, vce(robust)
