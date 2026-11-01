"""TEMPLATE for the hand-coded survey records (the real files survey_records.py and survey_records_wave2.py
are confidential and are NOT in this repository).

The record below is FICTIONAL and only shows the format.

One dict per paper form. Coding conventions (see README sheet of the workbook):
  - Money in BDT. "k" = 1,000; "lakh/lac/l" = 100,000. Ranges -> midpoint.
  - Rates (per day / per week / per month / per trip) are multiplied out; the multiplier is
    stated in the raw-text audit column.
  - oop (medical out-of-pocket) = fees + diagnostics + medicines + bed/procedure, OR the
    respondent's stated all-inclusive medical total when that is larger / the only figure.
  - Transport and attendant food are kept separate (non-medical) because SDG 3.8.2 excludes them.
  - None = not recorded / not answered.

Keys
  s serial | w ward (1 PaedUro 2 PaedNephro 3 MaleUro 4 MaleNephro 5 FemaleUro 6 FemaleNephro)
  sex 1M 2F | age years | rp respondent is patient (1) | rel relation of attendant
  d district | res 1 city 2 town 3 village | hh household size | occ occupation text | oc occ category
  br spending bracket 1..5 | sx exact monthly spending if given
  cond condition text | cc condition category | dia dialysis | nc nights this admission
  nk nights covered by the costs | on days since symptom onset
  fc first point of care (1 pharmacy 2 village doctor/informal 3 community clinic 4 govt hosp
     5 private clinic/hosp 6 this hospital directly 7 nothing) | path care path | np n providers before
  rs reason (1 referred 2 reputation 3 closest 4 cost 5 doctor known 6 own choice/other)
  pm pre-medication | ab antibiotic | pk painkiller | po other
  fe dg md bd cost items | rt reported total (medical) | oop final medical OOP | tr transport | fd food
  inf informal payment reported | ins scheme/insurance paid
  fin codes: S savings R relatives/friends L lender/NGO/bank U borrowed (source unspecified)
             A sold/pawned/mortgaged asset X help from abroad I current income N none
  dl days of work lost | jl job loss | dy delay | dr delay reason (1 funds 2 transport 3 other 4 unspecified)
  fm fewer medicines | st skipped test | le left early
  nx next choice (1 NIKDU 2 other facility 3 depends/doctor 4 unsure 9 missing)
  nr next reason (1 quality 2 familiar 3 cheap/govt 4 doctor advice 5 close 6 other)
  sc (superseded: cost period now derived from care_months in build_workbook.py) | q cost quality (1 itemised 2 partial 3 total only)
  us usable for cost analysis | cmp completion as ticked on form
  raw [fees, diag, med, bed, transport, food] as written + conversion
  iss list of (severity, issue, action)
"""

R = []
def r(**k):
    R.append(k)

# FICTIONAL example - not a real patient
r(s='0', w=3, sex=1, age=40, rp=1, rel='', d='Dhaka', res=1, hh=4, occ='Example occupation', oc=8, br=2, sx=None,
  cond='Example condition', cc=8, dia=0, nc=3, nk=3, on=30,
  fc=4, path='Example govt hospital -> this hospital', np=1, rs=1, pm=0, ab=0, pk=0, po=0,
  fe=500, dg=5000, md=2000, bd=1000, rt=None, oop=8500, tr=600, fd=400, inf=0, ins=0,
  fin='S', dl=5, jl=0, dy=0, dr=None, fm=0, st=0, le=0, nx=1, nr=3, sc=1, q=1, us=1, cmp='Yes',
  raw=['500', '5000', '2000', '1000', '300 x 2', '400'],
  iss=[])

# Wave 2+ records also carry: hosp (hospital code), vt (visit type), spec (specialty), cm (months of care)
