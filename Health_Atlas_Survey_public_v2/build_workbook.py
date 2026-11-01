"""Build the Health Atlas survey workbook (v2), a Stata/SPSS-ready CSV and a Stata do-file
from the hand-coded records in survey_records.py (wave 1, NIKDU inpatients) and
survey_records_wave2.py (wave 2, NIKDU outpatients).

Run:  python3 build_workbook.py     (needs numpy and openpyxl)
"""
import csv
import math
import statistics as st

import numpy as np
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula

from districts import DIVISION
from survey_records import R as R1
from survey_records_wave2 import R2, REFUSALS

for _x in R1:
    _x.setdefault('wave', 1)
for _x in R2:
    _x.setdefault('wave', 2)
R = R1 + R2

OUT_XLSX = 'Health_Atlas_Survey_Workbook_v2.xlsx'
OUT_CSV = 'survey_clean_v2.csv'
OUT_DO = 'survey_analysis_v2.do'
BUILD_DATE = '1 Nov 2026'

# Monthly consumption used when only a bracket is known (assumption, documented in README).
BRACKET_MID = {1: 12000, 2: 22500, 3: 45000, 4: 80000, 5: 120000}

# Months of care per wave 1 form (keyed by serial) are kept with the confidential data.
from care_months_wave1 import CARE_MONTHS

WARDS = {1: 'Paediatric Urology', 2: 'Paediatric Nephrology', 3: 'Male Urology', 4: 'Male Nephrology',
         5: 'Female Urology', 6: 'Female Nephrology', 7: 'NIKDU Outpatient', 8: 'NICVD Cardiology (inpatient)',
         9: 'NICVD Cardiac Surgery (inpatient)', 10: 'NICVD Outpatient', 11: 'Private cardiology practice'}
HOSPITALS = {1: 'NIKDU', 2: 'NICVD', 3: 'Private cardiology practice'}


def care_months(x):
    return x['cm'] if 'cm' in x else CARE_MONTHS[x['s']]

# ---------------------------------------------------------------- variable spec
# (name, label, codes/units, kind)  kind: 'in' = entered from the form, 'der' = derived
SPEC = [
    ('id', 'Row id', 'running number', 'in'),
    ('serial', 'Serial number on paper form', 'text', 'in'),
    ('wave', 'Survey wave', '1 NIKDU inpatients (Sep 2026)  2 NIKDU outpatients (Oct 2026)', 'in'),
    ('hospital', 'Hospital', '1 NIKDU  2 NICVD  3 private cardiology practice (author-related physician: disclose)', 'in'),
    ('ward', 'Ward / department', '; '.join(f'{k} {v}' for k, v in WARDS.items()), 'in'),
    ('specialty', 'Specialty', '1 urology  2 nephrology  3 cardiology/cardiac surgery (derived from ward 1-6, 8-11; ENTER for ward 7 outpatients from the department line)', 'der'),
    ('paediatric', 'Paediatric patient (<18: paediatric ward 1-2, or age <18 for other wards)', '0/1', 'der'),
    ('visit_type', 'Visit type', '1 inpatient  2 outpatient', 'in'),
    ('interview_date', 'Interview date', 'dd/mm/yyyy (NOT recorded in waves 1 or 2)', 'in'),
    ('sex', 'Patient sex', '1 male  2 female', 'in'),
    ('age', 'Patient age', 'years (decimals for infants)', 'in'),
    ('resp_patient', 'Respondent is the patient', '1 patient  0 attendant', 'in'),
    ('resp_relation', 'Attendant relation to patient', 'text', 'in'),
    ('district', 'Household district', 'BBS spelling (see Atlas_Districts)', 'in'),
    ('division', 'Household division', 'text', 'der'),
    ('residence', 'Residence type', '1 city  2 town  3 village', 'in'),
    ('hh_size', 'Household size', 'persons', 'in'),
    ('occupation', 'Main earner occupation (as said)', 'text', 'in'),
    ('occ_cat', 'Main earner occupation category', '1 agriculture/fishing  2 wage/manual labour  3 small business/trade  4 salaried/skilled  5 driver/transport  6 migrant worker abroad  7 pension/govt service  8 other/unclear', 'in'),
    ('spend_bracket', 'Monthly household spending bracket', '1 <15k  2 15-30k  3 30-60k  4 60-100k  5 >100k', 'in'),
    ('spend_exact', 'Monthly spending if an exact figure was given', 'BDT', 'in'),
    ('monthly_cons', 'Monthly consumption used for shares', 'BDT = spend_exact, else bracket midpoint', 'der'),
    ('cond_text', 'Condition (as recorded)', 'text', 'in'),
    ('cond_cat', 'Condition category', '1 kidney failure-dialysis  2 CKD/kidney damage  3 glomerular ("filter")  4 stone  5 obstruction/hydronephrosis  6 tumour/cancer  7 infection  8 other/undiagnosed  9 heart (cardiology)  10 cardiac surgery', 'in'),
    ('dialysis', 'Dialysis during episode', '0/1', 'in'),
    ('nights_current', 'Nights in current admission', 'nights', 'in'),
    ('nights_costed', 'Inpatient nights covered by the reported costs', 'nights (all stays inside the costed period)', 'in'),
    ('onset_days', 'Days since symptoms began', 'days (approx.)', 'in'),
    ('chronic', 'Symptoms for 90+ days', '0/1', 'der'),
    ('care_months', 'Months over which the reported costs were spent (first care for this disease -> interview)', 'months', 'in'),
    ('first_care', 'First point of care (C1)', '1 pharmacy/drug seller  2 village doctor/informal  3 community clinic  4 govt hospital  5 private clinic/hospital  6 this hospital directly  7 nothing at first', 'in'),
    ('first_informal', 'First care at pharmacy or informal provider (H3)', '0/1', 'der'),
    ('first_private', 'First care at private clinic/hospital', '0/1', 'der'),
    ('first_direct', 'Came to this hospital first', '0/1', 'der'),
    ('care_path', 'Care path (as recorded)', 'text', 'in'),
    ('n_providers', 'Providers seen before this hospital (C2)', 'count', 'in'),
    ('reason', 'Why this hospital (C3)', '1 referred  2 reputation  3 closest  4 cost  5 doctor known  6 own choice/other', 'in'),
    ('referred', 'Referred', '0/1', 'der'),
    ('premed', 'Took medicine before a prescription (C4)', '0/1', 'in'),
    ('premed_antibiotic', 'Pre-prescription antibiotic', '0/1', 'in'),
    ('premed_painkiller', 'Pre-prescription painkiller', '0/1', 'in'),
    ('premed_other', 'Pre-prescription other medicine', '0/1', 'in'),
    ('cost_fees', 'D1 Doctor/consultation fees', 'BDT', 'in'),
    ('cost_diag', 'D2 Diagnostics & imaging', 'BDT', 'in'),
    ('cost_med', 'D3 Medicines', 'BDT', 'in'),
    ('cost_bed', 'D4 Bed, procedure, operation (incl. dialysis sessions, blood, ICU)', 'BDT', 'in'),
    ('cost_items_sum', 'Sum of reported medical items D1-D4', 'BDT', 'der'),
    ('cost_reported_total', 'All-inclusive medical total stated by respondent', 'BDT', 'in'),
    ('oop_medical', 'Medical out-of-pocket spending (analysis variable)', 'BDT: max(items, stated medical total) - see Cost_Workings', 'in'),
    ('cost_transport', 'D5 Transport, all trips', 'BDT', 'in'),
    ('cost_food', 'D6 Food & lodging for attendants', 'BDT', 'in'),
    ('oop_nonmedical', 'Transport + food', 'BDT', 'der'),
    ('oop_total', 'Total economic cost (medical + non-medical)', 'BDT', 'der'),
    ('informal_pay', 'D7 Unbilled payment reported (tips/staff/brokers)', '0/1', 'in'),
    ('scheme_paid', 'D8 Insurance/employer/NGO/govt scheme paid', '0/1', 'in'),
    ('fin_savings', 'E1 Paid from savings', '0/1', 'in'),
    ('fin_borrow_rel', 'E1 Borrowed from relatives/friends', '0/1', 'in'),
    ('fin_borrow_lender', 'E1 Borrowed from lender/NGO/bank', '0/1', 'in'),
    ('fin_borrow_unspec', 'E1 Borrowed, source not recorded', '0/1', 'in'),
    ('fin_sold_asset', 'E1 Sold/pawned/mortgaged asset (incl. land)', '0/1', 'in'),
    ('fin_abroad', 'E1 Help from relatives abroad', '0/1', 'in'),
    ('fin_income', 'E1 Paid from current income', '0/1', 'in'),
    ('fin_distress', 'Any borrowing or asset sale', '0/1', 'der'),
    ('days_lost', 'E2 Work days lost (patient + attendants)', 'days', 'in'),
    ('job_loss', 'Someone lost their job during the illness', '0/1', 'in'),
    ('delay', 'E3 Delayed coming to hospital', '0/1', 'in'),
    ('delay_reason', 'Reason for delay', '1 funds  2 transport  3 other (non-cost)  4 not specified', 'in'),
    ('delay_cost', 'Delay attributed to cost (funds or unspecified)', '0/1', 'der'),
    ('fewer_meds', 'E3 Bought fewer medicines than prescribed', '0/1', 'in'),
    ('skipped_test', 'E3 Skipped a test', '0/1', 'in'),
    ('left_early', 'E3 Left earlier than advised', '0/1', 'in'),
    ('any_cost_coping', 'Any of E3 due to cost', '0/1', 'der'),
    ('next_choice', 'E4 Where next time', '1 same hospital  2 other facility  3 depends/doctor advice  4 unsure  9 missing', 'in'),
    ('next_reason', 'E4 Why', '1 quality  2 familiar  3 cheap/govt  4 doctor advice  5 close  6 other', 'in'),
    ('cost_scope', 'Length of the costed illness', '1 within 12 months  2 longer than 12 months (from care_months)', 'der'),
    ('cost_quality', 'Cost data quality', '1 fully itemised  2 partly itemised/estimated  3 total only', 'in'),
    ('usable_cost', 'Usable for cost analysis', '0/1', 'in'),
    ('form_complete', 'Completion as ticked on form', 'Yes/No/blank', 'in'),
    ('oop_per_month', 'Medical OOP per month of care', 'BDT = oop_medical / max(care_months,1)', 'der'),
    ('oop_12m', 'Medical OOP falling in one 12-month window', 'BDT = oop_medical x min(1, 12/care_months) (even spread)', 'der'),
    ('oop_share', '12-month medical OOP / annual household consumption (SDG 3.8.2 style)', 'fraction', 'der'),
    ('cat10', 'Catastrophic: 12-month OOP > 10% of annual consumption', '0/1', 'der'),
    ('cat25', 'Catastrophic: 12-month OOP > 25% of annual consumption', '0/1', 'der'),
    ('share_illness', 'Whole-illness medical OOP / one year of consumption', 'fraction', 'der'),
    ('cat25_illness', 'Whole-illness OOP > 25% of one year of consumption', '0/1', 'der'),
    ('ln_oop_medical', 'ln(medical OOP)', 'log BDT', 'der'),
    ('ln_oop_total', 'ln(total economic cost)', 'log BDT', 'der'),
    ('n_issues', 'Number of data issues logged', 'count (see Data_Issues)', 'der'),
]
COLS = [s[0] for s in SPEC]
COL = {name: get_column_letter(i + 1) for i, name in enumerate(COLS)}


def build_rows():
    rows = []
    for i, x in enumerate(R, start=1):
        fin = x['fin']
        items = [x[k] for k in ('fe', 'dg', 'md', 'bd') if x[k] is not None]
        mcons = x['sx'] or BRACKET_MID.get(x['br'])
        oop = x['oop']
        nonmed = (x['tr'] or 0) + (x['fd'] or 0)
        cm = care_months(x)
        w = x['w']
        if w in (1, 3, 5):
            spec = 1
        elif w in (2, 4, 6):
            spec = 2
        elif w >= 8:
            spec = 3
        else:
            spec = x['spec']
        paed = (1 if w in (1, 2) else 0) if w <= 6 else (None if x['age'] is None else int(x['age'] < 18))
        dy = x['dy']
        dcost = None if dy is None else int(dy == 1 and x['dr'] in (1, 4))
        coping = None if dy is None else int(dcost == 1 or x['fm'] == 1 or x['st'] == 1 or x['le'] == 1)
        oop12 = oop * min(1.0, 12.0 / cm)
        share = oop12 / (12 * mcons) if (mcons and x['us']) else None
        share_ill = oop / (12 * mcons) if (mcons and x['us']) else None
        fc = x['fc']
        row = {
            'id': i, 'serial': x['s'], 'wave': x['wave'], 'hospital': x.get('hosp', 1), 'ward': w,
            'specialty': spec, 'paediatric': paed,
            'visit_type': x.get('vt', 1), 'interview_date': None, 'sex': x['sex'], 'age': x['age'],
            'resp_patient': x['rp'], 'resp_relation': x['rel'] or None,
            'district': x['d'], 'division': DIVISION[x['d']], 'residence': x['res'], 'hh_size': x['hh'],
            'occupation': x['occ'], 'occ_cat': x['oc'], 'spend_bracket': x['br'], 'spend_exact': x['sx'],
            'monthly_cons': mcons, 'cond_text': x['cond'], 'cond_cat': x['cc'], 'dialysis': x['dia'],
            'nights_current': x['nc'], 'nights_costed': x['nk'], 'onset_days': x['on'],
            'chronic': 1 if x['on'] >= 90 else 0, 'care_months': cm,
            'first_care': fc, 'first_informal': None if fc is None else int(fc in (1, 2)),
            'first_private': None if fc is None else int(fc == 5),
            'first_direct': None if fc is None else int(fc == 6), 'care_path': x['path'], 'n_providers': x['np'],
            'reason': x['rs'], 'referred': None if x['rs'] is None else int(x['rs'] == 1),
            'premed': x['pm'], 'premed_antibiotic': x['ab'], 'premed_painkiller': x['pk'], 'premed_other': x['po'],
            'cost_fees': x['fe'], 'cost_diag': x['dg'], 'cost_med': x['md'], 'cost_bed': x['bd'],
            'cost_items_sum': sum(items) if items else None, 'cost_reported_total': x['rt'],
            'oop_medical': oop, 'cost_transport': x['tr'], 'cost_food': x['fd'],
            'oop_nonmedical': nonmed, 'oop_total': oop + nonmed,
            'informal_pay': x['inf'], 'scheme_paid': x['ins'],
            'fin_savings': int('S' in fin), 'fin_borrow_rel': int('R' in fin), 'fin_borrow_lender': int('L' in fin),
            'fin_borrow_unspec': int('U' in fin), 'fin_sold_asset': int('A' in fin), 'fin_abroad': int('X' in fin),
            'fin_income': int('I' in fin), 'fin_distress': int(any(c in fin for c in 'RLUA')),
            'days_lost': x['dl'], 'job_loss': x['jl'], 'delay': dy, 'delay_reason': x['dr'],
            'delay_cost': dcost,
            'fewer_meds': x['fm'], 'skipped_test': x['st'], 'left_early': x['le'],
            'any_cost_coping': coping,
            'next_choice': x['nx'], 'next_reason': x['nr'], 'cost_scope': 2 if cm > 12 else 1, 'cost_quality': x['q'],
            'usable_cost': x['us'], 'form_complete': x['cmp'],
            'oop_per_month': round(oop / max(cm, 1)), 'oop_12m': round(oop12),
            'oop_share': round(share, 4) if share is not None else None,
            'cat10': None if share is None else int(share > 0.10),
            'cat25': None if share is None else int(share > 0.25),
            'share_illness': round(share_ill, 4) if share_ill is not None else None,
            'cat25_illness': None if share_ill is None else int(share_ill > 0.25),
            'ln_oop_medical': round(math.log(oop), 4) if oop > 0 else None,
            'ln_oop_total': round(math.log(oop + nonmed), 4) if (oop + nonmed) > 0 else None,
            'n_issues': len(x['iss']),
        }
        if len(fin) == 0:  # financing not answered
            for k in ('fin_savings', 'fin_borrow_rel', 'fin_borrow_lender', 'fin_borrow_unspec',
                      'fin_sold_asset', 'fin_abroad', 'fin_income', 'fin_distress'):
                row[k] = None
        rows.append(row)
    return rows


# ---------------------------------------------------------------- statistics helpers
def betacf(a, b, x):
    m_max, eps = 200, 3e-14
    qab, qap, qam = a + b, a + 1, a - 1
    c, d = 1.0, 1 - qab * x / qap
    d = 1 / (d if abs(d) > 1e-300 else 1e-300)
    h = d
    for m in range(1, m_max + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1 + aa * d; d = 1 / (d if abs(d) > 1e-300 else 1e-300)
        c = 1 + aa / c if abs(1 + aa / c) > 1e-300 else 1e-300
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1 + aa * d; d = 1 / (d if abs(d) > 1e-300 else 1e-300)
        c = 1 + aa / c if abs(1 + aa / c) > 1e-300 else 1e-300
        de = d * c
        h *= de
        if abs(de - 1) < eps:
            break
    return h


def t_pvalue(t, df):
    x = df / (df + t * t)
    a, b = df / 2, 0.5
    bt = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1 - x))
    ib = bt * betacf(a, b, x) / a if x < (a + 1) / (a + b + 2) else 1 - bt * betacf(b, a, 1 - x) / b
    return ib  # two-sided p


def ols_hc1(rows, y, xs):
    data = [r for r in rows if r[y] is not None and all(r[v] is not None for v in xs)]
    Y = np.array([r[y] for r in data], float)
    X = np.column_stack([np.ones(len(data))] + [np.array([r[v] for r in data], float) for v in xs])
    n, k = X.shape
    XtX_inv = np.linalg.inv(X.T @ X)
    b = XtX_inv @ X.T @ Y
    e = Y - X @ b
    meat = X.T @ (X * (e ** 2)[:, None])
    V = XtX_inv @ meat @ XtX_inv * n / (n - k)
    se = np.sqrt(np.diag(V))
    r2 = 1 - (e @ e) / ((Y - Y.mean()) @ (Y - Y.mean()))
    out = []
    for name, bi, si in zip(['constant'] + xs, b, se):
        t = bi / si
        out.append((name, bi, si, t_pvalue(abs(t), n - k)))
    return out, n, r2


# ---------------------------------------------------------------- styling helpers
FONT = 'Arial'
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color='FFFFFF')
F_TITLE = Font(name=FONT, size=14, bold=True, color='1F3864')
F_SUB = Font(name=FONT, size=11, bold=True, color='1F3864')
F_NOTE = Font(name=FONT, size=9, italic=True, color='595959')
F_BLUE = Font(name=FONT, size=10, color='0000FF')
FILL_HEAD = PatternFill('solid', start_color='1F3864')
FILL_DER = PatternFill('solid', start_color='D9D9D9')
FILL_IN = PatternFill('solid', start_color='FFFF00')
FILL_EX = PatternFill('solid', start_color='E2EFDA')
FILL_HIGH = PatternFill('solid', start_color='F8CBAD')
FILL_MED = PatternFill('solid', start_color='FFE699')
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')


def style_header(ws, row, ncols, fill=FILL_HEAD):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = F_HEAD
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical='center')
        cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def base_font(ws):
    for row in ws.iter_rows():
        for cell in row:
            if cell.font is None or cell.font.name != FONT:
                cell.font = Font(name=FONT, size=cell.font.size if cell.font and cell.font.size else 10,
                                 bold=cell.font.bold if cell.font else False,
                                 italic=cell.font.italic if cell.font else False,
                                 color=cell.font.color if cell.font else None)


# ---------------------------------------------------------------- build
def main():
    rows = build_rows()
    N = len(rows)
    last = N + 1
    wb = Workbook()

    # ============ README
    ws = wb.active
    ws.title = 'README'
    n1 = sum(r['wave'] == 1 for r in rows)
    n2 = sum(r['wave'] == 2 for r in rows)
    lines = [
        ('Health Atlas of Bangladesh - Patient Cost Survey Workbook (v2)', F_TITLE),
        (f'v2 combines wave 1 (NIKDU inpatients, {n1} forms: serials 1-48, 51-60 and one unnumbered free-care case "NC") and wave 2 '
         f'(NIKDU outpatients, {n2} forms: serials 61-100). Total {N} forms. Built {BUILD_DATE}. Wave 1 rows are unchanged from v1 (29 Sep).', F_BASE),
        ('Serial numbers run 1-100 but serials 49 and 50 were never supplied and "NC" has no number, so the dataset has 99 rows, not 100.', F_BOLD),
        ('', F_BASE),
        ('Sheets', F_SUB),
        ('Survey_Data - one row per respondent, numeric codes, Stata/SPSS-safe variable names in row 1 (also saved as survey_clean_v2.csv). New column "wave".', F_BASE),
        ('Codebook - label, codes/units and whether each variable is entered from the form (white) or derived (grey).', F_BASE),
        ('Cost_Workings - audit trail: what the respondent said for each cost item, how it was converted, and the final coded value.', F_BASE),
        ('Data_Issues - every problem found, by form, with severity and what to do (wave 2 issues added).', F_BASE),
        ('Phase1_Summary - descriptive tables. Column C is live (formulas over all rows). Columns D-F are fixed snapshots: all forms, inpatients, outpatients.', F_BASE),
        ('Sensitivity - how the headline numbers move when the most doubtful wave 2 forms are dropped or re-coded.', F_BASE),
        ('H3_Test - Hypothesis 3 with all 99 forms, outpatient status controlled.', F_BASE),
        ('Fieldwork_Tracker - quotas with live counts, plus refusal tallies (response rate).', F_BASE),
        ('New_Entry - blank entry sheet with dropdowns for the next forms (private cardiology practice).', F_BASE),
        ('Atlas_Districts - 64 districts x 8 divisions with indicator columns and a live count of respondents per district.', F_BASE),
        ('', F_BASE),
        ('Cost coding rules (unchanged from v1)', F_SUB),
        ('1. All money in BDT. "k" = 1,000; "lakh/lac/l" = 100,000. Ranges (e.g. 10-15k) coded at the midpoint.', F_BASE),
        ('2. Rates were multiplied out: per day x nights, per week x weeks, per trip x trips, per month x months. Where the number of weeks was not written, the illness length was used and the assumption is written in Cost_Workings. A rate with no usable duration (e.g. "20tk a day") is left missing.', F_BASE),
        ("3. oop_medical = D1+D2+D3+D4, or the respondent's all-inclusive medical total when that is larger or is the only figure given.", F_BASE),
        ('4. Transport (D5) and attendant food (D6) are kept OUT of oop_medical (SDG 3.8.2 counts only medical payments). oop_total adds them.', F_BASE),
        ('5. Catastrophic spending: oop_12m (= oop_medical x min(1, 12/care_months)) / (12 x monthly consumption) above 10% or 25%. Monthly consumption = exact figure if given, else bracket midpoint (<15k -> 12,000; 15-30k -> 22,500; 30-60k -> 45,000; 60-100k -> 80,000; >100k -> 120,000).', F_BASE),
        ('6. Costs trace ONE disease from first care to the interview, including care before NIKDU. care_months records how long that took (estimated from the path notes). For wave 2 it is stored with each record in survey_records_wave2.py.', F_BASE),
        ('', F_BASE),
        ('Wave 2 (outpatient) coding notes', F_SUB),
        ('a. All wave 2 forms: hospital = 1 (NIKDU), ward = 7 (NIKDU Outpatient), visit_type = 2, nights_current = 0. nights_costed = inpatient nights inside the costed period where the path mentions an admission.', F_BASE),
        ('b. specialty is taken from the department line (serials 61-80 urology, 81-100 nephrology; blank or garbled department lines coded from the batch). paediatric = age < 18 for outpatients.', F_BASE),
        ('c. Serial 83 ticked NICVD but is a NIKDU nephrology outpatient - coded NIKDU. Serials 75 and 96 ticked "government hospital" first with 0 providers before - coded "this hospital directly" (NIKDU is a government hospital).', F_BASE),
        ('d. A blank coping question (serial 97) is coded missing, not 0; the do-file now respects that.', F_BASE),
        ('e. Refusal tallies were written for forms 71-80 (13 refusals) and 81-90 (5 refusals) - the first response-rate data in the survey.', F_BASE),
        ('', F_BASE),
        ('Colour legend', F_SUB),
        ('Yellow fill = cells for you to fill in (New_Entry, Atlas_Districts). Grey header = derived variable (leave blank on entry; the do-file recomputes). Green row = worked example. Blue text = hard-coded assumption.', F_BASE),
        ('', F_BASE),
        ('Disclosure (required for wave 3)', F_SUB),
        ("The next wave is at a private cardiology practice whose physician is related to the author. This must be disclosed in every report that uses those data.", F_BOLD),
        ('', F_BASE),
        ('Privacy', F_SUB),
        ('This file contains row-level patient data. Keep it on the encrypted drive; do NOT commit it to GitHub or upload it anywhere public. Publish only aggregated tables (medians, percentages).', F_BOLD),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=text)
        c.font = font
        c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.column_dimensions['A'].width = 130

    # ============ Survey_Data
    ws = wb.create_sheet('Survey_Data')
    for j, name in enumerate(COLS, start=1):
        ws.cell(row=1, column=j, value=name)
    style_header(ws, 1, len(COLS))
    for j, s in enumerate(SPEC, start=1):
        if s[3] == 'der':
            ws.cell(row=1, column=j).fill = PatternFill('solid', start_color='7F7F7F')
        ws.cell(row=1, column=j).comment = Comment(f'{s[1]}\n{s[2]}', 'codebook')
    money = {'cost_fees', 'cost_diag', 'cost_med', 'cost_bed', 'cost_items_sum', 'cost_reported_total', 'oop_medical',
             'cost_transport', 'cost_food', 'oop_nonmedical', 'oop_total', 'monthly_cons', 'spend_exact', 'oop_per_month', 'oop_12m'}
    for i, row in enumerate(rows, start=2):
        for j, name in enumerate(COLS, start=1):
            c = ws.cell(row=i, column=j, value=row[name])
            c.font = F_BASE
            if name in money:
                c.number_format = '#,##0'
            elif name in ('oop_share', 'share_illness'):
                c.number_format = '0.0%'
    ws.freeze_panes = 'C2'
    for j, name in enumerate(COLS, start=1):
        w = 11
        if name in ('cond_text', 'care_path', 'occupation'):
            w = 34
        elif name in money:
            w = 12
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.auto_filter.ref = f'A1:{get_column_letter(len(COLS))}{last}'

    # ============ Codebook
    ws = wb.create_sheet('Codebook')
    hdr = ['Variable', 'Label', 'Codes / units', 'Entered or derived', 'Column in Survey_Data']
    ws.append(hdr)
    style_header(ws, 1, len(hdr))
    for s in SPEC:
        ws.append([s[0], s[1], s[2], 'entered from form' if s[3] == 'in' else 'derived', COL[s[0]]])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = F_BASE
            c.alignment = WRAP
        if row[3].value == 'derived':
            for c in row:
                c.fill = FILL_DER
    set_widths(ws, [20, 48, 80, 18, 12])
    ws.freeze_panes = 'A2'

    # ============ Cost_Workings
    ws = wb.create_sheet('Cost_Workings')
    hdr = ['Serial', 'Wave', 'Ward', 'D1 fees (as said)', 'D2 diagnostics (as said)', 'D3 medicines (as said)', 'D4 bed/procedure (as said)',
           'D5 transport (as said)', 'D6 food (as said)', 'Fees', 'Diagnostics', 'Medicines', 'Bed/proc.', 'Stated medical total',
           'oop_medical', 'Transport', 'Food', 'Scope', 'Quality']
    ws.append(hdr)
    style_header(ws, 1, len(hdr))
    for x in R:
        ws.append([x['s'], x['wave'], WARDS[x['w']]] + x['raw'] + [x['fe'], x['dg'], x['md'], x['bd'], x['rt'], x['oop'], x['tr'], x['fd'],
                  f"{care_months(x)} months of care",
                  {1: 'itemised', 2: 'partial', 3: 'total only'}[x['q']]])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = F_BASE
            c.alignment = WRAP
        for c in row[9:17]:
            c.number_format = '#,##0'
        row[14].font = F_BOLD
    set_widths(ws, [7, 6, 18, 18, 30, 26, 34, 26, 20, 10, 11, 11, 11, 12, 12, 11, 10, 14, 11])
    ws.freeze_panes = 'D2'

    # ============ Data_Issues
    ws = wb.create_sheet('Data_Issues')
    hdr = ['Serial', 'Severity', 'Issue', 'Action']
    ws.append(hdr)
    style_header(ws, 1, len(hdr))
    general = [
        ('ALL', 'High', 'No interview date on any of the 99 forms (waves 1 and 2). Wave 2 was around mid-October 2026.', 'Write the date on every form from now on; reconstruct wave-1 dates from your diary if you can.'),
        ('ALL', 'High', 'Serials 49 and 50 are missing (no forms supplied). The count is 59 forms, not 60.', 'Find the two forms or record them as lost / refused.'),
        ('ALL', 'Medium', 'Costs correctly trace one disease from onset, but the length of that period (2 weeks to 9 years) is only recoverable from the path notes. care_months was estimated from them.', 'Add one line to the form: "Months since first care for this illness: __". Check the estimates in Survey_Data against your notes.'),
        ('ALL', 'Medium', 'Refusals not counted in wave 1 or for forms 61-70 and 91-100. Wave 2 tallies: 13 refusals for forms 71-80, 5 for 81-90 (20 completed of 38 approached = 53%; 43% for 71-80, 67% for 81-90).', 'Keep a tally of approached / refused / completed for every batch.'),
        ('ALL', 'Medium', '"Interview completion" left blank on 13 wave-1 forms and 3 wave-2 forms (89, 95 blank; 67 ticked "No" though complete).', 'Tick it before moving on.'),
        ('W2', 'Medium', 'Outpatient costs often given as weekly/monthly medicine rates; duration assumed from the illness length (serials 71, 72, 91, 93, 95, 98).', 'Write "x N weeks" next to every rate.'),
        ('W2', 'Medium', 'Several dialysis outpatients (83, 84, 91) report few or no dialysis charges - dialysis spending is under-reported.', 'Ask "how many sessions a week, since when, price per session".'),
        ('W2', 'Low', 'Department line sometimes holds other answers (90: "female"; 96: "NIKDU"; 66: "outpatient").', 'Write the clinic name (urology/nephrology OPD).'),
        ('ALL', 'Medium', 'First point of care (C1) often recorded as a narrative path rather than a ticked category.', 'Tick one C1 category AND write the path on a separate line.'),
        ('ALL', 'Medium', 'Several forms give rates (per day/week/month) without the number of days/weeks.', 'Always write rate x number (e.g. 300 x 12 days).'),
    ]
    issues = general + [(x['s'], sev, iss, act) for x in R for (sev, iss, act) in x['iss']]
    for it in issues:
        ws.append(list(it))
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = F_BASE
            c.alignment = WRAP
        fill = FILL_HIGH if row[1].value == 'High' else FILL_MED if row[1].value == 'Medium' else None
        if fill:
            row[1].fill = fill
    set_widths(ws, [8, 10, 90, 60])
    ws.freeze_panes = 'A2'

    # ============ Phase1_Summary
    ws = wb.create_sheet('Phase1_Summary')
    D = 'Survey_Data!'
    def col(name):
        return f"{D}${COL[name]}:${COL[name]}"
    def rng(name):
        return f"{D}${COL[name]}$2:${COL[name]}$1000"

    usable = [r for r in rows if r['usable_cost'] == 1]
    ep = [r for r in usable if r['cost_scope'] == 1]
    catrows = [r for r in usable if r['cat10'] is not None]
    oops = [r['oop_medical'] for r in usable]
    GROUPS = [('All forms', rows), ('Inpatients (wave 1)', [r for r in rows if r['visit_type'] == 1]),
              ('Outpatients (wave 2)', [r for r in rows if r['visit_type'] == 2])]

    ws['A1'] = 'Phase 1 descriptive results - patient cost survey (v2: inpatients + outpatients)'
    ws['A1'].font = F_TITLE
    ws['A2'] = ('Column C recalculates from Survey_Data when opened in Excel (all rows, so it updates as you add forms). '
                f'Columns D-F are fixed snapshots ({BUILD_DATE}) for all forms, inpatients and outpatients. Money in BDT.')
    ws['A2'].font = F_NOTE
    table = []

    def _safe(fn, rs):
        try:
            return fn(rs)
        except (ZeroDivisionError, st.StatisticsError, TypeError, ValueError):
            return None

    def add(label, formula, fn, fmt='#,##0', array=False):
        table.append((label, formula, [_safe(fn, g) for _, g in GROUPS], fmt, array))

    def section(title):
        table.append((title, None, None, None, False))

    def U(rs):
        return [r for r in rs if r['usable_cost'] == 1]

    def mean_of(rs, nm):
        v = [r[nm] for r in rs if r[nm] is not None]
        return st.mean(v)

    def med_of(rs, nm):
        v = [r[nm] for r in rs if r[nm] is not None]
        return st.median(v)

    def share(rs, nm, code):
        v = [r[nm] for r in rs if r[nm] is not None]
        return sum(x == code for x in v) / len(v)

    section('A. Sample')
    add('Forms coded', f'=COUNTA({col("serial")})-1', len, '0')
    add('Usable for cost analysis', f'=COUNTIF({col("usable_cost")},1)', lambda rs: len(U(rs)), '0')
    add('Costed illness within 12 months (n)', f'=COUNTIFS({col("usable_cost")},1,{col("cost_scope")},1)', lambda rs: sum(r['cost_scope'] == 1 for r in U(rs)), '0')
    add('Outpatients (share)', f'=COUNTIF({col("visit_type")},2)/COUNT({col("visit_type")})', lambda rs: share(rs, 'visit_type', 2), '0.0%')
    add('Paediatric patients <18 (share)', f'=AVERAGE({col("paediatric")})', lambda rs: mean_of(rs, 'paediatric'), '0.0%')
    add('Urology (share)', f'=COUNTIF({col("specialty")},1)/COUNT({col("specialty")})', lambda rs: share(rs, 'specialty', 1), '0.0%')
    add('Female patients (share)', f'=COUNTIF({col("sex")},2)/COUNT({col("sex")})', lambda rs: share(rs, 'sex', 2), '0.0%')
    add('Median age (years)', f'=MEDIAN({col("age")})', lambda rs: med_of(rs, 'age'), '0')
    add('Respondent is the patient (share)', f'=AVERAGE({col("resp_patient")})', lambda rs: mean_of(rs, 'resp_patient'), '0.0%')
    add('Village households (share)', f'=COUNTIF({col("residence")},3)/COUNT({col("residence")})', lambda rs: share(rs, 'residence', 3), '0.0%')
    add('Households spending <15k/month (share)', f'=COUNTIF({col("spend_bracket")},1)/COUNT({col("spend_bracket")})', lambda rs: share(rs, 'spend_bracket', 1), '0.0%')
    add('Households from outside Dhaka division (share)', f'=1-COUNTIF({col("division")},"Dhaka")/(COUNTA({col("division")})-1)',
        lambda rs: 1 - sum(r['division'] == 'Dhaka' for r in rs) / len(rs), '0.0%')
    add('Distinct districts represented', None, lambda rs: len({r['district'] for r in rs}), '0')
    add('Dialysis patients (n)', f'=COUNTIF({col("dialysis")},1)', lambda rs: sum(r['dialysis'] == 1 for r in rs), '0')

    section('B. Cost of the illness (usable forms)')
    add('Mean medical OOP', f'=AVERAGEIFS({col("oop_medical")},{col("usable_cost")},1)', lambda rs: mean_of(U(rs), 'oop_medical'))
    add('Median medical OOP', f'=MEDIAN(IF({rng("usable_cost")}=1,{rng("oop_medical")}))', lambda rs: med_of(U(rs), 'oop_medical'), array=True)
    add('25th percentile medical OOP', f'=PERCENTILE(IF({rng("usable_cost")}=1,{rng("oop_medical")}),0.25)',
        lambda rs: float(np.percentile([r['oop_medical'] for r in U(rs)], 25)), array=True)
    add('75th percentile medical OOP', f'=PERCENTILE(IF({rng("usable_cost")}=1,{rng("oop_medical")}),0.75)',
        lambda rs: float(np.percentile([r['oop_medical'] for r in U(rs)], 75)), array=True)
    add('Median medical OOP, illness within 12 months', f'=MEDIAN(IF(({rng("usable_cost")}=1)*({rng("cost_scope")}=1),{rng("oop_medical")}))',
        lambda rs: med_of([r for r in U(rs) if r['cost_scope'] == 1], 'oop_medical'), array=True)
    add('Median medical OOP, illness longer than 12 months', f'=MEDIAN(IF(({rng("usable_cost")}=1)*({rng("cost_scope")}=2),{rng("oop_medical")}))',
        lambda rs: med_of([r for r in U(rs) if r['cost_scope'] == 2], 'oop_medical'), array=True)
    add('Median total economic cost (incl. transport, food)', f'=MEDIAN(IF({rng("usable_cost")}=1,{rng("oop_total")}))', lambda rs: med_of(U(rs), 'oop_total'), array=True)
    add('Median transport cost (where recorded)', f'=MEDIAN(IF(({rng("usable_cost")}=1)*({rng("cost_transport")}<>""),{rng("cost_transport")}))',
        lambda rs: med_of(U(rs), 'cost_transport'), array=True)
    add('Median OOP as multiple of monthly household spending', f'=MEDIAN(IF(({rng("usable_cost")}=1)*({rng("monthly_cons")}<>""),{rng("oop_medical")}/{rng("monthly_cons")}))',
        lambda rs: st.median([r['oop_medical'] / r['monthly_cons'] for r in U(rs) if r['monthly_cons']]), '0.0"x"', array=True)
    add('Median medical OOP per month of care', f'=MEDIAN(IF({rng("usable_cost")}=1,{rng("oop_per_month")}))', lambda rs: med_of(U(rs), 'oop_per_month'), array=True)
    add('Median months of care covered by the costs', f'=MEDIAN({col("care_months")})', lambda rs: med_of(rs, 'care_months'), '0.0')
    add('Total medical OOP, all usable forms', f'=SUMIFS({col("oop_medical")},{col("usable_cost")},1)', lambda rs: sum(r['oop_medical'] for r in U(rs)))

    section('C. Where the medical money goes (fully itemised forms only, H2)')
    def q1(rs):
        return [r for r in U(rs) if r['cost_quality'] == 1 and r['oop_medical'] > 0]
    for nm, lab in [('cost_fees', 'Doctor fees'), ('cost_diag', 'Diagnostics & imaging'), ('cost_med', 'Medicines'), ('cost_bed', 'Bed / procedures / dialysis')]:
        add(f'{lab} - share of medical OOP',
            f'=SUMIFS({col(nm)},{col("cost_quality")},1,{col("usable_cost")},1)/SUMIFS({col("cost_items_sum")},{col("cost_quality")},1,{col("usable_cost")},1)',
            lambda rs, nm=nm: sum(r[nm] or 0 for r in q1(rs)) / sum(r['cost_items_sum'] for r in q1(rs)), '0.0%')
    add('Transport + food as share of total economic cost',
        f'=SUMIFS({col("oop_nonmedical")},{col("cost_quality")},1,{col("usable_cost")},1)/SUMIFS({col("oop_total")},{col("cost_quality")},1,{col("usable_cost")},1)',
        lambda rs: sum(r['oop_nonmedical'] for r in q1(rs)) / sum(r['oop_total'] for r in q1(rs)), '0.0%')
    add('Number of fully itemised forms', f'=COUNTIFS({col("cost_quality")},1,{col("usable_cost")},1,{col("oop_medical")},">0")', lambda rs: len(q1(rs)), '0')

    section('D. Catastrophic spending (SDG 3.8.2 style: medical OOP in a 12-month window vs annual consumption)')
    def C(rs):
        return [r for r in U(rs) if r['cat10'] is not None]
    add('Forms with budget data (n)', f'=COUNT({col("cat10")})', lambda rs: len(C(rs)), '0')
    add('OOP > 10% of annual consumption - all', f'=COUNTIFS({col("cat10")},1)/COUNT({col("cat10")})', lambda rs: mean_of(C(rs), 'cat10'), '0.0%')
    add('OOP > 25% of annual consumption - all', f'=COUNTIFS({col("cat25")},1)/COUNT({col("cat25")})', lambda rs: mean_of(C(rs), 'cat25'), '0.0%')
    add('OOP > 10% - illness within 12 months', f'=COUNTIFS({col("cat10")},1,{col("cost_scope")},1)/COUNTIFS({col("cat10")},"<>",{col("cost_scope")},1)',
        lambda rs: mean_of([r for r in C(rs) if r['cost_scope'] == 1], 'cat10'), '0.0%')
    add('OOP > 25% - illness within 12 months', f'=COUNTIFS({col("cat25")},1,{col("cost_scope")},1)/COUNTIFS({col("cat25")},"<>",{col("cost_scope")},1)',
        lambda rs: mean_of([r for r in C(rs) if r['cost_scope'] == 1], 'cat25'), '0.0%')
    add('Whole-illness OOP > 25% of ONE year of consumption (burden view)', f'=COUNTIFS({col("cat25_illness")},1)/COUNT({col("cat25_illness")})',
        lambda rs: mean_of(C(rs), 'cat25_illness'), '0.0%')
    add('Median OOP share of annual consumption', f'=MEDIAN(IF({rng("oop_share")}<>"",{rng("oop_share")}))', lambda rs: med_of(C(rs), 'oop_share'), '0.0%', array=True)

    section('E. Care-seeking path (H3)')
    for code, lab in [(1, 'Pharmacy / drug seller'), (2, 'Village doctor / informal'), (3, 'Community clinic'), (4, 'Government hospital'),
                      (5, 'Private clinic / hospital'), (6, 'This hospital directly'), (7, 'Nothing at first')]:
        add(f'First care: {lab} (n)', f'=COUNTIF({col("first_care")},{code})', lambda rs, code=code: sum(r['first_care'] == code for r in rs), '0')
    add('First care not recorded (n)', None, lambda rs: sum(r['first_care'] is None for r in rs), '0')
    add('Mean providers seen before arriving', f'=AVERAGE({col("n_providers")})', lambda rs: mean_of(rs, 'n_providers'), '0.0')
    add('3+ providers before arriving (share)', f'=COUNTIF({col("n_providers")},">=3")/COUNT({col("n_providers")})',
        lambda rs: st.mean(int(r['n_providers'] >= 3) for r in rs if r['n_providers'] is not None), '0.0%')
    add('Referred (share)', f'=AVERAGE({col("referred")})', lambda rs: mean_of(rs, 'referred'), '0.0%')
    add('Took medicine before any prescription (share)', f'=AVERAGE({col("premed")})', lambda rs: mean_of(rs, 'premed'), '0.0%')
    add('... of which antibiotic (n)', f'=COUNTIF({col("premed_antibiotic")},1)', lambda rs: sum(r['premed_antibiotic'] == 1 for r in rs), '0')

    section('F. How families paid (share of forms answering)')
    for nm, lab in [('fin_savings', 'Savings'), ('fin_borrow_rel', 'Borrowed from relatives/friends'), ('fin_borrow_lender', 'Borrowed from lender/NGO/bank'),
                    ('fin_borrow_unspec', 'Borrowed (source not recorded)'), ('fin_sold_asset', 'Sold/pawned/mortgaged asset'),
                    ('fin_abroad', 'Help from relatives abroad'), ('fin_income', 'Current income'),
                    ('fin_distress', 'ANY borrowing or asset sale (distress financing)')]:
        add(lab, f'=AVERAGE({col(nm)})', lambda rs, nm=nm: mean_of(rs, nm), '0.0%')
    add('Scheme / insurance paid anything (n)', f'=COUNTIF({col("scheme_paid")},1)', lambda rs: sum(r['scheme_paid'] == 1 for r in rs), '0')
    add('Unbilled (informal) payment reported (n)', f'=COUNTIF({col("informal_pay")},1)', lambda rs: sum(r['informal_pay'] == 1 for r in rs), '0')

    section('G. Consequences')
    add('Delayed coming to hospital because of cost (share)', f'=AVERAGE({col("delay_cost")})', lambda rs: mean_of(rs, 'delay_cost'), '0.0%')
    add('Skipped a test because of cost (n)', f'=COUNTIF({col("skipped_test")},1)', lambda rs: sum(r['skipped_test'] == 1 for r in rs), '0')
    add('Any cost-related coping (delay/fewer medicines/skipped test/left early)', f'=AVERAGE({col("any_cost_coping")})', lambda rs: mean_of(rs, 'any_cost_coping'), '0.0%')
    add('Median work days lost', f'=MEDIAN(IF({rng("days_lost")}<>"",{rng("days_lost")}))', lambda rs: med_of(rs, 'days_lost'), '0', array=True)
    add('Job lost during illness (n)', f'=COUNTIF({col("job_loss")},1)', lambda rs: sum(r['job_loss'] == 1 for r in rs), '0')
    add('Would return to the same hospital next time (share)', f'=COUNTIF({col("next_choice")},1)/(COUNTIFS({col("next_choice")},"<>9",{col("next_choice")},"<>")-1)',
        lambda rs: sum(r['next_choice'] == 1 for r in rs) / sum(r['next_choice'] not in (9, None) for r in rs), '0.0%')
    add('... because it is cheap / government (share of those giving a reason)', f'=COUNTIF({col("next_reason")},3)/COUNT({col("next_reason")})',
        lambda rs: share(rs, 'next_reason', 3), '0.0%')

    section('H. Median medical OOP by ward and condition (usable forms)')
    for code in range(1, 8):
        add(f'{WARDS[code]} (n = {sum(1 for r in usable if r["ward"] == code)})', f'=MEDIAN(IF(({rng("usable_cost")}=1)*({rng("ward")}={code}),{rng("oop_medical")}))',
            lambda rs, code=code: med_of([r for r in U(rs) if r['ward'] == code], 'oop_medical'), array=True)
    add('Dialysis patients', f'=MEDIAN(IF(({rng("usable_cost")}=1)*({rng("dialysis")}=1),{rng("oop_medical")}))',
        lambda rs: med_of([r for r in U(rs) if r['dialysis'] == 1], 'oop_medical'), array=True)
    for code, lab in [(1, 'Kidney failure / dialysis'), (2, 'CKD / kidney damage'), (3, 'Glomerular'), (4, 'Stone'), (5, 'Obstruction / hydronephrosis'),
                      (6, 'Tumour / cancer'), (7, 'Infection'), (8, 'Other / undiagnosed')]:
        add(f'Condition: {lab}', f'=MEDIAN(IF(({rng("usable_cost")}=1)*({rng("cond_cat")}={code}),{rng("oop_medical")}))',
            lambda rs, code=code: med_of([r for r in U(rs) if r['cond_cat'] == code], 'oop_medical'), array=True)
        add(f'   n with {lab.lower()}', f'=COUNTIFS({col("cond_cat")},{code},{col("usable_cost")},1)',
            lambda rs, code=code: sum(r['cond_cat'] == code for r in U(rs)), '0')

    r_i = 4
    hdr = ['Indicator', '', 'Live value (all rows)'] + [f'Snapshot: {g}' for g, _ in GROUPS]
    for j, h in enumerate(hdr, start=1):
        ws.cell(row=r_i, column=j, value=h)
    style_header(ws, r_i, len(hdr))
    r_i += 1
    snap_rows = []
    for label, formula, snaps, fmt, array in table:
        if formula is None and snaps is None:
            ws.cell(row=r_i, column=1, value=label).font = F_SUB
            r_i += 1
            continue
        ws.cell(row=r_i, column=1, value=label).font = F_BASE
        if formula:
            ref = f'C{r_i}'
            if array:
                ws[ref] = ArrayFormula(ref, formula)
            else:
                ws[ref] = formula
            ws[ref].number_format = fmt
            ws[ref].font = F_BASE
        else:
            ws.cell(row=r_i, column=3, value='(see snapshot)').font = F_NOTE
        for k, v in enumerate(snaps):
            c = ws.cell(row=r_i, column=4 + k, value=v if v is not None else '-')
            c.number_format = fmt
            c.font = Font(name=FONT, size=10, color='595959')
            c.alignment = Alignment(horizontal='right')
        snap_rows.append((label, snaps, fmt))
        r_i += 1
    ws.cell(row=r_i + 1, column=1,
            value=('Caution: convenience samples at one public tertiary kidney hospital (inpatients, then outpatients). They describe the high-cost '
                   'tail of Bangladeshi health spending, not the average household - use HIES 2022 for national/district incidence. '
                   'Outpatient costs are for the whole illness so far, not just today\'s visit.')).font = F_NOTE
    set_widths(ws, [70, 2, 16, 16, 18, 18])
    ws.freeze_panes = 'A5'
    SNAPSHOT = snap_rows

    # ============ Sensitivity
    ws = wb.create_sheet('Sensitivity')
    ws['A1'] = 'Sensitivity of headline numbers to doubtful wave 2 forms'
    ws['A1'].font = F_TITLE
    ws['A2'] = 'Each row re-computes the headline figures after one change. Fixed values (not live formulas).'
    ws['A2'].font = F_NOTE

    def headline(rs):
        u = [r for r in rs if r['usable_cost'] == 1]
        c = [r for r in u if r['cat10'] is not None]
        return [len(u), st.median(r['oop_medical'] for r in u), st.mean(r['oop_medical'] for r in u),
                st.mean(r['cat10'] for r in c), st.mean(r['cat25'] for r in c),
                st.mean(r['fin_distress'] for r in rs if r['fin_distress'] is not None),
                st.mean(r['delay_cost'] for r in rs if r['delay_cost'] is not None)]

    def recode(rs, serial, **chg):
        out = []
        for r in rs:
            if r['serial'] == serial:
                r = dict(r)
                fixed12 = chg.pop('oop_12m', None)
                r.update(chg)
                chg['oop_12m'] = fixed12
                if r['monthly_cons'] and r['usable_cost'] == 1:
                    r['oop_12m'] = fixed12 if fixed12 is not None else r['oop_medical'] * min(1, 12 / r['care_months'])
                    r['oop_share'] = r['oop_12m'] / (12 * r['monthly_cons'])
                    r['cat10'] = int(r['oop_share'] > 0.10)
                    r['cat25'] = int(r['oop_share'] > 0.25)
            out.append(r)
        return out

    op = [r for r in rows if r['visit_type'] == 2]
    scen = [
        ('Base case: all forms', rows, op),
        ('Drop outlier A (largest single cost)', [r for r in rows if r['serial'] != '85'], [r for r in op if r['serial'] != '85']),
        ('Outlier B: count only the current illness, not earlier unrelated care', recode(rows, '88', oop_medical=100000), recode(op, '88', oop_medical=100000)),
        ('One form: use the stated total instead of the item sum', recode(rows, '62', oop_medical=65000), recode(op, '62', oop_medical=65000)),
        ('One long illness: use the reported 12-month cost, not an even spread', recode(rows, '68', oop_12m=250000),
         recode(op, '68', oop_12m=250000)),
        ('Drop 3 forms with known under-reporting', [r for r in rows if r['serial'] not in ('69', '91', '93')],
         [r for r in op if r['serial'] not in ('69', '91', '93')]),
        ('Drop outliers A and B together', [r for r in rows if r['serial'] not in ('85', '88')], [r for r in op if r['serial'] not in ('85', '88')]),
    ]
    hdr = ['Scenario', 'Sample', 'Usable n', 'Median OOP', 'Mean OOP', 'Cat >10%', 'Cat >25%', 'Distress financing', 'Cost delay']
    for j, h in enumerate(hdr, start=1):
        ws.cell(row=4, column=j, value=h)
    style_header(ws, 4, len(hdr))
    rr = 5
    SENS = []
    for lab, all_rs, op_rs in scen:
        for samp, rs in (('All 99', all_rs), ('Outpatients', op_rs)):
            vals = headline(rs)
            SENS.append((lab, samp, vals))
            ws.cell(row=rr, column=1, value=lab if samp == 'All 99' else '').font = F_BASE
            ws.cell(row=rr, column=2, value=samp).font = F_BASE
            for j, (v, fmt) in enumerate(zip(vals, ['0', '#,##0', '#,##0', '0.0%', '0.0%', '0.0%', '0.0%']), start=3):
                c = ws.cell(row=rr, column=j, value=v)
                c.number_format = fmt
                c.font = F_BASE
            rr += 1
    ws.cell(row=rr + 1, column=1, value=('Reading: medians barely move; means are driven by two outliers. Report medians and shares, '
                                         'and footnote the two extreme forms.')).font = F_NOTE
    set_widths(ws, [62, 13, 9, 12, 12, 10, 10, 16, 11])

    # ============ H3_Test
    ws = wb.create_sheet('H3_Test')
    reg_rows = [dict(r, ln_nights=math.log(max(r['nights_current'], 1)), uro=int(r['specialty'] == 1), ln_care=math.log(max(r['care_months'], 0.5)),
                     outpatient=int(r['visit_type'] == 2))
                for r in rows if r['usable_cost'] == 1 and r['oop_medical'] > 0]
    controls = ['outpatient', 'dialysis', 'paediatric', 'uro', 'ln_care', 'ln_nights', 'spend_bracket']
    models = [
        ('M1: pharmacy / informal first (H3 as written)', ['first_informal'] + controls),
        ('M2: number of providers before NIKDU (fragmentation)', ['n_providers'] + controls),
        ('M3: private clinic first (vs government / direct)', ['first_private'] + controls),
        ('M4: self-medicated before a prescription', ['premed'] + controls),
    ]
    ws['A1'] = 'Hypothesis 3 - what waves 1 + 2 can say'
    ws['A1'].font = F_TITLE
    inf_serials = ', '.join(r['serial'] for r in rows if r['first_informal'] == 1)
    notes = [
        'H3: patients who first go to a pharmacy or unqualified provider spend more in total than those who go straight to a qualified provider.',
        f'Waves 1 + 2: only {sum(r["first_informal"] == 1 for r in rows)} of {N} respondents went first to a pharmacy or informal provider (serials {inf_serials}). '
        'Adding 40 outpatients added just one such case, so M1 still rests on a handful of people and cannot test H3.',
        'Why: both waves are at a national referral hospital. Patients reach NIKDU via government or private hospitals; people who stay with pharmacies '
        'are filtered out before they get here. Outpatients were expected to differ, but 39 of 40 also came through formal providers.',
        'What the data CAN test: "fragmented, multi-stop care raises costs" (M2), private-first vs public-first (M3), and self-medication (M4). '
        'All models now include an outpatient dummy, because outpatient illnesses cost far less.',
        'Dependent variable: ln(medical OOP). Coefficients are approx. % differences (exact: exp(b)-1). Robust (HC1) standard errors = Stata vce(robust). '
        'ln(months of care) is controlled so longer illnesses are not mistaken for costlier pathways. Exploratory only.',
    ]
    for i, t in enumerate(notes, start=2):
        c = ws.cell(row=i, column=1, value=t)
        c.font = F_BASE if i > 2 else F_BOLD
        c.alignment = WRAP
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=7)
        ws.row_dimensions[i].height = 30
    rr = 8
    ws.cell(row=rr, column=1, value='Descriptive comparison (usable forms, medical OOP, BDT)').font = F_SUB
    rr += 1
    for j, h in enumerate(['Group', 'n (all)', 'Median OOP', 'Mean OOP', 'n inpatients', 'Median IP', 'n outpatients', 'Median OP'], start=1):
        ws.cell(row=rr, column=j, value=h)
    style_header(ws, rr, 8)
    rr += 1
    groups = [('Pharmacy / village doctor first', lambda r: r['first_informal'] == 1),
              ('Community clinic first', lambda r: r['first_care'] == 3),
              ('Government hospital first', lambda r: r['first_care'] == 4),
              ('Private clinic / hospital first', lambda r: r['first_care'] == 5),
              ('Came to NIKDU directly', lambda r: r['first_care'] == 6),
              ('0-1 providers before NIKDU', lambda r: r['n_providers'] is not None and r['n_providers'] <= 1),
              ('2 providers before NIKDU', lambda r: r['n_providers'] == 2),
              ('3+ providers before NIKDU', lambda r: r['n_providers'] is not None and r['n_providers'] >= 3),
              ('Self-medicated before prescription', lambda r: r['premed'] == 1),
              ('Did not self-medicate', lambda r: r['premed'] == 0)]
    for lab, f in groups:
        v = [r['oop_medical'] for r in reg_rows if f(r)]
        vo = [r['oop_medical'] for r in reg_rows if f(r) and r['outpatient'] == 1]
        vi = [r['oop_medical'] for r in reg_rows if f(r) and r['outpatient'] == 0]
        for j, val in enumerate([lab, len(v), st.median(v) if v else None, st.mean(v) if v else None,
                                 len(vi), st.median(vi) if vi else None, len(vo), st.median(vo) if vo else None], start=1):
            c = ws.cell(row=rr, column=j, value=val)
            c.font = F_BASE
            if j in (3, 4, 6, 8):
                c.number_format = '#,##0'
        rr += 1
    rr += 1
    reg_out = []
    for title, xs in models:
        res, n, r2 = ols_hc1(reg_rows, 'ln_oop_medical', xs)
        reg_out.append((title, res, n, r2))
        ws.cell(row=rr, column=1, value=f'{title}   (n = {n}, R2 = {r2:.2f})').font = F_SUB
        rr += 1
        for j, h in enumerate(['Variable', 'Coef. (log)', 'Robust SE', 'p-value', 'approx. % effect', 'Reading'], start=1):
            ws.cell(row=rr, column=j, value=h)
        style_header(ws, rr, 6)
        rr += 1
        for name, b, se, p in res:
            reading = ''
            if name == xs[0]:
                reading = 'not distinguishable from zero' if p >= 0.10 else ('weak evidence' if p >= 0.05 else 'statistically significant at 5%')
            vals = [name, b, se, p, (math.exp(b) - 1) if name != 'constant' else None, reading]
            for j, val in enumerate(vals, start=1):
                c = ws.cell(row=rr, column=j, value=val)
                c.font = F_BOLD if name == xs[0] else F_BASE
                if j in (2, 3):
                    c.number_format = '0.000'
                elif j == 4:
                    c.number_format = '0.000'
                elif j == 5:
                    c.number_format = '0%'
            rr += 1
        rr += 1
    ws.cell(row=rr, column=1, value=('Controls: outpatient, dialysis, paediatric, urology (vs nephrology), ln(months of care), ln(nights this admission; 0 for outpatients), '
                                     'spending bracket (1-5, linear). Re-run with survey_analysis_v2.do after each new wave.')).font = F_NOTE
    set_widths(ws, [52, 12, 12, 12, 14, 32, 13, 12])

    # ============ Fieldwork_Tracker
    ws = wb.create_sheet('Fieldwork_Tracker')
    ws['A1'] = 'Fieldwork tracker - live counts from Survey_Data'
    ws['A1'].font = F_TITLE
    hdr = ['Stratum', 'Hospital code', 'Visit type code', 'Target', 'Done (live)', 'Remaining', 'Suggested dates']
    for j, h in enumerate(hdr, start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, len(hdr))
    strata = [('NIKDU inpatients (wave 1)', 1, 1, 60, 'done 23-28 Sep (59 forms; 49, 50 missing)'),
              ('NIKDU outpatients (wave 2)', 1, 2, 40, 'done Oct (approx. 18 Oct)'),
              ('Private cardiology practice - inpatients (wave 3)', 3, 1, 25, 'to do - DISCLOSE relationship'),
              ('Private cardiology practice - outpatients (wave 3)', 3, 2, 25, 'to do - DISCLOSE relationship')]
    for i, (lab, h, v, tgt, dates) in enumerate(strata, start=4):
        ws.cell(row=i, column=1, value=lab)
        ws.cell(row=i, column=2, value=h)
        ws.cell(row=i, column=3, value=v)
        c = ws.cell(row=i, column=4, value=tgt)
        c.font = F_BLUE
        ws.cell(row=i, column=5, value=f'=COUNTIFS({col("hospital")},B{i},{col("visit_type")},C{i})')
        ws.cell(row=i, column=6, value=f'=MAX(D{i}-E{i},0)')
        ws.cell(row=i, column=7, value=dates)
    ws.cell(row=8, column=1, value='Total').font = F_BOLD
    for cl in 'DEF':
        ws[f'{cl}8'] = f'=SUM({cl}4:{cl}7)'
        ws[f'{cl}8'].font = F_BOLD
    ws.cell(row=8, column=7, value='Split of the 50 private-practice forms is a placeholder - edit D6:D7.').font = F_NOTE
    tips = [
        'Refusal tallies (written on the forms): forms 71-80: 13 refused (10/23 = 43% response); forms 81-90: 5 refused (10/15 = 67%). No tally for 1-70 or 91-100.',
        'Keep a tally sheet per day: approached / refused / completed - needed for the response rate in the write-up.',
        'Write the interview date on every form (still missing on all 99).',
        'Private cardiology practice: the physician is related to the author. State this in every report, and do not let the practice choose which patients are interviewed.',
        'Cardiology: ask chest-pain patients what they did first - antacids/"gas" medicine from a pharmacy is common before a cardiac diagnosis. This is where H3 cases are most likely.',
        'Medicines on a weekly/monthly rate: always write "x number of weeks/months" (6 outpatient forms needed an assumption).',
    ]
    for i, t in enumerate(tips, start=10):
        ws.cell(row=i, column=1, value=t).font = F_BASE
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=7)
    ws['D3'].comment = Comment('Targets: 60 NIKDU inpatients + 40 NIKDU outpatients + 50 private cardiology practice = 150.', 'plan')
    for row in ws.iter_rows(min_row=4, max_row=8):
        for c in row:
            if c.font.name != FONT:
                c.font = F_BASE
            c.border = BORDER
    set_widths(ws, [46, 13, 14, 9, 12, 11, 18])

    # ============ New_Entry
    ws = wb.create_sheet('New_Entry')
    ws['A1'] = ('Enter one form per row from row 4. Yellow = type here. Grey = derived (leave blank; survey_analysis.do computes them). '
                'Row 3 (green) is a worked example - delete it before pasting into Survey_Data.')
    ws['A1'].font = F_BOLD
    for j, s in enumerate(SPEC, start=1):
        c = ws.cell(row=2, column=j, value=s[0])
        c.comment = Comment(f'{s[1]}\n{s[2]}', 'codebook')
    style_header(ws, 2, len(SPEC))
    for j, s in enumerate(SPEC, start=1):
        if s[3] == 'der':
            ws.cell(row=2, column=j).fill = PatternFill('solid', start_color='7F7F7F')
    example = dict(id=N + 1, serial='101', wave=3, hospital=3, ward=11, visit_type=2, specialty=3, interview_date='03/11/2026', sex=2, age=46, resp_patient=1,
                   district='Tangail', residence=3, hh_size=5, occupation='Farmer', occ_cat=1, spend_bracket=2, spend_exact=None,
                   cond_text='Chest pain, angina (example)', cond_cat=9, dialysis=0, nights_current=0, nights_costed=0, onset_days=120,
                   first_care=1, care_path='Pharmacy (painkillers, 2 months) -> Tangail Sadar -> NIKDU OPD', n_providers=2, reason=1,
                   premed=1, premed_antibiotic=0, premed_painkiller=1, premed_other=0, cost_fees=0, cost_diag=6500, cost_med=2400,
                   cost_bed=0, cost_reported_total=None, oop_medical=8900, cost_transport=1800, cost_food=300, informal_pay=0,
                   scheme_paid=0, fin_savings=1, fin_borrow_rel=0, fin_borrow_lender=0, fin_borrow_unspec=0, fin_sold_asset=0,
                   fin_abroad=0, fin_income=0, days_lost=3, job_loss=0, delay=1, delay_reason=1, fewer_meds=0, skipped_test=0,
                   left_early=0, next_choice=1, next_reason=3, cost_scope=1, cost_quality=1, usable_cost=1, form_complete='Yes')
    for j, s in enumerate(SPEC, start=1):
        c = ws.cell(row=3, column=j, value=example.get(s[0]) if s[3] == 'in' else None)
        c.fill = FILL_EX
        c.font = F_BASE
    for i in range(4, 104):
        for j, s in enumerate(SPEC, start=1):
            c = ws.cell(row=i, column=j)
            c.fill = FILL_IN if s[3] == 'in' else FILL_DER
            c.border = BORDER
    lists = {'wave': '1,2,3', 'hospital': '1,2,3', 'ward': ','.join(str(k) for k in WARDS), 'visit_type': '1,2', 'sex': '1,2', 'resp_patient': '0,1',
             'residence': '1,2,3', 'occ_cat': '1,2,3,4,5,6,7,8', 'spend_bracket': '1,2,3,4,5', 'cond_cat': '1,2,3,4,5,6,7,8,9,10',
             'first_care': '1,2,3,4,5,6,7', 'reason': '1,2,3,4,5,6', 'delay_reason': '1,2,3,4', 'next_choice': '1,2,3,4,9',
             'next_reason': '1,2,3,4,5,6', 'cost_scope': '1,2', 'cost_quality': '1,2,3', 'form_complete': 'Yes,No'}
    binary = [s[0] for s in SPEC if s[2] == '0/1' and s[3] == 'in']
    for b in binary:
        lists[b] = '0,1'
    for name, opts in lists.items():
        dv = DataValidation(type='list', formula1=f'"{opts}"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f'{COL[name]}3:{COL[name]}103')
    dvd = DataValidation(type='list', formula1="=Atlas_Districts!$A$4:$A$67", allow_blank=True)
    ws.add_data_validation(dvd)
    dvd.add(f'{COL["district"]}3:{COL["district"]}103')
    ws.freeze_panes = 'C3'
    for j in range(1, len(SPEC) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 12
    ws.column_dimensions[COL['care_path']].width = 34
    ws.column_dimensions[COL['cond_text']].width = 26

    # ============ Atlas_Districts
    ws = wb.create_sheet('Atlas_Districts')
    ws['A1'] = 'District atlas table - fill yellow cells from HIES 2022 (survey-weighted), DGHS Health Bulletin, BDHS 2022 and Census 2022'
    ws['A1'].font = F_BOLD
    hdr = ['district', 'division', 'survey_n (live)', 'pop_2022', 'ind1_mean_monthly_oop_bdt', 'ind2_oop_share_consumption',
           'ind3a_cat10', 'ind3b_cat25', 'ind4_impoverished', 'ind5_doctors_nurses_per10k', 'ind6_beds_per10k',
           'ind7_pharmacy_first_child', 'ind8_medicine_share_of_oop', 'poverty_rate_upper', 'urban_share', 'source_notes']
    for j, h in enumerate(hdr, start=1):
        ws.cell(row=3, column=j, value=h)
    style_header(ws, 3, len(hdr))
    order = ['Barishal', 'Chattogram', 'Dhaka', 'Khulna', 'Mymensingh', 'Rajshahi', 'Rangpur', 'Sylhet']
    dists = sorted(DIVISION.items(), key=lambda kv: (order.index(kv[1]), kv[0]))
    for i, (dname, div) in enumerate(dists, start=4):
        ws.cell(row=i, column=1, value=dname).font = F_BASE
        ws.cell(row=i, column=2, value=div).font = F_BASE
        ws.cell(row=i, column=3, value=f'=COUNTIF({col("district")},A{i})').font = F_BASE
        for j in range(4, len(hdr) + 1):
            c = ws.cell(row=i, column=j)
            c.fill = FILL_IN
            c.border = BORDER
            c.font = F_BASE
    ws['M3'].comment = Comment('Indicator 8 (travel time) is not in the released BDHS/HIES files. Suggested replacement: share of household OOP spent on medicines (HIES item codes) - it directly tests H2.', 'note')
    ws.cell(row=69, column=1, value='Total').font = F_BOLD
    ws['C69'] = '=SUM(C4:C67)'
    ws['C69'].font = F_BOLD
    ws.cell(row=70, column=1, value='Check: should equal forms coded').font = F_NOTE
    ws['C70'] = f'=COUNTA({col("district")})-1'
    ws.freeze_panes = 'C4'
    set_widths(ws, [18, 13, 13, 12, 16, 16, 12, 12, 14, 16, 13, 16, 16, 14, 12, 30])

    for sheet in wb.worksheets:
        base_font(sheet)
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT_XLSX)

    # ============ CSV
    with open(OUT_CSV, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(COLS)
        for row in rows:
            w.writerow(['' if row[c] is None else row[c] for c in COLS])

    write_do()

    # ============ console summary
    print('forms', N, 'usable', len(usable), 'single-episode', len(ep))
    for label, snaps, fmt in SNAPSHOT:
        print(f'{label[:70]:70s}', ' | '.join('-' if v is None else (f'{v:.3f}' if isinstance(v, float) and abs(v) < 50 else f'{v:,.0f}') for v in snaps))
    for title, res, n, r2 in reg_out:
        name, b, se, p = res[1]
        print(f'{title}: {name} b={b:.3f} se={se:.3f} p={p:.3f} n={n} r2={r2:.2f}')
    for lab, samp, vals in SENS:
        print('SENS', lab[:45], samp, vals)


def write_do():
    val_labels = {
        'wave': '1 "W1 NIKDU inpatients" 2 "W2 NIKDU outpatients" 3 "W3 private cardiology"',
        'hospital': '1 "NIKDU" 2 "NICVD" 3 "Private cardiology practice"',
        'ward': ' '.join(f'{k} "{v}"' for k, v in WARDS.items()),
        'visit_type': '1 "Inpatient" 2 "Outpatient"',
        'sex': '1 "Male" 2 "Female"',
        'residence': '1 "City" 2 "Town" 3 "Village"',
        'occ_cat': '1 "Agriculture/fishing" 2 "Wage/manual" 3 "Small business" 4 "Salaried/skilled" 5 "Driver/transport" 6 "Migrant abroad" 7 "Pension/govt" 8 "Other"',
        'spend_bracket': '1 "<15k" 2 "15-30k" 3 "30-60k" 4 "60-100k" 5 ">100k"',
        'cond_cat': '1 "Kidney failure-dialysis" 2 "CKD/kidney damage" 3 "Glomerular" 4 "Stone" 5 "Obstruction/hydronephrosis" 6 "Tumour/cancer" 7 "Infection" 8 "Other/undiagnosed" 9 "Heart (cardiology)" 10 "Cardiac surgery"',
        'first_care': '1 "Pharmacy" 2 "Village doctor/informal" 3 "Community clinic" 4 "Govt hospital" 5 "Private clinic/hospital" 6 "This hospital directly" 7 "Nothing"',
        'reason': '1 "Referred" 2 "Reputation" 3 "Closest" 4 "Cost" 5 "Doctor known" 6 "Own choice/other"',
        'delay_reason': '1 "Funds" 2 "Transport" 3 "Other (non-cost)" 4 "Not specified"',
        'next_choice': '1 "Same hospital" 2 "Other facility" 3 "Depends/doctor" 4 "Unsure" 9 "Missing"',
        'next_reason': '1 "Quality" 2 "Familiar" 3 "Cheap/govt" 4 "Doctor advice" 5 "Close" 6 "Other"',
        'cost_scope': '1 "Within 12 months" 2 "Longer than 12 months"',
        'cost_quality': '1 "Itemised" 2 "Partial" 3 "Total only"',
        'specialty': '1 "Urology" 2 "Nephrology" 3 "Cardiology"',
    }
    L = []
    L.append('* Health Atlas of Bangladesh - patient cost survey analysis')
    L.append(f'* v2 (waves 1 + 2, n = 99). Generated {BUILD_DATE}. Put this file next to survey_clean_v2.csv and run: do survey_analysis_v2.do')
    L.append('* CONFIDENTIAL row-level data: never commit the CSV or this log output with individual rows to a public place.')
    L.append('clear all')
    L.append('set more off')
    L.append('import delimited using "survey_clean_v2.csv", clear varnames(1) encoding("utf-8") stringcols(2)')
    L.append('')
    L.append('* ---- variable labels')
    for s in SPEC:
        L.append(f'label variable {s[0]} "{s[1][:80]}"')
    L.append('')
    L.append('* ---- value labels')
    for v, lab in val_labels.items():
        L.append(f'label define {v}_lbl {lab}, replace')
        L.append(f'label values {v} {v}_lbl')
    L.append('')
    L.append('* ---- recompute derived variables (so newly entered rows only need the "entered" columns)')
    L.append('replace specialty = cond(inlist(ward,1,3,5),1, cond(inlist(ward,2,4,6),2, cond(ward>=8 & ward<=11,3, specialty)))')
    L.append('replace paediatric = inlist(ward,1,2) if !missing(ward) & ward<=6')
    L.append('replace paediatric = (age<18) if !missing(age) & ward>6')
    L.append('replace monthly_cons = cond(!missing(spend_exact), spend_exact, cond(spend_bracket==1,12000, cond(spend_bracket==2,22500, cond(spend_bracket==3,45000, cond(spend_bracket==4,80000, cond(spend_bracket==5,120000,.))))))')
    L.append('replace chronic = onset_days>=90 if !missing(onset_days)')
    L.append('replace first_informal = inlist(first_care,1,2) if !missing(first_care)')
    L.append('replace first_informal = . if missing(first_care)')
    L.append('replace first_private = cond(missing(first_care), ., first_care==5)')
    L.append('replace first_direct = cond(missing(first_care), ., first_care==6)')
    L.append('replace referred = reason==1 if !missing(reason)')
    L.append('egen _items = rowtotal(cost_fees cost_diag cost_med cost_bed), missing')
    L.append('replace cost_items_sum = _items')
    L.append('drop _items')
    L.append('egen _nonmed = rowtotal(cost_transport cost_food)')
    L.append('replace oop_nonmedical = _nonmed')
    L.append('drop _nonmed')
    L.append('replace oop_total = oop_medical + oop_nonmedical')
    L.append('replace fin_distress = (fin_borrow_rel==1 | fin_borrow_lender==1 | fin_borrow_unspec==1 | fin_sold_asset==1) if !missing(fin_savings)')
    L.append('replace delay_cost = cond(missing(delay), ., (delay==1 & inlist(delay_reason,1,4)))')
    L.append('replace any_cost_coping = cond(missing(delay), ., (delay_cost==1 | fewer_meds==1 | skipped_test==1 | left_early==1))')
    L.append('replace cost_scope = cond(care_months>12,2,1) if !missing(care_months)')
    L.append('replace oop_per_month = oop_medical/max(care_months,1)')
    L.append('replace oop_12m = oop_medical*min(1,12/care_months)')
    L.append('replace oop_share = oop_12m/(12*monthly_cons) if usable_cost==1 & !missing(monthly_cons)')
    L.append('replace share_illness = oop_medical/(12*monthly_cons) if usable_cost==1 & !missing(monthly_cons)')
    L.append('replace cat25_illness = share_illness>0.25 if !missing(share_illness)')
    L.append('replace cat10 = oop_share>0.10 if !missing(oop_share)')
    L.append('replace cat25 = oop_share>0.25 if !missing(oop_share)')
    L.append('replace ln_oop_medical = ln(oop_medical) if oop_medical>0 & !missing(oop_medical)')
    L.append('replace ln_oop_total = ln(oop_total) if oop_total>0 & !missing(oop_total)')
    L.append('gen byte uro = specialty==1')
    L.append('gen ln_care = ln(max(care_months,0.5))')
    L.append('gen ln_nights = ln(max(nights_current,1))')
    L.append('gen byte outpatient = visit_type==2')
    L.append('')
    L.append('* ---- Phase 1 descriptives')
    L.append('tab ward visit_type')
    L.append('tab wave')
    L.append('tabstat oop_medical oop_total cost_transport if usable_cost==1, by(cost_scope) stat(n mean p50 p25 p75) format(%12.0fc)')
    L.append('tabstat oop_medical if usable_cost==1, by(ward) stat(n mean p50) format(%12.0fc)')
    L.append('* Inpatients vs outpatients')
    L.append('tabstat oop_medical oop_total if usable_cost==1, by(visit_type) stat(n mean p25 p50 p75) format(%12.0fc)')
    L.append('tabstat oop_share if usable_cost==1, by(visit_type) stat(n p25 p50 p75) format(%5.3f)')
    L.append('tabstat cat10 cat25 fin_distress delay_cost any_cost_coping, by(visit_type) stat(n mean) format(%5.3f)')
    L.append('ranksum oop_medical if usable_cost==1, by(visit_type)')
    L.append('tabstat oop_medical if usable_cost==1, by(cond_cat) stat(n p50) format(%12.0fc)')
    L.append('summarize cat10 cat25 if usable_cost==1')
    L.append('summarize cat10 cat25 cat25_illness if usable_cost==1 & cost_scope==1')
    L.append('tab first_care')
    L.append('summarize fin_* delay_cost any_cost_coping')
    L.append('* H2: shares of medical OOP (fully itemised forms)')
    L.append('preserve')
    L.append('keep if usable_cost==1 & cost_quality==1')
    L.append('collapse (sum) cost_fees cost_diag cost_med cost_bed cost_items_sum')
    L.append('foreach v in cost_fees cost_diag cost_med cost_bed {')
    L.append('    display "`v\' share = " %5.3f `v\'/cost_items_sum')
    L.append('}')
    L.append('restore')
    L.append('')
    L.append('* ---- Phase 2: survey regressions (H3). Robust SEs.')
    L.append('local ctrl outpatient dialysis paediatric uro ln_care ln_nights spend_bracket')
    L.append('reg ln_oop_medical first_informal `ctrl\' if usable_cost==1, vce(robust)')
    L.append('reg ln_oop_medical n_providers `ctrl\' if usable_cost==1, vce(robust)')
    L.append('reg ln_oop_medical first_private `ctrl\' if usable_cost==1, vce(robust)')
    L.append('reg ln_oop_medical premed `ctrl\' if usable_cost==1, vce(robust)')
    L.append('* Outpatients only (wave 2)')
    L.append('reg ln_oop_medical n_providers dialysis paediatric uro ln_care spend_bracket if usable_cost==1 & outpatient==1, vce(robust)')
    L.append('* Sensitivity: drop the two extreme-cost forms (outliers A and B; serials listed in the private data notes)')
    L.append('reg ln_oop_medical n_providers `ctrl\' if usable_cost==1 & !inlist(serial,"85","88"), vce(robust)')
    L.append('* After the private cardiology wave: add hospital dummy and condition fixed effects')
    L.append('* reg ln_oop_medical first_informal n_providers i.hospital i.visit_type i.cond_cat ln_nights i.spend_bracket ln_care if usable_cost==1, vce(robust)')
    L.append('* Robustness: single episodes only; total cost incl. transport; drop top 5% of OOP')
    L.append('* reg ln_oop_medical n_providers `ctrl\' if usable_cost==1 & cost_scope==1, vce(robust)')
    L.append('* reg ln_oop_total n_providers `ctrl\' if usable_cost==1, vce(robust)')
    L.append('* Catastrophic spending (linear probability model)')
    L.append('* reg cat25 n_providers first_informal `ctrl\' if usable_cost==1, vce(robust)')
    with open(OUT_DO, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')


if __name__ == '__main__':
    main()
