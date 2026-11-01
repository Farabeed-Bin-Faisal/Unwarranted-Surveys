"""Make the de-identified public dataset (survey_public_v2.csv) from the confidential records.

Run in the private v2 folder:  python3 make_public_dataset.py
Output: v1 (wave 1 only) and v2 (waves 1 + 2) versions of survey_public_*.csv, public_codebook_*.md, public_check_*.txt

Rules (see public_codebook.md):
  - No serial numbers, no free text, no exact age, no exact household size or occupation. District is kept.
  - Bands from the survey form are used where the form has them: patient 18 or older (yes/no),
    city/town/village, the five monthly spending brackets, the form's answer lists (first care, reason,
    medicine type, financing, coping).
  - Where the form has no band, a band is set here: nights admitted,
    months of care, days of work lost, providers seen (capped at 4+).
  - Money: exact, as coded.
  - Rows are shuffled and given new ids, so row order says nothing about serial numbers.
  - District, division and region are kept on every row (the project is a district atlas).
  - Privacy check on the other traits: every row must match at least K = 3 rows on (visit type, sex, 18+,
    region, residence), counting a blank as matching anything. Rows in smaller groups lose residence,
    then 18+, then sex.
"""
import collections
import csv
import random
import statistics as st

import build_workbook as bw

K = 3
SEED = 20261101
QI = ['visit_type', 'sex', 'adult', 'region', 'residence']
SUPPRESS = ('residence', 'adult', 'sex')


def round_money(v):
    """Costs are released exactly as coded (no rounding, no cap), at the author's request."""
    return None if v is None else int(round(v))


def band(v, cuts, labels):
    """cuts are upper bounds (inclusive); last label is open-ended."""
    if v is None:
        return None
    for c, lab in zip(cuts, labels):
        if v <= c:
            return lab
    return labels[-1]


def public_row(r):
    nights = r['nights_current']
    return {
        'wave': r['wave'], 'visit_type': r['visit_type'], 'specialty': r['specialty'],
        'sex': r['sex'], 'adult': None if r['age'] is None else int(r['age'] >= 18),
        'resp_patient': r['resp_patient'],
        'region': 1 if r['division'] == 'Dhaka' else 2, 'division': r['division'], 'district': r['district'],
        'residence': r['residence'], 'spend_bracket': r['spend_bracket'],
        'cond_cat': r['cond_cat'], 'dialysis': r['dialysis'],
        'nights_band': band(nights, [0, 7, 14, 30], [0, 1, 2, 3, 4]),
        'chronic': r['chronic'],
        'care_months_band': band(r['care_months'], [1, 3, 6, 12, 24], [1, 2, 3, 4, 5, 6]),
        'cost_scope': r['cost_scope'],
        'first_care': r['first_care'], 'n_providers': None if r['n_providers'] is None else min(r['n_providers'], 4),
        'reason': r['reason'], 'referred': r['referred'],
        'premed': r['premed'], 'premed_antibiotic': r['premed_antibiotic'],
        'premed_painkiller': r['premed_painkiller'], 'premed_other': r['premed_other'],
        'cost_fees': round_money(r['cost_fees']), 'cost_diag': round_money(r['cost_diag']),
        'cost_med': round_money(r['cost_med']), 'cost_bed': round_money(r['cost_bed']),
        'oop_medical': round_money(r['oop_medical']),
        'cost_transport': round_money(r['cost_transport']), 'cost_food': round_money(r['cost_food']),
        'oop_total': round_money(r['oop_total']),
        'informal_pay': r['informal_pay'], 'scheme_paid': r['scheme_paid'],
        'fin_savings': r['fin_savings'], 'fin_borrow_rel': r['fin_borrow_rel'], 'fin_borrow_lender': r['fin_borrow_lender'],
        'fin_borrow_unspec': r['fin_borrow_unspec'], 'fin_sold_asset': r['fin_sold_asset'], 'fin_abroad': r['fin_abroad'],
        'fin_income': r['fin_income'], 'fin_distress': r['fin_distress'],
        'days_lost_band': band(r['days_lost'], [0, 30, 90, 365], [0, 1, 2, 3, 4]),
        'delay': r['delay'], 'delay_reason': r['delay_reason'], 'delay_cost': r['delay_cost'],
        'fewer_meds': r['fewer_meds'], 'skipped_test': r['skipped_test'], 'left_early': r['left_early'],
        'any_cost_coping': r['any_cost_coping'],
        'next_choice': r['next_choice'], 'next_reason': r['next_reason'],
        'cost_quality': r['cost_quality'], 'usable_cost': r['usable_cost'],
        'oop_share': r['oop_share'],
        'cat10': r['cat10'], 'cat25': r['cat25'], 'cat25_illness': r['cat25_illness'],
    }


CODEBOOK = [
    ('id', 'Public row id (random order; not the form serial number)', 'number'),
    ('wave', 'Survey wave', '1 NIKDU inpatients (Sep 2026); 2 NIKDU outpatients (Oct 2026)'),
    ('visit_type', 'Visit type', '1 inpatient; 2 outpatient'),
    ('specialty', 'Department', '1 urology; 2 nephrology'),
    ('sex', 'Patient sex', '1 male; 2 female; blank = suppressed for privacy'),
    ('adult', 'Patient is 18 or older (form screening question)', '1 yes; 0 no (child); blank = missing or suppressed'),
    ('resp_patient', 'Respondent is the patient', '1 patient; 0 attendant'),
    ('region', 'Household region', '1 Dhaka division; 2 outside Dhaka division'),
    ('division', 'Household division', 'name (8 divisions)'),
    ('district', 'Household district (BBS spelling)', 'name (all rows)'),
    ('residence', 'Residence type (form)', '1 city; 2 town; 3 village; blank = not answered or suppressed for privacy'),
    ('spend_bracket', 'Monthly household spending (form brackets)', '1 under 15,000; 2 15,000-30,000; 3 30,000-60,000; 4 60,000-100,000; 5 above 100,000 BDT'),
    ('cond_cat', 'Condition category', '1 kidney failure/dialysis; 2 CKD/kidney damage; 3 glomerular; 4 stone; 5 obstruction/hydronephrosis; 6 tumour/cancer; 7 infection; 8 other/undiagnosed'),
    ('dialysis', 'Dialysis during the illness', '0/1'),
    ('nights_band', 'Nights in current admission', '0 none (outpatient); 1 1-7; 2 8-14; 3 15-30; 4 31+'),
    ('chronic', 'Symptoms for 90+ days', '0/1'),
    ('care_months_band', 'Months over which the reported costs were spent', '1 up to 1; 2 1-3; 3 3-6; 4 6-12; 5 12-24; 6 over 24'),
    ('cost_scope', 'Length of the costed illness', '1 within 12 months; 2 longer than 12 months'),
    ('first_care', 'First point of care (form)', '1 pharmacy/drug seller; 2 village doctor/informal; 3 community clinic; 4 govt hospital; 5 private clinic/hospital; 6 this hospital directly; 7 nothing at first'),
    ('n_providers', 'Providers seen before this hospital', '0-3; 4 = four or more'),
    ('reason', 'Why this hospital (form)', '1 referred; 2 reputation; 3 closest; 4 cost; 5 doctor known; 6 own choice/other'),
    ('referred', 'Referred', '0/1'),
    ('premed', 'Took medicine before a prescription', '0/1'),
    ('premed_antibiotic', '... antibiotic', '0/1'),
    ('premed_painkiller', '... painkiller', '0/1'),
    ('premed_other', '... other', '0/1'),
    ('cost_fees', 'Doctor/consultation fees', 'BDT, exact as coded'),
    ('cost_diag', 'Diagnostic tests and imaging', 'BDT, exact as coded'),
    ('cost_med', 'Medicines', 'BDT, exact as coded'),
    ('cost_bed', 'Bed, procedure, operation, dialysis', 'BDT, exact as coded'),
    ('oop_medical', 'Medical out-of-pocket cost (analysis variable)', 'BDT, exact as coded'),
    ('cost_transport', 'Transport, all trips', 'BDT, exact as coded'),
    ('cost_food', 'Food and lodging for attendants', 'BDT, exact as coded'),
    ('oop_total', 'Medical + transport + food', 'BDT, exact as coded'),
    ('informal_pay', 'Unbilled payment reported', '0/1'),
    ('scheme_paid', 'Insurance/employer/NGO/govt scheme paid', '0/1'),
    ('fin_savings', 'Paid from savings', '0/1'),
    ('fin_borrow_rel', 'Borrowed from relatives/friends', '0/1'),
    ('fin_borrow_lender', 'Borrowed from lender/NGO/bank', '0/1'),
    ('fin_borrow_unspec', 'Borrowed, source not recorded', '0/1'),
    ('fin_sold_asset', 'Sold/pawned an asset', '0/1'),
    ('fin_abroad', 'Help from relatives abroad', '0/1'),
    ('fin_income', 'Paid from current income', '0/1'),
    ('fin_distress', 'Any borrowing or asset sale', '0/1'),
    ('days_lost_band', 'Work days lost (patient + attendants)', '0 none; 1 1-30; 2 31-90; 3 91-365; 4 over 365'),
    ('delay', 'Delayed coming to hospital', '0/1'),
    ('delay_reason', 'Reason for delay', '1 funds; 2 transport; 3 other (non-cost); 4 not specified'),
    ('delay_cost', 'Delay because of cost', '0/1'),
    ('fewer_meds', 'Bought fewer medicines than prescribed', '0/1'),
    ('skipped_test', 'Skipped a test', '0/1'),
    ('left_early', 'Left earlier than advised', '0/1'),
    ('any_cost_coping', 'Any of the four because of cost', '0/1'),
    ('next_choice', 'Where next time', '1 same hospital; 2 other facility; 3 depends/doctor; 4 unsure; 9 missing'),
    ('next_reason', 'Why', '1 quality; 2 familiar; 3 cheap/govt; 4 doctor advice; 5 close; 6 other'),
    ('cost_quality', 'Cost data quality', '1 fully itemised; 2 partly itemised/estimated; 3 total only'),
    ('usable_cost', 'Usable for cost analysis', '0/1'),
    ('oop_share', '12-month medical cost / annual household spending', 'fraction'),
    ('cat10', 'Catastrophic: share > 10%', '0/1'),
    ('cat25', 'Catastrophic: share > 25%', '0/1'),
    ('cat25_illness', 'Whole-illness cost > 25% of one year of spending', '0/1'),
]


WAVE_TEXT = {1: '59 NIKDU inpatients (wave 1, Sep 2026)', 2: '40 NIKDU outpatients (wave 2, Oct 2026)'}


def make(waves, tag):
    rows = [r for r in bw.build_rows() if r['wave'] in waves]
    pub = [public_row(r) for r in rows]

    # privacy check: k-anonymity with local suppression (sdcMicro convention: a blank matches any value).
    # The row in the smallest group loses residence, then district, 18+, sex, division, region.
    def fk(i):
        a = pub[i]
        return sum(1 for b in pub if all(a[q] is None or b[q] is None or a[q] == b[q] for q in QI))
    blanked = collections.Counter()
    while True:
        ks = [fk(i) for i in range(len(pub))]
        bad = [i for i, k in enumerate(ks) if k < K]
        if not bad:
            break
        i = min(bad, key=lambda j: ks[j])
        for f in SUPPRESS:
            if pub[i][f] is not None:
                pub[i][f] = None
                blanked[f] += 1
                break
        else:
            raise SystemExit(f'row {i} cannot reach k={K}')
    n_dist = sum(p['district'] is not None for p in pub)
    dcount = collections.Counter(p['district'] for p in pub)
    log = [f'{len(pub)} rows; district kept for all {n_dist}; {len(dcount)} districts, {sum(1 for v in dcount.values() if v == 1)} with only one patient',
           f'values blanked so every row matches at least {K} rows on {QI}: ' + ', '.join(f'{f} {n}' for f, n in blanked.items()),
           f'smallest group after suppression: {min(fk(i) for i in range(len(pub)))}']

    random.Random(SEED + len(pub)).shuffle(pub)
    cols = ['id'] + list(pub[0].keys())
    assert [c[0] for c in CODEBOOK] == cols, 'codebook out of sync'
    with open(f'survey_public_{tag}.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(cols)
        for i, p in enumerate(pub, start=1):
            p['id'] = i
            w.writerow(['' if p[c] is None else p[c] for c in cols])

    def stats(rs):
        u = [r for r in rs if r['usable_cost'] == 1]
        c = [r for r in u if r['cat10'] is not None]
        return [len(u), st.median(r['oop_medical'] for r in u), round(st.mean(x['cat10'] for x in c), 3), round(st.mean(x['cat25'] for x in c), 3),
                round(st.mean(r['fin_distress'] for r in rs if r['fin_distress'] is not None), 3),
                round(st.mean(r['delay_cost'] for r in rs if r['delay_cost'] is not None), 3)]
    lines = [f'Privacy check: survey_public_{tag}.csv'] + ['  ' + l for l in log] + ['',
             'Exact (confidential) vs public file: n usable, median medical OOP, cat >10%, cat >25%, borrowed/sold assets, cost delay']
    groups = [('all', lambda r: True)]
    if len(waves) > 1:
        groups += [('inpatients', lambda r: r['visit_type'] == 1), ('outpatients', lambda r: r['visit_type'] == 2)]
    for lab, f in groups:
        lines.append(f'  {lab:12s} exact  {stats([r for r in rows if f(r)])}')
        lines.append(f'  {lab:12s} public {stats([p for p in pub if f(p)])}')
    open(f'public_check_{tag}.txt', 'w').write('\n'.join(lines) + '\n')

    md = [f'# Codebook: survey_public_{tag}.csv', '',
          'De-identified data from the Health Atlas patient cost survey at NIKDU, Dhaka: '
          + ' + '.join(WAVE_TEXT[w] for w in waves) + f'. {len(pub)} rows, one per patient. Blank = missing.', '',
          '## What was done to protect respondents', '',
          '- Removed: serial numbers, interview notes and all free text (condition description, care path, occupation, '
          'attendant relation, answers as written), exact age, household size, occupation, exact spending, interview completion.',
          '- Kept: household district, division and region.',
          '- Bands from the survey form: patient 18 or older; city / town / village; the five monthly spending brackets; the form\'s answer lists.',
          '- Bands set for this file (the form has none): nights admitted; months of care; work days lost; providers seen capped at 4+.',
          '- Costs are exact, as coded from the forms (rules in the README). `oop_medical` can exceed the item sum when the respondent gave a larger all-inclusive total.',
          '- District is kept on every row. Some districts have only one or two patients; see public_check for counts.',
          f'- On the other traits, every row matches at least {K} patients on visit type, sex, 18+, region and residence (a blank counts as matching anything). '
          'Where a patient would otherwise stand out, residence, 18+ or sex was blanked on that row: '
          + ', '.join(f'{f} {n}' for f, n in blanked.items()) + '.',
          '- Rows are in random order with new ids.', '',
          '| Variable | Meaning | Codes / units |', '|---|---|---|']
    md += [f'| `{n}` | {l} | {c} |' for n, l, c in CODEBOOK]
    open(f'public_codebook_{tag}.md', 'w').write('\n'.join(md) + '\n')
    print('\n'.join(lines))


def main():
    make((1,), 'v1')
    make((1, 2), 'v2')


if __name__ == '__main__':
    main()
