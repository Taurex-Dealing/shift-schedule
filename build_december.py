"""
Build Dec26 from the Excel master and add it to the live shift-schedule tool.

Staffing: same 12 as Nov26 (row layout identical to November26). 1 Dec 2026 = Tuesday:
Mon 30 Nov lives on the November26 sheet, so the first block is [_, 1..6]; the last block
spills Fri 1 - Sun 3 Jan (nulled here). No previous-sheet weekend pickup needed this month.

Also patches BOTH HTML files with the new month plumbing:
  month button + currentMonth default + getMonthInfo entry + embedded data + BUILD_TS.

Run modes:
  python build_december.py            -> DRY RUN: extract + validate, no writes
  python build_december.py --write    -> write data.json + v.txt + patch both HTML files
"""
import json, time, re, os, sys
import openpyxl

REPO = r'C:\Users\SeanConway\td-schedule'
XLSX = os.environ.get('SCHED_XLSX', r'C:\Global Management Folder\Compenstaions calculations 2026 new.xlsx')
DJ   = os.path.join(REPO, 'data.json')

# Row layout (same as Sep-Nov sheets): blanks at offsets 6, 10, 13 (departed staff).
EXCEL_ORDER = ["Hristo","Victor","Stanimir","Ognyan","Petar",None,
               "Matthew","Win","Sarang",None,"Adam Nor","Afnan",
               None,"Andrey","Sean"]
OUTPUT_ORDER = [n for n in EXCEL_ORDER if n]

# hdr row -> (label, expected December dates Mon..Sun; None = Nov / Jan spillover)
WEEKS = [
    (3,  "Dec 1-6",   [None, 1, 2, 3, 4, 5, 6]),
    (19, "Dec 7-13",  [7, 8, 9, 10, 11, 12, 13]),
    (35, "Dec 14-20", [14, 15, 16, 17, 18, 19, 20]),
    (51, "Dec 21-27", [21, 22, 23, 24, 25, 26, 27]),
    (67, "Dec 28-31", [28, 29, 30, 31, None, None, None]),
]
EXPECT_STAFF, EXPECT_WEEKEND, EXPECT_WEEKDAYS = 12, 8, 23

ROLE = {
    "Morning (7 MYT)": "night",
    "Morning Shift":   "morning",
    "Evening Shift":   "afternoon",
}

def clean(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s != '' else None

def extract():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb['December26']
    weeks, weekend_shifts, problems = [], [], []

    for hdr, label, dates in WEEKS:
        # sanity: Excel header day numbers must match expected December dates
        for i, c in enumerate(range(2, 9)):
            v = ws.cell(hdr, c).value
            v = int(v) if isinstance(v, (int, float)) else None
            if dates[i] is not None and v != dates[i]:
                raise SystemExit(f"Header row {hdr} col {c}: expected {dates[i]}, Excel has {v}")
        staff = []
        wknd = {5: {}, 6: {}}
        for off in range(1, 16):
            r = hdr + off
            name = clean(ws.cell(r, 1).value)
            if name != EXCEL_ORDER[off-1]:
                raise SystemExit(f"Row {r}: expected {EXCEL_ORDER[off-1]!r}, got {name!r}")
            if name is None:
                continue
            shifts = []
            for i, c in enumerate(range(2, 9)):
                if dates[i] is None:
                    shifts.append(None)  # Nov/Jan spillover -> null regardless of cell content
                else:
                    shifts.append(clean(ws.cell(r, c).value))
            for idx in (5, 6):
                if dates[idx] is not None and shifts[idx] is not None:
                    wknd[idx][name] = shifts[idx]
            staff.append({"name": name, "shifts": shifts})
        weeks.append({"week": label, "dates": dates, "staff": staff})
        for idx in (5, 6):
            if dates[idx] is None:
                continue
            slot = {"night": "", "morning": "", "afternoon": ""}
            for nm, sh in wknd[idx].items():
                role = ROLE.get(sh)
                if role is None:
                    problems.append(f"{label} {dates[idx]:02d}.12: unmapped weekend shift {sh!r} for {nm}")
                    continue
                if slot[role]:
                    problems.append(f"{label} {dates[idx]:02d}.12: role {role} double-booked ({slot[role]} & {nm})")
                slot[role] = nm
            weekend_shifts.append({"date": f"{dates[idx]:02d}.12", **slot})
    return {"weeks": weeks, "staff_order": OUTPUT_ORDER, "weekend_shifts": weekend_shifts}, problems

def validate(new):
    print("=== VALIDATION ===")
    ok = True
    sp = 0
    for w in new['weeks']:
        dts = [d for d in w['dates'] if d is not None]
        if dts != sorted(dts):
            print(f"  !! spillover/order issue in {w['week']}: {w['dates']}")
            sp += 1
    alld = sorted(d for w in new['weeks'] for d in w['dates'] if d is not None)
    if alld != list(range(1, 32)):
        print(f"  !! December dates not exactly 1..31: {alld}")
        sp += 1
    print(f"  spillover check: {'PASS' if sp==0 else 'FAIL'}")
    ok &= sp == 0
    n = len(new['staff_order'])
    print(f"  staff_order count: {n} (expect {EXPECT_STAFF}): {'PASS' if n==EXPECT_STAFF else 'FAIL'}")
    ok &= n == EXPECT_STAFF
    wk = len(new['weekend_shifts'])
    print(f"  weekend_shifts entries: {wk} (expect {EXPECT_WEEKEND}): {'PASS' if wk==EXPECT_WEEKEND else 'FAIL'}")
    ok &= wk == EXPECT_WEEKEND
    nights = [e['date'] for e in new['weekend_shifts'] if not e['night']]
    print(f"  every weekend day has a night (MYT) worker: {'PASS' if not nights else 'FAIL '+str(nights)}")
    ok &= not nights
    print(f"  per-staff weekday cell coverage (Mon-Fri non-null, expect {EXPECT_WEEKDAYS}):")
    for s_name in new['staff_order']:
        cnt = 0
        for w in new['weeks']:
            st = next((x for x in w['staff'] if x['name']==s_name), None)
            if st:
                cnt += sum(1 for v in st['shifts'][:5] if v is not None)
        flag = '' if cnt == EXPECT_WEEKDAYS else '   <-- CHECK'
        print(f"     {s_name:12} {cnt}{flag}")
        ok &= cnt == EXPECT_WEEKDAYS
    return ok

def patch_html(fp, data_json_str, ts):
    html = open(fp, encoding='utf-8').read()
    if 'data-month="Dec26"' not in html:
        html, c = re.subn(
            r'(<button class="month-btn)( active)?(" data-month="Nov26" onclick="switchMonth\(\'Nov26\'\)">Nov 2026</button>)',
            lambda m: m.group(1) + m.group(3) + '\n    <button class=\"month-btn active\" data-month=\"Dec26\" onclick=\"switchMonth(\'Dec26\')\">Dec 2026</button>',
            html, count=1)
        assert c == 1, f"{fp}: month button not inserted"
    if "var currentMonth = 'Dec26';" not in html:
        html, c = re.subn(r"var currentMonth = 'Nov26';", "var currentMonth = 'Dec26';", html)
        assert c == 1, f"{fp}: currentMonth not updated"
    if "monthKey === 'Dec26'" not in html:
        html, c = re.subn(
            r"(  if \(monthKey === 'Nov26'\))",
            lambda m: "  if (monthKey === 'Dec26') return { year: 2026, month: 11, name: 'December 2026', days: 31 };\n" + m.group(1),
            html, count=1)
        assert c == 1, f"{fp}: getMonthInfo not updated"
    html, c = re.subn(r'var EMBEDDED_DATA = \{.*?\n\};',
                      lambda m: 'var EMBEDDED_DATA = ' + data_json_str + ';', html, count=1, flags=re.DOTALL)
    assert c == 1, f"{fp}: EMBEDDED_DATA not replaced"
    html, c = re.subn(r"var BUILD_TS = '\d+';", f"var BUILD_TS = '{ts}';", html)
    assert c >= 1, f"{fp}: BUILD_TS not replaced"
    open(fp, 'w', encoding='utf-8').write(html)
    print(f"{os.path.basename(fp)} patched (button + currentMonth + getMonthInfo + data + BUILD_TS={ts})")

if __name__ == '__main__':
    new, problems = extract()
    if problems:
        print("!!! WEEKEND ROLE PROBLEMS:")
        for p in problems: print("   ", p)
        print()
    print(json.dumps(new['weekend_shifts'], indent=1))
    ok = validate(new)

    if '--write' in sys.argv:
        if problems or not ok:
            raise SystemExit("Refusing to write: resolve problems above first.")
        if 'SCHED_XLSX' in os.environ:
            raise SystemExit("Refusing to write: SCHED_XLSX override is for dry-runs only.")
        data = json.load(open(DJ, encoding='utf-8'))
        if 'Dec26' in data:
            print("NOTE: Dec26 already in data.json - replacing it.")
        data['Dec26'] = new
        json.dump(data, open(DJ, 'w', encoding='utf-8'), indent=2, ensure_ascii=False)
        print("\ndata.json written (Dec26 added; other months untouched).")
        ts = str(int(time.time()))
        open(os.path.join(REPO, 'v.txt'), 'w').write(ts)
        print("v.txt:", ts)
        data_json_str = json.dumps(data, indent=2, ensure_ascii=False)
        for hf in ['index.html', 'management.html']:
            patch_html(os.path.join(REPO, hf), data_json_str, ts)
        print("\n=== WRITE COMPLETE ===")
