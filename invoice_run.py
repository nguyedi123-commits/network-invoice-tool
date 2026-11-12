"""One run of the tool: uploaded invoice files -> Excel extract + remit files + Network Lines.csv, all in one zip.

Same steps as the old .exe (App/server.py process()); nothing is kept on the server.
"""
import csv
import io
import os
import re
import tempfile
import zipfile
from datetime import date, datetime

import nem_core as nc

LINES_COLUMNS = nc.NETWORK_LINE_COLUMNS + ['LINE_SEQ']


def safe_label(s):
    s = re.sub(r'[^A-Za-z0-9 ._-]', '', (s or '').strip())
    return s[:40] or 'Network Bills'


def _cell(v):
    if v is None:
        return ''
    if isinstance(v, (datetime, date)):
        return v.strftime('%Y-%m-%d')
    return v


def network_lines_csv(parsed):
    """Every charge line of the uploaded charge files (copy invoices left out, interest kept - the
    Snowflake view leaves it out). LINE_SEQ = order of the line in its file."""
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator='\n')
    w.writerow(LINES_COLUMNS)
    count = 0
    for pf in sorted(parsed, key=lambda p: (p.lnsp, p.file_ts)):
        for seq, r in enumerate(nc.network_lines(pf), 1):
            w.writerow([_cell(r.get(c)) for c in nc.NETWORK_LINE_COLUMNS] + [seq])
            count += 1
    return buf.getvalue(), count


def process_files(month, files):
    """month: label for the output names. files: list of (file name, file bytes).
    Returns dict: zip_name, zip_bytes, results (one per file), skipped, lines."""
    label = safe_label(month)
    parsed, problems, seen = [], [], set()
    # an original .csv wins over a copy of the same file without the extension
    for name, data in sorted(files, key=lambda f: (not f[0].lower().endswith('.csv'), f[0])):
        name = os.path.basename(name)
        base = re.sub(r'\.csv$', '', name, flags=re.I)
        if base in seen:
            problems.append((name, 'uploaded twice - second copy ignored'))
            continue
        seen.add(base)
        if '#NBREMITT#' in name.upper():
            problems.append((name, 'this is a remittance file (we send these), not an invoice file - skipped'))
            continue
        try:
            parsed.append(nc.parse(name, data.decode('utf-8-sig', errors='replace')))
        except ValueError as e:
            problems.append((name, str(e)))
        except Exception as e:  # unexpected layout
            problems.append((name, f'could not read file: {e}'))

    results, remits = [], {}
    for pf in sorted(parsed, key=lambda p: (p.lnsp, p.file_ts)):
        text = nc.remit_text(pf)
        if text:
            rname = nc.remit_name(pf)
            remits[rname] = text.encode('ascii', errors='replace')
        else:
            rname = 'No remit needed (GSL credit only)'
        rows = nc.excel_rows(pf, label)
        results.append({
            'Network': pf.lnsp, 'File': pf.file_name,
            'Type': 'Overdue notice' if pf.kind == 'REMIND' else ('Invoice - national' if pf.layout == 'NATIONAL' else 'Invoice - Vic/SA'),
            'Invoices': len(pf.invoices),
            'Ex GST': sum(r[7] for r in rows), 'GST': sum(r[8] for r in rows), 'Inc GST': sum(r[9] for r in rows),
            'GSL credit': float(pf.gsl_ex + pf.gsl_gst) if pf.gsl_lines() else None,
            'Remit pays': float(nc.remit_total_inc(pf)), 'Remit file': rname,
            'Invoice date': pf.invoice_date.strftime('%d/%m/%Y'),
            'Due date': pf.due_date.strftime('%d/%m/%Y') if pf.due_date else '',
            'Check': ' | '.join(pf.warnings) if pf.warnings else 'OK',
        })

    lines_text, lines = network_lines_csv(parsed)
    zbuf = io.BytesIO()
    folder = f'{label} Network Invoice'
    with zipfile.ZipFile(zbuf, 'w', zipfile.ZIP_DEFLATED) as z:
        if parsed:
            with tempfile.TemporaryDirectory() as tmp:
                xlsx = os.path.join(tmp, f'{label} Network Invoice.xlsx')
                nc.build_workbook(parsed, label, xlsx, problems)
                z.write(xlsx, f'{folder}/{label} Network Invoice.xlsx')
        for rname, data in remits.items():
            z.writestr(f'{folder}/Remit/{rname}', data)
        z.writestr(f'{folder}/Network Lines.csv', lines_text.encode('utf-8'))
    return {
        'zip_name': f'{folder}.zip', 'zip_bytes': zbuf.getvalue(), 'results': results,
        'skipped': [{'File': n, 'Reason': r} for n, r in problems], 'lines': lines, 'remits': len(remits),
    }
