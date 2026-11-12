"""Core logic: read NEM network billing files and build the Excel extract + remit files.

Handles:
  - NBCHARGES, national layout   (ACTEWP, AURORAP, CNRGYP, ENERGEXP, ENERGYAP, ERGONETP, INTEGP)
  - NBCHARGES, Vic/SA layout     (CITIPP, POWCP, SOLARISP, UNITED, EASTERN, UMPLP)
  - NBREMIND, overdue invoice notice (930/931/932)

Business rules agreed with the user (see memory note "network-bills-app-decisions"):
  - Invoice Date = date in the file header timestamp
  - Due Date     = latest due date on the invoices in the file
  - Remit pays every invoice in full; GSL credits (record 400) are NOT deducted
  - GSL credits get their own negative row in the Excel extract
  - Remit Paid Date is left blank
"""
import csv
import io
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation

RETAILER = 'EIABVIR'
CENT = Decimal('0.01')

EXCEL_COLUMNS = [
    'Vendor ID', 'Vendor Code', 'Vendor Name', 'PO #', 'File Name', 'Invoice No', 'File Location',
    'Amount (Ex GST)', 'GST', 'Amount (Inc GST)', 'Invoice Date', 'Due Date',
    'Receipt Acknowledged / PO created / GRd', 'Invoice Processed', 'Paid',
]

GSL_TEXT_RE = re.compile(r'\bGSL\b|GUARANTEED\s+(CUSTOMER\s+)?SERVICE|\bGCSS\b|PERFORMANCE\s+STANDARD', re.I)


def load_gsl_codes():
    import os
    import sys
    if getattr(sys, 'frozen', False):
        # .exe: use the editable list in App/ next to the exe, else the copy bundled inside the exe
        exe_dir = os.path.dirname(sys.executable)
        app_copy, local_copy = os.path.join(exe_dir, 'App', 'gsl_codes.txt'), os.path.join(exe_dir, 'gsl_codes.txt')
        bundled = os.path.join(getattr(sys, '_MEIPASS', exe_dir), 'gsl_codes.txt')
        if not os.path.exists(app_copy) and not os.path.exists(local_copy):
            # First run on a new PC: put an editable copy of the list next to the exe
            try:
                import shutil
                shutil.copyfile(bundled, local_copy)
            except OSError:
                pass
        candidates = [app_copy, local_copy, bundled]
    else:
        candidates = [os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gsl_codes.txt')]
    for path in candidates:
        try:
            with open(path, encoding='utf-8') as fh:
                return {l.strip().upper() for l in fh if l.strip() and not l.strip().startswith('#')}
        except OSError:
            continue
    return {'TASGSL1'}


GSL_CODES = load_gsl_codes()

NAME_RE = re.compile(r'^NEM#(NBCHARGES|NBREMIND)#([A-Z0-9]+)#([A-Z0-9]+)#(\d{14}V\d+)$', re.I)


def dec(s):
    s = (s or '').strip()
    if s == '':
        return Decimal('0')
    try:
        return Decimal(s)
    except InvalidOperation:
        raise ValueError(f'not a number: {s!r}')


def money(d):
    """Format a Decimal as a plain 2-decimal string (no thousands separator)."""
    return str(Decimal(d).quantize(CENT))


def is_date(s):
    return bool(re.fullmatch(r'20\d{6}', (s or '').strip()))


def to_date(s):
    return datetime.strptime(s[:8], '%Y%m%d')


@dataclass
class Invoice:
    number: str
    nmi: str
    checksum: str
    invoice_date: str
    due_date: str
    status: str
    ex: Decimal
    gst: Decimal
    inc: Decimal
    inc_text: str          # amount exactly as the network wrote it


@dataclass
class Line:
    record: str
    invoice: str
    line_id: str
    nmi: str
    checksum: str
    ex_text: str
    gst_text: str
    is_gsl: bool = False
    adj: str = ''
    detail: str = ''       # charge code / description fields, for the GSL list
    raw: list = field(default_factory=list)   # the full CSV row, for the Network Lines extract

    @property
    def ex(self):
        return dec(self.ex_text)

    @property
    def gst(self):
        return dec(self.gst_text)


@dataclass
class ParsedFile:
    file_name: str                 # without .csv
    kind: str                      # 'CHARGES' or 'REMIND'
    layout: str                    # 'NATIONAL', 'VICSA' or 'REMIND'
    lnsp: str
    header_ts: str                 # 14-digit timestamp from header row
    file_ts: str                   # e.g. 20261009003012V1 (from file name)
    invoices: list = field(default_factory=list)
    lines: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    # ---- totals -------------------------------------------------------
    def billable_invoices(self):
        return [i for i in self.invoices if 'COPY' not in i.status.upper()]

    @property
    def total_ex(self):
        return sum((i.ex for i in self.billable_invoices()), Decimal('0'))

    @property
    def total_gst(self):
        return sum((i.gst for i in self.billable_invoices()), Decimal('0'))

    @property
    def total_inc(self):
        return sum((i.inc for i in self.billable_invoices()), Decimal('0'))

    def gsl_lines(self):
        return [l for l in self.lines if l.is_gsl]

    def gsl_only_invoices(self):
        """Invoice numbers whose charge lines are all GSL credits (left out of the remit)."""
        lines_by_inv = {}
        for l in self.lines:
            lines_by_inv.setdefault(l.invoice, []).append(l)
        return {inv for inv, ls in lines_by_inv.items() if ls and all(l.is_gsl for l in ls)}

    @property
    def gsl_ex(self):
        return sum((l.ex for l in self.gsl_lines()), Decimal('0'))

    @property
    def gsl_gst(self):
        return sum((l.gst for l in self.gsl_lines()), Decimal('0'))

    @property
    def invoice_date(self):
        return to_date(self.header_ts)

    @property
    def due_date(self):
        if self.kind == 'REMIND':
            return self._remind_due
        dues = [i.due_date for i in self.billable_invoices() if is_date(i.due_date)]
        return to_date(max(dues)) if dues else None


def read_rows(text):
    rows = []
    for r in csv.reader(io.StringIO(text)):
        r = [c.strip() for c in r]
        while r and r[-1] == '':
            r.pop()
        if r:
            rows.append(r)
    return rows


def parse(file_name, text):
    """Parse one uploaded file. Returns ParsedFile, or raises ValueError with a plain message."""
    base = re.sub(r'\.csv$', '', file_name.strip(), flags=re.I)
    rows = read_rows(text.lstrip('﻿'))
    if not rows:
        raise ValueError('file is empty')
    first = rows[0][0]

    if first in ('800', '920'):
        raise ValueError('this is a remittance file (we send these), not an invoice file - skipped')
    if first not in ('010', '10', '930'):
        raise ValueError(f'unknown file type (first record is {first!r})')

    m = NAME_RE.match(base)
    hdr = rows[0]
    lnsp = hdr[1]
    header_ts = hdr[3] if len(hdr) > 3 else ''
    warnings = []
    if not re.fullmatch(r'\d{14}', header_ts):
        raise ValueError(f'header timestamp is not valid: {header_ts!r}')
    if hdr[2] != RETAILER:
        warnings.append(f'file is addressed to {hdr[2]}, not {RETAILER}')

    if m:
        file_ts = m.group(4)
        if m.group(2).upper() != lnsp.upper():
            warnings.append(f'file name says {m.group(2)} but file content says {lnsp}')
    else:
        file_ts = header_ts + 'V1'
        warnings.append('file name is not in the standard NEM#TYPE#FROM#TO#TIMESTAMP format; '
                        f'using header timestamp {file_ts} for the remit name')

    if first == '930':
        return _parse_remind(base, rows, lnsp, header_ts, file_ts, warnings)
    return _parse_charges(base, rows, lnsp, header_ts, file_ts, warnings)


def _parse_charges(base, rows, lnsp, header_ts, file_ts, warnings):
    inv020 = [r for r in rows if r[0] in ('020', '20')]
    if not inv020:
        raise ValueError('no invoice summary (020) rows found')
    vicsa = is_date(inv020[0][2])
    pf = ParsedFile(base, 'CHARGES', 'VICSA' if vicsa else 'NATIONAL', lnsp, header_ts, file_ts, warnings=warnings)

    for r in inv020:
        if vicsa:
            # 020,inv,invdate,due,retailer,abn,lnsp,abn,status,ex,gst,inc,firstline,lastline,count,...
            pf.invoices.append(Invoice(r[1], '', '', r[2], r[3], r[8], dec(r[9]), dec(r[10]), dec(r[11]), r[11]))
        else:
            # 020,inv,nmi,ck,invdate,due,lnspname,abn,retname,abn,status,ex,gst,inc,Y
            pf.invoices.append(Invoice(r[1], r[2], r[3], r[4], r[5], r[10], dec(r[11]), dec(r[12]), dec(r[13]), r[13]))

    for r in rows:
        if r[0] in ('010', '10', '011', '11', '020', '20', '021'):
            continue
        if r[-1] in ('Y', 'N') and len(r) >= 10:
            # all detail records end with ...,ex,gst,GSTindicator
            gsl = (r[0] == '400'
                   or any(c.upper() in GSL_CODES for c in r[9:-3])
                   or bool(GSL_TEXT_RE.search(' '.join(r[9:-3]))))
            detail = ' '.join(c for c in r[9:-3] if c)
            adj = r[6] if vicsa else r[5]          # N / C / R flag sits in a different field per layout
            pf.lines.append(Line(r[0], r[1], r[2], r[7], r[8], r[-3], r[-2], gsl, adj, detail, r))
        else:
            pf.warnings.append(f'unrecognised row skipped: {",".join(r)[:80]}')

    _check_charges(pf, rows)
    return pf


def _check_charges(pf, rows):
    copies = [i.number for i in pf.invoices if 'COPY' in i.status.upper()]
    if copies:
        pf.warnings.append(f'{len(copies)} "Copy Invoice" row(s) excluded from totals and remit')

    nums = [i.number for i in pf.invoices]
    if len(nums) != len(set(nums)):
        pf.warnings.append('the same invoice number appears more than once in this file')

    footer = [r for r in rows if r[0] in ('011', '11')]
    if not footer:
        pf.warnings.append('file has no footer (011) row - totals could not be cross-checked')
    else:
        f = footer[0]
        try:
            if pf.layout == 'NATIONAL':
                fex, fgst, finc = dec(f[-3]), dec(f[-2]), dec(f[-1])
                if (fex, fgst, finc) != (pf.total_ex, pf.total_gst, pf.total_inc):
                    pf.warnings.append(f'footer totals {fex}/{fgst}/{finc} do not match invoice totals '
                                       f'{pf.total_ex}/{pf.total_gst}/{pf.total_inc}')
            else:
                fex = sum((dec(x) for x in f[1:5]), Decimal('0'))
                if fex != pf.total_ex:
                    pf.warnings.append(f'footer total ex GST {fex} does not match invoice total {pf.total_ex}')
        except (ValueError, IndexError):
            pf.warnings.append('footer row could not be read')

    # line items should add up to each invoice
    by_inv = {}
    for l in pf.lines:
        by_inv.setdefault(l.invoice, Decimal('0'))
        by_inv[l.invoice] += l.ex
    bad = [i.number for i in pf.invoices if by_inv.get(i.number, Decimal('0')) != i.ex]
    if bad:
        pf.warnings.append(f'{len(bad)} invoice(s) where the charge lines do not add up to the invoice total, '
                           f'e.g. {bad[0]}')

    # GSL checks
    gsl = pf.gsl_lines()
    if gsl:
        only = pf.gsl_only_invoices()
        mixed = {l.invoice for l in gsl} - only
        parts = [f'{len(gsl)} GSL credit line(s), total {money(pf.gsl_ex + pf.gsl_gst)} incl. GST - '
                 'separate Excel row; listed on the "GSL for Operations" tab (tell Operations to adjust the customer bill)']
        if mixed:
            parts.append(f'credit taken off the remit on {len(mixed)} invoice(s) with normal charges')
        if only:
            parts.append(f'{len(only)} GSL-only invoice(s) left out of the remit')
        pf.warnings.append('; '.join(parts))
    # Other credit lines on event/interest charges that might be an unlabelled GSL
    odd = [l for l in pf.lines if not l.is_gsl and l.record in ('200', '900') and l.ex < 0 and l.adj != 'C']
    if odd:
        pf.warnings.append(f'CHECK: {len(odd)} negative event/interest charge line(s) not marked as GSL '
                           f'(e.g. invoice {odd[0].invoice}). If it is a GSL, add its code to App/gsl_codes.txt')

    if pf.layout == 'VICSA':
        no_nmi = [l.line_id for l in pf.lines if not l.nmi]
        if no_nmi:
            pf.warnings.append(f'{len(no_nmi)} charge line(s) have no meter number (NMI)')


def _parse_remind(base, rows, lnsp, header_ts, file_ts, warnings):
    pf = ParsedFile(base, 'REMIND', 'REMIND', lnsp, header_ts, file_ts, warnings=warnings)
    hdr = rows[0]
    due = hdr[5] if len(hdr) > 5 and is_date(hdr[5]) else header_ts[:8]
    pf._remind_due = to_date(due)
    for r in rows:
        if r[0] == '931':
            # 931,inv,nmi,ck,invdate,due,amount(inc GST),flag
            inc = dec(r[6])
            pf.invoices.append(Invoice(r[1], r[2], r[3], r[4], r[5], 'Overdue', Decimal('0'), Decimal('0'), inc, r[6]))
    if not pf.invoices:
        raise ValueError('overdue notice has no invoice (931) rows')
    # Overdue notice only gives the GST-inclusive amount; split assuming 10% GST
    inc = sum((i.inc for i in pf.invoices), Decimal('0'))
    ex = (inc / Decimal('1.1')).quantize(CENT)
    pf._remind_split = (ex, inc - ex)
    pf.warnings.append('overdue notice only shows amounts incl. GST - ex GST/GST split assumes 10% GST')
    ftr = [r for r in rows if r[0] == '932']
    if ftr and len(ftr[0]) > 2 and dec(ftr[0][2]) != inc:
        pf.warnings.append(f'footer total {ftr[0][2]} does not match sum of invoices {inc}')
    return pf


# ---------------------------------------------------------------------------
# Network Lines extract (same layout as the old Network_Raw tab, plus tracing columns)
# ---------------------------------------------------------------------------

LNSP_STATE = {'ACTEWP': 'ACT', 'AURORAP': 'TAS', 'CNRGYP': 'NSW', 'ENERGYAP': 'NSW', 'INTEGP': 'NSW',
              'ENERGEXP': 'QLD', 'ERGONETP': 'QLD', 'CITIPP': 'VIC', 'POWCP': 'VIC', 'SOLARISP': 'VIC',
              'UNITED': 'VIC', 'EASTERN': 'VIC', 'UMPLP': 'SA'}

NETWORK_LINE_COLUMNS = [
    'INVOICE_LINE', 'INVOICE_NUMBER', 'INVOICE_DATE', 'INVOICE_TYPE', 'NMI', 'STATE',
    'INVOICE_START_DATE', 'INVOICE_END_DATE', 'TARIFF_CODE', 'TARIFF_TYPE', 'UNIT', 'UNIT_MEASURE',
    'UNIT_RATE', 'TOTAL_CHARGE_EXC_GST', 'GST_AMOUNTS', 'REQUIRED_PAYMENT_DATE', 'NETWORKS',
    'INVOICE_NOTE', 'Average_perday', 'NETWORK_TARIFF', 'LINE_ID', 'OLD_INVOICE', 'CHARGE_GROUP', 'FILE_NAME',
]

# What each record type is, for the reconciliation
#   100 = network (NUOS) charges, 600 = Vic/SA metering, 200 = network fees (special reads, site visits...),
#   400 / GSL-coded lines = GSL payments, 900 = interest (ours to pay, not passed to customers)
def charge_group(line):
    if line.is_gsl or line.record == '400':
        return 'GSL'
    return {'100': 'Network charges', '600': 'Metering', '200': 'Fees', '900': 'Interest'}.get(line.record, 'Other')


def _num(s):
    try:
        return float(dec(s))
    except ValueError:
        return None


def network_lines(pf):
    """Every charge line of a charges file (100 network charges, 200 fees, 400 GSL, 600 metering, 900 interest)
    as dicts in NETWORK_LINE_COLUMNS order. Copy invoices are left out (they repeat an invoice already sent)."""
    if pf.kind != 'CHARGES':
        return []
    invs = {i.number: i for i in pf.invoices}
    copies = {i.number for i in pf.invoices if 'COPY' in i.status.upper()}
    out = []
    for l in pf.lines:
        if l.invoice in copies:
            continue
        r = l.raw
        g = (lambda i: r[i] if len(r) > i else '')
        old_inv = g(4) if pf.layout == 'VICSA' else g(3)
        if l.record == '100':
            if pf.layout == 'VICSA':
                tariff, start, end, tou, desc, qty, unit, rate = g(9), g(10), g(11), g(12), g(14), g(15), g(16), g(17)
            else:
                tariff, start, end, tou, desc, qty, unit, rate = g(9), g(11), g(12), g(13), g(15), g(16), g(17), g(18)
        else:
            # 200 / 400 / 600 / 900: code [11], description [12], date [13]
            tariff, tou, desc, end = '', g(11), g(12), g(13)
            start, qty, unit, rate = end, '', '', ''
            if l.record == '200':
                qty, unit, rate = g(14), g(15), g(16)
            elif l.record == '600':
                qty, unit, rate = g(16), 'DAYS', g(17)          # metering: [16] = number of days up to the date
                try:
                    days = int(float(qty))
                    if is_date(end) and 1 <= days <= 400:
                        start = (to_date(end) - timedelta(days=days - 1)).strftime('%Y%m%d')
                except ValueError:
                    pass
        inv = invs.get(l.invoice)
        sd, ed = (to_date(start) if is_date(start) else None), (to_date(end) if is_date(end) else None)
        ex = float(l.ex)
        perday = ex / ((ed - sd).days + 1) if sd and ed and ed >= sd else None
        out.append({
            'INVOICE_LINE': int(l.record) if l.record.isdigit() else l.record, 'INVOICE_NUMBER': l.invoice,
            'INVOICE_DATE': to_date(inv.invoice_date) if inv and is_date(inv.invoice_date) else pf.invoice_date,
            'INVOICE_TYPE': l.adj or 'N', 'NMI': l.nmi, 'STATE': LNSP_STATE.get(pf.lnsp.upper(), ''),
            'INVOICE_START_DATE': sd, 'INVOICE_END_DATE': ed, 'TARIFF_CODE': tou, 'TARIFF_TYPE': desc,
            'UNIT': _num(qty), 'UNIT_MEASURE': unit, 'UNIT_RATE': _num(rate),
            'TOTAL_CHARGE_EXC_GST': ex, 'GST_AMOUNTS': float(l.gst),
            'REQUIRED_PAYMENT_DATE': to_date(inv.due_date) if inv and is_date(inv.due_date) else None,
            'NETWORKS': pf.lnsp, 'INVOICE_NOTE': inv.status if inv else '', 'Average_perday': perday,
            'NETWORK_TARIFF': tariff, 'LINE_ID': l.line_id, 'OLD_INVOICE': old_inv, 'CHARGE_GROUP': charge_group(l), 'FILE_NAME': pf.file_name,
        })
    return out


def write_network_lines_sheet(wb, rows, title='Network Lines'):
    from openpyxl.styles import Font
    ws = wb.create_sheet(title)
    ws.append(NETWORK_LINE_COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True)
    for d in rows:
        ws.append([d.get(k) for k in NETWORK_LINE_COLUMNS])
    for col in ('C', 'G', 'H', 'P'):
        for c in ws[col][1:]:
            c.number_format = 'dd/mm/yyyy'
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    return ws


# ---------------------------------------------------------------------------
# Excel rows
# ---------------------------------------------------------------------------

def excel_rows(pf, location):
    """Return list of row lists for the Excel extract (one main row + optional GSL row)."""
    def row(ex, gst, inc):
        return ['Not Found', pf.lnsp, None, None, pf.file_name, pf.file_ts, location,
                float(ex.quantize(CENT)), float(gst.quantize(CENT)), float(inc.quantize(CENT)),
                pf.invoice_date, pf.due_date, None, None, None]

    if pf.kind == 'REMIND':
        ex, gst = pf._remind_split
        return [row(ex, gst, pf.total_inc)]

    gex, ggst = pf.gsl_ex, pf.gsl_gst
    out = []
    only = pf.gsl_only_invoices()
    has_normal = any(i.number not in only for i in pf.billable_invoices())
    if has_normal or not pf.gsl_lines():
        out.append(row(pf.total_ex - gex, pf.total_gst - ggst, pf.total_inc - gex - ggst))
    if pf.gsl_lines():
        out.append(row(gex, ggst, gex + ggst))
    return out


# ---------------------------------------------------------------------------
# Remit files
# ---------------------------------------------------------------------------

def remit_name(pf):
    return f'NEM#NBREMITT#{RETAILER}#{pf.lnsp}#{pf.file_ts}.csv'


def _remit(pf):
    """Return (rows, total incl. GST, number of detail rows).

    Invoices made up only of GSL credit lines are left out (no payment, no credit taken).
    GSL lines inside an invoice with normal charges are netted off (credit taken).
    """
    out = []
    gsl_only = pf.gsl_only_invoices()
    if pf.layout in ('NATIONAL', 'REMIND'):
        out.append(['800', pf.lnsp, RETAILER, pf.header_ts])
        total = Decimal('0')
        invs = [i for i in pf.billable_invoices() if i.number not in gsl_only]
        for i in invs:
            out.append(['810', i.number, i.nmi, i.checksum, money(i.inc), '', pf.file_ts])
            total += i.inc
        out.append(['820', str(len(invs)), money(total)])
        return out, total, len(invs)

    # Vic/SA: one 921 row per charge line, at the amount on the invoice (GSL credits netted off).
    out.append(['920', pf.lnsp, RETAILER, pf.header_ts])
    copy_invs = {i.number for i in pf.invoices if 'COPY' in i.status.upper()}
    tex = tgst = Decimal('0')
    n = 0
    for l in pf.lines:
        if l.invoice in copy_invs or l.invoice in gsl_only:
            continue
        ex, gst = l.ex, l.gst
        out.append(['921', l.invoice, l.line_id, l.nmi, l.checksum, money(ex), money(gst), money(ex + gst)])
        tex += ex
        tgst += gst
        n += 1
    out.append(['922', money(tex), money(tgst), money(tex + tgst), str(n)])
    return out, tex + tgst, n


def remit_text(pf):
    """NBREMITT file content (CRLF line endings), or None if there is nothing to remit (e.g. GSL-only file)."""
    rows, _, n = _remit(pf)
    if n == 0:
        return None
    return ''.join(','.join(r) + '\r\n' for r in rows)


def remit_total_inc(pf):
    """Total the remit pays (incl. GST), for display."""
    return _remit(pf)[1]


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

def build_workbook(parsed, location, path, problems=()):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = (location or 'Network Bills')[:31].replace('/', '-')
    ws.append(EXCEL_COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='305496')
        c.alignment = Alignment(wrap_text=True, vertical='center')
    for pf in sorted(parsed, key=lambda p: (p.lnsp, p.file_ts)):
        for r in excel_rows(pf, location):
            ws.append(r)
    widths = [11, 12, 16, 12, 50, 19, 12, 16, 13, 16, 12, 12, 18, 12, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row[7:10]:
            c.number_format = '#,##0.00;[Red]-#,##0.00'
        for c in row[10:12]:
            c.number_format = 'dd/mm/yyyy'
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    chk = wb.create_sheet('Checks')
    chk.append(['File', 'Type', 'Remit file', 'Remit total (Inc GST)', 'Note'])
    for c in chk[1]:
        c.font = Font(bold=True)
    for pf in sorted(parsed, key=lambda p: (p.lnsp, p.file_ts)):
        notes = pf.warnings or ['OK']
        for n in notes:
            rn = remit_name(pf) if remit_text(pf) else 'No remit (GSL credit only)'
            chk.append([pf.file_name, f'{pf.kind} / {pf.layout}', rn, float(remit_total_inc(pf)), n])
    for name, msg in problems:
        chk.append([name, 'NOT PROCESSED', '', None, msg])
    for i, w in enumerate([50, 20, 50, 18, 90], 1):
        chk.column_dimensions[get_column_letter(i)].width = w

    # GSL list for the Operations team (to adjust customer bills)
    gs = wb.create_sheet('GSL for Operations')
    gs.append(['Network', 'File', 'Invoice', 'Line ID', 'NMI', 'Checksum', 'Charge detail',
               'Ex GST', 'GST', 'Inc GST', 'Remit'])
    for c in gs[1]:
        c.font = Font(bold=True)
    for pf in sorted(parsed, key=lambda p: (p.lnsp, p.file_ts)):
        if pf.kind != 'CHARGES':
            continue
        only = pf.gsl_only_invoices()
        copies = {i.number for i in pf.invoices if 'COPY' in i.status.upper()}
        for l in pf.gsl_lines():
            if l.invoice in copies:
                continue
            gs.append([pf.lnsp, pf.file_name, l.invoice, l.line_id, l.nmi, l.checksum, l.detail,
                       float(l.ex), float(l.gst), float(l.ex + l.gst),
                       'Not on remit (GSL-only invoice)' if l.invoice in only else 'Credit taken off remit'])
    if gs.max_row == 1:
        gs.append(['No GSL credits in this run'])
    write_network_lines_sheet(wb, [d for pf in sorted(parsed, key=lambda p: (p.lnsp, p.file_ts))
                                   for d in network_lines(pf)])
    for row in gs.iter_rows(min_row=2):
        for c in row[7:10]:
            c.number_format = '#,##0.00;[Red]-#,##0.00'
    for i, w in enumerate([11, 50, 16, 10, 14, 9, 40, 11, 9, 11, 30], 1):
        gs.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)
