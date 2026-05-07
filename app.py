import io, os, re
from flask import Flask, render_template, jsonify, request, send_file
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process as rfprocess

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB

DEFAULT_TEMPLATE = (
    "Hi {first_name}, I wanted to reach out to you at {company}. "
    "We have previously connected with {contacts} from your company "
    "as part of our FEAAM outreach."
)

COMPANY_SUFFIXES = {
    'inc', 'llc', 'ltd', 'corp', 'co', 'gmbh', 'ag', 'sa', 'plc',
    'limited', 'incorporated', 'corporation', 'company', 'group',
    'holding', 'holdings', 'enterprises', 'solutions', 'services',
    'international', 'global', 'worldwide', 'consulting', 'the',
}


# ── Utilities ─────────────────────────────────────────────────────────────────

def clean_company(name):
    if not isinstance(name, str):
        return ''
    s = name.lower().strip()
    s = re.sub(r'[.,&\-/\\]+', ' ', s)
    words = [w for w in s.split() if w not in COMPANY_SUFFIXES]
    return ' '.join(words).strip()


def find_col(df, *keywords):
    """Return the first column name whose header contains any keyword."""
    kws = [k.lower().replace(' ', '').replace('_', '') for k in keywords]
    for col in df.columns:
        cl = str(col).lower().replace(' ', '').replace('_', '').replace('-', '')
        if any(k in cl for k in kws):
            return col
    return None


def extract_name(row, df):
    """Best-effort full-name extraction from a row."""
    for col in df.columns:
        cl = str(col).lower().strip()
        if cl in ('name', 'full name', 'fullname', 'contact name', 'contact'):
            v = str(row.get(col, '')).strip()
            if v and v.lower() != 'nan':
                return v

    fc = find_col(df, 'first name', 'firstname', 'first')
    lc = find_col(df, 'last name', 'lastname', 'last', 'surname')
    first = (str(row.get(fc, '')).strip() if fc else '').replace('nan', '').strip()
    last  = (str(row.get(lc, '')).strip() if lc else '').replace('nan', '').strip()
    if first or last:
        return f'{first} {last}'.strip()

    for col in df.columns:
        if 'name' in col.lower():
            v = str(row.get(col, '')).strip()
            if v and v.lower() != 'nan':
                return v
    return ''


def safe_str(v):
    s = str(v).strip()
    return '' if s.lower() == 'nan' else s


def read_excel(f):
    try:
        return pd.read_excel(f, engine='openpyxl')
    except Exception:
        f.seek(0)
        return pd.read_excel(f)


def format_contacts(contacts):
    if not contacts:
        return ''
    if len(contacts) == 1:
        return contacts[0]
    if len(contacts) == 2:
        return f'{contacts[0]} and {contacts[1]}'
    return ', '.join(contacts[:-1]) + f', and {contacts[-1]}'


def build_msg(template, first, name, company, contacts):
    return (template
            .replace('{first_name}', first or 'there')
            .replace('{name}', name or '')
            .replace('{company}', company or '')
            .replace('{contacts}', contacts))


def build_company_map(df, co_col, fuzzy_thresh):
    """Build normalized_company → {display, contacts[], clean} dict."""
    co_map = {}
    for _, row in df.iterrows():
        co_raw = safe_str(row.get(co_col, ''))
        if not co_raw:
            continue
        co_key = clean_company(co_raw)
        if not co_key:
            continue
        name = extract_name(row, df)
        if not name:
            continue
        if co_key not in co_map:
            co_map[co_key] = {'display': co_raw, 'contacts': []}
        if name not in co_map[co_key]['contacts']:
            co_map[co_key]['contacts'].append(name)
    return co_map


def lookup_company(co_raw, co_map, fuzzy_thresh):
    co_key = clean_company(co_raw)
    if not co_key:
        return [], '', 'none'
    if co_key in co_map:
        return co_map[co_key]['contacts'], co_map[co_key]['display'], 'exact'
    if co_map:
        res = rfprocess.extractOne(co_key, list(co_map.keys()), scorer=fuzz.token_sort_ratio)
        if res and res[1] >= fuzzy_thresh:
            return co_map[res[0]]['contacts'], co_map[res[0]]['display'], f'fuzzy ({res[1]}%)'
    return [], '', 'none'


# ── Styling helpers ───────────────────────────────────────────────────────────

def style_header(ws, ncols):
    fill = PatternFill(start_color='111520', end_color='111520', fill_type='solid')
    font = Font(color='E4E8F4', bold=True)
    for i in range(1, ncols + 1):
        c = ws.cell(row=1, column=i)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 22


def auto_width(ws, df):
    for i, col in enumerate(df.columns, 1):
        sample = [str(col)] + [str(df.iloc[r, i - 1]) for r in range(min(20, len(df)))]
        best = max(len(s) for s in sample)
        ws.column_dimensions[get_column_letter(i)].width = min(max(best + 2, 10), 60)


def highlight_rows(ws, df, col_name, match_val, bg, fg=None):
    if col_name not in df.columns:
        return
    idx = list(df.columns).index(col_name) + 1
    fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
    for r in range(2, len(df) + 2):
        if str(ws.cell(row=r, column=idx).value) == match_val:
            for c in ws[r]:
                c.fill = fill
            if fg:
                ws.cell(row=r, column=idx).font = Font(color=fg, bold=True)


def highlight_has_value_rows(ws, df, col_name, bg):
    if col_name not in df.columns:
        return
    idx = list(df.columns).index(col_name) + 1
    fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
    for r in range(2, len(df) + 2):
        if ws.cell(row=r, column=idx).value:
            for c in ws[r]:
                c.fill = fill


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/preview', methods=['POST'])
def preview():
    int_f = request.files.get('interested')
    org_f = request.files.get('organization')
    template = request.form.get('template', DEFAULT_TEMPLATE)
    fuzzy_thresh = max(0, min(100, int(request.form.get('fuzzy', 80))))

    if not int_f or not org_f:
        return jsonify({'ok': False, 'error': 'Both files are required'}), 400

    try:
        int_df = read_excel(int_f)
        org_df = read_excel(org_f)
    except Exception as e:
        return jsonify({'ok': False, 'error': f'Could not read file: {e}'}), 400

    int_co_col = find_col(int_df, 'company', 'organization', 'org', 'employer', 'firm', 'account')
    org_co_col = find_col(org_df, 'company', 'organization', 'org', 'employer', 'firm', 'account')

    if not int_co_col:
        return jsonify({'ok': False,
                        'error': f'No company column in Interested list. Columns found: {list(int_df.columns)}'}), 400
    if not org_co_col:
        return jsonify({'ok': False,
                        'error': f'No company column in Organization list. Columns found: {list(org_df.columns)}'}), 400

    co_map = build_company_map(int_df, int_co_col, fuzzy_thresh)

    rows_out = []
    for _, row in org_df.iterrows():
        co_raw = safe_str(row.get(org_co_col, ''))
        full_name = extract_name(row, org_df)
        first = full_name.split()[0] if full_name else 'there'

        contacts, matched_co, match_type = lookup_company(co_raw, co_map, fuzzy_thresh)
        contacts_str = format_contacts(contacts)
        msg = build_msg(template, first, full_name, co_raw or matched_co, contacts_str) if contacts else ''

        rows_out.append({
            'name': full_name,
            'company': co_raw,
            'matched_company': matched_co,
            'match_type': match_type,
            'contacts': contacts_str,
            'message': msg,
            'has_match': bool(contacts),
        })

    matched = sum(1 for r in rows_out if r['has_match'])
    return jsonify({
        'ok': True,
        'total': len(rows_out),
        'matched': matched,
        'unmatched': len(rows_out) - matched,
        'preview': rows_out[:100],
    })


@app.route('/api/generate', methods=['POST'])
def generate():
    int_f = request.files.get('interested')
    org_f = request.files.get('organization')
    template = request.form.get('template', DEFAULT_TEMPLATE)
    fuzzy_thresh = max(0, min(100, int(request.form.get('fuzzy', 80))))

    try:
        int_df = read_excel(int_f)
        org_df = read_excel(org_f)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400

    int_co_col = find_col(int_df, 'company', 'organization', 'org', 'employer', 'firm', 'account')
    org_co_col = find_col(org_df, 'company', 'organization', 'org', 'employer', 'firm', 'account')

    co_map = build_company_map(int_df, int_co_col, fuzzy_thresh)

    output_rows = []
    for _, row in org_df.iterrows():
        co_raw = safe_str(row.get(org_co_col, ''))
        full_name = extract_name(row, org_df)
        first = full_name.split()[0] if full_name else 'there'

        contacts, matched_co, _ = lookup_company(co_raw, co_map, fuzzy_thresh)
        contacts_str = format_contacts(contacts)
        msg = build_msg(template, first, full_name, co_raw or matched_co, contacts_str) if contacts else ''

        r = {col: safe_str(row[col]) for col in org_df.columns}
        r['Previous Contacts'] = contacts_str
        r['Message'] = msg
        output_rows.append(r)

    out_df = pd.DataFrame(output_rows)
    base_cols = [c for c in out_df.columns if c not in ('Previous Contacts', 'Message')]
    out_df = out_df[base_cols + ['Previous Contacts', 'Message']]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        out_df.to_excel(writer, index=False, sheet_name='FEAAM Messages')
        ws = writer.sheets['FEAAM Messages']
        style_header(ws, len(out_df.columns))
        auto_width(ws, out_df)
        highlight_has_value_rows(ws, out_df, 'Message', 'e8f5e9')
    buf.seek(0)

    return send_file(buf, download_name='feaam_messages.xlsx', as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/dedup-preview', methods=['POST'])
def dedup_preview():
    ref_f = request.files.get('reference')
    others = request.files.getlist('others')

    if not ref_f:
        return jsonify({'ok': False, 'error': 'Reference file required'}), 400

    try:
        ref_df = read_excel(ref_f)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400

    ref_email_col = find_col(ref_df, 'email', 'e-mail', 'mail')
    ref_name_col  = find_col(ref_df, 'full name', 'fullname', 'name')

    other_emails, other_names = set(), set()
    files_read = 0
    for f in others:
        try:
            df = read_excel(f)
            files_read += 1
            for col in df.columns:
                cl = col.lower()
                if 'email' in cl or 'mail' in cl:
                    vals = df[col].dropna().astype(str).str.lower().str.strip()
                    other_emails.update(v for v in vals if v != 'nan')
                if 'name' in cl:
                    vals = df[col].dropna().astype(str).str.lower().str.strip()
                    other_names.update(v for v in vals if v != 'nan')
        except Exception:
            pass

    dup_rows = []
    statuses = []
    for _, row in ref_df.iterrows():
        is_dup, reason = False, ''
        if ref_email_col:
            email = safe_str(row.get(ref_email_col, '')).lower()
            if email and email in other_emails:
                is_dup, reason = True, 'Email match'
        if not is_dup and ref_name_col:
            name = safe_str(row.get(ref_name_col, '')).lower()
            if name and name in other_names:
                is_dup, reason = True, 'Name match'
        statuses.append(is_dup)
        if is_dup:
            dup_rows.append({
                'name': safe_str(row.get(ref_name_col, '')) if ref_name_col else '',
                'email': safe_str(row.get(ref_email_col, '')) if ref_email_col else '',
                'reason': reason,
            })

    return jsonify({
        'ok': True,
        'total': len(ref_df),
        'files_checked': files_read,
        'duplicates': len(dup_rows),
        'unique': len(ref_df) - len(dup_rows),
        'preview': dup_rows[:50],
    })


@app.route('/api/dedup-download', methods=['POST'])
def dedup_download():
    ref_f = request.files.get('reference')
    others = request.files.getlist('others')

    try:
        ref_df = read_excel(ref_f)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400

    ref_email_col = find_col(ref_df, 'email', 'e-mail', 'mail')
    ref_name_col  = find_col(ref_df, 'full name', 'fullname', 'name')

    other_emails, other_names = set(), set()
    for f in others:
        try:
            df = read_excel(f)
            for col in df.columns:
                cl = col.lower()
                if 'email' in cl or 'mail' in cl:
                    vals = df[col].dropna().astype(str).str.lower().str.strip()
                    other_emails.update(v for v in vals if v != 'nan')
                if 'name' in cl:
                    vals = df[col].dropna().astype(str).str.lower().str.strip()
                    other_names.update(v for v in vals if v != 'nan')
        except Exception:
            pass

    statuses, reasons = [], []
    for _, row in ref_df.iterrows():
        is_dup, reason = False, ''
        if ref_email_col:
            email = safe_str(row.get(ref_email_col, '')).lower()
            if email and email in other_emails:
                is_dup, reason = True, 'Email match'
        if not is_dup and ref_name_col:
            name = safe_str(row.get(ref_name_col, '')).lower()
            if name and name in other_names:
                is_dup, reason = True, 'Name match'
        statuses.append('Duplicate' if is_dup else 'Unique')
        reasons.append(reason)

    out_df = ref_df.copy()
    out_df['Status'] = statuses
    out_df['Match Reason'] = reasons

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        out_df.to_excel(writer, index=False, sheet_name='Deduplication')
        ws = writer.sheets['Deduplication']
        style_header(ws, len(out_df.columns))
        auto_width(ws, out_df)
        highlight_rows(ws, out_df, 'Status', 'Duplicate', 'ffebee', 'c62828')
    buf.seek(0)

    return send_file(buf, download_name='feaam_deduplication.xlsx', as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/health')
def health():
    return jsonify({'ok': True, 'service': 'FEAAM Contact Matcher'})


if __name__ == '__main__':
    port = int(os.getenv('PORT', 5051))
    app.run(debug=True, port=port, host='0.0.0.0')
