import io, os, re, json, shutil
from datetime import datetime
from flask import Flask, render_template, jsonify, request, send_file
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process as rfprocess

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB

REFERENCE_DIR = os.path.join(os.path.dirname(__file__), 'reference_files')
MANIFEST_PATH = os.path.join(REFERENCE_DIR, '_manifest.json')

os.makedirs(REFERENCE_DIR, exist_ok=True)


# ── Manifest helpers ──────────────────────────────────────────────────────────

def load_manifest():
    if os.path.exists(MANIFEST_PATH):
        with open(MANIFEST_PATH, 'r') as f:
            return json.load(f)
    return []


def save_manifest(manifest):
    with open(MANIFEST_PATH, 'w') as f:
        json.dump(manifest, f, indent=2)


def save_to_reference(file_obj, original_name):
    """Save an uploaded file to the reference directory. Returns entry dict."""
    manifest = load_manifest()

    # Sanitize filename
    safe_name = re.sub(r'[^\w.\-]', '_', original_name)
    dest_path = os.path.join(REFERENCE_DIR, safe_name)

    # Read content
    file_obj.seek(0)
    content = file_obj.read()
    file_obj.seek(0)

    with open(dest_path, 'wb') as f:
        f.write(content)

    # Update manifest (upsert by filename)
    entry = {
        'name': safe_name,
        'original_name': original_name,
        'size': len(content),
        'uploaded_at': datetime.now().isoformat(timespec='seconds'),
    }
    manifest = [e for e in manifest if e['name'] != safe_name]
    manifest.append(entry)
    save_manifest(manifest)
    return entry


# ── Reference file routes ─────────────────────────────────────────────────────

@app.route('/api/reference-files', methods=['GET'])
def list_reference_files():
    manifest = load_manifest()
    # Sort newest first
    manifest.sort(key=lambda e: e.get('uploaded_at', ''), reverse=True)
    return jsonify({'ok': True, 'files': manifest})


@app.route('/api/reference-files', methods=['POST'])
def upload_reference_file():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'ok': False, 'error': 'No files provided'}), 400
    saved = []
    for f in files:
        if f.filename:
            entry = save_to_reference(f, f.filename)
            saved.append(entry)
    return jsonify({'ok': True, 'saved': saved})


@app.route('/api/reference-files/<filename>', methods=['DELETE'])
def delete_reference_file(filename):
    manifest = load_manifest()
    entry = next((e for e in manifest if e['name'] == filename), None)
    if not entry:
        return jsonify({'ok': False, 'error': 'File not found'}), 404
    path = os.path.join(REFERENCE_DIR, filename)
    if os.path.exists(path):
        os.remove(path)
    manifest = [e for e in manifest if e['name'] != filename]
    save_manifest(manifest)
    return jsonify({'ok': True})


@app.route('/api/reference-files/<filename>', methods=['GET'])
def get_reference_file(filename):
    path = os.path.join(REFERENCE_DIR, filename)
    if not os.path.exists(path):
        return jsonify({'ok': False, 'error': 'File not found'}), 404
    return send_file(path, as_attachment=True)


# ── Utilities ─────────────────────────────────────────────────────────────────

COMPANY_SUFFIXES = {
    'inc', 'llc', 'ltd', 'corp', 'co', 'gmbh', 'ag', 'sa', 'plc',
    'limited', 'incorporated', 'corporation', 'company', 'group',
    'holding', 'holdings', 'enterprises', 'solutions', 'services',
    'international', 'global', 'worldwide', 'consulting', 'the',
}


def clean_company(name):
    if not isinstance(name, str):
        return ''
    s = name.lower().strip()
    s = re.sub(r'[.,&\-/\\]+', ' ', s)
    words = [w for w in s.split() if w not in COMPANY_SUFFIXES]
    return ' '.join(words).strip()


def find_col(df, *keywords):
    kws = [k.lower().replace(' ', '').replace('_', '') for k in keywords]
    for col in df.columns:
        cl = str(col).lower().replace(' ', '').replace('_', '').replace('-', '')
        if any(k in cl for k in kws):
            return col
    return None


def extract_name(row, df):
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


def read_excel_from_path(path):
    try:
        return pd.read_excel(path, engine='openpyxl')
    except Exception:
        return pd.read_excel(path)


def format_contacts(contacts):
    if not contacts:
        return ''
    if len(contacts) == 1:
        return contacts[0]
    if len(contacts) == 2:
        return f'{contacts[0]} and {contacts[1]}'
    return ', '.join(contacts[:-1]) + f', and {contacts[-1]}'


DEFAULT_TEMPLATE = (
    "Hi {first_name}, I wanted to reach out to you at {company}. "
    "We have previously connected with {contacts} from your company "
    "as part of our FEAAM outreach."
)


def build_msg(first, name, company, contacts):
    return (DEFAULT_TEMPLATE
            .replace('{first_name}', first or 'there')
            .replace('{name}', name or '')
            .replace('{company}', company or '')
            .replace('{contacts}', contacts))


def build_company_map(df, co_col):
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


def lookup_company(co_raw, co_map, fuzzy_thresh=80):
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


def open_file(file_obj_or_name):
    """Accept either an uploaded FileStorage or a reference filename string."""
    if isinstance(file_obj_or_name, str):
        path = os.path.join(REFERENCE_DIR, file_obj_or_name)
        return read_excel_from_path(path)
    return read_excel(file_obj_or_name)


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


# ── Message Generator routes ──────────────────────────────────────────────────

def resolve_inputs(request, int_key='interested', org_key='organization'):
    """Return (int_df, org_df) from either uploaded files or reference filenames."""
    int_name = request.form.get('interested_ref')
    org_name = request.form.get('organization_ref')

    int_f = request.files.get(int_key)
    org_f = request.files.get(org_key)

    # Save any newly uploaded files to reference
    if int_f and int_f.filename:
        save_to_reference(int_f, int_f.filename)
        int_f.seek(0)
    if org_f and org_f.filename:
        save_to_reference(org_f, org_f.filename)
        org_f.seek(0)

    int_src = int_name if int_name else int_f
    org_src = org_name if org_name else org_f

    if not int_src or not org_src:
        return None, None, 'Both an Interested List and an Organization List are required.'

    try:
        int_df = open_file(int_src)
        org_df = open_file(org_src)
    except Exception as e:
        return None, None, str(e)

    return int_df, org_df, None


@app.route('/api/preview', methods=['POST'])
def preview():
    int_df, org_df, err = resolve_inputs(request)
    if err:
        return jsonify({'ok': False, 'error': err}), 400

    int_co_col = find_col(int_df, 'company', 'organization', 'org', 'employer', 'firm', 'account')
    org_co_col = find_col(org_df, 'company', 'organization', 'org', 'employer', 'firm', 'account')

    if not int_co_col:
        return jsonify({'ok': False,
                        'error': f'No company column in Interested list. Found: {list(int_df.columns)}'}), 400
    if not org_co_col:
        return jsonify({'ok': False,
                        'error': f'No company column in Organization list. Found: {list(org_df.columns)}'}), 400

    co_map = build_company_map(int_df, int_co_col)
    rows_out = []

    for _, row in org_df.iterrows():
        co_raw = safe_str(row.get(org_co_col, ''))
        full_name = extract_name(row, org_df)
        first = full_name.split()[0] if full_name else 'there'
        contacts, matched_co, match_type = lookup_company(co_raw, co_map)
        contacts_str = format_contacts(contacts)
        msg = build_msg(first, full_name, co_raw or matched_co, contacts_str) if contacts else ''
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
    return jsonify({'ok': True, 'total': len(rows_out), 'matched': matched,
                    'unmatched': len(rows_out) - matched, 'preview': rows_out[:100]})


@app.route('/api/generate', methods=['POST'])
def generate():
    int_df, org_df, err = resolve_inputs(request)
    if err:
        return jsonify({'ok': False, 'error': err}), 400

    int_co_col = find_col(int_df, 'company', 'organization', 'org', 'employer', 'firm', 'account')
    org_co_col = find_col(org_df, 'company', 'organization', 'org', 'employer', 'firm', 'account')
    co_map = build_company_map(int_df, int_co_col)

    output_rows = []
    for _, row in org_df.iterrows():
        co_raw = safe_str(row.get(org_co_col, ''))
        full_name = extract_name(row, org_df)
        first = full_name.split()[0] if full_name else 'there'
        contacts, matched_co, _ = lookup_company(co_raw, co_map)
        contacts_str = format_contacts(contacts)
        msg = build_msg(first, full_name, co_raw or matched_co, contacts_str) if contacts else ''
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


# ── Deduplication routes ──────────────────────────────────────────────────────

def collect_other_keys(request):
    """Collect email/name keys from all 'others' sources (uploaded or reference)."""
    other_emails, other_names = set(), set()

    # Uploaded files
    for f in request.files.getlist('others'):
        if not f.filename:
            continue
        save_to_reference(f, f.filename)
        f.seek(0)
        try:
            df = read_excel(f)
            _extract_keys(df, other_emails, other_names)
        except Exception:
            pass

    # Reference filenames
    for name in request.form.getlist('others_ref'):
        try:
            df = read_excel_from_path(os.path.join(REFERENCE_DIR, name))
            _extract_keys(df, other_emails, other_names)
        except Exception:
            pass

    return other_emails, other_names


def _extract_keys(df, emails, names):
    for col in df.columns:
        cl = col.lower()
        if 'email' in cl or 'mail' in cl:
            vals = df[col].dropna().astype(str).str.lower().str.strip()
            emails.update(v for v in vals if v != 'nan')
        if 'name' in cl:
            vals = df[col].dropna().astype(str).str.lower().str.strip()
            names.update(v for v in vals if v != 'nan')


@app.route('/api/dedup-preview', methods=['POST'])
def dedup_preview():
    ref_name = request.form.get('reference_ref')
    ref_f    = request.files.get('reference')

    if ref_f and ref_f.filename:
        save_to_reference(ref_f, ref_f.filename)
        ref_f.seek(0)

    ref_src = ref_name if ref_name else ref_f
    if not ref_src:
        return jsonify({'ok': False, 'error': 'Reference file required'}), 400

    try:
        ref_df = open_file(ref_src)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400

    ref_email_col = find_col(ref_df, 'email', 'e-mail', 'mail')
    ref_name_col  = find_col(ref_df, 'full name', 'fullname', 'name')
    other_emails, other_names = collect_other_keys(request)

    dup_rows = []
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
        if is_dup:
            dup_rows.append({
                'name':   safe_str(row.get(ref_name_col, '')) if ref_name_col else '',
                'email':  safe_str(row.get(ref_email_col, '')) if ref_email_col else '',
                'reason': reason,
            })

    return jsonify({'ok': True, 'total': len(ref_df),
                    'duplicates': len(dup_rows), 'unique': len(ref_df) - len(dup_rows),
                    'preview': dup_rows[:50]})


@app.route('/api/dedup-download', methods=['POST'])
def dedup_download():
    ref_name = request.form.get('reference_ref')
    ref_f    = request.files.get('reference')

    if ref_f and ref_f.filename:
        save_to_reference(ref_f, ref_f.filename)
        ref_f.seek(0)

    ref_src = ref_name if ref_name else ref_f
    try:
        ref_df = open_file(ref_src)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 400

    ref_email_col = find_col(ref_df, 'email', 'e-mail', 'mail')
    ref_name_col  = find_col(ref_df, 'full name', 'fullname', 'name')
    other_emails, other_names = collect_other_keys(request)

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


# ── Health / Index ────────────────────────────────────────────────────────────

@app.route('/api/health')
def health():
    return jsonify({'ok': True, 'service': 'FEAAM Contact Matcher'})


@app.route('/')
def index():
    return render_template('index.html')


if __name__ == '__main__':
    port = int(os.getenv('PORT', 5051))
    app.run(debug=True, port=port, host='0.0.0.0')
