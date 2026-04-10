"""
Universal Excel Converter
Converts EDI 834, FLX fixed-width, delimited text, and other formats to Excel.
Drop any file → auto-detect → download .xlsx
"""

import os
import re
import io
import gzip
import tempfile
import traceback
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB

UPLOAD_FOLDER = tempfile.mkdtemp()

# ─── FORMAT DETECTION ─────────────────────────────────────────────────────────

def detect_format(content, filename):
    """Auto-detect file format from content and filename."""
    fname = filename.lower()

    # Excel passthrough
    if fname.endswith(('.xlsx', '.xls', '.xlsb')):
        return 'excel'

    # CSV / TSV
    if fname.endswith('.csv'):
        return 'csv'
    if fname.endswith('.tsv'):
        return 'tsv'

    # Check content-based detection
    text = content if isinstance(content, str) else content.decode('utf-8', errors='replace')
    lines = text.strip().split('\n')

    # EDI 834 detection
    if text.strip().startswith('ISA') and '~' in text[:500]:
        return 'edi_834'

    # FLX fixed-width detection
    if fname.endswith('.txt') and 'flx' in fname.lower():
        return 'flx'

    # Generic fixed-width: consistent line lengths, no common delimiters
    if len(lines) > 5:
        lengths = [len(l) for l in lines[:20]]
        if len(set(lengths)) <= 2 and lengths[0] > 100:
            # Check if there are very few delimiters
            sample = '\n'.join(lines[:5])
            pipe_count = sample.count('|')
            comma_count = sample.count(',')
            tab_count = sample.count('\t')
            if max(pipe_count, comma_count, tab_count) < len(lines[:5]) * 2:
                return 'fixed_width'

    # Pipe-delimited
    if len(lines) > 1:
        pipe_counts = [l.count('|') for l in lines[:10]]
        if all(c > 2 for c in pipe_counts) and len(set(pipe_counts)) <= 2:
            return 'pipe'

    # Tab-delimited
    if len(lines) > 1:
        tab_counts = [l.count('\t') for l in lines[:10]]
        if all(c > 0 for c in tab_counts) and len(set(tab_counts)) <= 2:
            return 'tab'

    # Comma-delimited (last resort for text files)
    if len(lines) > 1:
        comma_counts = [l.count(',') for l in lines[:10]]
        if all(c > 0 for c in comma_counts) and len(set(comma_counts)) <= 3:
            return 'csv'

    return 'unknown'


# ─── EDI 834 PARSER ───────────────────────────────────────────────────────────

def parse_edi_834(content):
    """Parse EDI 834 enrollment/eligibility files."""
    if isinstance(content, bytes):
        content = content.decode('utf-8', errors='replace')

    # Detect segment terminator
    terminator = '~'
    segments = [s.strip() for s in content.split(terminator) if s.strip()]

    records = []
    current = {}
    in_member = False

    for seg in segments:
        elements = seg.split('*')
        seg_id = elements[0].strip()

        if seg_id == 'ISA':
            # Interchange header
            pass
        elif seg_id == 'GS':
            # Functional group header
            pass
        elif seg_id == 'ST':
            pass
        elif seg_id == 'BGN':
            # Beginning segment
            if len(elements) > 2:
                current['Transaction_Set_Purpose'] = elements[1] if len(elements) > 1 else ''
                current['Reference_ID'] = elements[2] if len(elements) > 2 else ''
                current['File_Date'] = elements[3] if len(elements) > 3 else ''
        elif seg_id == 'INS':
            if in_member and current:
                records.append(current.copy())
            current = {}
            in_member = True
            current['Subscriber_Indicator'] = elements[1] if len(elements) > 1 else ''
            current['Relationship_Code'] = elements[2] if len(elements) > 2 else ''
            current['Maintenance_Type'] = elements[3] if len(elements) > 3 else ''
            current['Maintenance_Reason'] = elements[4] if len(elements) > 4 else ''
            current['Benefit_Status'] = elements[5] if len(elements) > 5 else ''
            if len(elements) > 7:
                current['Employment_Status'] = elements[8] if len(elements) > 8 else ''
        elif seg_id == 'REF':
            qualifier = elements[1] if len(elements) > 1 else ''
            value = elements[2] if len(elements) > 2 else ''
            if qualifier == '0F':
                current['Subscriber_ID'] = value
            elif qualifier == '1L':
                current['Group_Number'] = value
            elif qualifier == '23':
                current['Client_Number'] = value
            elif qualifier == 'ZZ':
                current['Mutually_Defined_ID'] = value
            else:
                current[f'REF_{qualifier}'] = value
        elif seg_id == 'DTP':
            qualifier = elements[1] if len(elements) > 1 else ''
            date_val = elements[3] if len(elements) > 3 else ''
            if qualifier == '336':
                current['Employment_Date'] = format_edi_date(date_val)
            elif qualifier == '348':
                current['Benefit_Begin_Date'] = format_edi_date(date_val)
            elif qualifier == '349':
                current['Benefit_End_Date'] = format_edi_date(date_val)
            elif qualifier == '303':
                current['Eligibility_Begin_Date'] = format_edi_date(date_val)
            elif qualifier == '304':
                current['Eligibility_End_Date'] = format_edi_date(date_val)
            elif qualifier == '473':
                current['Coverage_Begin_Date'] = format_edi_date(date_val)
            elif qualifier == '474':
                current['Coverage_End_Date'] = format_edi_date(date_val)
            else:
                current[f'Date_{qualifier}'] = format_edi_date(date_val)
        elif seg_id == 'NM1':
            entity_code = elements[1] if len(elements) > 1 else ''
            if entity_code == 'IL':  # Insured/subscriber
                current['Last_Name'] = elements[3] if len(elements) > 3 else ''
                current['First_Name'] = elements[4] if len(elements) > 4 else ''
                current['Middle_Name'] = elements[5] if len(elements) > 5 else ''
                current['Name_Suffix'] = elements[7] if len(elements) > 7 else ''
                if len(elements) > 9:
                    current['SSN_or_ID'] = elements[9]
            elif entity_code == 'P5':  # Plan sponsor
                current['Sponsor_Name'] = elements[3] if len(elements) > 3 else ''
            elif entity_code == 'IN':  # Insurer
                current['Payer_Name'] = elements[3] if len(elements) > 3 else ''
            elif entity_code == '70':  # Prior coverage org
                current['Prior_Coverage_Org'] = elements[3] if len(elements) > 3 else ''
            elif entity_code == '31':  # Postal mailing address
                current['Mail_Last_Name'] = elements[3] if len(elements) > 3 else ''
                current['Mail_First_Name'] = elements[4] if len(elements) > 4 else ''
        elif seg_id == 'PER':
            contact_type = elements[1] if len(elements) > 1 else ''
            if contact_type == 'IP':  # Insured party
                for i in range(3, len(elements) - 1, 2):
                    qual = elements[i]
                    val = elements[i + 1] if i + 1 < len(elements) else ''
                    if qual == 'HP':
                        current['Home_Phone'] = val
                    elif qual == 'WP':
                        current['Work_Phone'] = val
                    elif qual == 'TE':
                        current['Phone'] = val
                    elif qual == 'EM':
                        current['Email'] = val
        elif seg_id == 'N3':
            current['Address_Line1'] = elements[1] if len(elements) > 1 else ''
            current['Address_Line2'] = elements[2] if len(elements) > 2 else ''
        elif seg_id == 'N4':
            current['City'] = elements[1] if len(elements) > 1 else ''
            current['State'] = elements[2] if len(elements) > 2 else ''
            current['Zip'] = elements[3] if len(elements) > 3 else ''
            current['Country'] = elements[4] if len(elements) > 4 else ''
        elif seg_id == 'DMG':
            current['DOB'] = format_edi_date(elements[2]) if len(elements) > 2 else ''
            current['Gender'] = elements[3] if len(elements) > 3 else ''
        elif seg_id == 'HD':
            current['Maintenance_Type_HD'] = elements[1] if len(elements) > 1 else ''
            current['Insurance_Line'] = elements[3] if len(elements) > 3 else ''
            current['Plan_Coverage'] = elements[4] if len(elements) > 4 else ''
            current['Coverage_Level'] = elements[5] if len(elements) > 5 else ''
        elif seg_id == 'ICM':
            current['Frequency'] = elements[1] if len(elements) > 1 else ''
            current['Wage_Amount'] = elements[2] if len(elements) > 2 else ''
        elif seg_id == 'LUI':
            current['Language'] = elements[2] if len(elements) > 2 else ''
        elif seg_id == 'EC':
            current['Employment_Class'] = elements[1] if len(elements) > 1 else ''

    # Don't forget the last record
    if in_member and current:
        records.append(current.copy())

    if not records:
        raise ValueError("No member records found in EDI 834 file")

    df = pd.DataFrame(records)

    # Reorder columns for readability
    priority_cols = ['Subscriber_ID', 'Group_Number', 'Last_Name', 'First_Name',
                     'Middle_Name', 'DOB', 'Gender', 'Subscriber_Indicator',
                     'Relationship_Code', 'Address_Line1', 'City', 'State', 'Zip',
                     'Phone', 'Email', 'Benefit_Begin_Date', 'Benefit_End_Date',
                     'Plan_Coverage', 'Coverage_Level']
    ordered = [c for c in priority_cols if c in df.columns]
    remaining = [c for c in df.columns if c not in ordered]
    df = df[ordered + remaining]

    return df


def format_edi_date(date_str):
    """Convert YYYYMMDD or date range to MM/DD/YYYY."""
    if not date_str:
        return ''
    # Handle date ranges (D8 format)
    date_str = date_str.replace('D8', '').strip()
    if len(date_str) == 8 and date_str.isdigit():
        try:
            return f"{date_str[4:6]}/{date_str[6:8]}/{date_str[:4]}"
        except:
            return date_str
    return date_str


# ─── FLX FIXED-WIDTH PARSER ──────────────────────────────────────────────────

# Standard FLX medical claims layout
FLX_LAYOUT = [
    ('Record_Type', 0, 1),
    ('Claim_Type', 1, 2),
    ('Subscriber_SSN', 2, 11),
    ('Member_ID', 11, 13),
    ('Patient_Last_Name', 13, 28),
    ('Patient_First_Name', 28, 40),
    ('Patient_MI', 40, 41),
    ('Patient_DOB', 41, 49),
    ('Patient_Sex', 49, 50),
    ('Patient_Relationship', 50, 52),
    ('Subscriber_Last_Name', 52, 67),
    ('Subscriber_First_Name', 67, 79),
    ('Patient_Address1', 79, 109),
    ('Patient_Address2', 109, 139),
    ('Patient_City', 139, 159),
    ('Patient_State', 159, 161),
    ('Patient_Zip', 161, 170),
    ('Group_Number', 170, 178),
    ('Group_Name', 178, 208),
    ('Subgroup_Number', 208, 216),
    ('Coverage_Type', 216, 217),
    ('Claim_Number', 217, 230),
    ('Claim_Line', 230, 232),
    ('Date_of_Service_From', 232, 240),
    ('Date_of_Service_To', 240, 248),
    ('Admission_Date', 248, 256),
    ('Discharge_Date', 256, 264),
    ('Provider_NPI', 264, 274),
    ('Provider_Tax_ID', 274, 283),
    ('Provider_Last_Name', 283, 313),
    ('Provider_First_Name', 313, 328),
    ('Provider_Address1', 328, 358),
    ('Provider_City', 358, 378),
    ('Provider_State', 378, 380),
    ('Provider_Zip', 380, 389),
    ('Place_of_Service', 389, 391),
    ('Type_of_Bill', 391, 395),
    ('Procedure_Code', 395, 403),
    ('Procedure_Modifier1', 403, 405),
    ('Procedure_Modifier2', 405, 407),
    ('Revenue_Code', 407, 411),
    ('Diagnosis_Code1', 411, 419),
    ('Diagnosis_Code2', 419, 427),
    ('Diagnosis_Code3', 427, 435),
    ('Diagnosis_Code4', 435, 443),
    ('DRG_Code', 443, 446),
    ('Units', 446, 453),
    ('Billed_Amount', 453, 464),
    ('Allowed_Amount', 464, 475),
    ('Paid_Amount', 475, 486),
    ('Copay_Amount', 486, 497),
    ('Coinsurance_Amount', 497, 508),
    ('Deductible_Amount', 508, 519),
    ('COB_Amount', 519, 530),
    ('Net_Total', 530, 541),
    ('Check_Date', 541, 549),
    ('Check_Number', 549, 564),
    ('Claim_Status', 564, 565),
    ('Run_Date', 565, 573),
    ('Batch_Number', 573, 581),
    ('Filler1', 581, 600),
    ('Claim_Reference', 600, 630),
]


def parse_flx(content):
    """Parse FLX fixed-width medical claims files."""
    if isinstance(content, bytes):
        content = content.decode('utf-8', errors='replace')

    lines = content.strip().split('\n')
    records = []

    for line in lines:
        if len(line.strip()) < 50:
            continue
        record = {}
        for field_name, start, end in FLX_LAYOUT:
            if start < len(line):
                val = line[start:min(end, len(line))].strip()
                record[field_name] = val
            else:
                record[field_name] = ''
        records.append(record)

    if not records:
        raise ValueError("No records found in FLX file")

    df = pd.DataFrame(records)

    # Format dates
    date_cols = ['Patient_DOB', 'Date_of_Service_From', 'Date_of_Service_To',
                 'Admission_Date', 'Discharge_Date', 'Check_Date', 'Run_Date']
    for col in date_cols:
        if col in df.columns:
            df[col] = df[col].apply(format_edi_date)

    # Format amounts
    amount_cols = ['Billed_Amount', 'Allowed_Amount', 'Paid_Amount', 'Copay_Amount',
                   'Coinsurance_Amount', 'Deductible_Amount', 'COB_Amount', 'Net_Total']
    for col in amount_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce') / 100  # Implied decimals

    # Format units
    if 'Units' in df.columns:
        df['Units'] = pd.to_numeric(df['Units'], errors='coerce') / 100

    # Remove empty filler columns
    df = df.drop(columns=[c for c in df.columns if c.startswith('Filler')], errors='ignore')

    return df


# ─── GENERIC FIXED-WIDTH PARSER ──────────────────────────────────────────────

def parse_fixed_width(content):
    """Parse generic fixed-width files by detecting column boundaries."""
    if isinstance(content, bytes):
        content = content.decode('utf-8', errors='replace')

    lines = content.strip().split('\n')
    if len(lines) < 2:
        raise ValueError("File has fewer than 2 lines")

    # Try to detect columns by finding consistent spaces
    # Use first 100 lines to detect boundaries
    sample_lines = lines[:100]
    max_len = max(len(l) for l in sample_lines)

    # Find positions where ALL lines have spaces
    space_positions = set(range(max_len))
    for line in sample_lines:
        padded = line.ljust(max_len)
        for i, ch in enumerate(padded):
            if ch != ' ':
                space_positions.discard(i)

    # Group consecutive space positions into gaps
    sorted_spaces = sorted(space_positions)
    boundaries = [0]
    if sorted_spaces:
        gap_start = sorted_spaces[0]
        for i in range(1, len(sorted_spaces)):
            if sorted_spaces[i] != sorted_spaces[i - 1] + 1:
                boundaries.append(sorted_spaces[i - 1] + 1)
                gap_start = sorted_spaces[i]
        boundaries.append(sorted_spaces[-1] + 1)

    if len(boundaries) < 2:
        # Fallback: just use the whole line
        df = pd.DataFrame({'Data': lines})
        return df

    # Extract columns
    records = []
    for line in lines:
        record = {}
        for i in range(len(boundaries)):
            start = boundaries[i]
            end = boundaries[i + 1] if i + 1 < len(boundaries) else len(line)
            record[f'Col_{i + 1}'] = line[start:end].strip()
        records.append(record)

    df = pd.DataFrame(records)

    # Use first row as header if it looks like one
    first_row = df.iloc[0]
    if all(isinstance(v, str) and not v.replace('.', '').replace('-', '').isdigit()
           for v in first_row.values if v):
        df.columns = [str(v).strip() if v else f'Col_{i}' for i, v in enumerate(first_row.values)]
        df = df.iloc[1:].reset_index(drop=True)

    return df


# ─── DELIMITED FILE PARSER ───────────────────────────────────────────────────

def parse_delimited(content, delimiter=','):
    """Parse delimited files (CSV, TSV, pipe-delimited)."""
    if isinstance(content, bytes):
        content = content.decode('utf-8', errors='replace')

    try:
        df = pd.read_csv(io.StringIO(content), sep=delimiter, dtype=str,
                         on_bad_lines='skip', engine='python')
    except Exception:
        # Try without header
        df = pd.read_csv(io.StringIO(content), sep=delimiter, dtype=str,
                         header=None, on_bad_lines='skip', engine='python')
    return df


# ─── EXCEL STYLING ────────────────────────────────────────────────────────────

def style_excel(wb, sheet_name='Sheet1'):
    """Apply professional styling to an Excel worksheet."""
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    # Header styling
    header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    header_font = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border

    # Data styling
    data_font = Font(name='Calibri', size=10)
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            # Alternate row shading
            if row_idx % 2 == 0:
                cell.fill = PatternFill(start_color='F2F7FB', end_color='F2F7FB', fill_type='solid')

    # Auto-width columns (cap at 40)
    for col_idx in range(1, ws.max_column + 1):
        max_width = 0
        col_letter = get_column_letter(col_idx)
        for row_idx in range(1, min(ws.max_row + 1, 200)):
            cell_val = ws.cell(row=row_idx, column=col_idx).value
            if cell_val:
                max_width = max(max_width, len(str(cell_val)))
        ws.column_dimensions[col_letter].width = min(max_width + 3, 40)

    # Freeze top row
    ws.freeze_panes = 'A2'

    # Auto-filter
    ws.auto_filter.ref = ws.dimensions

    return wb


# ─── MAIN CONVERSION ─────────────────────────────────────────────────────────

def convert_file(file_content, filename):
    """Main conversion entry point. Returns (output_bytes, output_filename, stats)."""

    # Handle gzipped files
    original_filename = filename
    if filename.lower().endswith('.gz'):
        file_content = gzip.decompress(file_content)
        filename = filename[:-3]  # Remove .gz

    # Detect format
    fmt = detect_format(file_content, filename)

    # Parse based on format
    if fmt == 'excel':
        # Passthrough — just return as-is
        return file_content, original_filename, {'format': 'Excel (passthrough)', 'rows': 'N/A'}

    elif fmt == 'edi_834':
        df = parse_edi_834(file_content)

    elif fmt == 'flx':
        df = parse_flx(file_content)

    elif fmt == 'fixed_width':
        df = parse_fixed_width(file_content)

    elif fmt == 'pipe':
        df = parse_delimited(file_content, delimiter='|')

    elif fmt == 'tab' or filename.lower().endswith('.tsv'):
        df = parse_delimited(file_content, delimiter='\t')

    elif fmt == 'csv':
        df = parse_delimited(file_content, delimiter=',')

    else:
        # Last resort: try comma, then tab, then pipe
        for delim in [',', '\t', '|']:
            try:
                df = parse_delimited(file_content, delimiter=delim)
                if len(df.columns) > 1:
                    break
            except:
                continue
        else:
            raise ValueError(f"Could not detect format for {filename}")

    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]

    # Remove completely empty columns
    df = df.dropna(axis=1, how='all')

    # Build output Excel
    output_name = os.path.splitext(filename)[0] + '.xlsx'
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')

    output.seek(0)
    wb = openpyxl.load_workbook(output)
    wb = style_excel(wb, 'Data')

    # Add summary sheet for EDI and FLX
    if fmt in ('edi_834', 'flx'):
        add_summary_sheet(wb, df, fmt)

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    stats = {
        'format': fmt.upper().replace('_', ' '),
        'rows': len(df),
        'columns': len(df.columns),
        'column_names': list(df.columns)
    }

    return final_output.getvalue(), output_name, stats


def add_summary_sheet(wb, df, fmt):
    """Add a summary/stats sheet."""
    ws = wb.create_sheet('Summary', 0)  # Insert at beginning

    header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    header_font = Font(name='Calibri', bold=True, color='FFFFFF', size=12)
    label_font = Font(name='Calibri', bold=True, size=11)
    value_font = Font(name='Calibri', size=11)

    row = 1
    ws.cell(row=row, column=1, value='File Summary').font = Font(name='Calibri', bold=True, size=14, color='1F4E79')
    row += 2

    # Basic stats
    stats = [
        ('Format', fmt.upper().replace('_', ' ')),
        ('Total Records', len(df)),
        ('Total Columns', len(df.columns)),
    ]

    if fmt == 'edi_834':
        if 'Subscriber_Indicator' in df.columns:
            subs = len(df[df['Subscriber_Indicator'] == 'Y'])
            deps = len(df[df['Subscriber_Indicator'] == 'N'])
            stats.append(('Subscribers', subs))
            stats.append(('Dependents', deps))
        if 'Relationship_Code' in df.columns:
            for code, label in [('18', 'Self'), ('01', 'Spouse'), ('19', 'Child')]:
                count = len(df[df['Relationship_Code'] == code])
                if count > 0:
                    stats.append((f'  {label}', count))
        if 'State' in df.columns:
            stats.append(('Unique States', df['State'].nunique()))
        if 'Sponsor_Name' in df.columns:
            stats.append(('Unique Sponsors', df['Sponsor_Name'].nunique()))

    elif fmt == 'flx':
        if 'Claim_Type' in df.columns:
            for ctype in df['Claim_Type'].unique():
                label = {'P': 'Professional', 'O': 'Outpatient', 'I': 'Inpatient', 'VB': 'Capitation'}.get(ctype, ctype)
                count = len(df[df['Claim_Type'] == ctype])
                stats.append((f'{label} Claims ({ctype})', count))
        for amt_col in ['Billed_Amount', 'Allowed_Amount', 'Paid_Amount', 'Net_Total']:
            if amt_col in df.columns:
                total = df[amt_col].sum()
                stats.append((f'Total {amt_col.replace("_", " ")}', f'${total:,.2f}'))

    for label, value in stats:
        ws.cell(row=row, column=1, value=label).font = label_font
        ws.cell(row=row, column=2, value=value).font = value_font
        row += 1

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 25

    return wb


# ─── FLASK ROUTES ─────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    try:
        file_content = file.read()
        output_bytes, output_name, stats = convert_file(file_content, file.filename)

        # Save to temp file
        output_path = os.path.join(UPLOAD_FOLDER, output_name)
        with open(output_path, 'wb') as f:
            f.write(output_bytes if isinstance(output_bytes, bytes) else output_bytes)

        return jsonify({
            'success': True,
            'filename': output_name,
            'stats': stats
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404
    return send_file(filepath, as_attachment=True, download_name=filename)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
