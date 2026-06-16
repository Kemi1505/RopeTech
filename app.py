import os
import pandas as pd
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from mailmerge import MailMerge
import tempfile
import traceback
import math

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('templates', exist_ok=True)

# Equipment configuration
EQUIPMENT_CONFIG = {
    # Shackle types – each has its own template and uses the same mapping file
    'Bow Shackle - Bolt & Nut': {
        'template': 'templates/Bow Shackle - Bolt Template.docx',
        'sheet': 'Bow Shackles | BN',
        'mapping_file': 'shackles/shackle-data.xlsx'
    },
    'Bow Shackle Screw Pin': {
        'template': 'templates/Bow Shackle - Screw Template.docx',
        'sheet': 'Bow Shackles | SP',
        'mapping_file': 'shackles/shackle-data.xlsx'
    },
    'Dee Shackle Bolt & Nut': {
        'template': 'templates/Dee Shackle - Bolt Template.docx',
        'sheet': 'Dee Shackles | BN',
        'mapping_file': 'shackles/shackle-data.xlsx'
    },
    'Dee Shackle Screw Pin': {
        'template': 'templates/Dee Shackle - Screw Template.docx',
        'sheet': 'Dee Shackles | SP',
        'mapping_file': 'shackles/shackle-data.xlsx'
    },
    # Other equipment – no mapping
    'Body Harness': {
        'template': 'templates/Body Harness Belt Template.docx',
        'sheet': 'Body Harness'
    },
    'Lanyard': {
        'template': 'templates/Lanyard Template.docx',
        'sheet': 'Lanyard'
    },
    'Flat Webbing Belts': {
        'template': 'templates/Flat Webbing Belts Template.docx',
        'sheet': 'Flat Belts'
    },
    'Chain Blocks': {
        'template': 'templates/Chain Blocks Template.docx',
        'sheet': 'Chain Block'
    },
    'Lever Hoists': {
        'template': 'templates/Lever Hoist Template.docx',
        'sheet': 'Lever Hoist'
    },
    'Plate Clamps': {
        'template': 'templates/Plate Clamp Template.docx',
        'sheet': 'Plate Clamp'
    },
    'Wire Ropes': {
        'template': 'templates/Wire Ropes Template.docx',
        'sheet': 'Wire Rope'
    }
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

def format_value(v):
    """
    Convert a value to a nicely formatted string:
    - NaN / None becomes empty string
    - Whole numbers (e.g., 1.0, 4.0) become '1', '4'
    - Other floats keep their decimal representation
    """
    if pd.isna(v) or v is None:
        return ''
    if isinstance(v, float):
        # If it's a whole number, convert to int to remove .0
        if v.is_integer():
            return str(int(v))
        else:
            return str(v)
    return str(v)

def load_mapping(mapping_file):
    """Read mapping Excel and return dict keyed by load (float), with formatted values."""
    df = pd.read_excel(mapping_file)
    df = df.where(pd.notnull(df), None)
    mapping = {}
    for _, row in df.iterrows():
        load_val = row.iloc[0]
        try:
            load = float(load_val)
        except:
            continue
        # Format all values in this row (remove .0 from whole numbers)
        formatted_row = {col: format_value(val) for col, val in row.items()}
        mapping[load] = formatted_row
    return mapping

def generate_certificates(excel_path, template_path, sheet_name, output_filename, mapping_file=None):
    """
    Generate certificates using mailmerge.
    If mapping_file is provided, it loads the mapping and adds those fields to each record.
    All numeric values are formatted to remove unnecessary .0.
    """
    # Read Excel without dtype conversion – we'll format manually
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df = df.where(pd.notnull(df), None)  # replace NaN with None
    raw_records = df.to_dict(orient='records')
    if not raw_records:
        raise ValueError("No data rows in selected sheet.")
    
    # Load mapping if needed
    if mapping_file:
        if not os.path.exists(mapping_file):
            raise FileNotFoundError(f"Mapping file not found: {mapping_file}")
        mapping = load_mapping(mapping_file)
    else:
        mapping = None
    
    # Prepare records for merge: each record gets formatted values and mapping values
    merge_records = []
    for rec in raw_records:
        # Get load value for mapping lookup (original numeric)
        load_val = rec.get('WLL')
        try:
            load = float(load_val) if load_val is not None else 0.0
        except:
            load = 0.0
        
        # Format the record (convert all values to string, no .0 for whole numbers)
        formatted_rec = {k: format_value(v) for k, v in rec.items()}
        
        # If mapping exists, add mapping values (already formatted)
        if mapping:
            mapping_vals = mapping.get(load, mapping.get(list(mapping.keys())[0], {}))
            formatted_rec.update(mapping_vals)
        
        merge_records.append(formatted_rec)
    
    # Use mailmerge to generate the document with all records
    output_dir = tempfile.mkdtemp()
    output_path = os.path.join(output_dir, output_filename)
    with MailMerge(template_path) as doc:
        doc.merge_templates(merge_records, separator='page_break')
        doc.write(output_path)
    return output_path

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        equipment = request.form.get('equipment')
        output_name = request.form.get('output_name', '').strip()
        
        if not output_name:
            flash('Please enter an output file name.')
            return redirect(url_for('index'))
        if not output_name.endswith('.docx'):
            output_name += '.docx'
        
        if 'excel_file' not in request.files:
            flash('No file uploaded.')
            return redirect(url_for('index'))
        
        file = request.files['excel_file']
        if file.filename == '':
            flash('No file selected.')
            return redirect(url_for('index'))
        if not allowed_file(file.filename):
            flash('Only .xlsx or .xls files are allowed.')
            return redirect(url_for('index'))
        if equipment not in EQUIPMENT_CONFIG:
            flash('Invalid equipment selected.')
            return redirect(url_for('index'))
        
        excel_filename = secure_filename(file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        file.save(excel_path)
        
        try:
            config = EQUIPMENT_CONFIG[equipment]
            template_path = config['template']
            sheet_name = config['sheet']
            mapping_file = config.get('mapping_file')  # None if not present
            
            output_path = generate_certificates(
                excel_path, template_path, sheet_name, output_name, mapping_file
            )
            
            return send_file(
                output_path,
                as_attachment=True,
                download_name=output_name,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        except Exception as e:
            traceback.print_exc()
            flash(str(e))
            return redirect(url_for('index'))
        finally:
            if os.path.exists(excel_path):
                os.remove(excel_path)
    
    return render_template('index.html', equipment_list=EQUIPMENT_CONFIG.keys())

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)