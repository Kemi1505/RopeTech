import os
import pandas as pd
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from mailmerge import MailMerge
import tempfile
import shutil

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

# Ensure folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('templates', exist_ok=True)

# Mapping: equipment -> (template_file, sheet_name)
EQUIPMENT_CONFIG = {
    'Body Harness': {
        'template': 'templates/Body Harness Belt Template.docx',
        'sheet': 'Body Harness'          # sheet name in Excel
    },
    'Lanyard': {
        'template': 'templates/Lanyard Template.docx',
        'sheet': 'Lanyard'          # sheet name in Excel
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

def generate_certificates(excel_path, template_path, sheet_name, output_filename):
    """Run mail merge and return path to generated file."""
    try:
        # Read Excel sheet
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        df = df.fillna('').astype(str)   # convert all to string
        records = df.to_dict(orient='records')
        
        if not records:
            raise ValueError("No data rows in selected sheet.")
        
        # Use a temporary output file
        output_dir = tempfile.mkdtemp()
        output_path = os.path.join(output_dir, output_filename)
        
        with MailMerge(template_path) as doc:
            doc.merge_templates(records, separator='page_break')
            doc.write(output_path)
        
        return output_path
    
    except Exception as e:
        raise Exception(f"Certificate generation failed: {str(e)}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 1. Get form data
        equipment = request.form.get('equipment')
        output_name = request.form.get('output_name', '').strip()
        
        # Validate output filename (add .docx if missing)
        if not output_name:
            flash('Please enter an output file name.')
            return redirect(url_for('index'))
        if not output_name.endswith('.docx'):
            output_name += '.docx'
        
        # 2. Check file upload
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
        
        # 3. Validate equipment selection
        if equipment not in EQUIPMENT_CONFIG:
            flash('Invalid equipment selected.')
            return redirect(url_for('index'))
        
        # 4. Save uploaded Excel to temporary file
        excel_filename = secure_filename(file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        file.save(excel_path)
        
        try:
            # 5. Get template and sheet from config
            config = EQUIPMENT_CONFIG[equipment]
            template_path = config['template']
            sheet_name = config['sheet']
            
            # 6. Generate certificates
            output_path = generate_certificates(excel_path, template_path, sheet_name, output_name)
            
            # 7. Send file to user for download
            return send_file(
                output_path,
                as_attachment=True,
                download_name=output_name,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        
        except Exception as e:
            flash(str(e))
            return redirect(url_for('index'))
        
        finally:
            # Clean up uploaded Excel file
            if os.path.exists(excel_path):
                os.remove(excel_path)
    
    # GET request: show form
    return render_template('index.html', equipment_list=EQUIPMENT_CONFIG.keys())

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=6000)