import os
import tempfile
from flask import Flask, render_template, request, send_file
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dd1750-secret-key')
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    if 'bom_file' not in request.files:
        return render_template('index.html', error="No BOM PDF uploaded")
    
    if 'template_file' not in request.files:
        return render_template('index.html', error="No DD1750 Template PDF uploaded")
    
    bom_file = request.files['bom_file']
    template_file = request.files['template_file']
    
    if bom_file.filename == '' or template_file.filename == '':
        return render_template('index.html', error="Both files must be selected")
    
    if not (allowed_file(bom_file.filename) and allowed_file(template_file.filename)):
        return render_template('index.html', error="Both files must be PDF format")
    
    try:
        admin_data = {
            'unit': request.form.get('unit', '').strip(),
            'packed_by': request.form.get('packed_by', '').strip(),
            'date': request.form.get('date', '').strip(),
            'requisition_no': request.form.get('requisition_no', '').strip(),
            'order_no': request.form.get('order_no', '').strip(),
            'num_boxes': request.form.get('num_boxes', '').strip(),
            'end_item': request.form.get('end_item', '').strip(),
            'model': request.form.get('model', '').strip(),
        }
        
        with tempfile.TemporaryDirectory() as tmpdir:
            bom_path = os.path.join(tmpdir, 'bom.pdf')
            template_path = os.path.join(tmpdir, 'template.pdf')
            output_path = os.path.join(tmpdir, 'DD1750.pdf')
            
            bom_file.save(bom_path)
            template_file.save(template_path)
            
            output_path, count = generate_dd1750_from_pdf(
                bom_pdf_path=bom_path,
                template_pdf_path=template_path,
                out_pdf_path=output_path,
                start_page=0,
                admin_data=admin_data
            )
            
            if count == 0:
                return render_template('index.html', error="No items found in BOM PDF")
            
            return send_file(output_path, as_attachment=True, download_name='DD1750_Filled.pdf')
    
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return render_template('index.html', error=f"Error: {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
