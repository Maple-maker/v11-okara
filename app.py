import os
import tempfile
from flask import Flask, render_template, request, send_file, flash
from dd1750_core import generate_dd1750_from_pdf

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    try:
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, 'out.pdf')
            _, count = generate_dd1750_from_pdf(
                request.files['bom_file'].stream.name,
                request.files['template_file'].stream.name,
                out
            )
            
            if count == 0:
                flash('No items found')
                return redirect('/')
            
            return send_file(out, as_attachment=True, download_name='DD1750.pdf')
    except Exception as e:
        flash(str(e))
        return redirect('/')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8000)))
