from flask import Flask, request, render_template
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'Keine Datei ausgewählt'
        file = request.files['file']
        if file.filename == '':
            return 'Keine Datei ausgewählt'
        if file and file.filename.endswith('.pptx'):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            details = analyze_pptx(filepath)
            return render_template('index.html', details=details)
    return render_template('index.html', details=None)

def analyze_pptx(filepath):
    prs = Presentation(filepath)
    details = []
    for i, layout in enumerate(prs.slide_master.slide_layouts):
        layout_info = {
            'layout_index': i,
            'layout_name': layout.name,
            'placeholders': []
        }
        for shape in layout.shapes:
            if not shape.is_placeholder:
                continue
            phf = shape.placeholder_format
            ph_type = phf.type
            ph_idx = phf.idx
            ph_name = shape.name
            ph_content = ""
            if shape.has_text_frame:
                ph_content = shape.text_frame.text
            layout_info['placeholders'].append({
                'idx': ph_idx,
                'name': ph_name,
                'type': ph_type,
                'content': ph_content
            })
        details.append(layout_info)
    return details

if __name__ == '__main__':
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run(debug=True)
