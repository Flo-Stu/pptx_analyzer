from flask import Flask, request, render_template
from pptx import Presentation
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
        for placeholder in layout.placeholders:
            layout_info['placeholders'].append({
                'idx': placeholder.placeholder_format.idx,
                'type': placeholder.placeholder_format.type
            })
        details.append(layout_info)
    return details

if __name__ == '__main__':
    app.run(debug=True)