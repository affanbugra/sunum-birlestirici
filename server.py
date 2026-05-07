from flask import Flask, request, send_file, jsonify
from pptx import Presentation
import copy
import io

try:
    from pypdf import PdfWriter, PdfReader
except ImportError:
    from PyPDF2 import PdfWriter, PdfReader

app = Flask(__name__)

@app.route('/merge', methods=['POST'])
def merge():
    files = request.files.getlist('files')

    if not files:
        return jsonify({'error': 'Dosya gönderilmedi'}), 400

    pptx_files = [f for f in files if f.filename.endswith('.pptx')]
    pdf_files = [f for f in files if f.filename.endswith('.pdf')]

    if not pptx_files and not pdf_files:
        return jsonify({'error': 'Desteklenmeyen dosya türü'}), 400

    results = {}

    if pptx_files:
        merged_pptx = Presentation()
        for file in pptx_files:
            prs = Presentation(file)
            for slide in prs.slides:
                layout = merged_pptx.slide_layouts[6]
                new_slide = merged_pptx.slides.add_slide(layout)
                for shape in slide.shapes:
                    el = copy.deepcopy(shape.element)
                    new_slide.shapes._spTree.insert(2, el)
        pptx_output = io.BytesIO()
        merged_pptx.save(pptx_output)
        pptx_output.seek(0)
        results['pptx'] = pptx_output

    if pdf_files:
        pdf_writer = PdfWriter()
        for file in pdf_files:
            pdf_reader = PdfReader(file)
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
        pdf_output = io.BytesIO()
        pdf_writer.write(pdf_output)
        pdf_output.seek(0)
        results['pdf'] = pdf_output

    if 'pptx' in results:
        return send_file(results['pptx'], download_name='merged.pptx', as_attachment=True)
    elif 'pdf' in results:
        return send_file(results['pdf'], download_name='merged.pdf', as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
