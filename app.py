import os
from flask import Flask, render_template, request, send_from_directory, flash, redirect, url_for, session
from werkzeug.utils import secure_filename
import openpyxl
from docxtpl import DocxTemplate
import zipfile
import uuid

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = 'supersecretkey'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    download_file = session.pop('download_file', None)
    return render_template('index.html', download_file=download_file)

def generate_documents(template_path, template_name, data_path, output_folder):
    """
    Generates DOCX files from a template and Excel data using docxtpl.
    Returns the number of documents generated for this template.
    """
    try:
        workbook = openpyxl.load_workbook(data_path)
        sheet = workbook.active

        # Normalize headers: remove spaces and convert to string
        header = [str(cell.value).replace(' ', '') if cell.value else '' for cell in sheet[1]]

        doc_count = 0
        for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            doc = DocxTemplate(template_path)
            context = dict(zip(header, row))

            doc.render(context)

            # Generate a unique filename for each document based on template name and row
            template_base_name = os.path.splitext(template_name)[0]
            output_filename = f"{template_base_name}_row_{i-1}.docx"
            output_path = os.path.join(output_folder, output_filename)
            doc.save(output_path)
            doc_count += 1

        return doc_count
    except Exception as e:
        import traceback
        print(f"An error occurred during document generation for template {template_name}: {e}")
        traceback.print_exc()
        return -1

@app.route('/upload', methods=['POST'])
def upload_files():
    # Use getlist to handle multiple files for the 'template_files' input
    template_files = request.files.getlist("template_files")
    data_file = request.files.get('data_file')

    if not template_files or not data_file or not template_files[0].filename:
        flash('Необходимо выбрать хотя бы один файл шаблона и один файл с данными.', 'error')
        return redirect(url_for('index'))

    # Save the single data file first
    data_filename = secure_filename(data_file.filename)
    if not data_filename.endswith('.xlsx'):
        flash('Файл с данными должен быть в формате .xlsx.', 'error')
        return redirect(url_for('index'))
    data_path = os.path.join(app.config['UPLOAD_FOLDER'], data_filename)
    data_file.save(data_path)

    # Prepare for generation
    output_folder = os.path.join(app.config['UPLOAD_FOLDER'], 'generated_docs')
    os.makedirs(output_folder, exist_ok=True)

    total_docs_generated = 0
    error_occurred = False
    saved_template_paths = []

    # Loop through each uploaded template file
    for template_file in template_files:
        if template_file and template_file.filename.endswith('.docx'):
            template_filename = secure_filename(template_file.filename)
            template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_filename)
            template_file.save(template_path)
            saved_template_paths.append(template_path)

            # Generate documents for the current template
            num_docs = generate_documents(template_path, template_filename, data_path, output_folder)

            if num_docs > 0:
                total_docs_generated += num_docs
            else:
                error_occurred = True
                flash(f'Произошла ошибка при обработке шаблона {template_filename}.', 'error')
        else:
            flash(f'Неверный формат файла шаблона: {template_file.filename}. Он должен быть .docx.', 'error')
            error_occurred = True

    # Cleanup saved template files and data file
    for path in saved_template_paths:
        os.remove(path)
    os.remove(data_path)

    if total_docs_generated > 0:
        # Create a zip file with all generated documents
        zip_filename = f"generated_documents_{uuid.uuid4().hex}.zip"
        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in os.listdir(output_folder):
                zipf.write(os.path.join(output_folder, file), file)

        # Cleanup generated docx files and the folder
        for file in os.listdir(output_folder):
            os.remove(os.path.join(output_folder, file))
        os.rmdir(output_folder)

        flash(f'Успешно сгенерировано {total_docs_generated} документов.', 'success')
        session['download_file'] = zip_filename
        return redirect(url_for('index'))
    else:
        # Cleanup the empty generated_docs folder if it exists
        if os.path.exists(output_folder):
            os.rmdir(output_folder)
        if not error_occurred:
             flash('Не удалось сгенерировать ни одного документа.', 'error')
        return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
