import os
from flask import Flask, request, render_template, redirect, url_for, send_file, send_from_directory
from werkzeug.utils import secure_filename
from datetime import datetime
import json

# Importa as funções do corretor
from corretor import load_gabarito, run_correction_parallel

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg'}
app.config['TEMPLATE_GABARITO_FOLDER'] = 'templates_gabarito/'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])
if not os.path.exists(app.config['TEMPLATE_GABARITO_FOLDER']):
    os.makedirs(app.config['TEMPLATE_GABARITO_FOLDER'])

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download_template/<filename>')
def download_template(filename):
    try:
        return send_from_directory(app.config['TEMPLATE_GABARITO_FOLDER'], filename, as_attachment=True)
    except FileNotFoundError:
        return "Arquivo não encontrado.", 404

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        professor_name = request.form['professor_name']
        exam_date = request.form['exam_date']
        turma = request.form['turma']
        num_exam_types = int(request.form['num_exam_types'])

        exam_answers_json = request.form['exam_answers']
        exam_answers_data = json.loads(exam_answers_json)

        folder_name = f"{professor_name}_{exam_date}_{turma}".replace(" ", "_").replace("/", "-")
        target_folder = os.path.join(app.config['UPLOAD_FOLDER'], folder_name)

        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

        temp_gabarito_path = os.path.join(target_folder, 'gabarito_dinamico.txt')
        with open(temp_gabarito_path, 'w', encoding='utf-8') as f:
            for exam_type, questions in exam_answers_data.items():
                for q_num, correct_answer_info in questions.items():
                    f.write(f"{exam_type}|{q_num}|{correct_answer_info['peso_questao']}|")
                    pesos_str = ",".join([f"{k}:{v}" for k, v in correct_answer_info['pesos_alternativas'].items()])
                    f.write(f"{pesos_str}|{correct_answer_info['correta']}\n")

        uploaded_files = request.files.getlist('gabarito_images')
        image_paths = []
        for file in uploaded_files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(target_folder, filename)
                file.save(filepath)
                image_paths.append(filepath)
            else:
                return f"Tipo de arquivo não permitido para {file.filename}", 400

        if not image_paths:
            return "Nenhuma imagem foi enviada ou arquivos inválidos.", 400

        try:
            gabarito_data = load_gabarito(temp_gabarito_path)

            output_excel_filename = f"resultados_{folder_name}.xlsx"
            output_excel_path = os.path.join(target_folder, output_excel_filename)

            # limite de workers pode ser ajustado: ex. workers=2
            run_correction_parallel(image_paths, gabarito_data, output_excel_path)

            return send_file(output_excel_path, as_attachment=True, download_name=output_excel_filename)

        except Exception as e:
            app.logger.error(f"Erro durante o processamento: {e}")
            return f"Ocorreu um erro no processamento: {e}", 500

    return redirect(url_for('index'))
