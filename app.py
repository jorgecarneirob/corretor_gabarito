import os
from flask import Flask, request, render_template, redirect, url_for, send_file, send_from_directory
from werkzeug.utils import secure_filename
import json
from io import BytesIO
from multiprocessing import Pool, cpu_count
import logging

# Importa as funções do seu script corretor
from corretor import load_gabarito, process_gabarito, export_to_excel

# Configuração de logging
logging.basicConfig(level=logging.INFO)
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg'}
app.config['TEMPLATE_GABARITO_FOLDER'] = 'templates_gabarito/'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])
if not os.path.exists(app.config['TEMPLATE_GABARITO_FOLDER']):
    os.makedirs(app.path.join(app.root_path, app.config['TEMPLATE_GABARITO_FOLDER']))

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
        
        exam_answers_json = request.form.get('exam_answers')
        if not exam_answers_json:
            return "Erro: Dados do gabarito ausentes.", 400
        
        try:
            exam_answers_data = json.loads(exam_answers_json)
        except json.JSONDecodeError:
            return "Erro: Dados do gabarito em formato inválido.", 400

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
        try:
            for file in uploaded_files:
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    filepath = os.path.join(target_folder, filename)
                    file.save(filepath)
                    image_paths.append(filepath)
                else:
                    raise ValueError(f"Tipo de arquivo não permitido para {file.filename}")

            if not image_paths:
                raise ValueError("Nenhuma imagem foi enviada ou arquivos inválidos.")

            gabarito_data = load_gabarito(temp_gabarito_path)
            if gabarito_data is None:
                raise Exception("Falha ao carregar o gabarito. Verifique o formato do arquivo.")

            # CHAMA A FUNÇÃO CORRETA AGORA
            logging.info("Iniciando correção em paralelo...")
            num_processes = min(cpu_count(), 4)
            
            # ATENÇÃO: PARA GERAR IMAGENS DE DEPURAÇÃO, DEIXE save_debug=True
            # PARA PRODUÇÃO, MUDE PARA save_debug=False
            save_debug = False 
            
            with Pool(processes=num_processes) as pool:
                resultados = pool.starmap(process_gabarito, [(path, save_debug) for path in image_paths])
            
            resultados = [r for r in resultados if r is not None]

            if not resultados:
                raise Exception("O processo de correção não gerou resultados.")

            output_excel_path = export_to_excel(resultados, gabarito_data, filename=os.path.join(target_folder, f"resultados_{folder_name}.xlsx"))

            if not output_excel_path:
                raise Exception("O processo de exportação não gerou um arquivo de resultados.")

            with open(output_excel_path, 'rb') as f:
                excel_bytes = BytesIO(f.read())
            
            # Limpeza de arquivos temporários e de depuração
            temp_files = [temp_gabarito_path, output_excel_path]
            temp_files.extend(image_paths)
            if save_debug:
                # Se save_debug for True, encontra e remove os arquivos de depuração também
                for path in image_paths:
                    base = os.path.splitext(os.path.basename(path))[0]
                    temp_files.extend([
                        os.path.join(target_folder, f"{base}_warp.png"),
                        os.path.join(target_folder, f"{base}_respostas.png")
                    ])

            for path in temp_files:
                if os.path.exists(path):
                    os.remove(path)
                    
            excel_bytes.seek(0)
            return send_file(excel_bytes, as_attachment=True, download_name=os.path.basename(output_excel_path), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e:
            logging.error(f"Erro durante o processamento: {e}", exc_info=True)
            return f"Ocorreu um erro no processamento: {e}", 500

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)