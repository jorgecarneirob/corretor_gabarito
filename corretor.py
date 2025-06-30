#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import cv2
import numpy as np
import os
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# ---------- FUNÇÕES DE ANÁLISE DA IMAGEM ----------

def load_and_preprocess_image(image_path):
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    if img is None:
        print(f"Erro: Não consegui abrir a imagem: {image_path}")
        return None
    img = cv2.resize(img, (674, 790))
    _, binary = cv2.threshold(img, 127, 255, cv2.THRESH_BINARY_INV)
    return binary

def find_alignment_rectangles(binary):
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    expected_width = 70
    expected_height = 46
    tolerance = 0.2

    min_width = expected_width * (1 - tolerance)
    max_width = expected_width * (1 + tolerance)
    min_height = expected_height * (1 - tolerance)
    max_height = expected_height * (1 + tolerance)

    rectangles = []

    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if min_width <= w <= max_width and min_height <= h <= max_height:
            rectangles.append((x, y, w, h))

    if len(rectangles) < 4:
        print(f"Erro: Encontrados apenas {len(rectangles)} retângulos de alinhamento. Esperado: 4.")
        return None

    rectangles = sorted(rectangles, key=lambda r: (r[1], r[0]))
    top = sorted(rectangles[:2], key=lambda r: r[0])
    bottom = sorted(rectangles[2:], key=lambda r: r[0])
    ordered = [top[0], top[1], bottom[0], bottom[1]]
    return ordered

def compute_grid_centers(corners, rows=18, cols=9):
    tl = np.array([corners[0][0] + corners[0][2] // 2, corners[0][1] + corners[0][3] // 2])
    tr = np.array([corners[1][0] + corners[1][2] // 2, corners[1][1] + corners[1][3] // 2])
    bl = np.array([corners[2][0] + corners[2][2] // 2, corners[2][1] + corners[2][3] // 2])
    br = np.array([corners[3][0] + corners[3][2] // 2, corners[3][1] + corners[3][3] // 2])

    grid_centers = []
    for i in range(rows):
        row = []
        left = tl + (bl - tl) * i / (rows - 1)
        right = tr + (br - tr) * i / (rows - 1)
        for j in range(cols):
            point = left + (right - left) * j / (cols - 1)
            row.append(tuple(point.astype(int)))
        grid_centers.append(row)
    return grid_centers

def read_binary_value(binary, points):
    value = ""
    for x, y in points:
        area = binary[max(y-2,0):min(y+3,binary.shape[0]), max(x-2,0):min(x+3,binary.shape[1])]
        mean = np.mean(area)
        value += "1" if mean > 127 else "0"
    return int(value, 2)

def read_answers(binary, grid_centers, debug=False):
    answers = []
    debug_image = cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)

    for q, row in enumerate(range(7, 17)):  # linhas 8 a 17
        max_mean = 0
        selected_letter = "-"
        for i, col in enumerate(range(3, 8)):  # colunas 4 a 8 (A-E)
            x, y = grid_centers[row][col]
            area = binary[max(y-2, 0):min(y+3, binary.shape[0]), max(x-2, 0):min(x+3, binary.shape[1])]
            mean = np.mean(area)
            if mean > max_mean:
                max_mean = mean
                selected_letter = chr(65 + i)

            if debug:
                cv2.circle(debug_image, (x, y), 5, (255, 0, 0), -1)
                cv2.putText(debug_image, chr(65 + i), (x - 10, y - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 0, 0), 1)

        if max_mean < 10:
            answers.append("-")
        else:
            answers.append(selected_letter)

    if debug:
        cv2.imwrite("debug_respostas.png", debug_image)
        print("Imagem debug_respostas.png salva.")
    return answers

def process_gabarito(image_path):
    binary = load_and_preprocess_image(image_path)
    if binary is None:
        return None

    markers = find_alignment_rectangles(binary)
    if markers is None:
        return None

    grid_centers = compute_grid_centers(markers)

    id_points = [grid_centers[4][col] for col in range(2, 8)]
    student_id = read_binary_value(binary, id_points)

    prova_points = [grid_centers[5][col] for col in range(2, 6)]
    prova_id = read_binary_value(binary, prova_points)

    respostas = read_answers(binary, grid_centers, debug=True)

    return {
        "Aluno_ID": student_id,
        "Prova_ID": prova_id,
        "Respostas": respostas
    }

# ---------- LEITURA DO GABARITO ----------

def load_gabarito(filepath):
    gabarito = {}
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            prova, questao, peso_q, pesos_alt_str, correta = line.split("|")
            peso_q = float(peso_q)
            pesos_alt = {}
            for pair in pesos_alt_str.split(","):
                letra, val = pair.split(":")
                pesos_alt[letra] = float(val)
            if prova not in gabarito:
                gabarito[prova] = {}
            gabarito[prova][questao] = {
                "peso_questao": peso_q,
                "pesos_alternativas": pesos_alt,
                "correta": correta
            }
    return gabarito

# ---------- EXCEL ----------

from openpyxl import Workbook
from openpyxl.styles import PatternFill

def export_to_excel(resultados, gabarito, filename="resultados.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    questoes = list(next(iter(gabarito.values())).keys())
    headers = ["Aluno", "Prova"] + questoes + ["Total"]
    ws.append(headers)

    vermelho_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    verde_fill = PatternFill(start_color="FF00FF00", end_color="FF00FF00", fill_type="solid") #verde
    amarelo_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")# amarelo

    # ----------------------------
    # Gabaritos oficiais de todas as provas
    # ----------------------------
    for prova_id in sorted(gabarito.keys()):
        prova_gabarito = gabarito[prova_id]
        gabarito_row = [f"GABARITO_P{prova_id}", ""]
        for q in questoes:
            correta = prova_gabarito[q]["correta"]
            gabarito_row.append(correta)
        gabarito_row.append("")
        ws.append(gabarito_row)

    # ----------------------------
    # Dados dos alunos
    # ----------------------------
    for r in resultados:
        aluno_id = r["Aluno_ID"]
        prova_id = str(r["Prova_ID"])
        respostas = r["Respostas"]

        prova_gabarito = gabarito.get(prova_id, {})
        total = 0
        row = [aluno_id, prova_id]

        for i, questao in enumerate(questoes):
            resposta_aluno = respostas[i] if i < len(respostas) else "-"
            info = prova_gabarito.get(questao, None)

            if info:
                peso_q = info["peso_questao"]
                peso_alt = info["pesos_alternativas"].get(resposta_aluno, 0.0)
                pontos = peso_q * peso_alt
                total += pontos
                correta = info["correta"]

                row.append(resposta_aluno)
            else:
                row.append("-")

        row.append(total)
        ws.append(row)

        # aplicar cores azul / verde / vermelho
        for i, questao in enumerate(questoes):
            cell = ws.cell(row=ws.max_row, column=i + 3)
            correta = prova_gabarito.get(questao, {}).get("correta")
            pesos_alt = prova_gabarito.get(questao, {}).get("pesos_alternativas", {})

            resposta_aluno = respostas[i] if i < len(respostas) else "-"

            if correta:
                if resposta_aluno == correta:
                    peso_alt = pesos_alt.get(resposta_aluno, 0.0)
                    if abs(peso_alt - 1.0) < 1e-6:
                        cell.fill = verde_fill  # acerto total
                    elif 0.1 < peso_alt < 1.0:
                        cell.fill = amarelo_fill  # acerto parcial (correta, mas peso parcial)
                    else:
                        cell.fill = vermelho_fill
                else:
                    # resposta diferente da correta
                    peso_alt = pesos_alt.get(resposta_aluno, 0.0)
                    if peso_alt > 0.0:
                        cell.fill = amarelo_fill  # acerto parcial
                    else:
                        cell.fill = vermelho_fill
            else:
                cell.fill = vermelho_fill


    wb.save(filename)
    print(f"Arquivo {filename} salvo com sucesso.")


# ---------- EXECUÇÃO ----------

def run_correction(image_paths, gabarito_data, output_excel_path):
    """
    Função principal para executar a correção.

    Args:
        image_paths (list): Lista de caminhos completos para as imagens dos gabaritos.
        gabarito_data (dict): O dicionário de gabarito carregado pela função load_gabarito.
        output_excel_path (str): Caminho completo onde o arquivo Excel de resultados será salvo.
    """
    resultados = []
    for path in image_paths:
        resultado = process_gabarito(path)
        if resultado:
            resultados.append(resultado)
        # Opcional: Remover a imagem após o processamento para economizar espaço se não for mais necessária
        # os.remove(path)

    export_to_excel(resultados, gabarito_data, filename=output_excel_path)
    return output_excel_path # Retorna o caminho do arquivo gerado
