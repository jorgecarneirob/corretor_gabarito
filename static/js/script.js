document.addEventListener('DOMContentLoaded', function() {
    const numExamTypesInput = document.getElementById('num_exam_types');
    const examAnswersContainer = document.getElementById('exam_answers_container');
    const correctionForm = document.getElementById('correctionForm');
    const messageDiv = document.getElementById('message');

    // Mapeamento do gabarito.txt para facilitar a criação do JSON
    // Este é um exemplo, você precisaria adaptar com base na sua lógica de gabarito
    const gabaritoExample = {
        'Q1': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 1.0, 'B': 0.0, 'C': 0.0, 'D': 0.0, 'E': 0.0}, 'correta': 'A' },
        'Q2': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 0.0, 'B': 1.0, 'C': 0.0, 'D': 0.0, 'E': 0.0}, 'correta': 'B' },
        'Q3': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 0.0, 'B': 0.0, 'C': 1.0, 'D': 0.0, 'E': 0.0}, 'correta': 'C' },
        'Q4': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 0.0, 'B': 0.0, 'C': 0.0, 'D': 1.0, 'E': 0.7}, 'correta': 'D' },
        'Q5': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 0.0, 'B': 0.0, 'C': 0.0, 'D': 1.0, 'E': 0.0}, 'correta': 'D' },
        'Q6': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 0.0, 'B': 0.0, 'C': 0.0, 'D': 0.0, 'E': 1.0}, 'correta': 'E' },
        'Q7': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 0.0, 'B': 0.0, 'C': 0.0, 'D': 0.0, 'E': 1.0}, 'correta': 'E' },
        'Q8': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 0.0, 'B': 0.0, 'C': 0.0, 'D': 1.0, 'E': 0.0}, 'correta': 'D' },
        'Q9': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 0.0, 'B': 0.0, 'C': 1.0, 'D': 0.0, 'E': 0.0}, 'correta': 'C' },
        'Q10': { 'peso_questao': 1.0, 'pesos_alternativas': {'A': 0.0, 'B': 0.0, 'C': 0.0, 'D': 1.0, 'E': 0.0}, 'correta': 'D' }
    };

    function generateExamAnswerFields() {
        examAnswersContainer.innerHTML = ''; // Limpa os campos existentes
        const numTypes = parseInt(numExamTypesInput.value);

        if (isNaN(numTypes) || numTypes < 1) return;

        for (let i = 1; i <= numTypes; i++) {
            const examDiv = document.createElement('div');
            examDiv.className = 'exam-type-group';
            examDiv.innerHTML = `<h3>Tipo de Prova ${i}</h3>`;

            // Para cada questão (Q1 a Q10), crie campos de entrada
            for (let q = 1; q <= 10; q++) {
                const qKey = `Q${q}`;
                const questionDiv = document.createElement('div');
                questionDiv.className = 'question-input-group';
                questionDiv.innerHTML = `
                    <label>Questão ${q}:</label>
                    <select name="exam_answers_${i}_${qKey}_correct" class="correct-answer-select" data-exam-type="${i}" data-question-key="${qKey}">
                        <option value="">Selecione</option>
                        <option value="A">A</option>
                        <option value="B">B</option>
                        <option value="C">C</option>
                        <option value="D">D</option>
                        <option value="E">E</option>
                    </select>
                    <input type="number" step="0.1" min="0" max="2" value="${gabaritoExample[qKey] ? gabaritoExample[qKey].peso_questao : 1.0}"
                           name="exam_answers_${i}_${qKey}_peso_questao" placeholder="Peso Q${q}" class="peso-questao-input">
                    `;
                examDiv.appendChild(questionDiv);
            }
            examAnswersContainer.appendChild(examDiv);
        }
    }

    numExamTypesInput.addEventListener('change', generateExamAnswerFields);
    generateExamAnswerFields(); // Gera os campos iniciais ao carregar a página

    correctionForm.addEventListener('submit', async function(event) {
        event.preventDefault(); // Impede o envio padrão do formulário

        messageDiv.textContent = 'Processando... Por favor, aguarde.';
        messageDiv.style.color = 'blue';

        const formData = new FormData(correctionForm);

        // Coletar as respostas certas de cada tipo de prova em um JSON
        const numTypes = parseInt(numExamTypesInput.value);
        const examAnswers = {};
        for (let i = 1; i <= numTypes; i++) {
            examAnswers[i] = {};
            for (let q = 1; q <= 10; q++) { // Assumindo 10 questões por prova
                const qKey = `Q${q}`;
                const correctSelect = document.querySelector(`[name="exam_answers_${i}_${qKey}_correct"]`);
                const pesoQuestaoInput = document.querySelector(`[name="exam_answers_${i}_${qKey}_peso_questao"]`);

                if (correctSelect && pesoQuestaoInput) {
                    const selectedCorrect = correctSelect.value;
                    const pesoQuestao = parseFloat(pesoQuestaoInput.value);

                    // Criar os pesos das alternativas baseado na resposta correta
                    const pesosAlternativas = {'A': 0.0, 'B': 0.0, 'C': 0.0, 'D': 0.0, 'E': 0.0};
                    if (selectedCorrect) {
                        pesosAlternativas[selectedCorrect] = 1.0; // Pontuação total para a correta
                    }

                    examAnswers[i][qKey] = {
                        'peso_questao': pesoQuestao,
                        'pesos_alternativas': pesosAlternativas,
                        'correta': selectedCorrect
                    };
                }
            }
        }
        formData.append('exam_answers', JSON.stringify(examAnswers));

        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            if (response.ok) {
                // Se a resposta for um arquivo, inicie o download
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = response.headers.get('Content-Disposition').split('filename=')[1].replace(/"/g, ''); // Extrai o nome do arquivo
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                messageDiv.textContent = 'Correção concluída! O arquivo Excel foi baixado.';
                messageDiv.style.color = 'green';
//                correctionForm.reset(); // Limpa o formulário após o sucesso
                generateExamAnswerFields(); // Regenera campos de gabarito
            } else {
                const errorText = await response.text();
                messageDiv.textContent = `Erro: ${errorText}`;
                messageDiv.style.color = 'red';
            }
        } catch (error) {
            console.error('Erro ao enviar o formulário:', error);
            messageDiv.textContent = 'Erro ao conectar com o servidor. Tente novamente.';
            messageDiv.style.color = 'red';
        }
    });
});