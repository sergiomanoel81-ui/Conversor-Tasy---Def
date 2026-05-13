// Mapeamento de exames com tipos de dados
const MAPEAMENTO_EXAMES = {
    "36486": { colunas: ["C_Hb", "hb", "HB"], nome: "Hemoglobina (Hb)", tipo: "NUMBER" },
    "36485": { colunas: ["C_Ht", "ht", "HT"], nome: "Hematócrito (Ht)", tipo: "NUMBER" },
    "36452": { colunas: ["UREI"], nome: "Uréia Pré", tipo: "NUMBER" },
    "36581": { colunas: ["UPD"], nome: "Uréia Pós-Diálise", tipo: "NUMBER" },
    "36434": { colunas: ["CREA", "crea"], nome: "Creatinina", tipo: "NUMBER" },
    "36433": { colunas: ["CALCIO", "cálcio"], nome: "Cálcio", tipo: "NUMBER" },
    "36435": { colunas: ["FOSFS"], nome: "Fósforo", tipo: "NUMBER" },
    "36461": { colunas: ["Na"], nome: "Sódio", tipo: "NUMBER" },
    "36436": { colunas: ["POTAS"], nome: "Potássio (K)", tipo: "NUMBER" },
    "36437": { colunas: ["TGP"], nome: "TGP (ALT)", tipo: "NUMBER" },
    "36438": { colunas: ["GLIC"], nome: "Glicose", tipo: "NUMBER" },
    "36447": { colunas: ["CTOT"], nome: "Colesterol Total", tipo: "NUMBER" },
    "36448": { colunas: ["ALU_SER"], nome: "Alumínio Sérico", tipo: "NUMBER" },
    "36449": { colunas: ["TRIG"], nome: "Triglicerídeos", tipo: "NUMBER" },
    "36450": { colunas: ["T4"], nome: "T4 Total", tipo: "NUMBER" },
    "36453": { colunas: ["Hb_A1c"], nome: "Hemoglobina Glicada (HbA1c)", tipo: "VARCHAR2" },
    "36455": { colunas: ["TSH"], nome: "TSH", tipo: "NUMBER" },
    "36456": { colunas: ["ANTI_HBS"], nome: "Anti-HBs (Hepatite B)", tipo: "NUMBER" },
    "36457": { colunas: ["VITD25OH"], nome: "Vitamina D (25-OH)", tipo: "NUMBER" },
    "36483": { colunas: ["PLAQ"], nome: "Plaquetas", tipo: "NUMBER" },
    "36502": { colunas: ["PTH_DB"], nome: "PTH (Paratormônio)", tipo: "NUMBER" },
    "36518": { colunas: ["Ferritina", "FERRITINA"], nome: "Ferritina", tipo: "NUMBER" },
    "36520": { colunas: ["IST"], nome: "Índice Saturação Transferrina", tipo: "NUMBER" },
    "36522": { colunas: ["FA"], nome: "Fosfatase Alcalina", tipo: "NUMBER" },
    "36523": { colunas: ["PT"], nome: "Proteínas Totais", tipo: "NUMBER" },
    "36567": { colunas: ["FER"], nome: "Ferro Sérico", tipo: "NUMBER" },
    "36574": { colunas: ["HCV"], nome: "Anti-HCV (Hepatite C)", tipo: "VARCHAR2" },
    "36578": { colunas: ["AAU"], nome: "Anticorpo Anti-HBs", tipo: "VARCHAR2" },
    "36579": { colunas: ["HDL"], nome: "HDL Colesterol", tipo: "NUMBER" },
    "36580": { colunas: ["Col_LDL"], nome: "LDL Colesterol", tipo: "NUMBER" },
    "36582": { colunas: ["CTT"], nome: "Capacidade Total de Fixação", tipo: "NUMBER" },
    "36584": { colunas: ["ALB"], nome: "Albumina", tipo: "NUMBER" },
    "36585": { colunas: ["GLB"], nome: "Globulinas", tipo: "NUMBER" },
    "36587": { colunas: ["Rel_Alb_Gl"], nome: "Relação Albumina/Globulina", tipo: "NUMBER" }
};

const ORDEM_COLUNAS = [
    "NM_PACIENTE", "NR_ATENDIMENTO", "DT_RESULTADO", "DS_PROTOCOLO", "CD_ESTABELECIMENTO",
    "NR_EXAME_36433", "NR_EXAME_36434", "NR_EXAME_36435", "NR_EXAME_36436",
    "NR_EXAME_36437", "NR_EXAME_36438", "NR_EXAME_36439", "NR_EXAME_36447",
    "NR_EXAME_36448", "NR_EXAME_36449", "NR_EXAME_36450", "NR_EXAME_36452",
    "NR_EXAME_36453", "NR_EXAME_36455", "NR_EXAME_36456", "NR_EXAME_36457",
    "NR_EXAME_36461", "NR_EXAME_36483", "NR_EXAME_36485", "NR_EXAME_36486",
    "NR_EXAME_36501", "NR_EXAME_36502", "NR_EXAME_36518", "NR_EXAME_36520",
    "NR_EXAME_36522", "NR_EXAME_36523", "NR_EXAME_36567", "NR_EXAME_36574",
    "NR_EXAME_36578", "NR_EXAME_36579", "NR_EXAME_36580", "NR_EXAME_36581",
    "NR_EXAME_36582", "NR_EXAME_36584", "NR_EXAME_36585", "NR_EXAME_36587",
    "NR_EXAME_36588"
];

const ESTABELECIMENTOS = {
    "1": "MATRIZ",
    "3": "MONTE_SERRAT",
    "4": "CONVENIOS",
    "5": "RIO_VERMELHO",
    "7": "SANTO_ESTEVAO"
};

// Variáveis globais
let dadosTasy = null;
let dadosBasicos = null;
let dadosComplementares = null;
let colunasBasicos = [];
let colunasComplementares = [];
let resultadoAnalise = null;
let vinculacoesManuais = {};

// Estrutura de auditoria preenchida durante o processamento
let auditoria = {
    colunasOrfas: [],          // colunas das planilhas do lab não mapeadas
    examesCobertura: {},        // {codigoExame: {preenchidos, vazios, total, origemBasicos, origemComplementares}}
    porPaciente: [],            // [{nome, atendimento, preenchidos:[{cod,nome,valor,origem}], vazios:[{cod,nome}], numberZerados:[{cod,nome,textoOriginal}]}]
    numberZerados: [],          // lista achatada: campos NUMBER que receberam texto e ficaram vazios
    totalPreenchidos: 0,
    totalVazios: 0,
    origemBasicos: 0,
    origemComplementares: 0
};

// Conjunto de todas as colunas mapeadas (achatado), usado para detectar órfãs
const COLUNAS_MAPEADAS = new Set();
Object.values(MAPEAMENTO_EXAMES).forEach(cfg => {
    cfg.colunas.forEach(c => COLUNAS_MAPEADAS.add(c.toLowerCase()));
});
// Colunas estruturais conhecidas do lab que não são valores de exame
const COLUNAS_ESTRUTURAIS_LAB = new Set(['nome', 'dthr_os', 'os', 'cod', 'codigo', 'data', 'paciente']);

// Normalizar nome para comparação
function normalizarNome(nome) {
    if (!nome) return '';
    return nome.toString().trim().toUpperCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// Calcular similaridade entre nomes (Levenshtein simplificado)
function calcularSimilaridade(str1, str2) {
    const s1 = normalizarNome(str1);
    const s2 = normalizarNome(str2);
    
    if (s1 === s2) return 100;
    
    // Verificar se um contém o outro
    if (s1.includes(s2) || s2.includes(s1)) return 85;
    
    // Levenshtein distance simplificado
    const len1 = s1.length;
    const len2 = s2.length;
    const matrix = [];

    for (let i = 0; i <= len2; i++) {
        matrix[i] = [i];
    }
    for (let j = 0; j <= len1; j++) {
        matrix[0][j] = j;
    }

    for (let i = 1; i <= len2; i++) {
        for (let j = 1; j <= len1; j++) {
            if (s2.charAt(i - 1) === s1.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1,
                    matrix[i][j - 1] + 1,
                    matrix[i - 1][j] + 1
                );
            }
        }
    }

    const distance = matrix[len2][len1];
    const maxLen = Math.max(len1, len2);
    const similarity = ((maxLen - distance) / maxLen) * 100;
    
    return Math.round(similarity);
}

// Limpar valor numérico
function limparValorNumerico(valor) {
    if (valor === null || valor === undefined || valor === '') {
        return '';
    }
    
    if (typeof valor === 'number') {
        return valor;
    }
    
    const valorStr = valor.toString().trim();
    if (valorStr === '') return '';
    
    // Remove caracteres não numéricos, mantém apenas dígitos, vírgula, ponto e sinal negativo
    let numLimpo = valorStr.replace(/[^0-9.,\-]/g, '');
    
    // Se não sobrou nada numérico, retornar vazio
    if (!numLimpo || numLimpo === '-') return '';
    
    // Substituir vírgula por ponto
    numLimpo = numLimpo.replace(',', '.');
    
    // Tentar converter
    const numero = parseFloat(numLimpo);
    
    // Se conversão falhou, retornar vazio (IMPORTANTE: evita erro no Tasy)
    if (isNaN(numero)) return '';
    
    return numero;
}

// Ler arquivo Excel/CSV. Retorna { rows, headers } para permitir auditoria de colunas.
function lerArquivo(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { raw: false });
                // Cabeçalhos preservando ordem (linha 1 da planilha)
                const arr = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false });
                const headers = Array.isArray(arr[0]) ? arr[0].filter(h => h !== undefined && h !== null && h !== '') : [];
                resolve({ rows: jsonData, headers });
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Navegação entre fases
function irParaFase(numeroFase) {
    // Esconder todas as fases
    document.querySelectorAll('.phase').forEach(phase => {
        phase.classList.remove('active');
    });
    
    // Mostrar fase selecionada
    document.getElementById(`fase${numeroFase}`).classList.add('active');
    
    // Atualizar stepper
    document.querySelectorAll('.step').forEach(step => {
        const stepNum = parseInt(step.dataset.step);
        step.classList.remove('active', 'completed');
        
        if (stepNum === numeroFase) {
            step.classList.add('active');
        } else if (stepNum < numeroFase) {
            step.classList.add('completed');
        }
    });
}

function voltarFase(numeroFase) {
    irParaFase(numeroFase);
}

// Upload de arquivos
document.getElementById('fileTasy').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (file) {
        try {
            const { rows } = await lerArquivo(file);
            dadosTasy = rows;
            const box = document.getElementById('boxTasy');
            box.classList.add('uploaded');
            box.querySelector('.status-badge').textContent = `✓ ${dadosTasy.length} registros`;
            box.querySelector('.status-badge').className = 'status-badge badge-success';
            verificarArquivosCarregados();
        } catch (error) {
            alert('Erro ao ler arquivo Tasy: ' + error.message);
        }
    }
});

document.getElementById('fileBasicos').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (file) {
        try {
            const { rows, headers } = await lerArquivo(file);
            dadosBasicos = rows;
            colunasBasicos = headers;
            const box = document.getElementById('boxBasicos');
            box.classList.add('uploaded');
            box.querySelector('.status-badge').textContent = `✓ ${dadosBasicos.length} registros`;
            box.querySelector('.status-badge').className = 'status-badge badge-success';
            verificarArquivosCarregados();
        } catch (error) {
            alert('Erro ao ler arquivo Básicos: ' + error.message);
        }
    }
});

document.getElementById('fileComplementares').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (file) {
        try {
            const { rows, headers } = await lerArquivo(file);
            dadosComplementares = rows;
            colunasComplementares = headers;
            const box = document.getElementById('boxComplementares');
            box.classList.add('uploaded');
            box.querySelector('.status-badge').textContent = `✓ ${dadosComplementares.length} registros`;
            box.querySelector('.status-badge').className = 'status-badge badge-success';
            verificarArquivosCarregados();
        } catch (error) {
            alert('Erro ao ler arquivo Complementares: ' + error.message);
        }
    }
});

function verificarArquivosCarregados() {
    const estabelecimento = document.getElementById('estabelecimento').value;
    const protocolo = document.getElementById('protocolo').value;
    
    const todosCarregados = dadosTasy && 
                           (dadosBasicos || dadosComplementares) && 
                           estabelecimento && 
                           protocolo;
    
    document.getElementById('btnAnalisar').disabled = !todosCarregados;
}

document.getElementById('estabelecimento').addEventListener('change', verificarArquivosCarregados);
document.getElementById('protocolo').addEventListener('input', verificarArquivosCarregados);

// Analisar arquivos
document.getElementById('btnAnalisar').addEventListener('click', analisarArquivos);

function analisarArquivos() {
    // Unir dados do laboratório
    const dadosLab = [];
    const nomesLab = new Set();
    
    if (dadosBasicos) {
        dadosBasicos.forEach(row => {
            const nomeNorm = normalizarNome(row.nome);
            if (!nomesLab.has(nomeNorm)) {
                dadosLab.push(row);
                nomesLab.add(nomeNorm);
            }
        });
    }
    
    if (dadosComplementares) {
        dadosComplementares.forEach(row => {
            const nomeNorm = normalizarNome(row.nome);
            if (!nomesLab.has(nomeNorm)) {
                dadosLab.push(row);
                nomesLab.add(nomeNorm);
            }
        });
    }
    
    // Analisar cruzamentos
    const encontrados = [];
    const divergencias = [];
    const semResultados = [];
    
    dadosTasy.forEach(pacienteTasy => {
        const nomeTasy = pacienteTasy.nm_paciente || pacienteTasy.NM_PACIENTE || '';
        const nomeTasyNorm = normalizarNome(nomeTasy);
        
        // Buscar match exato
        const matchExato = dadosLab.find(lab => 
            normalizarNome(lab.nome) === nomeTasyNorm
        );
        
        if (matchExato) {
            encontrados.push({
                tasy: pacienteTasy,
                lab: matchExato,
                tipo: 'exato'
            });
        } else {
            // Buscar match similar (>= 80% similaridade)
            let melhorMatch = null;
            let melhorSimilaridade = 0;
            
            dadosLab.forEach(lab => {
                const similaridade = calcularSimilaridade(nomeTasy, lab.nome);
                if (similaridade >= 80 && similaridade > melhorSimilaridade) {
                    melhorMatch = lab;
                    melhorSimilaridade = similaridade;
                }
            });
            
            if (melhorMatch) {
                divergencias.push({
                    tasy: pacienteTasy,
                    lab: melhorMatch,
                    similaridade: melhorSimilaridade,
                    vinculado: false
                });
            } else {
                semResultados.push(pacienteTasy);
            }
        }
    });
    
    resultadoAnalise = {
        encontrados,
        divergencias,
        semResultados,
        totalTasy: dadosTasy.length
    };
    
    mostrarDashboard();
    irParaFase(2);
}

function mostrarDashboard() {
    const dashboard = document.getElementById('dashboard');
    dashboard.innerHTML = `
        <div class="card success">
            <div class="card-icon">✅</div>
            <div class="card-number">${resultadoAnalise.encontrados.length}</div>
            <div class="card-title">Pacientes Encontrados</div>
        </div>
        
        <div class="card warning">
            <div class="card-icon">⚠️</div>
            <div class="card-number">${resultadoAnalise.divergencias.length}</div>
            <div class="card-title">Divergências de Nome</div>
        </div>
        
        <div class="card danger">
            <div class="card-icon">❌</div>
            <div class="card-number">${resultadoAnalise.semResultados.length}</div>
            <div class="card-title">Sem Resultados</div>
        </div>
        
        <div class="card">
            <div class="card-icon">📊</div>
            <div class="card-number">${resultadoAnalise.totalTasy}</div>
            <div class="card-title">Total no Tasy</div>
        </div>
    `;
    
    // Mostrar divergências se houver
    if (resultadoAnalise.divergencias.length > 0) {
        document.getElementById('divergenciasSection').style.display = 'block';
        mostrarDivergencias();
    }
    
    // Mostrar pacientes sem resultados
    if (resultadoAnalise.semResultados.length > 0) {
        document.getElementById('semResultadosSection').style.display = 'block';
        mostrarSemResultados();
    }
    
    verificarConfirmacao();
}

function mostrarDivergencias() {
    const lista = document.getElementById('listaDivergencias');
    lista.innerHTML = '';
    
    resultadoAnalise.divergencias.forEach((div, index) => {
        const item = document.createElement('div');
        item.className = 'divergence-item';
        item.id = `div-${index}`;
        
        const nomeTasy = div.tasy.nm_paciente || div.tasy.NM_PACIENTE;
        const atendimento = div.tasy.nr_atendimento || div.tasy.NR_ATENDIMENTO;
        
        item.innerHTML = `
            <div class="divergence-names">
                <div class="name-box" style="background: #fff3cd;">
                    <div class="name-label">📋 Tasy - Atend. ${atendimento}</div>
                    <div class="name-value">${nomeTasy}</div>
                </div>
                <div class="arrow">→</div>
                <div class="name-box" style="background: #d1ecf1;">
                    <div class="name-label">🧪 Laboratório (${div.similaridade}% similar)</div>
                    <div class="name-value">${div.lab.nome}</div>
                </div>
            </div>
            <div class="divergence-actions">
                <button class="btn btn-success btn-small" onclick="vincularNome(${index})">
                    ✓ Vincular
                </button>
                <button class="btn btn-primary btn-small" onclick="ignorarNome(${index})">
                    ✗ Ignorar
                </button>
            </div>
        `;
        
        lista.appendChild(item);
    });
}

function vincularNome(index) {
    const div = resultadoAnalise.divergencias[index];
    div.vinculado = true;
    
    const nomeTasy = normalizarNome(div.tasy.nm_paciente || div.tasy.NM_PACIENTE);
    const nomeLab = normalizarNome(div.lab.nome);
    vinculacoesManuais[nomeTasy] = nomeLab;
    
    const item = document.getElementById(`div-${index}`);
    item.classList.add('resolved');
    item.querySelector('.divergence-actions').innerHTML = `
        <span class="status-badge badge-success">✓ Vinculado</span>
    `;
    
    verificarConfirmacao();
}

function ignorarNome(index) {
    const div = resultadoAnalise.divergencias[index];
    div.vinculado = false;
    
    // Mover para sem resultados
    resultadoAnalise.semResultados.push(div.tasy);
    resultadoAnalise.divergencias.splice(index, 1);
    
    // Atualizar interface
    mostrarDashboard();
}

function mostrarSemResultados() {
    const lista = document.getElementById('listaSemResultados');
    let html = '<ul style="list-style: none; padding: 0;">';
    
    resultadoAnalise.semResultados.forEach(pac => {
        const nome = pac.nm_paciente || pac.NM_PACIENTE;
        const atend = pac.nr_atendimento || pac.NR_ATENDIMENTO;
        html += `
            <li style="padding: 10px; border-bottom: 1px solid #dee2e6; display: flex; justify-content: space-between;">
                <span><strong>Atend. ${atend}:</strong> ${nome}</span>
                <span class="status-badge badge-danger">Sem resultado</span>
            </li>
        `;
    });
    
    html += '</ul>';
    lista.innerHTML = html;
}

function verificarConfirmacao() {
    // Verificar se todas divergências foram resolvidas
    const todasResolvidas = resultadoAnalise.divergencias.every(div => div.vinculado);
    document.getElementById('btnConfirmar').disabled = !todasResolvidas && resultadoAnalise.divergencias.length > 0;
}

// Confirmar e processar
document.getElementById('btnConfirmar').addEventListener('click', function() {
    prepararProcessamento();
    irParaFase(3);
});

function prepararProcessamento() {
    const totalProcessar = resultadoAnalise.encontrados.length + 
                          resultadoAnalise.divergencias.filter(d => d.vinculado).length;
    
    const summary = document.getElementById('summaryBox');
    summary.innerHTML = `
        <div class="summary-item">
            <span class="summary-label">Total no Tasy:</span>
            <span class="summary-value">${resultadoAnalise.totalTasy}</span>
        </div>
        <div class="summary-item">
            <span class="summary-label">Serão processados:</span>
            <span class="summary-value" style="color: #28a745;">${totalProcessar}</span>
        </div>
        <div class="summary-item">
            <span class="summary-label">Excluídos (sem resultado):</span>
            <span class="summary-value" style="color: #dc3545;">${resultadoAnalise.semResultados.length}</span>
        </div>
        <div class="summary-item">
            <span class="summary-label">Estabelecimento:</span>
            <span class="summary-value">${ESTABELECIMENTOS[document.getElementById('estabelecimento').value]}</span>
        </div>
        <div class="summary-item">
            <span class="summary-label">Protocolo:</span>
            <span class="summary-value">${document.getElementById('protocolo').value}</span>
        </div>
    `;
}

// Processar e gerar arquivos
document.getElementById('btnProcessar').addEventListener('click', processarArquivos);

async function processarArquivos() {
    const btn = document.getElementById('btnProcessar');
    btn.disabled = true;
    btn.textContent = '⏳ Processando...';
    
    document.getElementById('progressBar').style.display = 'block';
    document.getElementById('actionButtons3').style.display = 'none';
    
    try {
        // Reset da auditoria antes de processar
        auditoria = {
            colunasOrfas: [],
            examesCobertura: {},
            porPaciente: [],
            numberZerados: [],
            totalPreenchidos: 0,
            totalVazios: 0,
            origemBasicos: 0,
            origemComplementares: 0
        };
        detectarColunasOrfas();

        // Preparar dados para processamento
        const pacientesProcessar = [];
        
        // Adicionar encontrados
        resultadoAnalise.encontrados.forEach(item => {
            pacientesProcessar.push({
                tasy: item.tasy,
                nomeLabNormalizado: normalizarNome(item.lab.nome)
            });
        });
        
        // Adicionar vinculados
        resultadoAnalise.divergencias.forEach(item => {
            if (item.vinculado) {
                pacientesProcessar.push({
                    tasy: item.tasy,
                    nomeLabNormalizado: normalizarNome(item.lab.nome)
                });
            }
        });
        
        const estabelecimento = document.getElementById('estabelecimento').value;
        const protocolo = document.getElementById('protocolo').value;
        
        // Processar
        const resultados = [];
        const total = pacientesProcessar.length;
        
        for (let i = 0; i < pacientesProcessar.length; i++) {
            const item = pacientesProcessar[i];
            const percentual = Math.round(((i + 1) / total) * 90);
            atualizarProgresso(percentual);
            
            await new Promise(resolve => setTimeout(resolve, 10)); // Pequeno delay para UI
            
            const linha = processarPaciente(item.tasy, item.nomeLabNormalizado, estabelecimento, protocolo);
            resultados.push(linha);
        }
        
        atualizarProgresso(95);
        
        // Gerar arquivos
        const nomeArquivo = await gerarArquivoExcel(resultados, estabelecimento);
        await gerarLog(resultados, nomeArquivo);
        
        atualizarProgresso(100);

        // Mostrar resultado
        document.getElementById('resultSection').style.display = 'block';
        document.getElementById('btnDownloadExcel').style.display = 'inline-block';
        document.getElementById('btnDownloadLog').style.display = 'inline-block';

        // Renderizar relatório de auditoria na tela
        mostrarAuditoria();
        
    } catch (error) {
        alert('Erro ao processar: ' + error.message);
        btn.disabled = false;
        btn.textContent = '🚀 Gerar Arquivos';
    }
}

function processarPaciente(pacienteTasy, nomeLabNormalizado, estabelecimento, protocolo) {
    const nomePaciente = pacienteTasy.nm_paciente || pacienteTasy.NM_PACIENTE;
    const atendimento = pacienteTasy.nr_atendimento || pacienteTasy.NR_ATENDIMENTO;

    const linha = {
        NM_PACIENTE: nomePaciente,
        NR_ATENDIMENTO: atendimento,
        DT_RESULTADO: '',
        DS_PROTOCOLO: protocolo,
        CD_ESTABELECIMENTO: estabelecimento
    };

    const auditPaciente = {
        nome: nomePaciente,
        atendimento: atendimento,
        preenchidos: [],
        vazios: [],
        numberZerados: []
    };

    Object.keys(MAPEAMENTO_EXAMES).forEach(codigoExame => {
        const config = MAPEAMENTO_EXAMES[codigoExame];
        const dados = encontrarDadosExame(codigoExame, nomeLabNormalizado);
        const valor = dados.valor !== null ? dados.valor : '';
        linha[`NR_EXAME_${codigoExame}`] = valor;

        if (!linha.DT_RESULTADO && dados.data) {
            linha.DT_RESULTADO = dados.data;
        }

        // Garante slot na cobertura
        if (!auditoria.examesCobertura[codigoExame]) {
            auditoria.examesCobertura[codigoExame] = {
                nome: config.nome,
                tipo: config.tipo,
                preenchidos: 0,
                vazios: 0,
                origemBasicos: 0,
                origemComplementares: 0,
                numberZerados: 0
            };
        }
        const cov = auditoria.examesCobertura[codigoExame];

        if (valor !== '' && valor !== null && valor !== undefined) {
            cov.preenchidos++;
            auditoria.totalPreenchidos++;
            auditPaciente.preenchidos.push({ codigo: codigoExame, nome: config.nome, valor: valor, origem: dados.origem });
            if (dados.origem === 'basicos') { cov.origemBasicos++; auditoria.origemBasicos++; }
            else if (dados.origem === 'complementares') { cov.origemComplementares++; auditoria.origemComplementares++; }
        } else {
            cov.vazios++;
            auditoria.totalVazios++;
            auditPaciente.vazios.push({ codigo: codigoExame, nome: config.nome });
        }

        if (dados.numberZerado) {
            cov.numberZerados++;
            const item = {
                paciente: nomePaciente,
                atendimento: atendimento,
                codigo: codigoExame,
                nome: config.nome,
                textoOriginal: dados.textoOriginal,
                origem: dados.origem
            };
            auditoria.numberZerados.push(item);
            auditPaciente.numberZerados.push(item);
        }
    });

    auditoria.porPaciente.push(auditPaciente);
    return linha;
}

function encontrarDadosExame(codigoExame, nomeLabNormalizado) {
    const config = MAPEAMENTO_EXAMES[codigoExame];
    if (!config) return { valor: null, data: null, origem: null };

    const datasources = [
        { dados: dadosBasicos, origem: 'basicos' },
        { dados: dadosComplementares, origem: 'complementares' }
    ].filter(d => d.dados !== null);

    for (const { dados: data, origem } of datasources) {
        for (const row of data) {
            if (normalizarNome(row.nome) === nomeLabNormalizado) {
                for (const coluna of config.colunas) {
                    if (row[coluna] !== undefined && row[coluna] !== null && row[coluna] !== '') {
                        const valorOriginal = row[coluna];
                        let valorFinal = valorOriginal;

                        if (config.tipo === 'NUMBER') {
                            valorFinal = limparValorNumerico(valorFinal);
                            // Se limpeza falhou, NUMBER recebeu texto: deixar vazio e registrar
                            if (valorFinal === '') {
                                return {
                                    valor: null,
                                    data: row.dthr_os || null,
                                    origem: origem,
                                    numberZerado: true,
                                    textoOriginal: valorOriginal
                                };
                            }
                        } else if (config.tipo === 'VARCHAR2') {
                            const textoStr = valorFinal.toString().trim();
                            valorFinal = textoStr.length > 1020 ? textoStr.substring(0, 1020) : textoStr;
                        }

                        return {
                            valor: valorFinal,
                            data: row.dthr_os || null,
                            origem: origem
                        };
                    }
                }
            }
        }
    }

    return { valor: null, data: null, origem: null };
}

// Detecta colunas das planilhas do laboratório que não estão mapeadas
function detectarColunasOrfas() {
    const orfas = [];
    const registrar = (header, planilha) => {
        if (!header) return;
        const lower = header.toString().toLowerCase().trim();
        if (!lower) return;
        if (COLUNAS_ESTRUTURAIS_LAB.has(lower)) return;
        if (COLUNAS_MAPEADAS.has(lower)) return;
        orfas.push({ coluna: header, planilha });
    };
    colunasBasicos.forEach(h => registrar(h, 'basicos'));
    colunasComplementares.forEach(h => registrar(h, 'complementares'));
    auditoria.colunasOrfas = orfas;
}

function atualizarProgresso(percentual) {
    const fill = document.getElementById('progressFill');
    fill.style.width = percentual + '%';
    fill.textContent = percentual + '%';
}

async function gerarArquivoExcel(resultados, estabelecimento) {
    const wb = XLSX.utils.book_new();

    // Aba principal: Exames
    const wsExames = XLSX.utils.json_to_sheet(resultados, { header: ORDEM_COLUNAS });
    XLSX.utils.book_append_sheet(wb, wsExames, 'Exames');

    // Aba: Auditoria - Cobertura por Exame
    const cobertura = Object.keys(MAPEAMENTO_EXAMES).map(codigo => {
        const cov = auditoria.examesCobertura[codigo];
        if (!cov) return null;
        const total = cov.preenchidos + cov.vazios;
        const pct = total > 0 ? Math.round((cov.preenchidos / total) * 100) : 0;
        return {
            CODIGO: codigo,
            EXAME: cov.nome,
            TIPO: cov.tipo,
            PREENCHIDOS: cov.preenchidos,
            VAZIOS: cov.vazios,
            TOTAL: total,
            'COBERTURA_%': pct,
            ORIGEM_BASICOS: cov.origemBasicos,
            ORIGEM_COMPLEMENTARES: cov.origemComplementares,
            NUMBER_ZERADOS: cov.numberZerados
        };
    }).filter(Boolean);
    const wsCobertura = XLSX.utils.json_to_sheet(cobertura);
    XLSX.utils.book_append_sheet(wb, wsCobertura, 'Auditoria_Cobertura');

    // Aba: Auditoria - Matriz Paciente x Exame
    // Formato planilha intuitiva: colunas nomeadas pelo exame, valores ✓ / em branco / ⚠
    const matriz = auditoria.porPaciente.map(p => {
        const linha = {
            'Atendimento': p.atendimento,
            'Paciente': p.nome,
            'Preenchidos': p.preenchidos.length,
            'Vazios': p.vazios.length
        };
        Object.keys(MAPEAMENTO_EXAMES).forEach(codigo => {
            const preenchido = p.preenchidos.find(e => e.codigo === codigo);
            const zerado = p.numberZerados.find(e => e.codigo === codigo);
            const nomeColuna = `${MAPEAMENTO_EXAMES[codigo].nome} (${codigo})`;
            if (zerado) linha[nomeColuna] = '⚠ texto';
            else if (preenchido) linha[nomeColuna] = preenchido.valor;
            else linha[nomeColuna] = '';
        });
        return linha;
    });
    const wsMatriz = XLSX.utils.json_to_sheet(matriz);
    // Ajusta largura das colunas para legibilidade
    const colWidths = [
        { wch: 12 }, { wch: 35 }, { wch: 12 }, { wch: 10 }
    ];
    Object.keys(MAPEAMENTO_EXAMES).forEach(() => colWidths.push({ wch: 22 }));
    wsMatriz['!cols'] = colWidths;
    XLSX.utils.book_append_sheet(wb, wsMatriz, 'Auditoria_Matriz');

    // Aba: Campos NUMBER zerados (recebido texto)
    if (auditoria.numberZerados.length > 0) {
        const wsNumber = XLSX.utils.json_to_sheet(auditoria.numberZerados.map(n => ({
            NM_PACIENTE: n.paciente,
            NR_ATENDIMENTO: n.atendimento,
            CODIGO: n.codigo,
            EXAME: n.nome,
            ORIGEM: n.origem,
            TEXTO_RECEBIDO: n.textoOriginal
        })));
        XLSX.utils.book_append_sheet(wb, wsNumber, 'Auditoria_NumberZerados');
    }

    // Aba: Colunas órfãs do laboratório
    if (auditoria.colunasOrfas.length > 0) {
        const wsOrfas = XLSX.utils.json_to_sheet(auditoria.colunasOrfas.map(o => ({
            COLUNA: o.coluna,
            PLANILHA: o.planilha,
            OBSERVACAO: 'Coluna presente no arquivo do laboratório mas não mapeada para código Tasy'
        })));
        XLSX.utils.book_append_sheet(wb, wsOrfas, 'Auditoria_ColunasOrfas');
    }

    const hoje = new Date();
    const dataFormatada = hoje.toISOString().split('T')[0].replace(/-/g, '');
    const nomeEstabelecimento = ESTABELECIMENTOS[estabelecimento];
    const nomeArquivo = `Exames_${nomeEstabelecimento}_${dataFormatada}.xls`;

    // Exportar em formato BIFF8 (.xls - Excel 97-2003)
    const wbout = XLSX.write(wb, { bookType: 'biff8', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.ms-excel' });

    document.getElementById('btnDownloadExcel').onclick = function() {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = nomeArquivo;
        a.click();
    };

    return nomeArquivo;
}

async function gerarLog(resultados, nomeArquivoExcel) {
    const hoje = new Date();
    const dataHora = hoje.toLocaleString('pt-BR');
    
    let log = `═══════════════════════════════════════════════════════════════
LOG DE PROCESSAMENTO - SISTEMA IMPORTAÇÃO TASY
═══════════════════════════════════════════════════════════════

Data/Hora: ${dataHora}
Arquivo Excel Gerado: ${nomeArquivoExcel}
Estabelecimento: ${ESTABELECIMENTOS[document.getElementById('estabelecimento').value]}
Protocolo: ${document.getElementById('protocolo').value}

═══════════════════════════════════════════════════════════════
RESUMO DO PROCESSAMENTO
═══════════════════════════════════════════════════════════════

Total de pacientes no Tasy: ${resultadoAnalise.totalTasy}
Pacientes processados: ${resultados.length}
Pacientes excluídos (sem resultado): ${resultadoAnalise.semResultados.length}

─────────────────────────────────────────────────────────────

✅ PACIENTES PROCESSADOS COM SUCESSO (${resultados.length})
─────────────────────────────────────────────────────────────
`;

    resultados.forEach((pac, i) => {
        log += `\n${i + 1}. Atend. ${pac.NR_ATENDIMENTO} - ${pac.NM_PACIENTE}`;
    });

    if (resultadoAnalise.semResultados.length > 0) {
        log += `\n\n═══════════════════════════════════════════════════════════════
❌ PACIENTES SEM RESULTADOS - EXCLUÍDOS DO ARQUIVO (${resultadoAnalise.semResultados.length})
═══════════════════════════════════════════════════════════════
`;

        resultadoAnalise.semResultados.forEach((pac, i) => {
            const nome = pac.nm_paciente || pac.NM_PACIENTE;
            const atend = pac.nr_atendimento || pac.NR_ATENDIMENTO;
            log += `\n${i + 1}. Atend. ${atend} - ${nome}`;
            log += `\n   MOTIVO: Não encontrado em nenhuma planilha do laboratório`;
        });
    }

    if (Object.keys(vinculacoesManuais).length > 0) {
        log += `\n\n═══════════════════════════════════════════════════════════════
⚠️ VINCULAÇÕES MANUAIS REALIZADAS (${Object.keys(vinculacoesManuais).length})
═══════════════════════════════════════════════════════════════
`;

        let i = 1;
        for (const [nomeTasy, nomeLab] of Object.entries(vinculacoesManuais)) {
            log += `\n${i}. Tasy: ${nomeTasy}`;
            log += `\n   Lab:  ${nomeLab}`;
            i++;
        }
    }

    // ═══════════════════════════════════════════════
    // SEÇÃO DE AUDITORIA
    // ═══════════════════════════════════════════════
    log += `\n\n═══════════════════════════════════════════════════════════════
📊 AUDITORIA - COBERTURA POR EXAME
═══════════════════════════════════════════════════════════════

Total de campos preenchidos:    ${auditoria.totalPreenchidos}
Total de campos vazios:         ${auditoria.totalVazios}
Origem - Exames Básicos:        ${auditoria.origemBasicos}
Origem - Exames Complementares: ${auditoria.origemComplementares}

─────────────────────────────────────────────────────────────
Cobertura por exame (preenchidos / total processados):
─────────────────────────────────────────────────────────────
`;

    Object.keys(MAPEAMENTO_EXAMES).forEach(codigo => {
        const cov = auditoria.examesCobertura[codigo];
        if (!cov) return;
        const total = cov.preenchidos + cov.vazios;
        const pct = total > 0 ? Math.round((cov.preenchidos / total) * 100) : 0;
        const nomeExame = cov.nome.padEnd(38);
        const cnt = `${cov.preenchidos}/${total}`.padStart(8);
        log += `\n  [${codigo}] ${nomeExame} ${cnt}  (${String(pct).padStart(3)}%)`;
        if (cov.numberZerados > 0) {
            log += `  ⚠️ ${cov.numberZerados} NUMBER zerado(s)`;
        }
    });

    // Detalhes por paciente
    log += `\n\n═══════════════════════════════════════════════════════════════
👥 AUDITORIA - DETALHE POR PACIENTE
═══════════════════════════════════════════════════════════════
`;

    auditoria.porPaciente.forEach((p, i) => {
        log += `\n\n${i + 1}. Atend. ${p.atendimento} - ${p.nome}`;
        log += `\n   ✅ ${p.preenchidos.length} preenchido(s) | ❌ ${p.vazios.length} vazio(s)`;
        if (p.preenchidos.length > 0) {
            log += `\n   Exames carregados:`;
            p.preenchidos.forEach(e => {
                log += `\n     • [${e.codigo}] ${e.nome} = ${e.valor} (origem: ${e.origem || 'n/d'})`;
            });
        }
        if (p.numberZerados.length > 0) {
            log += `\n   ⚠️ Campos NUMBER zerados (recebido texto):`;
            p.numberZerados.forEach(z => {
                log += `\n     • [${z.codigo}] ${z.nome} ← "${z.textoOriginal}"`;
            });
        }
    });

    // Campos NUMBER que receberam texto
    if (auditoria.numberZerados.length > 0) {
        log += `\n\n═══════════════════════════════════════════════════════════════
⚠️ AUDITORIA - CAMPOS NUMBER ZERADOS POR TEXTO (${auditoria.numberZerados.length})
═══════════════════════════════════════════════════════════════
`;
        auditoria.numberZerados.forEach((n, i) => {
            log += `\n${i + 1}. Atend. ${n.atendimento} - ${n.paciente}`;
            log += `\n   Exame: [${n.codigo}] ${n.nome}`;
            log += `\n   Texto recebido: "${n.textoOriginal}"`;
            log += `\n   Origem: ${n.origem || 'n/d'}\n`;
        });
    }

    // Colunas órfãs
    if (auditoria.colunasOrfas.length > 0) {
        log += `\n\n═══════════════════════════════════════════════════════════════
🔎 AUDITORIA - COLUNAS DO LABORATÓRIO NÃO MAPEADAS (${auditoria.colunasOrfas.length})
═══════════════════════════════════════════════════════════════

As colunas abaixo existem nas planilhas do lab mas NÃO estão no
MAPEAMENTO_EXAMES e por isso não foram exportadas:
`;
        auditoria.colunasOrfas.forEach((o, i) => {
            log += `\n${i + 1}. "${o.coluna}" (planilha: ${o.planilha})`;
        });
    }

    log += `\n\n═══════════════════════════════════════════════════════════════
ℹ️ OBSERVAÇÕES TÉCNICAS
═══════════════════════════════════════════════════════════════

⚠️ IMPORTANTE: Campos numéricos com valores não numéricos foram tratados:

- Campos NUMBER que receberam TEXTO do laboratório → deixados VAZIOS
- Campos VARCHAR2 aceitam até 1020 caracteres (truncados se exceder)

Exemplo: 
  Campo: NR_EXAME_36580 (LDL Colesterol) - Tipo: NUMBER
  Valor recebido: "Impossivel liberacao devido a hipertrigliceridemia"
  Processado: (vazio)
  Motivo: Campo NUMBER não aceita texto

Esta proteção EVITA ERROS na importação do Tasy por incompatibilidade
de tipo de dados.

`;

    log += `\n═══════════════════════════════════════════════════════════════
FIM DO LOG
═══════════════════════════════════════════════════════════════
`;

    // Salvar para download
    const blob = new Blob([log], { type: 'text/plain;charset=utf-8' });
    const dataFormatada = hoje.toISOString().split('T')[0].replace(/-/g, '');
    const nomeLog = `LOG_Processamento_${dataFormatada}.txt`;
    
    document.getElementById('btnDownloadLog').onclick = function() {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = nomeLog;
        a.click();
    };
}

// Renderiza o relatório de auditoria na tela (Fase 3, após processar)
function mostrarAuditoria() {
    const container = document.getElementById('auditoriaContainer');
    if (!container) return;

    const totalCampos = auditoria.totalPreenchidos + auditoria.totalVazios;
    const pctGeral = totalCampos > 0 ? Math.round((auditoria.totalPreenchidos / totalCampos) * 100) : 0;

    // Cards resumo
    let html = `
        <h3 style="margin: 30px 0 15px; color: #495057;">📊 Relatório de Auditoria</h3>
        <div class="dashboard">
            <div class="card success">
                <div class="card-icon">✅</div>
                <div class="card-number">${auditoria.totalPreenchidos}</div>
                <div class="card-title">Campos Preenchidos (${pctGeral}%)</div>
            </div>
            <div class="card warning">
                <div class="card-icon">⚪</div>
                <div class="card-number">${auditoria.totalVazios}</div>
                <div class="card-title">Campos Vazios</div>
            </div>
            <div class="card">
                <div class="card-icon">📋</div>
                <div class="card-number">${auditoria.origemBasicos}</div>
                <div class="card-title">Origem: Básicos</div>
            </div>
            <div class="card">
                <div class="card-icon">🔬</div>
                <div class="card-number">${auditoria.origemComplementares}</div>
                <div class="card-title">Origem: Complementares</div>
            </div>
        </div>
    `;

    // Resumo simples por paciente (sem a matriz larga)
    html += `
        <h4 style="margin: 25px 0 10px; color: #495057;">Resumo por Paciente</h4>
        <div class="alert alert-info">
            <span>📊</span>
            <div>A matriz detalhada <strong>Paciente × Exame</strong> está na aba <strong>Auditoria_Matriz</strong> do arquivo <code>.xls</code> baixado — formato planilha, fácil de filtrar e imprimir.</div>
        </div>
        <div style="max-height: 400px; overflow-y: auto; background: #f8f9fa; border-radius: 10px; padding: 15px;">
        <table class="audit-table">
            <thead><tr><th>Atend.</th><th>Paciente</th><th>Preenchidos</th><th>Vazios</th><th>NUMBER zerados</th></tr></thead>
            <tbody>
    `;
    auditoria.porPaciente.forEach(p => {
        const total = p.preenchidos.length + p.vazios.length;
        const pct = total > 0 ? Math.round((p.preenchidos.length / total) * 100) : 0;
        const cor = pct >= 80 ? '#28a745' : pct >= 40 ? '#ffc107' : '#dc3545';
        html += `<tr>
            <td>${p.atendimento}</td>
            <td>${p.nome}</td>
            <td style="color:#28a745; font-weight:600;">${p.preenchidos.length}/${total} <span style="color:${cor};">(${pct}%)</span></td>
            <td style="color:#6c757d;">${p.vazios.length}</td>
            <td>${p.numberZerados.length > 0 ? `<span class="status-badge badge-danger">${p.numberZerados.length}</span>` : '—'}</td>
        </tr>`;
    });
    html += `</tbody></table></div>`;

    // NUMBER zerados
    if (auditoria.numberZerados.length > 0) {
        html += `
            <h4 style="margin: 25px 0 10px; color: #721c24;">⚠️ Campos NUMBER Zerados (${auditoria.numberZerados.length})</h4>
            <div class="alert alert-warning">
                <span>ℹ️</span>
                <div>Os campos abaixo são <strong>NUMBER</strong> no Tasy mas o laboratório enviou texto. Foram deixados vazios para evitar erro de importação.</div>
            </div>
            <div style="max-height: 250px; overflow-y: auto; background: #f8f9fa; border-radius: 10px; padding: 15px;">
            <table class="audit-table">
                <thead><tr><th>Atend.</th><th>Paciente</th><th>Código</th><th>Exame</th><th>Origem</th><th>Texto recebido</th></tr></thead>
                <tbody>
        `;
        auditoria.numberZerados.forEach(n => {
            html += `<tr>
                <td>${n.atendimento}</td>
                <td>${n.paciente}</td>
                <td><code>${n.codigo}</code></td>
                <td>${n.nome}</td>
                <td>${n.origem || '—'}</td>
                <td style="color:#721c24;">"${n.textoOriginal}"</td>
            </tr>`;
        });
        html += `</tbody></table></div>`;
    }

    // Colunas órfãs
    if (auditoria.colunasOrfas.length > 0) {
        html += `
            <h4 style="margin: 25px 0 10px; color: #856404;">🔎 Colunas do Lab Não Mapeadas (${auditoria.colunasOrfas.length})</h4>
            <div class="alert alert-info">
                <span>ℹ️</span>
                <div>Colunas presentes nas planilhas do laboratório que <strong>não estão no mapeamento</strong> e por isso não foram exportadas. Se forem relevantes, adicione-as ao <code>MAPEAMENTO_EXAMES</code>.</div>
            </div>
            <div style="background: #f8f9fa; border-radius: 10px; padding: 15px;">
            <table class="audit-table">
                <thead><tr><th>Coluna</th><th>Planilha</th></tr></thead>
                <tbody>
        `;
        auditoria.colunasOrfas.forEach(o => {
            html += `<tr><td><code>${o.coluna}</code></td><td>${o.planilha}</td></tr>`;
        });
        html += `</tbody></table></div>`;
    }

    container.innerHTML = html;
    container.style.display = 'block';
}
