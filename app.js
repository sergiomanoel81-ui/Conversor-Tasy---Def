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
let resultadoAnalise = null;
let vinculacoesManuais = {};
let camposVaziosPorTipo = []; // Rastrear campos deixados vazios

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

// Ler arquivo Excel/CSV
function lerArquivo(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { raw: false });
                resolve(jsonData);
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
            dadosTasy = await lerArquivo(file);
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
            dadosBasicos = await lerArquivo(file);
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
            dadosComplementares = await lerArquivo(file);
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
        
    } catch (error) {
        alert('Erro ao processar: ' + error.message);
        btn.disabled = false;
        btn.textContent = '🚀 Gerar Arquivos';
    }
}

function processarPaciente(pacienteTasy, nomeLabNormalizado, estabelecimento, protocolo) {
    const nomePaciente = pacienteTasy.nm_paciente || pacienteTasy.NM_PACIENTE;
    const atendimento = pacienteTasy.nr_atendimento || pacienteTasy.NR_ATENDIMENTO;
    
    // Buscar dados nos laboratórios
    let dataResultado = '';
    
    const linha = {
        NM_PACIENTE: nomePaciente,
        NR_ATENDIMENTO: atendimento,
        DT_RESULTADO: '',
        DS_PROTOCOLO: protocolo,
        CD_ESTABELECIMENTO: estabelecimento
    };
    
    // Buscar valores de todos os exames
    Object.keys(MAPEAMENTO_EXAMES).forEach(codigoExame => {
        const dados = encontrarDadosExame(codigoExame, nomeLabNormalizado);
        linha[`NR_EXAME_${codigoExame}`] = dados.valor !== null ? dados.valor : '';
        
        if (!linha.DT_RESULTADO && dados.data) {
            linha.DT_RESULTADO = dados.data;
        }
    });
    
    return linha;
}

function encontrarDadosExame(codigoExame, nomeLabNormalizado) {
    const config = MAPEAMENTO_EXAMES[codigoExame];
    if (!config) return { valor: null, data: null };
    
    const datasources = [dadosBasicos, dadosComplementares].filter(d => d !== null);
    
    for (const data of datasources) {
        for (const row of data) {
            if (normalizarNome(row.nome) === nomeLabNormalizado) {
                for (const coluna of config.colunas) {
                    if (row[coluna] !== undefined && row[coluna] !== null && row[coluna] !== '') {
                        let valorFinal = row[coluna];
                        
                        if (config.tipo === 'NUMBER') {
                            valorFinal = limparValorNumerico(valorFinal);
                            // Se limpeza falhou, deixar vazio
                            if (valorFinal === '') {
                                return { valor: null, data: row.dthr_os || null };
                            }
                        } else if (config.tipo === 'VARCHAR2') {
                            // VARCHAR2 aceita até 1020 caracteres no Tasy
                            const textoStr = valorFinal.toString().trim();
                            if (textoStr.length > 1020) {
                                // Truncar se exceder limite
                                valorFinal = textoStr.substring(0, 1020);
                            } else {
                                valorFinal = textoStr;
                            }
                        }
                        
                        return {
                            valor: valorFinal,
                            data: row.dthr_os || null
                        };
                    }
                }
            }
        }
    }
    
    return { valor: null, data: null };
}

function atualizarProgresso(percentual) {
    const fill = document.getElementById('progressFill');
    fill.style.width = percentual + '%';
    fill.textContent = percentual + '%';
}

async function gerarArquivoExcel(resultados, estabelecimento) {
    const ws = XLSX.utils.json_to_sheet(resultados, { header: ORDEM_COLUNAS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Exames');
    
    const hoje = new Date();
    const dataFormatada = hoje.toISOString().split('T')[0].replace(/-/g, '');
    const nomeEstabelecimento = ESTABELECIMENTOS[estabelecimento];
    const nomeArquivo = `Exames_${nomeEstabelecimento}_${dataFormatada}.xlsx`;
    
    // Salvar para download
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    
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
