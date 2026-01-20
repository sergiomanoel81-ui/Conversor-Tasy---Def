// Mapeamento de exames (mesmo do sistema principal)
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

let arquivoGerado = null;
let arquivoOriginalBasicos = null;
let arquivoOriginalComplementares = null;
let resultadoValidacao = null;

// Normalizar nome
function normalizarNome(nome) {
    if (!nome) return '';
    return nome.toString().trim().toUpperCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// Normalizar valor para comparação
function normalizarValor(valor, tipo) {
    if (valor === null || valor === undefined || valor === '') {
        return null;
    }
    
    if (tipo === 'NUMBER') {
        // Converter para número
        if (typeof valor === 'number') {
            return Number(valor.toFixed(2));
        }
        
        const valorStr = valor.toString().trim();
        if (valorStr === '') return null;
        
        // Limpar e extrair número
        let numLimpo = valorStr.replace(/[^0-9.,\-]/g, '');
        if (!numLimpo || numLimpo === '-') return null;
        
        numLimpo = numLimpo.replace(',', '.');
        const numero = parseFloat(numLimpo);
        
        if (isNaN(numero)) return null;
        
        return Number(numero.toFixed(2));
    } else {
        // VARCHAR2 - comparação de texto
        return valor.toString().trim().toUpperCase();
    }
}

// Comparar valores
function valoresIguais(valor1, valor2, tipo) {
    const v1 = normalizarValor(valor1, tipo);
    const v2 = normalizarValor(valor2, tipo);
    
    // Ambos vazios/null
    if (v1 === null && v2 === null) return true;
    
    // Um vazio e outro não
    if (v1 === null || v2 === null) return false;
    
    if (tipo === 'NUMBER') {
        // Para números, considerar iguais se diferença < 0.01
        return Math.abs(v1 - v2) < 0.01;
    } else {
        // Para texto, comparação direta
        return v1 === v2;
    }
}

// Ler arquivo
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

// Upload de arquivos
document.getElementById('fileGerado').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (file) {
        try {
            arquivoGerado = await lerArquivo(file);
            const box = document.getElementById('boxGerado');
            box.classList.add('uploaded');
            box.querySelector('.status-badge').textContent = `✓ ${arquivoGerado.length} registros`;
            box.querySelector('.status-badge').className = 'status-badge badge-success';
            verificarArquivos();
        } catch (error) {
            alert('Erro ao ler arquivo gerado: ' + error.message);
        }
    }
});

document.getElementById('fileOriginalBasicos').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (file) {
        try {
            arquivoOriginalBasicos = await lerArquivo(file);
            const box = document.getElementById('boxOriginalBasicos');
            box.classList.add('uploaded');
            box.querySelector('.status-badge').textContent = `✓ ${arquivoOriginalBasicos.length} registros`;
            box.querySelector('.status-badge').className = 'status-badge badge-success';
            verificarArquivos();
        } catch (error) {
            alert('Erro ao ler arquivo original básicos: ' + error.message);
        }
    }
});

document.getElementById('fileOriginalComplementares').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (file) {
        try {
            arquivoOriginalComplementares = await lerArquivo(file);
            const box = document.getElementById('boxOriginalComplementares');
            box.classList.add('uploaded');
            box.querySelector('.status-badge').textContent = `✓ ${arquivoOriginalComplementares.length} registros`;
            box.querySelector('.status-badge').className = 'status-badge badge-success';
            verificarArquivos();
        } catch (error) {
            alert('Erro ao ler arquivo original complementares: ' + error.message);
        }
    }
});

function verificarArquivos() {
    const todosCarregados = arquivoGerado && 
                           (arquivoOriginalBasicos || arquivoOriginalComplementares);
    
    document.getElementById('btnValidar').disabled = !todosCarregados;
}

// Validar arquivos
document.getElementById('btnValidar').addEventListener('click', validarArquivos);

async function validarArquivos() {
    const btn = document.getElementById('btnValidar');
    btn.disabled = true;
    btn.textContent = '⏳ Validando...';
    
    const progressBar = document.getElementById('progressBar');
    progressBar.style.display = 'block';
    
    try {
        // Unir dados originais
        const dadosOriginais = [];
        
        if (arquivoOriginalBasicos) {
            arquivoOriginalBasicos.forEach(row => {
                const nomeNorm = normalizarNome(row.nome);
                const existente = dadosOriginais.find(d => normalizarNome(d.nome) === nomeNorm);
                if (!existente) {
                    dadosOriginais.push({ ...row });
                } else {
                    // Mesclar dados
                    Object.assign(existente, row);
                }
            });
        }
        
        if (arquivoOriginalComplementares) {
            arquivoOriginalComplementares.forEach(row => {
                const nomeNorm = normalizarNome(row.nome);
                const existente = dadosOriginais.find(d => normalizarNome(d.nome) === nomeNorm);
                if (!existente) {
                    dadosOriginais.push({ ...row });
                } else {
                    // Mesclar dados
                    Object.assign(existente, row);
                }
            });
        }
        
        atualizarProgresso(20);
        
        // Validar cada linha do arquivo gerado
        const validacoes = [];
        let totalOk = 0;
        let totalDivergente = 0;
        let totalFaltando = 0;
        
        const total = arquivoGerado.length;
        
        for (let i = 0; i < arquivoGerado.length; i++) {
            const linhaGerada = arquivoGerado[i];
            const percentual = 20 + Math.round(((i + 1) / total) * 70);
            atualizarProgresso(percentual);
            
            await new Promise(resolve => setTimeout(resolve, 5)); // Pequeno delay para UI
            
            const nomePaciente = linhaGerada.NM_PACIENTE;
            const atendimento = linhaGerada.NR_ATENDIMENTO;
            const nomeNormalizado = normalizarNome(nomePaciente);
            
            // Buscar paciente nos dados originais
            const pacienteOriginal = dadosOriginais.find(d => 
                normalizarNome(d.nome) === nomeNormalizado
            );
            
            if (!pacienteOriginal) {
                // Paciente não encontrado nos originais
                totalFaltando++;
                validacoes.push({
                    atendimento,
                    paciente: nomePaciente,
                    status: 'PACIENTE_NAO_ENCONTRADO',
                    tipo: 'faltando'
                });
                continue;
            }
            
            // Validar cada exame
            Object.keys(MAPEAMENTO_EXAMES).forEach(codigoExame => {
                const config = MAPEAMENTO_EXAMES[codigoExame];
                const colunaGerada = `NR_EXAME_${codigoExame}`;
                const valorGerado = linhaGerada[colunaGerada];
                
                // Buscar valor original
                let valorOriginal = null;
                for (const colOriginal of config.colunas) {
                    if (pacienteOriginal[colOriginal] !== undefined && 
                        pacienteOriginal[colOriginal] !== null && 
                        pacienteOriginal[colOriginal] !== '') {
                        valorOriginal = pacienteOriginal[colOriginal];
                        break;
                    }
                }
                
                // Comparar
                const iguais = valoresIguais(valorOriginal, valorGerado, config.tipo);
                
                const vOrigNorm = normalizarValor(valorOriginal, config.tipo);
                const vGerNorm = normalizarValor(valorGerado, config.tipo);
                
                // Determinar status
                let status = 'OK';
                let tipo = 'ok';
                
                if (!iguais) {
                    if (vOrigNorm !== null && vGerNorm === null) {
                        status = 'FALTANDO';
                        tipo = 'faltando';
                        totalFaltando++;
                    } else if (vOrigNorm === null && vGerNorm !== null) {
                        status = 'EXTRA';
                        tipo = 'divergente';
                        totalDivergente++;
                    } else {
                        status = 'DIVERGENTE';
                        tipo = 'divergente';
                        totalDivergente++;
                    }
                } else if (vOrigNorm !== null) {
                    totalOk++;
                }
                
                // Só adicionar se houver valor ou divergência
                if (vOrigNorm !== null || vGerNorm !== null || status !== 'OK') {
                    validacoes.push({
                        atendimento,
                        paciente: nomePaciente,
                        exame: config.nome,
                        codigoExame,
                        valorOriginal: vOrigNorm,
                        valorGerado: vGerNorm,
                        valorOriginalDisplay: valorOriginal || '',
                        valorGeradoDisplay: valorGerado || '',
                        status,
                        tipo
                    });
                }
            });
        }
        
        atualizarProgresso(100);
        
        resultadoValidacao = {
            validacoes,
            totalOk,
            totalDivergente,
            totalFaltando,
            totalValidado: validacoes.length
        };
        
        mostrarResultado();
        
    } catch (error) {
        alert('Erro ao validar: ' + error.message);
        btn.disabled = false;
        btn.textContent = '🔍 Iniciar Validação';
    }
}

function atualizarProgresso(percentual) {
    const fill = document.getElementById('progressFill');
    fill.style.width = percentual + '%';
    fill.textContent = percentual + '%';
}

function mostrarResultado() {
    document.getElementById('resultSection').classList.add('show');
    
    // Dashboard
    const dashboard = document.getElementById('dashboard');
    dashboard.innerHTML = `
        <div class="card success">
            <div class="card-icon">✅</div>
            <div class="card-number">${resultadoValidacao.totalOk}</div>
            <div class="card-title">Valores OK</div>
        </div>
        
        <div class="card danger">
            <div class="card-icon">❌</div>
            <div class="card-number">${resultadoValidacao.totalDivergente}</div>
            <div class="card-title">Divergentes</div>
        </div>
        
        <div class="card warning">
            <div class="card-icon">⚠️</div>
            <div class="card-number">${resultadoValidacao.totalFaltando}</div>
            <div class="card-title">Faltando</div>
        </div>
        
        <div class="card">
            <div class="card-icon">📊</div>
            <div class="card-number">${resultadoValidacao.totalValidado}</div>
            <div class="card-title">Total Validado</div>
        </div>
    `;
    
    // Alerta
    const alertSection = document.getElementById('alertSection');
    if (resultadoValidacao.totalDivergente > 0 || resultadoValidacao.totalFaltando > 0) {
        alertSection.innerHTML = `
            <div class="alert alert-danger">
                <span>⚠️</span>
                <div>
                    <strong>Atenção!</strong> Foram encontradas ${resultadoValidacao.totalDivergente} divergências 
                    e ${resultadoValidacao.totalFaltando} valores faltando. 
                    Revise os detalhes na tabela abaixo.
                </div>
            </div>
        `;
    } else {
        alertSection.innerHTML = `
            <div class="alert alert-success">
                <span>✅</span>
                <div>
                    <strong>Perfeito!</strong> Todos os valores conferem com os dados originais do laboratório. 
                    O arquivo gerado está 100% correto!
                </div>
            </div>
        `;
    }
    
    // Tabela
    preencherTabela();
    
    // Configurar filtros
    configurarFiltros();
    
    // Botão de download
    document.getElementById('btnDownloadRelatorio').style.display = 'inline-block';
    document.getElementById('btnDownloadRelatorio').addEventListener('click', gerarRelatorio);
}

function preencherTabela() {
    const tbody = document.getElementById('corpoTabela');
    tbody.innerHTML = '';
    
    resultadoValidacao.validacoes.forEach(val => {
        const tr = document.createElement('tr');
        tr.className = `filtro-${val.tipo}`;
        
        let statusHtml = '';
        let valorOriginalDisplay = val.valorOriginalDisplay || '-';
        let valorGeradoDisplay = val.valorGeradoDisplay || '-';
        
        if (val.status === 'OK') {
            statusHtml = '<span class="status-ok">✅ OK</span>';
        } else if (val.status === 'DIVERGENTE') {
            statusHtml = '<span class="status-error">❌ DIVERGE</span>';
        } else if (val.status === 'FALTANDO') {
            statusHtml = '<span class="status-missing">⚠️ FALTA</span>';
        } else if (val.status === 'EXTRA') {
            statusHtml = '<span class="status-error">❌ EXTRA</span>';
        } else if (val.status === 'PACIENTE_NAO_ENCONTRADO') {
            statusHtml = '<span class="status-error">❌ PAC. NÃO ENCONTRADO</span>';
            valorOriginalDisplay = '-';
            valorGeradoDisplay = '-';
        }
        
        tr.innerHTML = `
            <td>${val.atendimento}</td>
            <td>${val.paciente}</td>
            <td>${val.exame || '-'}</td>
            <td style="font-family: monospace;">${valorOriginalDisplay}</td>
            <td style="font-family: monospace;">${valorGeradoDisplay}</td>
            <td>${statusHtml}</td>
        `;
        
        tbody.appendChild(tr);
    });
}

function configurarFiltros() {
    const filterOk = document.getElementById('filterOk');
    const filterDivergente = document.getElementById('filterDivergente');
    const filterFaltando = document.getElementById('filterFaltando');
    
    function aplicarFiltros() {
        const mostrarOk = filterOk.checked;
        const mostrarDivergente = filterDivergente.checked;
        const mostrarFaltando = filterFaltando.checked;
        
        document.querySelectorAll('#corpoTabela tr').forEach(tr => {
            const isOk = tr.classList.contains('filtro-ok');
            const isDivergente = tr.classList.contains('filtro-divergente');
            const isFaltando = tr.classList.contains('filtro-faltando');
            
            let mostrar = false;
            if (isOk && mostrarOk) mostrar = true;
            if (isDivergente && mostrarDivergente) mostrar = true;
            if (isFaltando && mostrarFaltando) mostrar = true;
            
            tr.style.display = mostrar ? '' : 'none';
        });
    }
    
    filterOk.addEventListener('change', aplicarFiltros);
    filterDivergente.addEventListener('change', aplicarFiltros);
    filterFaltando.addEventListener('change', aplicarFiltros);
}

function gerarRelatorio() {
    const hoje = new Date();
    const dataHora = hoje.toLocaleString('pt-BR');
    
    let relatorio = `═══════════════════════════════════════════════════════════════
RELATÓRIO DE VALIDAÇÃO - SISTEMA TASY
═══════════════════════════════════════════════════════════════

Data/Hora: ${dataHora}

═══════════════════════════════════════════════════════════════
RESUMO DA VALIDAÇÃO
═══════════════════════════════════════════════════════════════

✅ Valores OK (conferem): ${resultadoValidacao.totalOk}
❌ Valores Divergentes: ${resultadoValidacao.totalDivergente}
⚠️ Valores Faltando: ${resultadoValidacao.totalFaltando}
📊 Total Validado: ${resultadoValidacao.totalValidado}

`;

    if (resultadoValidacao.totalDivergente === 0 && resultadoValidacao.totalFaltando === 0) {
        relatorio += `═══════════════════════════════════════════════════════════════
✅ VALIDAÇÃO 100% APROVADA
═══════════════════════════════════════════════════════════════

Todos os valores do arquivo gerado conferem EXATAMENTE com os 
dados originais do laboratório. O arquivo está pronto para 
importação no Tasy com segurança total!

`;
    } else {
        relatorio += `═══════════════════════════════════════════════════════════════
⚠️ ATENÇÃO - DIVERGÊNCIAS ENCONTRADAS
═══════════════════════════════════════════════════════════════

`;

        // Listar divergências
        const divergencias = resultadoValidacao.validacoes.filter(v => 
            v.status === 'DIVERGENTE' || v.status === 'FALTANDO' || v.status === 'EXTRA'
        );
        
        divergencias.forEach((div, i) => {
            relatorio += `\n${i + 1}. ATEND. ${div.atendimento} - ${div.paciente}`;
            relatorio += `\n   Exame: ${div.exame}`;
            relatorio += `\n   Original: ${div.valorOriginalDisplay || '(vazio)'}`;
            relatorio += `\n   Gerado: ${div.valorGeradoDisplay || '(vazio)'}`;
            relatorio += `\n   Status: ${div.status}`;
            relatorio += `\n`;
        });
    }

    relatorio += `\n═══════════════════════════════════════════════════════════════
DETALHAMENTO COMPLETO
═══════════════════════════════════════════════════════════════\n`;

    resultadoValidacao.validacoes.forEach((val, i) => {
        relatorio += `\n${i + 1}. ${val.status === 'OK' ? '✅' : val.status === 'DIVERGENTE' || val.status === 'EXTRA' ? '❌' : '⚠️'}`;
        relatorio += ` Atend. ${val.atendimento} - ${val.paciente}`;
        if (val.exame) {
            relatorio += `\n   Exame: ${val.exame}`;
            relatorio += `\n   Original: ${val.valorOriginalDisplay || '-'}`;
            relatorio += `\n   Gerado: ${val.valorGeradoDisplay || '-'}`;
        }
    });

    relatorio += `\n\n═══════════════════════════════════════════════════════════════
FIM DO RELATÓRIO
═══════════════════════════════════════════════════════════════
`;

    // Download
    const blob = new Blob([relatorio], { type: 'text/plain;charset=utf-8' });
    const dataFormatada = hoje.toISOString().split('T')[0].replace(/-/g, '');
    const nomeArquivo = `Relatorio_Validacao_${dataFormatada}.txt`;
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nomeArquivo;
    a.click();
}
