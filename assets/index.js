
let inputPlanilha1 = document.querySelector("#planilha");
let inputPlanilha2 = document.querySelector("#planilha2");
let statusPlanilha = document.querySelector("#statusPlanilha");
let botao = document.querySelector("#compara");
let resultado = document.querySelector("#resultado");
let resultado2 = document.querySelector("#resultado2");

let numerosPlanilha1 = [];
let numerosPlanilha2 = [];

function normalizarNumero(valor) {
    return String(valor ?? "")
        .trim()
        .replace(/\D/g, "");
}

function saoEquivalentes(numA, numB) {
    if (numA === numB) return true;

    // Remove os zeros à esquerda para comparar apenas os dígitos significativos
    let limpoA = numA.replace(/^0+/, "");
    let limpoB = numB.replace(/^0+/, "");

    // Trava de segurança para não comparar números curtos
    if (limpoA.length < 6 || limpoB.length < 6) return false;

    return limpoA.includes(limpoB) || limpoB.includes(limpoA);
}

// --- 4. Função genérica para ler qualquer arquivo Excel ---
// --- 4. Função genérica para ler arquivos Excel e CSV ---
function lerPlanilha(arquivo, callbackSucesso, callbackErro) {
    let leitor = new FileReader();

    leitor.onload = function (eventoLeitura) {
        try {
            let dadosArquivo = eventoLeitura.target.result;
            let pasta = XLSX.read(dadosArquivo, { type: "array" });
            let primeiraAba = pasta.Sheets[pasta.SheetNames[0]];

            let linhas = XLSX.utils.sheet_to_json(primeiraAba, {
                header: 1,
                defval: "",
                raw: false
            });

            if (linhas.length === 0) {
                callbackSucesso([]);
                return;
            }

            // 1. Procura em qual linha está o cabeçalho e em qual coluna está o "processo"
            let indiceLinhaCabecalho = -1;
            let indiceColunaProcesso = -1;

            for (let i = 0; i < Math.min(10, linhas.length); i++) {
                let cabecalho = linhas[i].map(c => String(c).trim().toLowerCase());
                let col = cabecalho.findIndex(c => c.includes("processo"));
                if (col !== -1) {
                    indiceLinhaCabecalho = i;
                    indiceColunaProcesso = col;
                    break;
                }
            }

            let numerosExtraidos = [];

            if (indiceColunaProcesso !== -1) {
                // Se encontrou a coluna de processo (seja "numeroProcesso" ou "Processo"), lê apenas ela!
                for (let i = indiceLinhaCabecalho + 1; i < linhas.length; i++) {
                    let normalizado = normalizarNumero(linhas[i][indiceColunaProcesso]);
                    if (normalizado.length >= 6) {
                        numerosExtraidos.push(normalizado);
                    }
                }
            } else {
                // Fallback para folhas sem cabeçalho padronizado
                for (let i = 0; i < linhas.length; i++) {
                    linhas[i].forEach(function (celula) {
                        let normalizado = normalizarNumero(celula);
                        if (normalizado.length >= 6) {
                            numerosExtraidos.push(normalizado);
                        }
                    });
                }
            }

            callbackSucesso(numerosExtraidos);
        } catch (erro) {
            callbackErro(erro);
        }
    };

    leitor.onerror = callbackErro;
    leitor.readAsArrayBuffer(arquivo);
}

// --- 5. Evento: Carregar Planilha 1 ---
inputPlanilha1.addEventListener("change", function (evento) {
    let arquivo = evento.target.files[0];
    if (!arquivo) return;

    lerPlanilha(
        arquivo,
        function (numeros) {
            //esse numeros é o nnumerosExtraidos
            numerosPlanilha1 = numeros;
            statusPlanilha.textContent = `Planilha 1 carregada: ${numeros.length} números encontrados.`;
            resultado.textContent = "";
        },
        function () {
            statusPlanilha.textContent = "Erro ao ler a Planilha 1.";
        }
    );
});

// --- 6. Evento: Carregar Planilha 2 ---
inputPlanilha2.addEventListener("change", function (evento) {
    let arquivo = evento.target.files[0];
    if (!arquivo) return;

    lerPlanilha(
        arquivo,
        function (numeros) {
            numerosPlanilha2 = numeros;
            statusPlanilha.textContent = `Planilha 2 carregada: ${numeros.length} números encontrados.`;
            resultado.textContent = "";
        },
        function () {
            statusPlanilha.textContent = "Erro ao ler a Planilha 2.";
        }
    );
});

// --- 7. Evento: Comparar as duas planilhas ---
botao.addEventListener("click", function () {
    // Validação: ambas precisam estar carregadas
    if (numerosPlanilha1.length === 0 || numerosPlanilha2.length === 0) {
        resultado.textContent = "Selecione as duas planilhas antes de comparar.";
        resultado.style.color = "orange";
        return;
    }

   let comuns = []
   let rejeitados = []
    
   //cada valor de numerosPlanilha2 vai ser o num2 em cada iteração
    numerosPlanilha2.forEach(function (num2){

        let encontrou = numerosPlanilha1.some(function(num1) {
            return saoEquivalentes(num1,num2);
        })

      if (encontrou) {
        comuns.push(num2);
    } else {
        rejeitados.push(num2);
    }

    })

//o spread espalha os itens para que o array se torne limpo e sem duplicados
    let comunsSemDuplicadas = [...new Set(comuns)];
    let rejeitadosSemDuplicados = [...new Set(rejeitados)];
   // 1. Exibe os comuns
if (comunsSemDuplicadas.length > 0) {
    resultado.textContent = `Foram encontrados ${comunsSemDuplicadas.length} números em comum: ${comunsSemDuplicadas.join(", ")}`;
    resultado.style.color = "green";
} else {
    resultado.textContent = "Nenhum número em comum foi encontrado entre as duas planilhas.";
    resultado.style.color = "red";
}

// 2. Exibe os rejeitados (apenas processos reais >= 7)
let rejeitadosValidos = rejeitadosSemDuplicados.filter(num => num.length >= 7);

if (rejeitadosValidos.length > 0) {
    resultado2.textContent = `Processos não encontrados na Planilha 1 (${rejeitadosValidos.length}): ${rejeitadosValidos.join(", ")}`;
    resultado2.style.color = "#c0392b";
} else {
    resultado2.textContent = "Todos os processos da Planilha 2 foram encontrados na Planilha 1!";
    resultado2.style.color = "green";
}
});



