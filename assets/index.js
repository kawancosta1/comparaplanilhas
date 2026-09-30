let inputPlanilha = document.querySelector("#planilha");
let statusPlanilha = document.querySelector("#statusPlanilha");

let botao = document.querySelector("#compara");
let inputNumero = document.querySelector("#numero");
let resultado = document.querySelector("#resultado");

let numerosDaPlanilha = [];


/*
    Transforma:

    11434/08  em  1143408
    11434-08  em  1143408
    11.434/08 em  1143408
*/
function normalizarNumero(valor) {
    return String(valor ?? "")
        .trim()
        .replace(/\D/g, "");
}


inputPlanilha.addEventListener("change", function (evento) {
    // esse target é onde aconteceu o  evento e o evento é a planplan e o files[0] é o arquivo que mandamos, pois o files é um array, então se mandassemosmais de um arquivo daria para acaessar todos
    let arquivo = evento.target.files[0];

    if (!arquivo) {
        statusPlanilha.textContent =
            "Nenhuma planilha selecionada.";

        return;
    }

    let leitor = new FileReader();

    leitor.onload = function (eventoLeitura) {
        try {
            let dadosArquivo = eventoLeitura.target.result;

            let pastaDeTrabalho = XLSX.read(dadosArquivo, {
                type: "array"
            });

            let nomePrimeiraAba =
                pastaDeTrabalho.SheetNames[0];

            let primeiraAba =
                pastaDeTrabalho.Sheets[nomePrimeiraAba];

            let linhas = XLSX.utils.sheet_to_json(
                primeiraAba,
                {
                    header: 1,
                    defval: "",
                    raw: false
                }
            );

            console.log("Conteúdo da planilha:");
            console.table(linhas);

            /*
                Percorre todas as linhas.
                Depois percorre todas as células de cada linha.
            */
            numerosDaPlanilha = [];

            linhas.forEach(function (linha) {
                linha.forEach(function (celula) {
                    let numeroNormalizado =
                        normalizarNumero(celula);

                    if (numeroNormalizado !== "") {
                        numerosDaPlanilha.push(
                            numeroNormalizado
                        );
                    }
                });
            });

            console.log(
                "Números encontrados:",
                numerosDaPlanilha
            );

            statusPlanilha.textContent =
                "Planilha carregada: " +
                numerosDaPlanilha.length +
                " valores numéricos encontrados.";

            resultado.textContent = "";
        } catch (erro) {
            console.error("Erro ao ler planilha:", erro);

            statusPlanilha.textContent =
                "Não foi possível ler a planilha.";
        }
    };

    leitor.onerror = function () {
        statusPlanilha.textContent =
            "Ocorreu um erro ao abrir o arquivo.";
    };

    leitor.readAsArrayBuffer(arquivo);
});


botao.addEventListener("click", function () {
    if (numerosDaPlanilha.length === 0) {
        resultado.textContent =
            "Primeiro escolha uma planilha.";

        resultado.style.color = "orange";

        return;
    }

    let numeroDigitado =
        normalizarNumero(inputNumero.value);

    if (numeroDigitado === "") {
        resultado.textContent =
            "Digite um número para pesquisar.";

        resultado.style.color = "orange";

        return;
    }

    let foiEncontrado =
        numerosDaPlanilha.includes(numeroDigitado);

    if (foiEncontrado) {
        resultado.textContent =
            inputNumero.value + " foi encontrado.";

        resultado.style.color = "green";
    } else {
        resultado.textContent =
            inputNumero.value + " não foi encontrado.";

        resultado.style.color = "red";
    }
});


inputNumero.addEventListener("keydown", function (evento) {
    if (evento.key === "Enter") {
        botao.click();
    }
});