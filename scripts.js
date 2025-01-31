// Função para alternar exibição do status baseado no select
function toggleDetails(selectElement) {
    let statusDetails = selectElement.closest('tr').querySelector('.status-details');
    statusDetails.style.display = selectElement.value === "Sim" ? "block" : "none";
}

// Função para gerar PDF e enviar os dados para o Excel
function gerarPDF() {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();

    html2canvas(document.getElementById("ratForm")).then(canvas => {
        const imgData = canvas.toDataURL("image/png");
        const imgWidth = 190;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;

        pdf.addImage(imgData, "PNG", 10, 10, imgWidth, imgHeight);
        pdf.save("Relatorio_RAT.pdf");

        // Chamar a função para enviar os dados ao Excel
        enviarParaExcel();
    });
}

// Função para enviar os dados para o Excel via Sheet Monkey
function enviarParaExcel() {
    let tabela = document.querySelector("table");
    let linhas = tabela.querySelectorAll("tr");

    let dados = {}; // Objeto para armazenar os dados

    linhas.forEach((linha, index) => {
        if (index === 0) return; // Ignorar cabeçalho

        let colunas = linha.querySelectorAll("td, th");

        // Verifica se há colunas suficientes
        if (colunas.length < 4) return;

        let servico = colunas[1].textContent.trim();
        let selectElement = colunas[2].querySelector("select");
        let assinaturaElement = colunas[3].querySelector("input");

        // Verificar se os elementos existem antes de acessar
        if (!selectElement) {
            console.warn(`Aviso: Nenhum <select> encontrado na linha ${index}`);
            return;
        }
        
        let status = selectElement.value;
        let assinatura = assinaturaElement ? assinaturaElement.value : "";

        // Garantir que os campos não estão vazios antes de enviar
        if (servico && status) {
            dados[`Serviço ${index}`] = servico;
            dados[`Status ${index}`] = status;
            dados[`Rubrica ${index}`] = assinatura;
        }
    });

    // Se não houver dados, não envia
    if (Object.keys(dados).length === 0) {
        alert("Nenhum dado válido para enviar.");
        return;
    }

    // Exibir os dados no console para verificar antes do envio
    console.log("Enviando dados:", JSON.stringify(dados));

    // Enviar para a API do Sheet Monkey
    fetch("https://api.sheetmonkey.io/form/fxS1JNaQbCiAZ1yJmCmhgp", {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify(dados)
    })
    .then(response => response.text())  // Ler a resposta como texto
    .then(data => {
        console.log("Resposta da API:", data);
        alert("Dados enviados com sucesso para o Excel!");
    })
    .catch(error => {
        console.error("Erro ao enviar:", error);
        alert("Erro ao enviar os dados. Verifique o console.");
    });
}
