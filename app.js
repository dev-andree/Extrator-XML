const fs = require('fs');
const path = require('path');
const xml2js = require('xml2js');
const XLSX = require('xlsx');

// Função para ler e extrair dados do XML
async function extractDataFromXML(xmlPath) {
    try {
        // Lê o arquivo XML
        const xmlContent = fs.readFileSync(xmlPath, 'utf-8');

        // Parse do XML para JSON
        const parser = new xml2js.Parser({ explicitArray: false, preserveChildrenOrder: true });
        const result = await parser.parseStringPromise(xmlContent);

        // Acessando dados da NFe (ajuste a estrutura conforme o XML que você possui)
        const notaFiscal = result['nfeProc'] || {};
        const infNFe = notaFiscal.NFe?.infNFe || {};

        // Extração dos dados principais
        const nNF = infNFe.ide?.nNF || ''; // Número da Nota
        const dEmi = infNFe.ide?.dhEmi || ''; // Data de Emissão
        const xNome = infNFe.emit?.xNome || ''; // Nome do Emitente

        // Verifica se o campo 'det' existe e pode ser tratado como array
        const produtos = Array.isArray(infNFe.det) ? infNFe.det : (infNFe.det ? [infNFe.det] : []);

        // Extração dos dados dos produtos
        const extractedProducts = produtos.map(item => {
            const isService = item.imposto?.ISSQN ? true : false; // Verifica se o item é um serviço
            return {
                nNF, // Número da Nota
                dEmi, // Data de Emissão
                xNome, // Nome do Emitente
                tipo: isService ? 'Serviço' : 'Produto', // Define "Produto" ou "Serviço"
                xProd: item.prod?.xProd || '', // Nome do Produto
                qCom: item.prod?.qCom || '', // Quantidade Comercial
                uCom: item.prod?.uCom || '', // Unidade Comercial
                vUnCom: item.prod?.vUnCom || '', // Valor Unitário
                vProd: item.prod?.vProd || '', // Valor Total do Produto
            };
        });

        return extractedProducts;
    } catch (err) {
        console.error("Erro ao extrair dados do XML:", err.message);
        return null;
    }
}

// Função para criar ou atualizar a planilha
function createOrUpdateSpreadsheet(data, outputFile) {
    try {
        // Criação ou leitura da planilha existente
        const planilhaDir = path.join(__dirname, 'PLANILHA');
        if (!fs.existsSync(planilhaDir)) {
            fs.mkdirSync(planilhaDir);
        }

        const outputPath = path.join(planilhaDir, outputFile);

        let workbook;
        let worksheet;

        if (fs.existsSync(outputPath)) {
            // Se a planilha já existir, lê a planilha existente
            workbook = XLSX.readFile(outputPath);
            worksheet = workbook.Sheets[workbook.SheetNames[0]]; // Seleciona a primeira aba
        } else {
            // Caso contrário, cria um novo workbook e worksheet
            workbook = XLSX.utils.book_new();
            worksheet = XLSX.utils.aoa_to_sheet([
                ['Número da Nota', 'Emitente', 'Data de Emissão', 'Tipo (Produto/Serviço)', 'Nome do Produto/Serviço', 'Classificação', 'Quantidade', 'Unidade', 'Valor Unitário', 'Valor Total']
            ]);
            XLSX.utils.book_append_sheet(workbook, worksheet, 'NFe Data');
        }

        // Função para classificar como Consumível ou Patrimonial
        const classifyItem = (nomeProduto) => {
            const patrimonialKeywords = ['periférico', 'móvel', 'equipamento', 'hardware']; // Exemplos de palavras que indicam um item patrimonial
            const isPatrimonial = patrimonialKeywords.some(keyword => nomeProduto.toLowerCase().includes(keyword));
            return isPatrimonial ? 'Patrimonial' : 'Consumível';
        };

        // Adiciona os produtos à planilha
        data.forEach(item => {
            const classificacao = classifyItem(item.xProd); // Classifica o item como Consumível ou Patrimonial
            const newRow = [
                item.nNF,
                item.xNome,
                item.dEmi,
                item.tipo, // Adiciona a informação de Produto ou Serviço
                item.xProd,
                classificacao, // Nova coluna de classificação
                item.qCom,
                item.uCom,
                item.vUnCom,
                item.vProd
            ];
            XLSX.utils.sheet_add_aoa(worksheet, [newRow], { origin: -1 }); // Adiciona no final da planilha
        });

        // Salva a planilha no mesmo caminho
        XLSX.writeFile(workbook, outputPath);
        console.log("Planilha atualizada com sucesso em:", outputPath);
    } catch (err) {
        console.error("Erro ao criar ou atualizar a planilha:", err.message);
    }
}

// Função principal (entry point)
(async () => {
    const xmlDir = path.join(__dirname, 'XML'); // Caminho da pasta XML
    const outputFile = 'dados_nfe.xlsx'; // Nome do arquivo de saída

    // Lê todos os arquivos da pasta XML
    fs.readdir(xmlDir, async (err, files) => {
        if (err) {
            console.error("Erro ao ler a pasta XML:", err.message);
            return;
        }

        // Filtra arquivos XML
        const xmlFiles = files.filter(file => file.endsWith('.xml'));

        if (xmlFiles.length === 0) {
            console.log("Nenhum arquivo XML encontrado na pasta XML.");
            return;
        }

        // Processa cada arquivo XML encontrado
        for (const xmlFile of xmlFiles) {
            const xmlPath = path.join(xmlDir, xmlFile);
            console.log(`Iniciando extração de dados do arquivo XML: ${xmlFile}...`);

            const extractedData = await extractDataFromXML(xmlPath);

            if (extractedData && extractedData.length > 0) {
                console.log("Dados extraídos com sucesso:", extractedData);
                console.log("Atualizando planilha...");
                createOrUpdateSpreadsheet(extractedData, outputFile);
            } else {
                console.log(`Não foi possível extrair os dados ou nenhum produto foi encontrado no arquivo ${xmlFile}.`);
            }
        }
    });
})();
