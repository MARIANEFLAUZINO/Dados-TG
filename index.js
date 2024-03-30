const Excel = require('exceljs');
const path = require('path');

// Constantes contendo os dados para gerar os registros falsos
const FAKE_DATA = {
    YEARS: [2019, 2020, 2021],
    MONTHS: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
    SELLERS: ['Lucimar Sasso', 'Mariane Flauzino', 'Mateus Alves'],
    REGIONS: ['Sul', 'Suldeste', 'Centro-Oeste', 'Norte', 'Nordeste'],
    PRODUCTS: [{ item: 'Pacote Office', price: 450, acc: 35 }, { item: 'Power BI', price: 500, acc: 40 }, { item: 'Microsoft Azure', price: 600, acc: 50 }],
    PAYMENT_METHODS: ['Cartão de Débito', 'Cartão de Crédito', 'Boleto']
}

// Função para obter um elemento aleatório de um array
function getRandomElement(array) {
    return array[Math.floor(Math.random() * array.length)];
}

// Função para gerar uma data aleatória dentro de um mês e ano específicos
function getRandomDate(month, year) {
    const lastDayOfMonth = new Date(year, month, 0).getDate();
    const randomDay = Math.floor(Math.random() * lastDayOfMonth) + 1;
    return `${randomDay < 10 ? '0' : ''}${randomDay}/${month < 10 ? '0' : ''}${month}/${year}`;
}

// Calcula o preço de um produto com base no ano de venda e no aumento acumulado.
function calcProductPrice(product, year) {
    const { price, acc } = product;
    // Calcula quantos anos se passaram desde 2019
    const yearsSinceBase = year - 2019;

    // Calcula o preço com base no aumento acumulado
    const adjustedPrice = price + (acc * yearsSinceBase);

    // Retorna o preço ajustado e o nome do item
    return { price: adjustedPrice, item: product.item };
}

// Array para armazenar os registros falsos gerados
const fakeData = [];
// Limite de registros a serem gerados
const limit = 80 * 1000;

// Loop para gerar registros falsos
for (let index = 0; index < limit; index++) {
    // Gerar um registro falso
    const row = {
        year: getRandomElement(FAKE_DATA.YEARS),
        month: getRandomElement(FAKE_DATA.MONTHS),
        seller: getRandomElement(FAKE_DATA.SELLERS),
        region: getRandomElement(FAKE_DATA.REGIONS),
        paymentMethod: getRandomElement(FAKE_DATA.PAYMENT_METHODS)
    };
    // Calcular o preço do produto para o ano selecionado
    const product = calcProductPrice(getRandomElement(FAKE_DATA.PRODUCTS), row.year);
    // Adicionar informações do produto ao registro falso
    row.product = product.item;
    row.price = product.price;
    // Gerar uma data aleatória para o registro
    row.date = getRandomDate(FAKE_DATA.MONTHS.indexOf(row.month) + 1, row.year);
    // Adicionar o registro falso ao array de dados falsos
    fakeData.push(row);
}

// Criar um novo arquivo Excel
const workbook = new Excel.Workbook();
// Adicionar uma nova planilha ao arquivo Excel
const worksheet = workbook.addWorksheet('Dados');

// Definir os cabeçalhos das colunas na planilha
const columns = [
    { header: 'Vendedor', key: 'seller', width: 20 },
    { header: 'Produto', key: 'product', width: 10 },
    { header: 'Preço', key: 'price', width: 10 },
    { header: 'Forma de Pagamento', key: 'paymentMethod', width: 20 },
    { header: 'Data', key: 'date', width: 20 },
    { header: 'Região', key: 'region', width: 20 },
    { header: 'Mês', key: 'month', width: 20 },
    { header: 'Ano', key: 'year', width: 20 }
];
// Definir as colunas na planilha
worksheet.columns = columns;

// Preencher a planilha com os dados falsos gerados
fakeData.forEach(row => {
    worksheet.addRow(row);
});

// Salvar o arquivo Excel no disco
const filePath = path.join(__dirname, 'dados.xlsx');

workbook.xlsx.writeFile(filePath)
    .then(() => {
        // Exibir mensagem de sucesso após a exportação do arquivo Excel
        console.log(`Arquivo Excel exportado com sucesso para: ${filePath}`);
    })
    .catch(error => {
        // Exibir mensagem de erro, se houver, durante a exportação do arquivo Excel
        console.error('Erro ao exportar para Excel:', error);
    });