const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');
const { createObjectCsvWriter } = require('csv-writer');
const XLSX = require('xlsx');
const csv = require('csv-parser');

// Ścieżka do folderu z plikami HTML
const folderPath = path.join('C:', 'Users', 'GPD', 'Desktop', 'pages');

// Funkcja do ekstrakcji danych z pliku HTML
function extractDataFromHtml(fileContent) {
  const $ = cheerio.load(fileContent);
  const pharmacies = [];

  // Sprawdź, czy akordeony są obecne
  const accordionTabs = $('cez-accordion-tab');
  console.log(`Found ${accordionTabs.length} accordion tabs`);

  if (accordionTabs.length === 0) {
    console.log('No accordion tabs found. Please check the selector or the page structure.');
    return pharmacies;
  }

  accordionTabs.each((index, tab) => {
    const $tab = $(tab);
    const id = $tab.find('cez-simple-cell[label="ID Apteki"] .simple-cell-value').text().trim();
    const name = $tab.find('cez-simple-cell[label="Nazwa"] .simple-cell-value').text().trim();
    const address = $tab.find('cez-simple-cell[label="Adres apteki"] .simple-cell-value').text().trim();
    const status = $tab.find('cez-simple-cell[label="Status"] .simple-cell-value').text().trim();
    const type = $tab.find('cez-simple-cell[label="Rodzaj apteki"] .simple-cell-value').text().trim();
    const owner = $tab.find('cez-simple-cell[label="Właściciel"] .simple-cell-value').text().trim();
    const phone = $tab.find('cez-simple-cell[label="Telefon"] .simple-cell-value').text().trim();
    const email = $tab.find('cez-simple-cell[label="Email"] .simple-cell-value').text().trim();

    pharmacies.push({ id, name, address, status, type, owner, phone, email });
  });

  return pharmacies;
}

// Funkcja do zapisania danych do pliku CSV
async function saveToCSV(data) {
  const csvWriter = createObjectCsvWriter({
    path: 'pharmacies.csv',
    header: [
      { id: 'id', title: 'ID' },
      { id: 'name', title: 'Nazwa' },
      { id: 'address', title: 'Adres' },
      { id: 'status', title: 'Status' },
      { id: 'type', title: 'Rodzaj' },
      { id: 'owner', title: 'Właściciel' },
      { id: 'phone', title: 'Telefon' },
      { id: 'email', title: 'E-mail' },
    ]
  });

  await csvWriter.writeRecords(data);
  console.log('Dane zostały zapisane do pliku pharmacies.csv');
}

// Funkcja do konwersji CSV na XLSX
function convertCsvToXlsx(csvFilePath, xlsxFilePath) {
  const rows = [];

  // Odczytaj CSV
  fs.createReadStream(csvFilePath)
    .pipe(csv())
    .on('data', (row) => rows.push(row))
    .on('end', () => {
      // Utwórz arkusz Excel
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();

      // Dodaj arkusz do książki
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

      // Zapisz książkę do pliku XLSX
      XLSX.writeFile(wb, xlsxFilePath);
      console.log(`Plik XLSX zapisany jako ${xlsxFilePath}`);
    });
}

// Główna funkcja
async function main() {
  let allData = [];

  for (let i = 1; i <= 142; i++) {
    const filePath = path.join(folderPath, `page${i}.html`);
    
    if (!fs.existsSync(filePath)) {
      console.log(`Plik ${filePath} nie istnieje. Pomijanie.`);
      continue;
    }

    const fileContent = fs.readFileSync(filePath, 'utf8');
    const data = extractDataFromHtml(fileContent);
    allData = allData.concat(data);
  }

  if (allData.length > 0) {
    await saveToCSV(allData);
    convertCsvToXlsx('pharmacies.csv', 'pharmacies.xlsx');
  } else {
    console.log('No data extracted.');
  }
}

main().catch(err => console.error(err));