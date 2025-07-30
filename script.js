// Global state
let vendorMapping = {};    // vendorName -> { mpk, group }
let lastNumberMap = {};    // key (mpk|group) -> highest number found (integer)
let dataset = [];          // collected invoice info for Excel export

// Storage for the last generated label and modified PDF bytes
let lastLabel = '';
let modifiedPdfBytes = null;

// Configure PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

const guidelinesInput = document.getElementById('guidelinesFile');
const invoiceInput = document.getElementById('invoiceFile');
const processBtn = document.getElementById('processBtn');
const outputDiv = document.getElementById('output');
const errorDiv = document.getElementById('error');
const downloadSection = document.getElementById('downloadSection');
const downloadBtn = document.getElementById('downloadBtn');
const downloadModifiedBtn = document.getElementById('downloadModifiedBtn');

// Read guidelines Excel and prepare mappings
guidelinesInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  readGuidelines(file);
});

// Process invoice PDF when button clicked
processBtn.addEventListener('click', () => {
  const pdfFile = invoiceInput.files[0];
  if (!pdfFile) {
    showError('Proszę wybrać plik PDF z fakturą.');
    return;
  }
  // Ensure guidelines loaded
  if (Object.keys(vendorMapping).length === 0) {
    const confirmContinue = confirm(
      'Nie wczytano wytycznych. Czy chcesz kontynuować bez nich? Wartości MPK i numerów zostaną ustawione domyślnie.'
    );
    if (!confirmContinue) return;
  }
  processInvoice(pdfFile);
});

// Download Excel on click
downloadBtn.addEventListener('click', () => {
  if (dataset.length === 0) {
    showError('Brak danych do zapisania. Przetwórz co najmniej jedną fakturę.');
    return;
  }
  downloadExcel();
});

// Download modified PDF on click
downloadModifiedBtn.addEventListener('click', () => {
  if (!modifiedPdfBytes) {
    showError('Brak zmodyfikowanej faktury do pobrania. Przetwórz fakturę ponownie.');
    return;
  }
  const blob = new Blob([modifiedPdfBytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  // Nazwa pliku zawiera MPK i numer
  a.download = `faktura_z_opisem_${lastLabel.replace(/\s+/g, '_')}.pdf`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

function readGuidelines(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    try {
      const workbook = XLSX.read(data, { type: 'array' });
      // Attempt to use sheet "Koszty - przyklady" as primary mapping
      const sheetName = workbook.SheetNames.find((name) => name.toLowerCase().includes('przyklady'));
      if (sheetName) {
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        json.forEach((row) => {
          const vendor = (row['Nazwa kontrahenta'] || '').toString().trim();
          const label = (row['Etykieta'] || '').toString().trim();
          if (!vendor || !label) return;
          const parts = label.split(';');
          let group = '';
          let mpk = '';
          let number = '';
          parts.forEach((part) => {
            part = part.trim();
            if (part.toUpperCase().startsWith('MPK')) {
              mpk = part.toUpperCase();
            } else if (/\d+\/\d+/.test(part)) {
              group = part;
            } else if (/\d+\/\d{4}/.test(part)) {
              number = part;
            }
          });
          if (vendor) {
            vendorMapping[vendor.toLowerCase()] = { mpk, group };
          }
          // Update last number map
          if (mpk && group && number) {
            const key = mpk + '|' + group;
            const numMatch = number.match(/(\d+)\/(\d{4})/);
            if (numMatch) {
              const num = parseInt(numMatch[1], 10);
              const year = parseInt(numMatch[2], 10);
              const current = lastNumberMap[key];
              if (!current || num > current.value) {
                lastNumberMap[key] = { value: num, year: year };
              }
            }
          }
        });
      }
      alert('Wczytano wytyczne.');
    } catch (err) {
      showError('Błąd podczas wczytywania wytycznych: ' + err.message);
    }
  };
  reader.onerror = () => {
    showError('Nie udało się odczytać pliku z wytycznymi.');
  };
  reader.readAsArrayBuffer(file);
}

async function processInvoice(pdfFile) {
  hideError();
  outputDiv.classList.add('hidden');
  try {
    const arrayBuffer = await pdfFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = '';
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const strings = content.items.map((item) => item.str);
      fullText += strings.join('\n') + '\n';
    }
    // Extract vendor, buyer NIP and invoice number
    const vendor = extractVendor(fullText);
    const nipBuyer = extractNIPBuyer(fullText);
    const invoiceNumber = extractInvoiceNumber(fullText);
    // Determine MPK and group
    const mapping = findVendorMapping(vendor);
    let mpk = mapping.mpk || 'MPK000';
    let group = mapping.group || '0/0';
    // Determine next sequential number for this mpk and group
    let nextNumberObj = generateNextNumber(mpk, group);
    // Compose label for the invoice (Etykieta)
    const numberFormatted = String(nextNumberObj.value).padStart(3, '0') + '/' + nextNumberObj.year;
    const etykieta = `${group};${mpk};${numberFormatted}`;
    // Update dataset and UI
    const record = {
      'Nazwa kontrahenta': vendor,
      'NIP nabywcy': nipBuyer,
      'Numer faktury': invoiceNumber,
      'MPK': mpk,
      'Grupa': group,
      'Numer kolejny': numberFormatted,
      'Etykieta': etykieta,
    };
    dataset.push(record);
    displayOutput(record);
    // Create modified PDF with label in header
    const labelDisplay = `${group} – ${mpk} – ${numberFormatted}`;
    lastLabel = labelDisplay;
    modifiedPdfBytes = await createModifiedPdf(arrayBuffer, labelDisplay);
    downloadSection.classList.remove('hidden');
  } catch (err) {
    showError('Błąd podczas przetwarzania faktury: ' + err.message);
  }
}

function extractVendor(text) {
  // Try to extract after 'Sprzedawca:' up to next newline and before 'NIP'
  const sprzedawcaRegex = /Sprzedawca:?\s*\n?([^\n]+)\n/i;
  let match = text.match(sprzedawcaRegex);
  if (match) {
    return match[1].trim();
  }
  // Fallback: return first non-empty line containing 'sp.' or 'Sp.' as company marker
  const lines = text.split(/\n/);
  for (let line of lines) {
    if (/sp\.?/i.test(line) && !/Nabywca|NIP/i.test(line)) {
      return line.trim();
    }
  }
  return 'Nie znaleziono';
}

function extractNIPBuyer(text) {
  // look for 'NIP:' after 'Nabywca'
  const nabywcaIndex = text.search(/Nabywca/i);
  let searchArea = text;
  if (nabywcaIndex >= 0) {
    searchArea = text.slice(nabywcaIndex);
  }
  const nipRegex = /NIP[:\s]*([0-9]{10})/;
  let match = searchArea.match(nipRegex);
  if (match) {
    return match[1];
  }
  // fallback: find any 10-digit number that looks like NIP
  const fallback = text.match(/([0-9]{10})/);
  return fallback ? fallback[1] : 'Brak';
}

function extractInvoiceNumber(text) {
  // look for patterns like FZ 328/01/2023 or 18/11/2023 or 328/01/2023
  // We'll first search for two letters + numbers pattern
  const regexFull = /([A-Z]{1,3}\s*\d+[\/-]\d+[\/-]\d{2,4})/;
  let match = text.match(regexFull);
  if (match) {
    return match[1].replace(/\s+/g, ' ').trim();
  }
  // fallback to digits/digits/4-digit year
  const simple = /(\d+[\/-]\d+[\/-]\d{4})/;
  match = text.match(simple);
  return match ? match[1] : 'Nieznany';
}

function findVendorMapping(vendor) {
  if (!vendor) return { mpk: '', group: '' };
  const key = vendor.toLowerCase().trim();
  return vendorMapping[key] || { mpk: '', group: '' };
}

function generateNextNumber(mpk, group) {
  const now = new Date();
  const year = now.getFullYear();
  const key = mpk + '|' + group;
  let entry = lastNumberMap[key];
  if (entry) {
    // If entry year is previous year, reset to 1
    if (entry.year === year) {
      entry.value = entry.value + 1;
    } else {
      entry.value = 1;
      entry.year = year;
    }
  } else {
    entry = { value: 1, year: year };
  }
  // Save back to map
  lastNumberMap[key] = { value: entry.value, year: entry.year };
  return entry;
}

function displayOutput(record) {
  outputDiv.classList.remove('hidden');
  outputDiv.innerHTML = '';
  const table = document.createElement('table');
  const headerRow = document.createElement('tr');
  Object.keys(record).forEach((key) => {
    const th = document.createElement('th');
    th.textContent = key;
    headerRow.appendChild(th);
  });
  table.appendChild(headerRow);
  const dataRow = document.createElement('tr');
  Object.values(record).forEach((value) => {
    const td = document.createElement('td');
    td.textContent = value;
    dataRow.appendChild(td);
  });
  table.appendChild(dataRow);
  outputDiv.appendChild(table);
}

/**
 * Create a modified copy of the original PDF with a prominent label
 * at the top of the first page. Uses pdf-lib to embed text.
 * @param {ArrayBuffer} originalBuffer The original PDF as ArrayBuffer.
 * @param {string} label The text to draw, e.g. "3/8 – MPK610 – 181/2025".
 * @returns {Promise<Uint8Array>} The modified PDF bytes.
 */
async function createModifiedPdf(originalBuffer, label) {
  try {
    const pdfDoc = await PDFLib.PDFDocument.load(originalBuffer);
    const pages = pdfDoc.getPages();
    if (pages.length === 0) return new Uint8Array(originalBuffer);
    const firstPage = pages[0];
    const { width, height } = firstPage.getSize();
    // Embed a standard bold font
    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
    const fontSize = 20;
    // Choose a striking colour (red) for visibility
    const color = PDFLib.rgb(0.8, 0.0, 0.0);
    // Position: margin of 40 points from top-left
    const x = 50;
    const y = height - 40;
    firstPage.drawText(label, {
      x: x,
      y: y,
      size: fontSize,
      font: font,
      color: color,
    });
    const pdfBytes = await pdfDoc.save();
    return pdfBytes;
  } catch (err) {
    console.error('createModifiedPdf error:', err);
    return new Uint8Array(originalBuffer);
  }
}

function downloadExcel() {
  // Convert dataset array to worksheet and then to workbook
  const ws = XLSX.utils.json_to_sheet(dataset);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Raport');
  XLSX.writeFile(wb, 'raport_faktury.xlsx');
}

function showError(msg) {
  errorDiv.textContent = msg;
  errorDiv.classList.remove('hidden');
}

function hideError() {
  errorDiv.textContent = '';
  errorDiv.classList.add('hidden');
}