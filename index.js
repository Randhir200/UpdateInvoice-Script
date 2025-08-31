const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const { toWords } = require("number-to-words");

// === Parameters ===
const inputFolder = "../../BINOD/Dummy";     // path for udpated invoices
const outputFolder = "../../BINOD/Dummy/Updated_Invoices";
const newInvoiceDate = "01/09/2025";  // new invoice date
const lastMonth = "August";              // Bill of the month
const lastMonthNum = "08";             // month number
const newYear = "2025";
let initalInvoiceAmt = 10503;     // Last invoice amount
let initialInvoiceNum = 11069;    // Last invoice number

// agar output folder nahi hai to bana do
if (!fs.existsSync(outputFolder)) {
  fs.mkdirSync(outputFolder);
}

function amountToWords(num) {
  const [rupeesStr, paiseStr] = num.toString().split(".");
  const rupees = parseInt(rupeesStr, 10);
  const paise = paiseStr ? parseInt(paiseStr.padEnd(2, "0").slice(0, 2), 10) : 0;

  let words = "";
  if (rupees > 0) {
    words += toWords(rupees) + " rupees";
  }
  if (paise > 0) {
    words += " and " + toWords(paise) + " paisa";
  }
  return (
    words.charAt(0).toUpperCase() +
    words.slice(1) +
    " only."
  );
}


async function updateInvoice(filePath, outputPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet(1);

  // F5: Invoice no. +1
  let invCell = sheet.getCell("F5").value;
  if (invCell) {
    let numMatch = invCell.toString().match(/(\d+)/);
    if (numMatch) {
    //   let num = parseInt(numMatch[0], 10) + 1;
    initialInvoiceNum = initialInvoiceNum + 1;
    sheet.getCell("F5").value = invCell.toString().replace(numMatch[0], initialInvoiceNum);
    }
  }

  // I5: Invoice Date
  sheet.getCell("I5").value = `Invoice Date: ${newInvoiceDate}`;

  // F10: Bill of the month
  sheet.getCell("F10").value = `Bill of the month ${lastMonth} ${newYear}(01-${lastMonthNum}-${newYear} to 31-${lastMonthNum}-${newYear})`;
  
  //C16 String vlaue
  sheet.getCell("C16").value = `Bill of the month ${lastMonth}`
  // C17: From/To Dates
  sheet.getCell("C17").value = `(From Date -01-${lastMonthNum}-${newYear} TO Date 31-${lastMonthNum}-${newYear})`;

  // J13: Amount +1
  let amt = parseFloat(sheet.getCell("J13").value) || 0;
  //   let newAmt = amt + 1;
  //Increase by 1 amount
  initalInvoiceAmt = initalInvoiceAmt+1;
  sheet.getCell("J13").value = initalInvoiceAmt;
  // J39: update
  sheet.getCell("J39").value = initalInvoiceAmt;
  
  // J40 & J41: 9% tax
  let tax1 = (initalInvoiceAmt * 9) / 100;
  let tax2 = (initalInvoiceAmt * 9) / 100;
  sheet.getCell("J40").value = tax1;
  sheet.getCell("J41").value = tax2;
  
  // J43: total
  let total = initalInvoiceAmt + tax1 + tax2;
  sheet.getCell("J43").value = total;
  // B43: total in words
  sheet.getCell("B43").value = amountToWords(total);


  await workbook.xlsx.writeFile(outputPath);
  console.log("Updated:", path.basename(outputPath));
}

async function processAllInvoices() {
  const files = fs.readdirSync(inputFolder).filter(f => f.endsWith(".xlsx"));

  for (const file of files) {
    const inputPath = path.join(inputFolder, file);
    const outputPath = path.join(outputFolder, file);
    await updateInvoice(inputPath, outputPath);
  }
  console.log("✅ All invoices updated!");
}

processAllInvoices();
