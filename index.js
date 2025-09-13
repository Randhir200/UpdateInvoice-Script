const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const { toWords } = require("number-to-words");

// === Parameters ===
const inputFolder = "../../BINOD/September 2025";     // path for udpated invoices // D:\BINOD\September 2025
const outputFolder = "../../BINOD/September 2025/Updated_Invoices";  // New path where you want updated sheet
const newInvoiceDate = "01/10/2025";  // new invoice date
const lastMonth = "September";              // Bill of the month
const lastMonthNum = "09"             // month number
const newYear = "2025";
let lastMonthAmt = 10527;     // Last invoice amount
let lastMonthInvNum = 11103;    // Last invoice number
let lastMonthTotalDays = 30 //

// if folder is not present then create
if (!fs.existsSync(outputFolder)) {
  fs.mkdirSync(outputFolder);
}

function amountToWords(num) {
  // Round to 2 decimal places safely
  const fixedNum = Number(num).toFixed(2); // "12427.7599999" => 12427.56
  const [rupeesStr, paiseStr] = fixedNum.split(".");
  const rupees = parseInt(rupeesStr, 10);
  const paise = paiseStr ? parseInt(paiseStr, 10) : 0;

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
    lastMonthInvNum = lastMonthInvNum + 1;
    sheet.getCell("F5").value = invCell.toString().replace(numMatch[0], lastMonthInvNum);
    }
  }

  // I5: Invoice Date
  sheet.getCell("I5").value = `Invoice Date: ${newInvoiceDate}`;

  // F10: Bill of the month
  sheet.getCell("F10").value = `Bill of the month ${lastMonth} ${newYear}(01-${lastMonthNum}-${newYear} to ${lastMonthTotalDays}-${lastMonthNum}-${newYear})`;
  
  //C16 String vlaue
  sheet.getCell("C16").value = `Bill of the month ${lastMonth}`
  // C17: From/To Dates
  sheet.getCell("C17").value = `(From Date -01-${lastMonthNum}-${newYear} TO Date ${lastMonthTotalDays}-${lastMonthNum}-${newYear})`;

  // J13: Amount +1
  // let amt = parseFloat(sheet.getCell("J13").value) || 0;
  //   let newAmt = amt + 1;
  //Amount increase by 1
  lastMonthAmt = lastMonthAmt+1;
  sheet.getCell("J13").value = lastMonthAmt;
  // J39: update
  sheet.getCell("J39").value = lastMonthAmt;
  
  // J40 & J41: 9% tax
  let tax1 = (lastMonthAmt * 9) / 100;
  let tax2 = (lastMonthAmt * 9) / 100;
  sheet.getCell("J40").value = tax1;
  sheet.getCell("J41").value = tax2;
  
  // J43: total
  let total = lastMonthAmt + tax1 + tax2;
  sheet.getCell("J43").value = total;
  // B43: total in words
  const totalAmntInWord = amountToWords(total)
  sheet.getCell("B43").value = totalAmntInWord;


  await workbook.xlsx.writeFile(outputPath);
  console.log("Updated:", path.basename(outputPath));
}

async function processAllInvoices() {
  const files = fs.readdirSync(inputFolder).filter(f => f.endsWith(".xlsx"));

  for (const file of files) {
    const inputPath = path.join(inputFolder, file);
    console.log(file)
    // Replace "August" with "September" (ya jo bhi target month hai)
    const month = file.split(" ")[0];
    const outputFile = file.replace(month, lastMonth);
    const outputPath = path.join(outputFolder, outputFile);
    await updateInvoice(inputPath, outputPath);
  }
  console.log("✅ All invoices updated!");
}

processAllInvoices();
