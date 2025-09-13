const { toWords } = require("number-to-words");

function amountToWords(num) {
  // Round to 2 decimal places safely
  const fixedNum = Number(num).toFixed(2); // "12427.76"
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


console.log("amount=>", amountToWords(12427.759999999998))