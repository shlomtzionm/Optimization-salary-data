
let fileInput = document.getElementById("fileInput");
fileInput.addEventListener("change", handleExcel);

function handleExcel(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheet = convertSheetToJson("ראשי", workbook);
    const N = convertSheetToJson("נ מרוכז", workbook);

    const indexOfType = findIndex(N[0], "שם הרכיב");
    const indexPriceInN = findIndex(N[0], "מחיר");
    const indexAmountInN = findIndex(N[0], "כמות");
  const indexMonthlyInN = findIndex(N[0], "תשלום");

let variables = {N, firstSheet,indexOfType,indexPriceInN,indexAmountInN,indexMonthlyInN}


    handleTravelRegular(variables);
    handleTravelRegularDiscount(variables)
    handleTravel75(N, firstSheet,indexOfType)
    handleTravelRegularSenior(variables)
    handleTravelDiscountSenior(variables)
    console.log(N, firstSheet);
    console.log(errors)
    deleteAllNAs(firstSheet)
    createNewExcel(firstSheet);
  };

  reader.readAsArrayBuffer(file);
}

let errors = []



function convertSheetToJson(sheetName, workbook) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet){ errors.push(`couldn't find ${sheetName} sheet`) 
  } else {
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    return jsonData;
  }
}

function handleTravelRegular(variables) {
  const firstSheet = variables.firstSheet
  const indexPriceInFirstSheet = findIndex(firstSheet[0], "נסיעות תעריף");
  const indexAmountInFirstSheet = findIndex(firstSheet[0], "נסיעות כמות");
  const indexMonthlyInFirstSheet = findIndex(firstSheet[0], "נסיעות סכום");

  logic("נסיעות",variables,indexPriceInFirstSheet,indexAmountInFirstSheet,indexMonthlyInFirstSheet,errors)

}

function handleTravelRegularDiscount(variables) {
  const firstSheet = variables.firstSheet
  const indexPriceInFirstSheet = findIndex(firstSheet[0], "נסיעות מ תעריף");
  const indexAmountInFirstSheet = findIndex(firstSheet[0], "נסיעות מ כמות");
  const indexMonthlyInFirstSheet = findIndex(firstSheet[0], "נסיעות מ סכום");

 logic("נסיעות מ", variables,indexPriceInFirstSheet,indexAmountInFirstSheet,indexMonthlyInFirstSheet,errors)
}


function handleTravelRegularSenior(variables) {
  const firstSheet = variables.firstSheet
  const indexPriceInFirstSheet = findIndex(firstSheet[0], "נסיעות ו תעריף");
  const indexAmountInFirstSheet = findIndex(firstSheet[0], "נסיעות ו כמות");
  const indexMonthlyInFirstSheet = findIndex(firstSheet[0], "נסיעות ו סכום");

 logic("נסיעות ותיק", variables,indexPriceInFirstSheet,indexAmountInFirstSheet,indexMonthlyInFirstSheet,errors)
}

function handleTravelDiscountSenior(variables) {
  const firstSheet = variables.firstSheet
  const N = variables.N

  const indexPriceInFirstSheet = findIndex(firstSheet[0], "נסיעות ו מ סכום");

  for (i = 1; i < variables.N.length - 1; i++) {
    if (N[i][variables.indexOfType] === "נסיעות מ ותיק") {
      const workerNumber = N[i][0];
debugger
      for (j = 0; j < firstSheet.length - 1; j++) {
        if (firstSheet[j][0] === workerNumber) {
          firstSheet[j][indexPriceInFirstSheet] = variables.N[i][variables.indexMonthlyInN];
          debugger
        }
      }
    }
  }
}

function handleTravel75(n, firstSheet,indexOfType) {
  const indexPriceInFirstSheet = findIndex(firstSheet[0], "נסיעות 75");

  for (i = 1; i < n.length - 1; i++) {
    if (n[i][indexOfType] === "נסיעות 75") {
      const workerNumber = n[i][0];
debugger
      for (j = 0; j < firstSheet.length - 1; j++) {
        if (firstSheet[j][0] === workerNumber) {
          firstSheet[j][indexPriceInFirstSheet] = 1;
          debugger
        }
      }
    }
  }
}

function findIndex(array, value) {
  for (let i = 0; i < array.length; i++) {
    if (array[i] === value) {
      return i;
    }
  }
  errors.push(`couldn't find ${value} column`);
  return -1
}

function deleteAllNAs(data) {
  for (let i = 1; i < data.length - 1; i++) {
    for (let j = 0; j < data[i].length - 1; j++) {
      if (data[i][j] === "#N/A") {
        data[i][j] = "";
      }
    }
  }
}

function createNewExcel(data) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet);
  XLSX.writeFile(workbook, "חדש.xlsx");
}


function logic(thisType,variables, indexPriceInFirstSheet, indexAmountInFirstSheet, indexMonthlyInFirstSheet, errors) {
  const firstSheet = variables.firstSheet
  const n = variables.N
  for (let i = 1; i < n.length-1; i++) { 
    if (n[i][variables.indexOfType] === thisType) {
      const workerNumber = n[i][0];
      let found = false;
      for (let j = 0; j < firstSheet.length-1; j++) { 
        if (firstSheet[j][0] === workerNumber) {
          found = true;
          if (n[i][variables.indexAmountInN]) {
            firstSheet[j][indexPriceInFirstSheet] = n[i][variables.indexPriceInN];
            firstSheet[j][indexAmountInFirstSheet] = n[i][variables.indexAmountInN];
          } else if (n[i][variables.indexMonthlyInN]) {
            firstSheet[j][indexMonthlyInFirstSheet] = n[i][variables.indexMonthlyInN];
          } else {
            errors.push(`${workerNumber} suppose to have ${thisType} but don't have any price`);
          }
        }
      }
      if (!found) {
        errors.push(`couldn't find ${workerNumber} in the main sheet`);
      }
    }
  }
  if (!n.some(row => row[variables.indexOfType] === thisType)) {
    errors.push(`couldn't find a worker with ${thisType}`);
  }
}
