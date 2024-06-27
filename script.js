let fileInput = document.getElementById("fileInput");
fileInput.addEventListener("change", handleExcel);

let errors = [];

function handleExcel(e) {
  const file = e.target.files[0];
  if (!file) {
    errors.push("No file selected");
    handleErrors(errors);
    return;
  }

  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheet = convertSheetToJson("ראשי", workbook);
    const N = convertSheetToJson("נ מרוכז", workbook);

    if (!firstSheet || !N) {
      handleErrors(errors);
      return;
    }

    const variables = {
      N,
      firstSheet,
      indexWorkDays: findIndex(firstSheet[0], "ימי עבודה בפועל"),
      indexOfType: findIndex(N[0], "שם הרכיב"),
      indexPriceInN: findIndex(N[0], "מחיר"),
      indexMonthlyInN: findIndex(N[0], "תשלום"),
      indexPriceInFirstSheet: null,
      indexAmountInFirstSheet: null,
      indexMonthlyInFirstSheet: null
    };

    checkIfEveryoneHaveTravel(variables.firstSheet, variables.N)
    handleTravelRegular(variables);
    handleTravelRegularDiscount(variables);
    handleTravel75(variables);
    handleTravelRegularSenior(variables);
    handleTravelDiscountSenior(variables);
    handleTravelExtra(variables);

    deleteAllNAs(firstSheet);
    createNewExcel(firstSheet, errors);
    handleErrors(errors);
  };

  reader.readAsArrayBuffer(file);
}

function convertSheetToJson(sheetName, workbook) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    errors.push(`Couldn't find ${sheetName} sheet`);
    return null;
  }
  return XLSX.utils.sheet_to_json(sheet, { header: 1 });
}

function handleTravelRegular(variables) {
  updateIndexesForType(variables, "נסיעות תעריף", "נסיעות כמות", "נסיעות סכום");
  logic("נסיעות", variables, 20, 11, 225);
}

function handleTravelRegularDiscount(variables) {
  updateIndexesForType(variables, "נסיעות מ תעריף", "נסיעות מ כמות", "נסיעות מ סכום");
  logic("נסיעות מ", variables, 8, 11, 99);
}

function handleTravelRegularSenior(variables) {
  updateIndexesForType(variables, "נסיעות ו תעריף", "נסיעות ו כמות", "נסיעות ו סכום");
  logic("נסיעות ותיק", variables, 20, 5.5, 112.5);
}

function handleTravelExtra(variables) {
  updateIndexesForType(variables, "תוספת נסיעות תעריף", "תוספת נסיעות כמות","תוספת נסיעות סכום");

  const { N, firstSheet, indexOfType, indexPriceInFirstSheet, indexAmountInFirstSheet, indexMonthlyInFirstSheet, indexWorkDays, indexPriceInN, indexMonthlyInN } = variables;

  N.forEach((row, i) => {
    if (i === 0 || i === N.length - 1) return;

    if (row[indexOfType] === "תוספת נסיעות") {
      const workerNumber = row[0];
      let found = false;

      firstSheet.forEach((firstSheetRow, j) => {
        if (j === 0 || j === firstSheet.length - 1) return;

        if (firstSheetRow[0] === workerNumber) {
          found = true;
          if (row[findIndex(N[0],"מחיר")]) {
            firstSheetRow[indexPriceInFirstSheet] = row[indexPriceInN];
            firstSheetRow[indexAmountInFirstSheet] = firstSheetRow[indexWorkDays];
            debugger
          } else if (row[indexMonthlyInN]) {
            firstSheetRow[indexMonthlyInFirstSheet] = row[indexMonthlyInN];
          } else {
            errors.push(`${workerNumber} has "תוספת נסיעות" but no price`);
          }
        }
      });

      if (!found) errors.push(`Couldn't find ${workerNumber} in the main sheet`);
    }
  });

  if (!N.some(row => row[indexOfType] === "תוספת נסיעות")) {
    errors.push(`Couldn't find a worker with תוספת נסיעות`);
  }
}

function handleTravelDiscountSenior(variables) {
  const { N, firstSheet, indexOfType, indexPriceInFirstSheet, indexAmountInFirstSheet, indexMonthlyInFirstSheet, indexWorkDays, indexPriceInN, indexMonthlyInN } = variables;

  N.forEach((row, i) => {
    if (i === 0 || i === N.length - 1) return;

    if (row[indexOfType] === "נסיעות מ ותיק") {
      const workerNumber = row[0];
      let found = false;

      firstSheet.forEach((firstSheetRow, j) => {
        if (j === 0 || j === firstSheet.length - 1) return;

        if (firstSheetRow[0] === workerNumber) {
          found = true;
          if (firstSheetRow[indexWorkDays] <= 8) {
            firstSheetRow[findIndex(firstSheet[0],"נסיעות ו מ סכום")] = 11;
          } else if (firstSheetRow[indexWorkDays] > 8) {
            firstSheetRow[findIndex(firstSheet[0],"נסיעות ו מ סכום")] = 44.5;
          } else {
            errors.push(`${workerNumber} has "נסיעות מ ותיק" but no price`);
          }
        }
      });

      if (!found) errors.push(`Couldn't find ${workerNumber} in the main sheet`);
    }
  });

  if (!N.some(row => row[indexOfType] === "נסיעות מ ותיק")) {
    errors.push(`Couldn't find a worker with  נסיעות מ ותיק`);
  }
}

function handleTravel75(variables) {
  const { N, firstSheet, indexOfType } = variables;
  const indexPriceInFirstSheet = findIndex(firstSheet[0], "נסיעות 75");

  N.forEach((row, i) => {
    if (i === 0 || i === N.length - 1) return;

    if (row[indexOfType] === "נסיעות 75") {
      const workerNumber = row[0];

      firstSheet.forEach((firstSheetRow, j) => {
        if (j === 0 || j === firstSheet.length - 1) return;

        if (firstSheetRow[0] === workerNumber) {
          firstSheetRow[indexPriceInFirstSheet] = 1;
        }
      });
    }
  });
}

function findIndex(array, value) {
  const index = array.indexOf(value);
  if (index === -1) {
    errors.push(`Couldn't find ${value} column`);
  }
  return index;
}

function deleteAllNAs(data) {
  data.forEach((row, i) => {
    if (i === 0 || i === data.length - 1) return;

    row.forEach((cell, j) => {
      if (cell === "#N/A") row[j] = "";
    });
  });
}

function createNewExcel(data, errors) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Data");

  const errorsArray = errors.map(error => [error]);
  const worksheetErrors = XLSX.utils.aoa_to_sheet(errorsArray);
  XLSX.utils.book_append_sheet(workbook, worksheetErrors, "Errors");

  XLSX.writeFile(workbook, "חדש.xlsx");
}

function logic(thisType, variables, lessThen, price, mPrice) {
  const { N, firstSheet, indexOfType, indexWorkDays } = variables;

  N.forEach((row, i) => {
    if (i === 0 || i === N.length - 1) return;

    if (row[indexOfType] === thisType) {
      const workerNumber = row[0];
      let found = false;

      firstSheet.forEach((firstSheetRow, j) => {
        if (j === 0 || j === firstSheet.length - 1) return;

        if (firstSheetRow[0] === workerNumber) {
          found = true;
          if (firstSheetRow[indexWorkDays] <= lessThen) {
            byDay(variables, j, price);
          } else if (firstSheetRow[indexWorkDays] > lessThen) {
            byMonth(variables, j, mPrice);
          } else {
            errors.push(`${workerNumber} has ${thisType} but no price`);
          }
        }
      });

      if (!found) errors.push(`Couldn't find ${workerNumber} in the main sheet`);
    }
  });

  if (!N.some(row => row[indexOfType] === thisType)) {
    errors.push(`Couldn't find a worker with ${thisType}`);
  }
}

function byDay(variables, j, price) {
  const { firstSheet, indexPriceInFirstSheet, indexAmountInFirstSheet, indexWorkDays } = variables;
  firstSheet[j][indexPriceInFirstSheet] = price;
  firstSheet[j][indexAmountInFirstSheet] = firstSheet[j][indexWorkDays];
}

function byMonth(variables, j, mPrice) {
  const { firstSheet, indexMonthlyInFirstSheet } = variables;
  firstSheet[j][indexMonthlyInFirstSheet] = mPrice;
}

function updateIndexesForType(variables, priceHeader, amountHeader, monthlyHeader) {
  const { firstSheet } = variables;
  variables.indexPriceInFirstSheet = findIndex(firstSheet[0], priceHeader);
  variables.indexAmountInFirstSheet = findIndex(firstSheet[0], amountHeader);
  variables.indexMonthlyInFirstSheet = findIndex(firstSheet[0], monthlyHeader);
}

function handleErrors(errors) {
  const errorContainer = document.getElementById("errorContainer");
  if (errors.length === 0) {
    errorContainer.innerHTML = "<p>No errors found</p>";
    return;
  }
  let table = `<table id="errorTable"><thead><tr><th>Error Messages</th></tr></thead><tbody>`;
  errors.forEach(error => {
    table += `<tr><td>${error}</td></tr>`;
  });
  table += `</tbody></table>`;

  errorContainer.innerHTML = table;
}

function checkIfEveryoneHaveTravel(firstSheet,n){
  for (let i = 1; i < firstSheet.length; i++) {
    const workerNumber = firstSheet[i][0];
    let found = false;

    for (let j = 1; j < n.length-1; j++) {
      if (n[j][0] === workerNumber) {
        found = true;
        break; // Exit the loop early if found
      }
    }

    if (!found) {
      errors.push(`${workerNumber} isn't in נ מרוכז`);
    }
  }
}