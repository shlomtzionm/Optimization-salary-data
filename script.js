
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

    handleTravelRegular(N, firstSheet,indexOfType,indexPriceInN,indexAmountInN,indexMonthlyInN);
    handleTravel75(N, firstSheet,indexOfType)
    console.log(N, firstSheet);
    createNewExcel(firstSheet);
  };

  reader.readAsArrayBuffer(file);
}

function convertSheetToJson(sheetName, workbook) {
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  return jsonData;
}

function handleTravelRegular(n, firstSheet,indexOfType,indexPriceInN,indexAmountInN,indexMonthlyInN) {
  const indexPriceInFirstSheet = findIndex(firstSheet[0], "נסיעות תעריף");
  const indexAmountInFirstSheet = findIndex(firstSheet[0], "נסיעות כמות");
  const indexMonthlyInFirstSheet = findIndex(firstSheet[0], "נסיעות סכום");

  for (i = 1; i < n.length - 1; i++) {
    if (n[i][indexOfType] === "נסיעות") {
      const workerNumber = n[i][0];
      for (j = 0; j < firstSheet.length - 1; j++) {
        if (firstSheet[j][0] === workerNumber) {
          if(n[i][indexAmountInN]){
            firstSheet[j][indexPriceInFirstSheet] = n[i][indexPriceInN];
            firstSheet[j][indexAmountInFirstSheet] = n[i][indexAmountInN]
            } else {
firstSheet[j][indexMonthlyInFirstSheet] = n[i][indexMonthlyInN]
            }
        }
      }
    }
  }
}

// function handleTravelRegular(n, firstSheet,indexOfType) {
//   const indexPriceInN = findIndex(n[0], "מחיר");
//   const indexAmountInN = findIndex(n[0], "כמות");
//   const indexPriceInFirstSheet = findIndex(firstSheet[0], "נסיעות תעריף");
//   const indexAmountInFirstSheet = findIndex(firstSheet[0], "נסיעות כמות");
//   const indexMonthlyInN = findIndex(n[0], "תשלום");
//   const indexMonthlyInFirstSheet = findIndex(firstSheet[0], "נסיעות סכום");

//   for (i = 1; i < n.length - 1; i++) {
//     if (n[i][indexOfType] === "נסיעות") {
//       const workerNumber = n[i][0];
//       for (j = 0; j < firstSheet.length - 1; j++) {
//         if (firstSheet[j][0] === workerNumber) {
//           if(n[i][indexAmountInN]){
//             firstSheet[j][indexPriceInFirstSheet] = n[i][indexPriceInN];
//             firstSheet[j][indexAmountInFirstSheet] = n[i][indexAmountInN]
//             } else {
// firstSheet[j][indexMonthlyInFirstSheet] = n[i][indexAmountInN]
//             }
//         }
//       }
//     }
//   }
// }

function handleTravel75(n, firstSheet,indexOfType) {
  const indexPriceInFirstSheet = findIndex(firstSheet[0], "נסיעות 75");

  for (i = 1; i < n.length - 1; i++) {
    if (n[i][indexOfType] === "75 נסיעות") {
      const workerNumber = n[i][0];

      for (j = 0; j < firstSheet.length - 1; j++) {
        if (firstSheet[j][0] === workerNumber) {
          firstSheet[j][indexPriceInFirstSheet] = 1;
        }
      }
    }
  }
}

function findIndex(array, value) {
  for (let i = 0; i < array.length-1; i++) {
    if (array[i] === value) {
      return i;
    }
  }
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
