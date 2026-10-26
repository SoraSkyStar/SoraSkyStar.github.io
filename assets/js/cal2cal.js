const dropZone = document.getElementById("drop-zone");
const staffID = document.getElementById("input-staff-id");
const outputInfo = document.getElementById("process-output");
const btnProcess = document.getElementById("btn-process");

dropZone.addEventListener("drop", dropHandler);
staffID.addEventListener("input", updateIDValue);
btnProcess.addEventListener("click", startProcess);

let legalStaffID = /[0-9]{7}/;
let trainingShift = [
  "16U",
  "08U",
  "07H",
  "13H",
  "07U",
  "14U",
  "08U",
  "15H",
  "08H",
];
let shiftFile = [];
let shiftCSV = [];

window.addEventListener("drop", (e) => {
  e.preventDefault();
});

dropZone.addEventListener("dragover", (e) => {
  const fileItems = [...e.dataTransfer.items].filter(
    (item) => item.kind === "file",
  );
  if (fileItems.length > 0) {
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  }
});

window.addEventListener("dragover", (e) => {
  e.preventDefault();
});

function displayFilesList(files) {
  listFiles.innerText = "";
  for (const file of files) {
    listFiles.innerText += `${file.name}\n`;
  }
}

function dropHandler(ev) {
  ev.preventDefault();
  shiftFile = [...ev.dataTransfer.items]
    .map((item) => item.getAsFile())
    .filter((file) => file.name.includes(".xlsm"));
  dropZone.innerText = shiftFile[0].name;
}

function updateIDValue(e) {
  // console.log(staffID.value);
  // console.log(legalStaffID.test(staffID.value));
  if (legalStaffID.test(staffID.value)) {
    staffID.className = "pass";
    setOutput("", true);
  } else {
    staffID.className = "error";
    setOutput("無效ID", false);
  }
}

async function startProcess(e) {
  if (staffID.className != "pass") {
    setOutput("無效ID", false);
    return;
  }
  try {
    const buffer = await shiftFile[0].arrayBuffer();
    console.log("Successfully Open file.");

    const fileWb = new ExcelJS.Workbook();
    await fileWb.xlsx.load(buffer);
    console.log("ExcelJS Successfully load from buffer.");

    const ws = fileWb.worksheets[0];
    console.log(`Open WS ${ws.name}, ${ws.rowCount}.`);

    const colID = ws.getColumn("C");
    // console.log(`${colID.values}`);

    const rowID = colID.values.indexOf(parseInt(staffID.value));
    console.log(rowID);
    if (rowID < 0) {
      setOutput("找無ID", false);
      return;
    }

    const rowCalendar = ws.getRow(3); //    取得日期
    console.log(rowCalendar.values);
    const rowShift = ws.getRow(rowID); //   取得班表
    console.log(rowShift.values);

    shiftCSV = [["Subject", "Start Date", "Start Time"]]; // 清空陣列
    rowCalendar.values.forEach((date, index) => {
      if (date instanceof Date) {
        // 將班表代號直接寫入時間date變數
        const shiftCode = rowShift.values[index];
        if (shiftCode.trim().length == 3) {
          date = parseShiftCode(shiftCode, date);

          // date.getMonth() Jan gives 0
          const startDate = `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
          const startHours = date.getHours();
          const startMinutes = date.getMinutes();
          const startTime = `${formatHRs(startHours)}:${formatMINs(startMinutes)} ${startHours >= 12 ? "PM" : "AM"}`;

          const shift = [shiftCode.toSrting(), startDate, startTime];
          shiftCSV.push(shift);
        }
      }
    });
    console.log(shiftCSV.length, shiftCSV);

    // write to file
    const outputCSV = `shift_${staffID.value}.csv`;
    downloadCSV(outputCSV, shiftCSV);
    setOutput(`已寫入${outputCSV}`, true);
  } catch (error) {
    console.log(error);
    setOutput("讀取檔案錯誤，請檢查檔案", false);
  }
}

function formatHRs(hrs) {
  const hr = hrs > 12 ? hrs - 12 : hrs;
  return hr < 10 ? `0${hr}` : `${hr}`; // 補足位數 字串格式
}

function formatMINs(mins) {
  return mins > 10 ? `${mins}` : `0${mins}`; // 補足位數 字串格式
}

function parseShiftCode(shiftCode, date) {
  shiftCode = shiftCode.trim();
  // 先把訓練班表轉通用班表
  if (shiftCode[0] == "T") { // 訓練班表
    shiftCode = trainingShift[parseInt(shiftCode.slice(1)) - 1];
  } 
  if (shiftCode.includes("H") || shiftCode.includes("I")) { // 整點班表
    date.setHours(parseInt(shiftCode.slice(0, 2)), 0);
  } else if (shiftCode.includes("U") || shiftCode.includes("V")) { // 半班表
    date.setHours(parseInt(shiftCode.slice(0, 2)), 30);
  } else if(shiftCode == "001"){ // 辦公室班表
    date.setHours(8,0); // 設定成最早八點整
  }
  return date;
}

function setOutput(string, result) {
  outputInfo.innerText = string;
  if (result) {
    outputInfo.className = "";
  } else {
    outputInfo.className = "error";
  }
}

function arrayToCSV(rows) {
  return rows
    .map((row) =>
      row
        .map((value) => {
          // 處理 null、undefined
          const text = value == null ? "" : String(value);

          // CSV 欄位含逗號、雙引號或換行時，需要加上雙引號
          return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
        })
        .join(","),
    )
    .join("\r\n");
}

function downloadCSV(filename, rows) {
  const csv = "\uFEFF" + arrayToCSV(rows);

  const blob = new Blob([csv], {
    type: "text/csv;charset=utf-8;",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");

  a.href = url;
  a.download = filename.endsWith(".csv") ? filename : `${filename}.csv`;

  document.body.appendChild(a);
  a.click();
  a.remove();

  // 延後釋放 URL，避免下載尚未開始就被清除
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
