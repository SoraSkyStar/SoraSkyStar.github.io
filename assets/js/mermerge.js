const dropZone = document.getElementById("drop-zone");
const listFiles = document.getElementById("files-list");

dropZone.addEventListener("drop", dropHandler);

window.addEventListener("drop", (e) => {
  e.preventDefault();
});

const monthName = [
  "JAN",
  "FEB",
  "MAR",
  "APR",
  "MAY",
  "JUN",
  "JUL",
  "AUG",
  "SEP",
  "OCT",
  "NOV",
  "DEC",
];

const allTaskColumns = [
  { header: "Pkg ID", key: "pkgID", width: 20 },
  { header: "Pkg Serial", key: "pkgSerial", width: 10 },
  { header: "Pkg Task Type", key: "pkgTaskType", width: 12 },
  { header: "Task ID", key: "taskID", width: 23 },
  { header: "Proc Rev", key: "procRev", width: 8 },
  { header: "WO", key: "wo", width: 10 },
  { header: "WO Task", key: "woTask", width: 8 },
  { header: "Task Title", key: "taskTitle", width: 65 },
  { header: "Pkg Task Status", key: "pkgTaskStatus", width: 12 },
];

const partActColumns = [
  { header: "Package", key: "package", width: 20 },
  { header: "Model", key: "model", width: 10 },
  { header: "Type", key: "type", width: 5 },
  { header: "Work Order", key: "workOrder", width: 10 },
  { header: "WO Task", key: "woTask", width: 8 },
  { header: "Task Title", key: "taskTitle", width: 65 },
  { header: "MQ Trx", key: "mqTrx", width: 20 },
  { header: "Part", key: "part", width: 25 },
  { header: "Description", key: "description", width: 30 },
  { header: "Qty", key: "qty", width: 8 },
  { header: "Unit", key: "unit", width: 8 },
  { header: "Avail Status", key: "availStatus", width: 8 },
];

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

async function dropHandler(ev) {
  ev.preventDefault();
  const files = [...ev.dataTransfer.items]
    .map((item) => item.getAsFile())
    .filter((file) => file.name.includes(".xlsx"));
  displayFilesList(files);

  const fileCatalog = [
    ["All Tasks", allTaskColumns],
    ["Part Activity", partActColumns],
  ];
  for (const fileType of fileCatalog) {
    const rows = await filterAndMergeRows(
      classfyFiles(files, fileType[0]),
      fileType[1],
    );
    if (rows.length > 0) {
      await writeRow2File(fileType[0], fileType[0], fileType[1], rows);
    }
  }
}

function classfyFiles(files, filename) {
  console.log(`Select filename contain ${filename}`);
  let targetFiles = [];
  for (const file of files) {
    if (file.name.includes(filename)) {
      targetFiles.push(file);
    }
  }
  console.log(targetFiles);
  return targetFiles;
}

async function filterAndMergeRows(files, columns) {
  const rows = [];

  for (const file of files) {
    console.log(`Trying to open ${file.name}.`);

    const buffer = await file.arrayBuffer();

    const fileWb = new ExcelJS.Workbook();
    await fileWb.xlsx.load(buffer);

    const ws = fileWb.worksheets[0];

    console.log(`Open ${file.name}, ${ws.name}, ${ws.actualRowCount}.`);
    // 先取得每個標題的位子
    
    if (ws.actualRowCount <= 1){
      continue;
    }

    const titleRow = ws.getRow(1);
    
    const titleIndex = columns.map((column) => [
      column.key,
      titleRow.values.indexOf(column.header)
    ]);

    for ( let rowNum = 2; rowNum <= ws.actualRowCount; rowNum++ ){
      const row = ws.getRow(rowNum);
      const newRow = {};
      console.log(`第 ${rowNum} 列:`, row.values);
      titleIndex.forEach(([key, index]) => {
        newRow[key] = (index !== -1 ? row.values[index] : null);
      });
      rows.push(newRow);
    }
    rows.push({}); // 隔開工包
  }
  rows.pop();
  return rows;
}

async function writeRow2File(filename, wbname, titleFormat, rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(wbname, {
    views:[{state:'frozen', ySplit: 1}]
  });

  ws.columns = titleFormat;

  rows.forEach((row) => {
    ws.addRow(row);
  })

  const today = new Date();
  filename = `${filename}_${monthName[today.getMonth()]}-${today.getDate()}-${today.getFullYear()}`;

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");

  a.href = url;
  a.download = filename.toLowerCase().endsWith(".xlsx")
    ? filename
    : `${filename}.xlsx`;

  document.body.appendChild(a);
  a.click();

  a.remove();
  URL.revokeObjectURL(url);
}
