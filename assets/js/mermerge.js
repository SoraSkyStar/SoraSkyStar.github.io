const dropZone = document.getElementById("drop-zone");
const listFiles = document.getElementById("files-list");

dropZone.addEventListener("drop", dropHandler);

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

async function dropHandler(ev) {
  ev.preventDefault();
  const files = [...ev.dataTransfer.items]
    .map((item) => item.getAsFile())
    .filter((file) => file.name.includes(".xlsx"));
  //   console.log(files);
  displayFilesList(files);
  
  const fileCatalog = ["All Tasks", "Part Activity"]
    for (const element of fileCatalog){
        const rows = await mergeRows(classfyFiles(files, element));
        await writeRow2File(`${element}.xlsx`, element, rows);
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

async function mergeRows(files) {
  const rows = [];

  for (const file of files) {
    console.log(`Trying to open ${file.name}.`);

    const buffer = await file.arrayBuffer();

    const fileWb = new ExcelJS.Workbook();
    await fileWb.xlsx.load(buffer);

    const ws = fileWb.worksheets[0];

    console.log(`Open ${file.name}, ${ws.name}, ${ws.rowCount}.`);

    ws.eachRow((row, rowNumber) => {
      console.log(`第 ${rowNumber} 列:`, row.values);
      rows.push(row.values);
    });
  }

  console.log(rows);

  return rows;
}

async function writeRow2File(filename, wbname, rows) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(wbname);

    ws.addRows(rows);

    const buffer = await wb.xlsx.writeBuffer();

    const blob = new Blob(
        [buffer],
        {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
    );

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