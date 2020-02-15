let Excel = require("exceljs");

const fs = require("fs");

let wb = new Excel.Workbook();
let path = require("path");
let filePath = path.resolve(__dirname, "Spec.xlsx");

wb.xlsx.readFile(filePath).then(function() {
  let sh = wb.getWorksheet("M");

  // writing to cell -> sh.getRow(1).getCell(2).value = 32;
  // writing to file wb.xlsx.writeFile("sample2.xlsx");

  let myArray = [];
  console.log(sh.rowCount);
  for (i = 3; i <= sh.rowCount; i++) {
    myArray.push({
      title: `${sh.getRow(i).getCell(1).value}`,
      M850W: `${sh.getRow(i).getCell(2).value}`,
      M830W: `${sh.getRow(i).getCell(3).value}`,
      M850S: `${sh.getRow(i).getCell(4).value}`,
      M830S: `${sh.getRow(i).getCell(5).value}`,
      M80W: `${sh.getRow(i).getCell(6).value}`,
      M80typeA: `${sh.getRow(i).getCell(7).value}`,
      M80typeB: `${sh.getRow(i).getCell(8).value}`,
      E80typeA: `${sh.getRow(i).getCell(9).value}`,
      E80typeB: `${sh.getRow(i).getCell(10).value}`,
      C80: `${sh.getRow(i).getCell(11).value}`
    });
  }

  // const text = arr.map(JSON.stringify).reduce((prev, next) => `${prev}\n${next}`);
  const arrayName = "SpecificationM";
  let writefilePath = path.resolve(__dirname, "Specification-M.js");

  const myArrayText = myArray.map(JSON.stringify);
  const completeText = `const ${arrayName} = [${myArrayText}]`;
  fs.writeFileSync(writefilePath, completeText, "utf-8");
});
