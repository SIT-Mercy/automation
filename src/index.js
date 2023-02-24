import * as sheetEngine from "sheet-engine/src/index.js"
const path = "sheet.xlsx"
const doc = sheetEngine.parseXlsxDocument(path)
const sheet = doc.name2Sheet.get("16335")
console.log(sheet.rowLength)
console.log(sheet.onRow(16336))
