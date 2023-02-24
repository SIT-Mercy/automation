import * as sheetEngine from "sheet-engine/dist/index.js"
import * as db from "mercydb/dist/db.js"

class StudentListSheetProvider {
  constructor() {
    this.name = "Student List Reader"
  }

  create(context) {
    return StudentListSheetLoader(context.sql)
  }
}

class StudentListSheetLoader {
  constructor(sql) {
    this.sql = sql
  }
  async load(sheet) {
    function findStartRowNumber() {
      for (let num = 1; num <= sheet.rowLength; num++) {
        const row = sheet.rowOn(num)
        if (parseInt(row[0]) !== undefined) {
          return num
        }
      }
      throw new Error("Can't find start row number")
    }
    const startRow = findStartRowNumber()
    for (let row = startRow; row <= sheet.rowLength; row++) {
      const college = sheet.on(row, "B")
      const studentID = sheet.on(row, "D")
      const name = sheet.on(row, "E")
      const student = await db.queryStudentByID(this.sql, studentID)
      if (student === undefined) {
        await db.addStudent(this.sql, {
          studentID: studentID,
          name: name,
          poorLevel: 0,
          currentPoint: 0,
          creationTime: new Date(),
          phoneNumber: student.phoneNumber,
          college: college,
        })
      } else if (student.college !== college || student.name !== name) {
        student.college = college
        student.name = name
        await student.saveChanges()
      }
    }
  }
}

const doc = sheetEngine.parseXlsxDocument(path)
const sheet = doc.name2Sheet.get("16335")
console.log(sheet.rowLength)
console.log(sheet.onRow(16336))
