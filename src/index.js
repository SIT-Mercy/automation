import { parseXlsxDocument } from "sheet-engine"
import { db } from "mercydb"

class StudentListSheetProvider {
  constructor() {
    this.name = "Student List Reader"
  }

  create(context) {
    return new StudentListSheetLoader(context.sql)
  }
}

class StudentListSheetLoader {
  constructor(sql) {
    this.sql = sql
  }
  async load(sheet) {
    function findStartRowNumber() {
      for (let num = 1; num <= sheet.rowLength; num++) {
        const row = sheet.onRow(num)
        if (!isNaN(parseInt(row[0]))) {
          return num + 1
        }
      }
      throw new Error("Can't find start row number")
    }
    const startRow = findStartRowNumber()
    for (let row = startRow; row <= sheet.rowLength; row++) {
      const college = sheet.at(row, "B")
      const studentID = sheet.at(row, "D")
      const name = sheet.at(row, "E")
      if (college === undefined || studentID === undefined || name === undefined) {
        return
      }
      const student = await db.queryStudentByID(this.sql, studentID)
      if (!student) {
        await db.addStudent(this.sql, {
          studentID: studentID,
          name: name,
          poorLevel: 0,
          currentPoint: 0,
          creationTime: new Date(),
          phoneNumber: null,
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
const path = ".\\StudentList.xlsx"
const doc = parseXlsxDocument(path)
const sheet = doc.get("16335")
const prov = new StudentListSheetProvider()
import postgres from "postgres"
const sql = postgres({
  host: "localhost",
  port: 5432,
  database: "sit_mercy",
  username: "sit_mercy",
  password: "sit_mercy",
});
const loader = prov.create({
  sql: sql
})
await loader.load(sheet)
await sql.end()