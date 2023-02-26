export default class StudentListSheetProvider {
  constructor() {
    this.name = "Student List Reader"
    this.type = "StudentList"
  }

  create(context) {
    return new StudentListSheetLoader()
  }
}

class StudentListSheetLoader {
  /**
   * Return {
   *  studentID: string,
   *  college: string,
   *  name: string
   * }
   */
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
    let students = []
    for (let row = startRow; row <= sheet.rowLength; row++) {
      const college = sheet.at(row, "B")
      const studentID = sheet.at(row, "D")
      const name = sheet.at(row, "E")
      if (college === undefined || studentID === undefined || name === undefined) {
        return
      }
      students.push({
        college,
        studentID,
        name
      })
    }
    return students
  }
}