import java.io.FileOutputStream

import org.apache.poi.openxml4j.exceptions.InvalidFormatException
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import org.json._
import scala.io.Source

case class RuleExecution(RuleID : Int, RunOn : String, RunNumber : Int, RunStatus : Boolean)

object JsonToExcelOperation {

  def main (args : Array[String]): Unit = {

    val fileContents = Source.fromFile(args(0)).getLines.mkString
    val jsonParser : JSONArray = new JSONArray(fileContents)

    val columns = List ( "Rule ID", "Run On", "Run number", "Status")

    val workbook : Workbook = new XSSFWorkbook()
    val createHelper : CreationHelper = workbook.getCreationHelper()
    val sheet : Sheet = workbook.createSheet("Employee")

    val headerFont : Font = workbook.createFont()
    headerFont.setBold(true)
    headerFont.setFontHeightInPoints(14.toShort)
    headerFont.setColor(IndexedColors.BLACK.getIndex())


    val headerCellStyle : CellStyle = workbook.createCellStyle()
    headerCellStyle.setFont(headerFont)

    val headerRow : Row = sheet.createRow(0)


    var counter = 0

    columns.foreach( columnName => {
      val cell : Cell = headerRow.createCell(counter)
      cell.setCellValue(columnName)
      cell.setCellStyle(headerCellStyle)
      counter += 1
    })

    var rowNum = 1
    counter = 0

    val redCellStyle : CellStyle = workbook.createCellStyle()
    redCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex())
    redCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)

    val greenCellStyle : CellStyle = workbook.createCellStyle()
    greenCellStyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex())
    greenCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)


    jsonParser.forEach((node) => {

      val jsonObj : JSONObject = node.asInstanceOf[JSONObject]
      val someMap = jsonObj.toMap()

      val row : Row = sheet.createRow(rowNum)

      row.createCell(0, CellType.NUMERIC).setCellValue(someMap.get("RULE_ID").asInstanceOf[Int])
      row.createCell(1, CellType.STRING).setCellValue(someMap.get("RUNON").asInstanceOf[String])
      row.createCell(2, CellType.NUMERIC).setCellValue(someMap.get("RUNNUMBER").asInstanceOf[Int])

      val statusCell : Cell = row.createCell(3, CellType.BOOLEAN)

      val isSuccessful = someMap.get("STATUS").asInstanceOf[Boolean]

      if (isSuccessful == true)
          statusCell.setCellStyle(greenCellStyle)
      else
        statusCell.setCellStyle(redCellStyle)


      statusCell.setCellValue(someMap.get("STATUS").asInstanceOf[Boolean])

      rowNum += 1
    })

    val fileOut : FileOutputStream = new FileOutputStream(args(1))

    workbook.write(fileOut)

    fileOut.close()

    workbook.close()

  }
}