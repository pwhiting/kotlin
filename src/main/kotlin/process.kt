package org.whitings.excel

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream

/**
 * Writes the value "TEST" to the cell at the first row and first column of worksheet.
 */
fun writeToExcelFile(filepath: String) {
    //Instantiate Excel workbook:
    val xlWb = XSSFWorkbook()
    //Instantiate Excel worksheet:
    val xlWs = xlWb.createSheet()

    //Row index specifies the row in the worksheet (starting at 0):
    val rowNumber = 0
    //Cell index specifies the column within the chosen row (starting at 0):
    val columnNumber = 0

    //Write text value to cell located at ROW_NUMBER / COLUMN_NUMBER:
    xlWs.createRow(rowNumber).createCell(columnNumber).setCellValue("TEST")

    //Write file:
    val outputStream = FileOutputStream(filepath)
    xlWb.write(outputStream)
    xlWb.close()
}

/**
 * Reads the value from the cell at the first row and first column of worksheet.
 */
fun readFromExcelFile(filepath: String) {
    val inputStream = FileInputStream(filepath)
    //Instantiate Excel workbook using existing file:
    var xlWb = WorkbookFactory.create(inputStream)

    //Row index specifies the row in the worksheet (starting at 0):
    val rowNumber = 0
    //Cell index specifies the column within the chosen row (starting at 0):
    val columnNumber = 0

    try {
        //Get reference to first sheet:
        val xlWs = xlWb.getSheet("Natalie") ?: throw IllegalArgumentException("no Natalie worksheet in $filepath")

        // read the headers
        for (cell in xlWs.first().iterator()) {
            println("Cell $cell")
        }

        for (row in xlWs.iterator()) {
            println("row");
        }

        println(xlWs.getRow(rowNumber).getCell(columnNumber))
        xlWb.createSheet()

        val fileOut = FileOutputStream("pete.xlsx")
        xlWb.write(fileOut)
        fileOut.close()

    } catch (e: Exception) {
        println("couldn't process $filepath $e")

    }
}

fun main(args: Array<String>) {
    val filepath = "./test.xlsx"
  //  println("Writing file $filepath")
  //  writeToExcelFile(filepath)
    println("Reading file $filepath")
    readFromExcelFile(filepath)
}