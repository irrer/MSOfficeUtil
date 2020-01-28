
package edu.umro.MSOfficeUtil.Excel

import java.io.File
import java.io.FileInputStream
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import scala.Left
import scala.Right

/**
 * Read an Excel spreadsheet.
 */
object ExcelUtil {

    def cellToString(cell: Cell): String = {
        if (cell == null) ""
        else cell.toString
    }

    def cellList(row: Row): Seq[Cell] = {
        import scala.collection.JavaConversions._
        row.cellIterator.toSeq
    }

    def rowList(sheet: Sheet): Seq[Row] = {
        import scala.collection.JavaConversions._
        sheet.rowIterator.toSeq
    }

    def sheetList(workbook: Workbook): Seq[Sheet] = {
        import scala.collection.JavaConversions._
        workbook.sheetIterator.toSeq
    }

    private def readHSSF(file: File): Either[String, Workbook] = {
        try {
            Right(new HSSFWorkbook(new FileInputStream(file)).asInstanceOf[Workbook])
        }
        catch {
            case t: Throwable => Left("HSSF (pre 2007 / *.xls) read failed: " + t)
        }
    }

    private def readXSSF(file: File): Either[String, Workbook] = {
        try {
            Right(new XSSFWorkbook(file).asInstanceOf[Workbook])
        }
        catch {
            case t: Throwable => Left("XSSF (post 2007 / *.xlsx) read failed: " + t)
        }
    }

    /**
     * Read an Excel spreadsheet.  Use the file name extension as a hint, but try to read it in either the
     * older (xls pre-2007) format or the newer format.  If all fails, then return an error message as Left.
     */
    def read(file: File): Either[String, Workbook] = {

        if (file.getName.toLowerCase.endsWith(".xlsx")) {
            readXSSF(file) match {
                case Left(err) => readHSSF(file)
                case Right(workbook) => Right(workbook)
            }
        }
        else {
            readHSSF(file) match {
                case Left(err) => readXSSF(file)
                case Right(workbook) => Right(workbook)
            }
        }
    }

    def main(args: Array[String]): Unit = {

        for (fileName <- List("""D:\tmp\aqa\extract\Summary U of M of Data 2016-06-14.xlsx""", """D:\tmp\aqa\extract\NEW_CODE_VA_TB1314_20151121data_Standard_Dec10th.xls""")) {
            println("\n\nReading file " + fileName)
            val cellMap = read(new File(fileName))
            //  cellMap.right.get.map(c => println(c._1 + " : " + c._2))
        }
    }

}

