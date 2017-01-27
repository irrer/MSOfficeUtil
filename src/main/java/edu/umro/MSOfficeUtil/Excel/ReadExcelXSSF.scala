
package edu.umro.MSOfficeUtil.Excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell
import java.io.File

/**
 * Read the newer style *.xlsx Excel spreadsheet.
 */
object ReadExcelXSSF {

    private def getRowList(sheet: XSSFSheet): Seq[XSSFRow] = {
        def getRow(rowIndex: Int): Option[XSSFRow] = {
            try {
                Some(sheet.getRow(rowIndex))
            }
            catch {
                case t: Throwable => None
            }
        }
        val rowList = (sheet.getFirstRowNum to sheet.getLastRowNum).map(r => getRow(r)).flatten
        rowList
    }

    private def flattenSeqMap[A, B](seq: Seq[Map[A, B]]): Map[A, B] = seq.foldLeft(Map[A, B]())((a, c) => a ++ c)

    private def getCellList(sheetIndex: Int, row: XSSFRow): ReadExcel.CellMapT = {
        def getCell(cellIndex: Int): Option[XSSFCell] = {
            try {
                row.getCell(cellIndex) match {
                    case cell: XSSFCell => Some(cell)
                    case _ => None
                }
            }
            catch {
                case t: Throwable => None
            }
        }
        if (row == null) ReadExcel.emptyCellMap
        else {
            val cellList = (row.getFirstCellNum to row.getLastCellNum).map(c => getCell(c)).flatten
            val cellMap = cellList.map(c => (new ReadExcel.CellCoord(sheetIndex, c.getRowIndex, c.getColumnIndex), c)).toMap
            cellMap
        }
    }

    private def getCellList(sheetIndex: Int, sheet: XSSFSheet): ReadExcel.CellMapT = {
        val seq = getRowList(sheet).map(row => getCellList(sheetIndex, row))
        flattenSeqMap(seq)
    }

    private def getCellList(workbook: XSSFWorkbook): ReadExcel.CellMapT = {
        val seq = (0 until workbook.getNumberOfSheets).map(s => getCellList(s, workbook.getSheetAt(s)))
        flattenSeqMap(seq)
    }

    def read(file: File): Either[String, ReadExcel.CellMapT] = {
        try {
            Right(getCellList(new XSSFWorkbook(file)))
        }
        catch {
            case t: Throwable => Left("XSSF (post 2007 *.xlsx) read failed: " + t)
        }
    }

    def main(args: Array[String]): Unit = {

        import java.io.File
        import java.io.FileInputStream

        val fileName = """D:\tmp\aqa\extract\Summary U of M of Data 2016-06-14.xlsx"""

        val workbook = new XSSFWorkbook(new File(fileName))

        val cellMap = getCellList(workbook)
        cellMap.map(c => println(c._1 + " : " + c._2))
    }

}

