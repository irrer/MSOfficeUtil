
package edu.umro.MSOfficeUtil.Excel

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.ss.usermodel.CellType
import java.io.File
import java.io.FileInputStream

/**
 * Read the older style *.xls Excel spreadsheet.
 */
object ReadExcelHSSF {

    def getRowList(sheet: HSSFSheet): Seq[HSSFRow] = {
        def getRow(rowIndex: Int): Option[HSSFRow] = {
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

    def cellToString(cell: HSSFCell): String = {

        if (cell == null) "null"
        else {
            cell.getCellTypeEnum match {
                case CellType._NONE =>
                    "none"
                case CellType.NUMERIC =>
                    cell.getNumericCellValue.toString
                case CellType.STRING =>
                    cell.getStringCellValue
                case CellType.FORMULA =>
                    cell.getStringCellValue
                case CellType.BLANK =>
                    "blank"
                case CellType.BOOLEAN =>
                    cell.getBooleanCellValue.toString
                case CellType.ERROR =>
                    "error"
                case _ =>
                    "Unknown cell type"
            }
        }
    }

    def flattenSeqMap[A, B](seq: Seq[Map[A, B]]): Map[A, B] = seq.foldLeft(Map[A, B]())((a, c) => a ++ c)

    def getCellList(sheetIndex: Int, row: HSSFRow): ReadExcel.CellMapT = {
        def getCell(cellIndex: Int): Option[HSSFCell] = {
            try {
                row.getCell(cellIndex) match {
                    case cell: HSSFCell => Some(cell)
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

    def getCellList(sheetIndex: Int, sheet: HSSFSheet): ReadExcel.CellMapT = {
        val seq = getRowList(sheet).map(row => getCellList(sheetIndex, row))
        flattenSeqMap(seq)
    }

    def getCellList(workbook: HSSFWorkbook): ReadExcel.CellMapT = {
        val seq = (0 until workbook.getNumberOfSheets).map(s => getCellList(s, workbook.getSheetAt(s)))
        flattenSeqMap(seq)
    }

    def read(file: File): Either[String, ReadExcel.CellMapT] = {
        try {
            Right(getCellList(new HSSFWorkbook(new FileInputStream(file))))
        }
        catch {
            case t: Throwable => Left("HSSF (pre 2007 *.xlsx) read failed: " + t)
        }
    }

    def main(args: Array[String]): Unit = {

        import java.io.File
        import java.io.FileInputStream

        val fileName = """D:\tmp\aqa\extract\NEW_CODE_VA_TB1314_20151121data_Standard_Dec10th.xls"""

        val fis = new FileInputStream(new File(fileName))

        val workbook = new HSSFWorkbook(fis)

        val cellMap = getCellList(workbook)
        cellMap.map(c => println(c._1 + " : " + c._2))
    }

}
