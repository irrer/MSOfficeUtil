
package edu.umro.MSOfficeUtil.Excel

import java.io.File
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Cell

/**
 * Read an Excel spreadsheet.
 */
object ReadExcel {

    def cellToString(cell: Cell): String = {
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

    case class CellCoord(sheet: Int, row: Int, col: Int) {
        override def toString = "s: " + sheet + " r:" + row + " c:" + col
    }

    type CellMapT = Map[CellCoord, Cell]

    val emptyCellMap: CellMapT = Map[CellCoord, Cell]()

    /**
     * Read an Excel spreadsheet.  Use the file name extension as a hint, but try to read it in either the
     * older (xls pre-2007) format or the newer format.  If all fails, then return an error message as Left.
     */
    def read(file: File): Either[String, ReadExcel.CellMapT] = {
        if (file.getName.toLowerCase.endsWith(".xlsx")) {
            ReadExcelXSSF.read(file) match {
                case Left(err) => ReadExcelHSSF.read(file)
                case Right(cellList) => Right(cellList)
            }
        }
        else {
            ReadExcelHSSF.read(file) match {
                case Left(err) => ReadExcelXSSF.read(file)
                case Right(cellList) => Right(cellList)
            }
        }
    }

    def main(args: Array[String]): Unit = {
        for (fileName <- List("""D:\tmp\aqa\extract\Summary U of M of Data 2016-06-14.xlsx""", """D:\tmp\aqa\extract\NEW_CODE_VA_TB1314_20151121data_Standard_Dec10th.xls""")) {
            println("\n\nReading file " + fileName)
            val cellMap = read(new File(fileName))
            cellMap.right.get.map(c => println(c._1 + " : " + c._2))
        }
    }

}

