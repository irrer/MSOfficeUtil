package edu.umro.MSOfficeUtil.Excel

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.{Cell, Row, Sheet, Workbook}
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io.{File, FileInputStream}

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
    } catch {
      case t: Throwable => Left("HSSF (pre 2007 / *.xls) read failed: " + t)
    }
  }

  private def readXSSF(file: File): Either[String, Workbook] = {
    try {
      Right(new XSSFWorkbook(file).asInstanceOf[Workbook])
    } catch {
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
        case Left(_)         => readHSSF(file)
        case Right(workbook) => Right(workbook)
      }
    } else {
      readHSSF(file) match {
        case Left(_)         => readXSSF(file)
        case Right(workbook) => Right(workbook)
      }
    }
  }

}
