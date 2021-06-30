package edu.umro.MSOfficeUtil.Excel
/*
 * Copyright 2021 Regents of the University of Michigan
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.{Cell, Row, Sheet, Workbook}
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io.{File, FileInputStream}

/**
  * Read an Excel spreadsheet.
  *
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
    *
    * @param file File to read.
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
