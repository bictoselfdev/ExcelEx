package com.example.excelex.poi

import android.graphics.Bitmap
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFDrawing
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.FileNotFoundException
import java.io.FileOutputStream
import java.io.IOException
import java.util.Calendar
import java.util.Date

abstract class ExcelBase {

    lateinit var xssfWorkbook: XSSFWorkbook

    abstract fun create(): Workbook // 상속 받아서 입맛에 맞게 재정의 해줘~

    fun saveExcel(workbook: Workbook, path: String) {
        val excelFile = File(path)
        if (excelFile.exists()) excelFile.delete()

        try {
            val fileOut = FileOutputStream(excelFile)
            workbook.write(fileOut)
            fileOut.close()
        } catch (e: FileNotFoundException) {
            e.printStackTrace()
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }

    fun setDefaultWidth(sheet: Sheet, width: Int) {
        // initial value : 8 (2048 / 256)
        sheet.defaultColumnWidth = width
    }

    fun setDefaultHeight(sheet: Sheet, height: Short) {
        // initial value : 300
        sheet.defaultRowHeight = height
    }

    fun setColumnWidth(sheet: Sheet, columnIndex: Int, width: Int) {
        // initial value : 2048 (8 * 256)
        sheet.setColumnWidth(columnIndex, width)
    }

    fun setRowHeight(sheet: Sheet, rowIndex: Int, height: Short) {
        // initial value : 300
        val row = sheet.getRow(rowIndex)
        if (row != null) row.height = height
    }

    fun setCell(sheet: Sheet, rowIndex: Int, columnIndex: Int, value: Any) {
        var row = sheet.getRow(rowIndex)
        if (row == null) row = sheet.createRow(rowIndex)

        var cell = row.getCell(columnIndex)
        if (cell == null) cell = row.createCell(columnIndex)

        when (value) {
            is RichTextString -> cell.setCellValue(value)
            is Int -> cell.setCellValue(value.toDouble())
            is Float -> cell.setCellValue(value.toDouble())
            is Double -> cell.setCellValue(value)
            is Boolean -> cell.setCellValue(value)
            is Date -> cell.setCellValue(value)
            is Calendar -> cell.setCellValue(value)
            else -> cell.setCellValue(value.toString())
        }
    }

    fun setCell(sheet: Sheet, rowIndex: Int, columnIndex: Int, value: Any, style: XSSFCellStyle) {
        var row = sheet.getRow(rowIndex)
        if (row == null) row = sheet.createRow(rowIndex)

        var cell = row.getCell(columnIndex)
        if (cell == null) cell = row.createCell(columnIndex)

        when (value) {
            is RichTextString -> cell.setCellValue(value)
            is Int -> cell.setCellValue(value.toDouble())
            is Float -> cell.setCellValue(value.toDouble())
            is Double -> cell.setCellValue(value)
            is Boolean -> cell.setCellValue(value)
            is Date -> cell.setCellValue(value)
            is Calendar -> cell.setCellValue(value)
            else -> cell.setCellValue(value.toString())
        }
        cell.cellStyle = style
    }

    fun setCellWithMerged(
        sheet: Sheet,
        rowStartIndex: Int,
        rowEndIndex: Int,
        columnStartIndex: Int,
        columnEndIndex: Int,
        value: Any,
        style: XSSFCellStyle
    ) {
        var row = sheet.getRow(rowStartIndex)
        if (row == null) row = sheet.createRow(rowStartIndex)

        var cell = row.getCell(columnStartIndex)
        if (cell == null) cell = row.createCell(columnStartIndex)

        when (value) {
            is RichTextString -> cell.setCellValue(value)
            is Double -> cell.setCellValue(value)
            is Boolean -> cell.setCellValue(value)
            is Date -> cell.setCellValue(value)
            is Calendar -> cell.setCellValue(value)
            else -> cell.setCellValue("$value")
        }
        cell.cellStyle = style

        val cellRangeAddress = CellRangeAddress(rowStartIndex, rowEndIndex, columnStartIndex, columnEndIndex)
        sheet.addMergedRegion(cellRangeAddress)

        RegionUtil.setBorderTop(cell.cellStyle.borderTop, cellRangeAddress, sheet)
        RegionUtil.setBorderBottom(cell.cellStyle.borderBottom, cellRangeAddress, sheet)
        RegionUtil.setBorderLeft(cell.cellStyle.borderLeft, cellRangeAddress, sheet)
        RegionUtil.setBorderRight(cell.cellStyle.borderRight, cellRangeAddress, sheet)
    }

    fun setCellImage(sheet: Sheet, rowStartIndex: Int, rowEndIndex: Int, columnStartIndex: Int, columnEndIndex: Int, bitmap: Bitmap) {

        // Convert bitmap to byte array
        val stream = ByteArrayOutputStream()
        bitmap.compress(Bitmap.CompressFormat.PNG, 100, stream)
        val imageData = stream.toByteArray()

        // Add picture to the workbook. (Picture is assigned unique ID)
        val pictureID = xssfWorkbook.addPicture(imageData, Workbook.PICTURE_TYPE_PNG)

        // Use 'ClientAnchor' to position on sheet.
        val helper = sheet.workbook.creationHelper
        val anchor = helper.createClientAnchor()
        anchor.row1 = rowStartIndex
        anchor.row2 = rowEndIndex
        anchor.setCol1(columnStartIndex)
        anchor.setCol2(columnEndIndex)

        // Attach picture to the sheet.
        val drawing = sheet.createDrawingPatriarch() as XSSFDrawing
        val picture = drawing.createPicture(anchor, pictureID)

        // java.lang.NoClassDefFoundError: Failed resolution of: java/awt/Dimension
         picture.resize()
    }

    fun setCellFormula(sheet: Sheet, rowIndex: Int, columnIndex: Int, formula: String) {
        var row = sheet.getRow(rowIndex)
        if (row == null) row = sheet.createRow(rowIndex)

        var cell = row.getCell(columnIndex)
        if (cell == null) cell = row.createCell(columnIndex)

        cell.cellFormula = formula
    }
}