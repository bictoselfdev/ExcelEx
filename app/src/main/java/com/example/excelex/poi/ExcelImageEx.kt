package com.example.excelex.poi

import android.graphics.Bitmap
import android.graphics.BitmapFactory
import com.example.excelex.MainApplication
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ExcelImageEx : ExcelBase() {

    companion object {
        const val ASSET_IMAGE_KOREA = "korea.jpg"
        const val ASSET_IMAGE_USA = "usa.png"
    }

    override fun create(): Workbook {
        xssfWorkbook = XSSFWorkbook()

        val sheetA = xssfWorkbook.createSheet()
        addContent(sheetA)

        return xssfWorkbook
    }

    private fun addContent(sheet: Sheet) {
        setCell(sheet, 1, 1, "한국", getTitleStyle())
        setCellImage(sheet, 1, 2, getBitmapFromAssets(ASSET_IMAGE_KOREA))

        setCell(sheet, 2, 1, "미국", getTitleStyle())
        setCellImage(sheet, 2, 2, getBitmapFromAssets(ASSET_IMAGE_USA))

        setColumnWidth(sheet, 0, 1000)
        setColumnWidth(sheet, 1, 5500)
        setColumnWidth(sheet, 2, 5500)
        setRowHeight(sheet, 0, 50)
        setRowHeight(sheet, 1, 2000)
        setRowHeight(sheet, 2, 2000)
    }

    private fun getBitmapFromAssets(name: String): Bitmap {
        val inputStream = MainApplication.getContext().assets.open(name)
        return BitmapFactory.decodeStream(inputStream)
    }

    private fun getTitleStyle(): XSSFCellStyle {
        val titleStyle = xssfWorkbook.createCellStyle()

        // Font
        val font = xssfWorkbook.createFont()
        font.fontHeightInPoints = 18.toShort()
        font.bold = true
        titleStyle.setFont(font)

        // Gravity
        titleStyle.setAlignment(HorizontalAlignment.CENTER)
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER)

        // Foreground Color
        titleStyle.fillForegroundColor = IndexedColors.LEMON_CHIFFON.index
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)

        // Border
        titleStyle.setBorderBottom(BorderStyle.THIN)
        titleStyle.setBorderTop(BorderStyle.THIN)
        titleStyle.setBorderRight(BorderStyle.THIN)
        titleStyle.setBorderLeft(BorderStyle.THIN)
        titleStyle.bottomBorderColor = IndexedColors.BLACK.getIndex()
        titleStyle.topBorderColor = IndexedColors.BLACK.getIndex()
        titleStyle.rightBorderColor = IndexedColors.BLACK.getIndex()
        titleStyle.leftBorderColor = IndexedColors.BLACK.getIndex()

        return titleStyle
    }
}