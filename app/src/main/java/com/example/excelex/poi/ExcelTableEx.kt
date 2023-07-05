package com.example.excelex.poi

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ExcelTableEx : ExcelBase() {

    override fun create(): Workbook {
        xssfWorkbook = XSSFWorkbook()

        val sheetA = xssfWorkbook.createSheet("6월 1일")
        setDefaultWidth(sheetA, 10)
        setDefaultHeight(sheetA, 450)

        setTitle(sheetA)
        addContent(sheetA)

        val sheetB = xssfWorkbook.createSheet("6월 2일")
        setDefaultWidth(sheetB, 10)
        setDefaultHeight(sheetB, 450)

        setTitle(sheetB)
        addContent(sheetB)

        return xssfWorkbook
    }

    private fun setTitle(sheet: Sheet) {
        setCellWithMerged(sheet, 1, 1, 1, 2, "구분", getTitleStyle())
        setCellWithMerged(sheet, 1, 1, 3, 4, "항목", getTitleStyle())
        setCell(sheet, 1, 5, "목표", getTitleStyle())
        setCell(sheet, 1, 6, "결과", getTitleStyle())
    }

    private fun addContent(sheet: Sheet) {
        setCellWithMerged(sheet, 2, 3, 1, 2, "무산소 운동", getDefaultStyle())
        setCellWithMerged(sheet, 2, 2, 3, 4, "푸시업", getDefaultStyle())
        setCellWithMerged(sheet, 3, 3, 3, 4, "스쿼트", getDefaultStyle())
        setCell(sheet, 2, 5, "10", getDefaultStyle())
        setCell(sheet, 3, 5, "20", getDefaultStyle())
        setCellWithMerged(sheet, 2, 3, 6, 6, "실패", getDefaultStyle())

        setCellWithMerged(sheet, 4, 4, 1, 2, "유산소 운동", getDefaultStyle())
        setCellWithMerged(sheet, 4, 4, 3, 4, "달리기", getDefaultStyle())
        setCell(sheet, 4, 5, "1Km", getDefaultStyle())
        setCell(sheet, 4, 6, "실패", getDefaultStyle())

        setCellWithMerged(sheet, 5, 6, 1, 6, "특이사항 없음", getDefaultStyle())
    }

    private fun getTitleStyle(): XSSFCellStyle {
        val titleStyle = xssfWorkbook.createCellStyle()

        // Font
        val font = xssfWorkbook.createFont()
        font.fontHeightInPoints = 11.toShort()
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

    private fun getDefaultStyle(): XSSFCellStyle {
        val defaultStyle = xssfWorkbook.createCellStyle()

        // Font
        val font = xssfWorkbook.createFont()
        font.fontHeightInPoints = 11.toShort()
        defaultStyle.setFont(font)

        // Gravity
        defaultStyle.setAlignment(HorizontalAlignment.CENTER)
        defaultStyle.setVerticalAlignment(VerticalAlignment.CENTER)

        // Border
        defaultStyle.setBorderBottom(BorderStyle.THIN)
        defaultStyle.setBorderTop(BorderStyle.THIN)
        defaultStyle.setBorderRight(BorderStyle.THIN)
        defaultStyle.setBorderLeft(BorderStyle.THIN)

        return defaultStyle
    }
}