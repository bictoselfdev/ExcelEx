package com.example.excelex.poi

import com.example.excelex.MainApplication
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook


class ExcelCopyFormEx : ExcelBase() {

    companion object {
        const val ASSET_SAMPLE = "sample.xlsx"
    }

    override fun create(): Workbook {
        val inputStream = MainApplication.getContext().assets.open(ASSET_SAMPLE)
        xssfWorkbook = XSSFWorkbook(inputStream) // sample.xlsx 불러 오기

        val sheetA = xssfWorkbook.cloneSheet(0, "홍길동") // copy form
        addContentA(sheetA)

        val sheetB = xssfWorkbook.cloneSheet(0, "홍길순") // copy form
        addContentB(sheetB)

        //xssfWorkbook.removeSheetAt(0) // 양식 삭제

        return xssfWorkbook
    }

    private fun addContentA(sheet: Sheet) {
        setCell(sheet, 12, 1, "마우스")
        setCell(sheet, 12, 4, "ea")
        setCell(sheet, 12, 5, 3)
        setCell(sheet, 12, 6, 125000)
        setCellFormula(sheet, 12, 7, "IF(ISBLANK(F13),\"\",F13*G13)")

        setCell(sheet, 13, 1, "키보드")
        setCell(sheet, 13, 4, "ea")
        setCell(sheet, 13, 5, 2)
        setCell(sheet, 13, 6, 210000)
        setCellFormula(sheet, 13, 7, "IF(ISBLANK(F13),\"\",F13*G13)")

        setCellFormula(sheet, 25, 5, "SUM(F13:F25)")
        setCellFormula(sheet, 25, 6, "SUM(G13:G25)")
        setCellFormula(sheet, 25, 7, "SUM(H13:H25)")
    }

    private fun addContentB(sheet: Sheet) {
        setCell(sheet, 12, 1, "헤드셋")
        setCell(sheet, 12, 4, "ea")
        setCell(sheet, 12, 5, 1)
        setCell(sheet, 12, 6, 99000)
        setCellFormula(sheet, 12, 7, "IF(ISBLANK(F13),\"\",F13*G13)")

        setCell(sheet, 13, 1, "노트북")
        setCell(sheet, 13, 4, "ea")
        setCell(sheet, 13, 5, 2)
        setCell(sheet, 13, 6, 710000)
        setCellFormula(sheet, 13, 7, "IF(ISBLANK(F13),\"\",F14*G14)")

        setCell(sheet, 25, 8, "")
        setCellFormula(sheet, 25, 5, "SUM(F13:F25)")
        setCellFormula(sheet, 25, 6, "SUM(G13:G25)")
        setCellFormula(sheet, 25, 7, "SUM(H13:H25)")
    }
}