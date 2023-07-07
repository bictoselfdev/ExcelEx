package com.example.excelex.poi

import android.graphics.Bitmap
import android.graphics.BitmapFactory
import com.example.excelex.MainApplication
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ExcelImageEx : ExcelBase() {

    companion object {
        const val ASSET_IMAGE_CAT = "cat.jpg"
        const val ASSET_IMAGE_DOG = "dog.jpg"
    }

    override fun create(): Workbook {
        xssfWorkbook = XSSFWorkbook()

        val sheetA = xssfWorkbook.createSheet()
        addContent(sheetA)

        return xssfWorkbook
    }

    private fun addContent(sheet: Sheet) {
        setCellImage(sheet, 1, 12, 1, 8, getBitmapFromAssets(ASSET_IMAGE_CAT))
        setCellImage(sheet, 14, 30, 1, 12, getBitmapFromAssets(ASSET_IMAGE_DOG))
    }

    private fun getBitmapFromAssets(name: String): Bitmap {
        val inputStream = MainApplication.getContext().assets.open(name)
        return BitmapFactory.decodeStream(inputStream)
    }
}