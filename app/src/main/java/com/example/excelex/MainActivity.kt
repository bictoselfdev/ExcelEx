package com.example.excelex

import android.os.Bundle
import androidx.appcompat.app.AppCompatActivity
import androidx.databinding.DataBindingUtil
import com.example.excelex.databinding.ActivityMainBinding
import com.example.excelex.poi.ExcelCopyFormEx
import com.example.excelex.poi.ExcelImageEx
import com.example.excelex.poi.ExcelTableEx

class MainActivity : AppCompatActivity() {

    private lateinit var binding: ActivityMainBinding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = DataBindingUtil.setContentView(this, R.layout.activity_main)

        binding.btnTableEx.setOnClickListener {
            val tableEx = ExcelTableEx()
            val workBook = tableEx.create()
            tableEx.saveExcel(workBook, "/sdcard/Download/tableEx.xlsx")
        }

        binding.btnCopyFormEx.setOnClickListener {
            val copyFormEx = ExcelCopyFormEx()
            val workBook = copyFormEx.create()
            copyFormEx.saveExcel(workBook, "/sdcard/Download/copyFormEx.xlsx")
        }

        binding.btnImageEx.setOnClickListener {
            val imageEx = ExcelImageEx()
            val workBook = imageEx.create()
            imageEx.saveExcel(workBook, "/sdcard/Download/imageEx.xlsx")
        }
    }
}