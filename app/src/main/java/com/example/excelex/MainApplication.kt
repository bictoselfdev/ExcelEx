package com.example.excelex

import android.app.Application
import android.content.Context

class MainApplication : Application() {

    init {
        instance = this

        /************************************************************************************
         * XML Parser 를 위한 System properties,
         * Apache Poi v3.xx 버전 사용 시 해당 properties 설정 하지 않은 경우 FactoryConfigurationError 발생
         * Apache Poi v5.x.x 버전 사용 시 해당 properties 설정을 제외 해도 문제 없었음
         * 정확한 이유를 찾지 못해서 유지
         *************************************************************************************/

        System.setProperty(
            "org.apache.poi.javax.xml.stream.XMLInputFactory",
            "com.fasterxml.aalto.stax.InputFactoryImpl"
        )

        System.setProperty(
            "org.apache.poi.javax.xml.stream.XMLOutputFactory",
            "com.fasterxml.aalto.stax.OutputFactoryImpl"
        )

        System.setProperty(
            "org.apache.poi.javax.xml.stream.XMLEventFactory",
            "com.fasterxml.aalto.stax.EventFactoryImpl"
        )
    }

    companion object {
        lateinit var instance: MainApplication

        fun getContext(): Context {
            return instance.applicationContext
        }
    }
}