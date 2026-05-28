package org.zetom1.com;

import org.zetom1.com.service.ExcelWatermarkPoiSrvImpl;

public class Main {
    public static void main(String[] args) {
        System.out.println("Iniciando la generación del documento Excel...");
        
        ExcelWatermarkPoiSrvImpl service = new ExcelWatermarkPoiSrvImpl();
        service.generateDocument("WatermarkExample.xlsx");
    }
}