package org.zetom1.com.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.zetom1.com.watemark.WatermarkGenerator;
import org.zetom1.com.watemark.WatermarkGenerator.WatermarkConfig;

import java.io.FileOutputStream;
import java.util.Map;
import java.util.TreeMap;

public class ExcelWatermarkPoiSrvImpl {

    private static final Logger logger = LoggerFactory.getLogger(ExcelWatermarkPoiSrvImpl.class);

    // Quitamos el 'static' y pasamos el nombre del archivo como parámetro
    public void generateDocument(String fileName) {
        // try-with-resources principal
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Employee Data");

            populateSheet(sheet);

            // 1. Configuramos y aplicamos la marca de agua
            WatermarkConfig config = WatermarkConfig.defaultConfig();
            WatermarkGenerator.addWatermark(workbook, config);

            // 2. try-with-resources anidado para el FileOutputStream (cierre automático seguro)
            try (FileOutputStream out = new FileOutputStream(fileName)) {
                workbook.write(out);
                logger.info("El archivo {} fue escrito exitosamente en el disco.", fileName);
            }

        } catch (Exception e) {
            logger.error("Error al generar el documento de Excel: ", e);
        }
    }

    private void populateSheet(XSSFSheet sheet) {
        // Usamos el operador diamante <> (Java 7+) y ordenamos por Integer en vez de String
        Map<Integer, Object[]> data = new TreeMap<>();
        data.put(1, new Object[] { "ID", "NAME", "LASTNAME" });
        for (int i = 1; i < 100; i++) {
            data.put(i + 1, new Object[] { i, "Name" + i, "LastName" + i });
        }

        int rownum = 0;
        // Iteramos directamente sobre los valores, ya que no necesitamos la llave (key)
        for (Object[] objArr : data.values()) {
            Row row = sheet.createRow(rownum++);
            int cellnum = 1; 
            
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                
                // Pattern Matching para instanceof (Java 16+)
                if (obj instanceof String strVal) {
                    cell.setCellValue(strVal);
                } else if (obj instanceof Integer intVal) {
                    cell.setCellValue(intVal);
                }
            }
        }
    }
}