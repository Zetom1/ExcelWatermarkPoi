package org.zetom1.com.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.zetom1.com.watemark.WatermarkGenerator;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class ExcelWatermarkPoiSrvImpl {

    public static Integer generateDocument() {
        try (//Blank workbook
             XSSFWorkbook workbook = new XSSFWorkbook()) {
            //Create a blank sheet
            XSSFSheet sheet = workbook.createSheet("Employee Data");

            //This data needs to be written (Object[])
            Map<String, Object[]> data = new TreeMap<String, Object[]>();
            data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
            for (int i=1; i<100; i++) {
                data.put(String.valueOf(i + 1), new Object[]{i, "Name" + i, "LastName" + i});
            }
            //Iterate over data and write to sheet
            Set<String> keyset = data.keySet();
            int rownum = 0;
            for (String key : keyset)
            {
                Row row = sheet.createRow(rownum++);
                Object [] objArr = data.get(key);
                int cellnum = 1;
                Cell cell = row.createCell(cellnum++);
                for (Object obj : objArr)
                {
                    cell = row.createCell(cellnum++);
                    if(obj instanceof String)
                        cell.setCellValue((String)obj);
                    else if(obj instanceof Integer)
                        cell.setCellValue((Integer)obj);
                }
            }

            try
            {
                //Write the workbook in file system
                FileOutputStream out = new FileOutputStream(new File("WatermarkExample.xlsx"));
                WatermarkGenerator.addWatermark(workbook);
                workbook.write(out);
                out.close();
                System.out.println("WatermarkExample.xlsx written successfully on disk.");
            }
            catch (Exception e)
            {
                e.printStackTrace();
            }
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return 2;
    }
}
