package org.zetom1.com;

import org.zetom1.com.service.ExcelWatermarkPoiSrvImpl;

public class Main {
    public static void main(String []args){
        System.out.println("This program generate an Excel document with a printable and in document watermark");
        System.out.println("");
        try{
            ExcelWatermarkPoiSrvImpl.generateDocument();
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }
}
