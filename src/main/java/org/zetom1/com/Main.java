package org.zetom1.com;

import org.zetom1.com.facade.ExcelWatermarkPoiFacadeImpl;

public class Main {
    public static void main(String []args){
        System.out.println("This program generate an Excel document with a printable and in document watermark");
        System.out.println("");
        try{
            ExcelWatermarkPoiFacadeImpl.generateDocument();
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }
}
