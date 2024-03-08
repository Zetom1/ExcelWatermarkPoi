package org.zetom1.com.facade;

import com.sun.istack.internal.NotNull;
import org.zetom1.com.service.ExcelWatermarkPoiSrvImpl;

public class ExcelWatermarkPoiFacadeImpl {


    @NotNull
    public static void generateDocument() {
        ExcelWatermarkPoiSrvImpl.generateDocument();
    }
}
