package org.zetom1.com.watemark;

import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.Transparency;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import javax.imageio.ImageIO;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.ImageUtils;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

public class WatermarkGenerator {

    private static final Logger logger = LoggerFactory.getLogger(WatermarkGenerator.class);

    // 1. Uso de Records (Java 14+) para la configuración
    public record WatermarkConfig(String text, int textSize, String font, int width, int height, 
                                  int axisX, int axisY, String format) {
        public static WatermarkConfig defaultConfig() {
            // Corregí el tamaño a 900x900 basado en el comportamiento real de tu código original
            return new WatermarkConfig("CONFIDENTIAL", 80, "Arial", 900, 900, -25, 700, "png");
        }
    }

    public static void addWatermark(XSSFWorkbook workbook, WatermarkConfig config) {
        try {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                insertPrintableWatermark(workbook, workbook.getSheetAt(i), i, config);
            }
            insertWatermarkInSheet(workbook, config);
            logger.info("Watermark applied successfully to {} sheets.", workbook.getNumberOfSheets());
        } catch (Exception e) {
            logger.error("Error applying watermark to workbook", e);
            throw new RuntimeException("Failed to generate watermark", e);
        }
    }

// -------------------------------------------------------------------------
    // NUEVO MÉTODO SOBRECARGADO PARA SXSSF (Big Data)
    // -------------------------------------------------------------------------
    public static void addWatermark(org.apache.poi.xssf.streaming.SXSSFWorkbook sxssfWorkbook, WatermarkConfig config) {
        try {
            // 1. Extraemos el XSSFWorkbook subyacente (¡Justo como recordabas!)
            XSSFWorkbook xssfWorkbook = sxssfWorkbook.getXSSFWorkbook();
            
            for (int i = 0; i < sxssfWorkbook.getNumberOfSheets(); i++) {
                // 2. Extraemos la hoja XSSF pura correspondiente a esta hoja SXSSF
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(i);
                
                // 3. Aplicamos la marca de agua imprimible a la hoja pura
                insertPrintableWatermark(xssfWorkbook, xssfSheet, i, config);
            }
            
            // 4. Inyectamos la imagen repetitiva en el documento base
            insertWatermarkInSheet(xssfWorkbook, config);
            
            logger.info("Marca de agua aplicada exitosamente en formato SXSSF (Streaming).");
        } catch (Exception e) {
            logger.error("Error al aplicar la marca de agua al documento SXSSF", e);
            throw new RuntimeException("Fallo al generar la marca de agua en SXSSF", e);
        }
    }

    private static void insertPrintableWatermark(XSSFWorkbook workbook, XSSFSheet sheet, int sheetIndex, WatermarkConfig config) throws Exception {
        Header header = sheet.getHeader();
        header.setCenter("&G"); // &G means Graphic

        BufferedImage printableWatermark = createWatermarkImage(config);
        
        // 2. Uso de try-with-resources para manejo de memoria
        try (ByteArrayOutputStream imageOs = new ByteArrayOutputStream()) {
            ImageIO.write(printableWatermark, config.format(), imageOs);
            int pictureIdx = workbook.addPicture(imageOs.toByteArray(), Workbook.PICTURE_TYPE_PNG);
            addPrintableWatermark(sheet, pictureIdx, "printableWatermark", sheetIndex, "CH");
        }
    }

    private static void insertWatermarkInSheet(XSSFWorkbook workbook, WatermarkConfig config) throws IOException {
        BufferedImage image = createWatermarkImage(config);
        int numberOfSheets = workbook.getNumberOfSheets();

        try (ByteArrayOutputStream imageOs = new ByteArrayOutputStream()) {
            ImageIO.write(image, config.format(), imageOs);
            int pictureIdx = workbook.addPicture(imageOs.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG);
            XSSFPictureData poixmlDocumentPart = workbook.getAllPictures().get(pictureIdx);
            
            for (int i = 0; i < numberOfSheets; i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                PackagePartName ppn = poixmlDocumentPart.getPackagePart().getPartName();
                String relType = XSSFRelation.IMAGES.getRelation();
                PackageRelationship pr = sheet.getPackagePart().addRelationship(ppn, TargetMode.INTERNAL, relType, null);
                sheet.getCTWorksheet().addNewPicture().setId(pr.getId());
            }
        }
    }

    private static BufferedImage createWatermarkImage(WatermarkConfig config) {
        Font font = new Font(config.font(), Font.PLAIN, config.textSize());
        
        // Limpié la lógica de AWT para que sea más directa
        BufferedImage tempImage = new BufferedImage(1, 1, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = tempImage.createGraphics();
        
        BufferedImage image = g2d.getDeviceConfiguration()
                .createCompatibleImage(config.width(), config.height(), Transparency.TRANSLUCENT);
        g2d.dispose();

        Graphics2D g = image.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.setColor(Color.LIGHT_GRAY);
        g.setFont(font);
        g.shear(0.3, -0.7);
        g.drawString(config.text(), config.axisX(), config.axisY());
        g.dispose();

        return image;
    }

    private static void addPrintableWatermark(XSSFSheet sheet, int pictureIdx,
                                              String pictureTitle, int vmlIdx, String headerPos) throws Exception {
        OPCPackage opcpackage = sheet.getWorkbook().getPackage();
        PackagePartName partname = PackagingURIHelper.createPartName("/xl/drawings/vmlDrawing" + vmlIdx + ".vml");
        PackagePart part = opcpackage.createPart(partname, "application/vnd.openxmlformats-officedocument.vmlDrawing");
        
        VmlDrawing vmldrawing = new VmlDrawing(part);
        XSSFPictureData picData = sheet.getWorkbook().getAllPictures().get(pictureIdx);
        String rIdPic = vmldrawing.addRelation(null, XSSFRelation.IMAGES, picData).getRelationship().getId();

        java.awt.Dimension imageDimension;
        try (ByteArrayInputStream is = new ByteArrayInputStream(picData.getData())) {
            imageDimension = ImageUtils.getImageDimension(is, picData.getPictureType());
        }

        vmldrawing.setRIdPic(rIdPic);
        vmldrawing.setPictureTitle(pictureTitle);
        vmldrawing.setImageDimension(imageDimension);
        vmldrawing.setHeaderPos(headerPos);

        String rIdExtLink = sheet.addRelation(null, XSSFRelation.VML_DRAWINGS, vmldrawing).getRelationship().getId();
        sheet.getCTWorksheet().addNewLegacyDrawingHF().setId(rIdExtLink);
    }

    // 3. Clase Interna para el VML (Mantenida por retrocompatibilidad)
    static class VmlDrawing extends POIXMLDocumentPart {
        String rIdPic = "";
        String pictureTitle = "";
        java.awt.Dimension imageDimension = null;
        String headerPos = "";

        VmlDrawing(PackagePart part) {
            super(part);
        }

        // Setters...
        void setRIdPic(String rIdPic) { this.rIdPic = rIdPic; }
        void setPictureTitle(String pictureTitle) { this.pictureTitle = pictureTitle; }
        void setHeaderPos(String headerPos) { this.headerPos = headerPos; }
        void setImageDimension(java.awt.Dimension imageDimension) { this.imageDimension = imageDimension; }

        @Override
        protected void commit() throws IOException {
            PackagePart part = getPackagePart();
            
            // 4. USO DE TEXT BLOCKS (Java 15+): Adiós a la concatenación infernal
            String xmlContent = """
                <xml xmlns:v="urn:schemas-microsoft-com:vml"
                     xmlns:o="urn:schemas-microsoft-com:office:office"
                     xmlns:x="urn:schemas-microsoft-com:office:excel">
                 <o:shapelayout v:ext="edit">
                  <o:idmap v:ext="edit" data="1"/>
                 </o:shapelayout>
                 <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75"
                  o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
                  <v:stroke joinstyle="miter"/>
                  <v:formulas>
                   <v:f eqn="if lineDrawn pixelLineWidth 0"/>
                   <v:f eqn="sum @0 1 0"/>
                   <v:f eqn="sum 0 0 @1"/>
                   <v:f eqn="prod @2 1 2"/>
                   <v:f eqn="prod @3 21600 pixelWidth"/>
                   <v:f eqn="prod @3 21600 pixelHeight"/>
                   <v:f eqn="sum @0 0 1"/>
                   <v:f eqn="prod @6 1 2"/>
                   <v:f eqn="prod @7 21600 pixelWidth"/>
                   <v:f eqn="sum @8 21600 0"/>
                   <v:f eqn="prod @7 21600 pixelHeight"/>
                   <v:f eqn="sum @10 21600 0"/>
                  </v:formulas>
                  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
                  <o:lock v:ext="edit" aspectratio="t"/>
                 </v:shapetype>
                 <v:shape id="%s" o:spid="_x0000_s1025" type="#_x0000_t75"
                  style='position:absolute;margin-left:0;margin-top:0;width:%dpx;height:%dpx;z-index:1'>
                  <v:imagedata o:relid="%s" o:title="%s"/>
                  <o:lock v:ext="edit" rotation="t"/>
                 </v:shape>
                </xml>
                """.formatted(
                    headerPos, 
                    (int) imageDimension.getWidth(), 
                    (int) imageDimension.getHeight(), 
                    rIdPic, 
                    pictureTitle
                );

            try (OutputStream out = part.getOutputStream()) {
                XmlObject doc = XmlObject.Factory.parse(xmlContent);
                doc.save(out, DEFAULT_XML_OPTIONS);
            } catch (Exception ex) {
                logger.error("Failed to parse and save VML XML", ex);
                throw new IOException("Failed to save VML", ex);
            }
        }
    }
}