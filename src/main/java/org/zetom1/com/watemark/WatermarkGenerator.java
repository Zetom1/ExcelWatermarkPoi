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

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

public class WatermarkGenerator {

    //Text configuration
    private static final String WATERMARKTEXT = "CONFIDENTIAL";
    private static final int TEXTSIZE = 80;
    private static final String FONT = "Arial";

    //Image Configuration
    //Size of image
    private static final int HEIGHT = 100;
    private static final int WIDTH = 100;
    //Image distribution in the sheet
    private static final int AXISX = -25;
    private static final int AXISY = 700;
    private static final String FORMAT = "png";
    //This configuration work only for a document in portrait and the text configuration

    public static void addWatermark(XSSFWorkbook workbook) throws IOException {
        //The iterator is used in VMLDRAWING
        for(int i = 0; i < workbook.getNumberOfSheets(); i++)
            insertPrintableWatermark(workbook,workbook.getSheetAt(i),i);
        insertWatermarkInSheet(workbook);
    }

    private static void insertPrintableWatermark(XSSFWorkbook workbook, XSSFSheet sheet, int n) {

        Header header;
        int pictureIdx;

        header = sheet.getHeader();
        header.setCenter("&G"); // &G means Graphic

        BufferedImage printableWatermark = createWatermarkImage();
        ByteArrayOutputStream imageOs =  new ByteArrayOutputStream();

        try {
            ImageIO.write(printableWatermark,FORMAT,imageOs);
            pictureIdx = workbook.addPicture(imageOs.toByteArray(), Workbook.PICTURE_TYPE_PNG);
            addPrintableWatermark(sheet,pictureIdx,"printableWatermark",n,"CH");

        } catch (IOException e) {
            e.printStackTrace();
        }catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    //Agrega la imagen creada como marca de agua
    private static void insertWatermarkInSheet(XSSFWorkbook workbook) throws IOException{
        BufferedImage image = createWatermarkImage();

        final int numberOfSheets = workbook.getNumberOfSheets();
        ByteArrayOutputStream imageOs = new ByteArrayOutputStream();

        //Carga la imagen generada
        ImageIO.write(image, "png", imageOs);
        int pictureIdx = workbook.addPicture(imageOs.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG);

        XSSFPictureData poixmlDocumentPart = workbook.getAllPictures().get(pictureIdx);
        //Funcionalidad para agregar la marca de agua a cada hoja en caso de que el documento tenga mas de una hoja
        for (int i = 0; i <numberOfSheets; i ++) {// Obitene cada hoja de la tabla
            XSSFSheet sheet = workbook.getSheetAt(i);
            PackagePartName ppn = poixmlDocumentPart.getPackagePart().getPartName();
            String relType = XSSFRelation.IMAGES.getRelation();
            PackageRelationship pr = sheet.getPackagePart().addRelationship(ppn, TargetMode.INTERNAL, relType, null);
            sheet.getCTWorksheet().addNewPicture().setId(pr.getId());
        }
    }

    //Creacion de imagen usando texto
    private static BufferedImage createWatermarkImage(){

        //Caracteristicas del texto
        Font font = new Font(FONT, Font.PLAIN, TEXTSIZE);

        BufferedImage image = new BufferedImage(WIDTH, HEIGHT, BufferedImage.TYPE_INT_RGB);
        //Transparencia para imagen
        Graphics2D g = image.createGraphics();
        image = g.getDeviceConfiguration().createCompatibleImage(900, 900, Transparency.TRANSLUCENT);
        g.dispose();

        //Caracteristicas de la imagen
        g = image.createGraphics();
        g.setColor(Color.LIGHT_GRAY);
        g.setFont(font);
        g.shear(0.3, -0.7);//Coloca la inclinacion de la imagen

        //Suavizado de fondo
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

        //Dibuja el texto
        g.drawString(WATERMARKTEXT, AXISX, AXISY);


        g.dispose();
        return image;
    }

    private static void addPrintableWatermark(XSSFSheet sheet, int pictureIdx,
                                              String pictureTitle, int vmlIdx, String headerPos) throws Exception {
        OPCPackage opcpackage = sheet.getWorkbook().getPackage();

        //creating /xl/drawings/vmlDrawing1.vml
        PackagePartName partname = PackagingURIHelper.createPartName("/xl/drawings/vmlDrawing" + vmlIdx+ ".vml");
        PackagePart part = opcpackage.createPart(partname, "application/vnd.openxmlformats-officedocument.vmlDrawing");
        //creating new VmlDrawing
        WatermarkGenerator.VmlDrawing vmldrawing = new WatermarkGenerator.VmlDrawing(part);

        //creating the relation to the picture in /xl/drawings/_rels/vmlDrawing1.vml.rels
        XSSFPictureData picData = sheet.getWorkbook().getAllPictures().get(pictureIdx);
        String rIdPic = vmldrawing.addRelation(null, XSSFRelation.IMAGES, picData).getRelationship().getId();

        //get image dimension
        ByteArrayInputStream is = new ByteArrayInputStream(picData.getData());
        java.awt.Dimension imageDimension = ImageUtils.getImageDimension(is, picData.getPictureType());
        is.close();

        //updating the VmlDrawing
        vmldrawing.setRIdPic(rIdPic);
        vmldrawing.setPictureTitle(pictureTitle);
        vmldrawing.setImageDimension(imageDimension);
        vmldrawing.setHeaderPos(headerPos);

        //creating the relation to /xl/drawings/vmlDrawing1.xml in /xl/worksheets/_rels/sheet1.xml.rels
        String rIdExtLink = sheet.addRelation(null, XSSFRelation.VML_DRAWINGS, vmldrawing).getRelationship().getId();

        //creating the <legacyDrawingHF r:id="..."/> in /xl/worksheets/sheetN.xml
        sheet.getCTWorksheet().addNewLegacyDrawingHF().setId(rIdExtLink);

    }

    //class for VmlDrawing
    static class VmlDrawing extends POIXMLDocumentPart {

        String rIdPic = "";
        String pictureTitle = "";
        java.awt.Dimension imageDimension = null;
        String headerPos = "";

        VmlDrawing(PackagePart part) {
            super(part);
        }

        void setRIdPic(String rIdPic) {
            this.rIdPic = rIdPic;
        }

        void setPictureTitle(String pictureTitle) {
            this.pictureTitle = pictureTitle;
        }

        void setHeaderPos(String headerPos) {
            this.headerPos = headerPos;
        }

        void setImageDimension(java.awt.Dimension imageDimension) {
            this.imageDimension = imageDimension;
        }

        @Override
        protected void commit() throws IOException {
            PackagePart part = getPackagePart();
            OutputStream out = part.getOutputStream();
            try {
                XmlObject doc = XmlObject.Factory.parse(

                        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\""
                                +" xmlns:o=\"urn:schemas-microsoft-com:office:office\""
                                +" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">"
                                +" <o:shapelayout v:ext=\"edit\">"
                                +"  <o:idmap v:ext=\"edit\" data=\"1\"/>"
                                +" </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\""
                                +"  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">"
                                +"  <v:stroke joinstyle=\"miter\"/>"
                                +"  <v:formulas>"
                                +"   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>"
                                +"   <v:f eqn=\"sum @0 1 0\"/>"
                                +"   <v:f eqn=\"sum 0 0 @1\"/>"
                                +"   <v:f eqn=\"prod @2 1 2\"/>"
                                +"   <v:f eqn=\"prod @3 21600 pixelWidth\"/>"
                                +"   <v:f eqn=\"prod @3 21600 pixelHeight\"/>"
                                +"   <v:f eqn=\"sum @0 0 1\"/>"
                                +"   <v:f eqn=\"prod @6 1 2\"/>"
                                +"   <v:f eqn=\"prod @7 21600 pixelWidth\"/>"
                                +"   <v:f eqn=\"sum @8 21600 0\"/>"
                                +"   <v:f eqn=\"prod @7 21600 pixelHeight\"/>"
                                +"   <v:f eqn=\"sum @10 21600 0\"/>"
                                +"  </v:formulas>"
                                +"  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>"
                                +"  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>"
                                +" </v:shapetype><v:shape id=\"" + headerPos + "\" o:spid=\"_x0000_s1025\" type=\"#_x0000_t75\""
                                +"  style='position:absolute;margin-left:0;margin-top:0;"
                                +"width:" + (int)imageDimension.getWidth() + "px;height:" + (int)imageDimension.getHeight() + "px;"
                                +"z-index:1'>"
                                +"  <v:imagedata o:relid=\""+ rIdPic + "\" o:title=\"" + pictureTitle + "\"/>"
                                +"  <o:lock v:ext=\"edit\" rotation=\"t\"/>"
                                +" </v:shape></xml>"

                );
                doc.save(out, DEFAULT_XML_OPTIONS);
                out.close();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }

    }
}