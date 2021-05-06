/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.uniwue.web;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.spire.doc.FileFormat;
import com.spire.doc.PictureWatermark;
import de.uniwue.algorithm.geometry.regions.type.RegionType;
import static de.uniwue.web.Table.forverseTableCells;
import de.uniwue.web.io.FileDatabase;
import de.uniwue.web.model.Book;
import de.uniwue.web.model.Page;
import de.uniwue.web.model.PageAnnotations;
import de.uniwue.web.model.Point;
import de.uniwue.web.model.Region;
import de.uniwue.web.model.TextLine;
import java.awt.Rectangle;
import java.awt.geom.GeneralPath;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;
import javax.imageio.ImageIO;
import nu.pattern.OpenCV;
import org.apache.http.HttpEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColumns;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;

/**
 *
 * @author alamgir
 */
public class DocxCreator {

    public static HashMap<String, ParagraphAlignment> regionIdToAlignment = new HashMap<>();
    public static String callBackURL = "";
    public static String callBackFileName = "";
    public static String pageName = "77";
    public static String jsonPath = "77.xml";
//        String imagePath = "1.png";
//    public static String bookPath = "/home/alamgir/Desktop/";
//    public static String outPath = "/home/alamgir/Desktop/output";
//    public static String bookName = "output";

    public static String bookPath = "F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\";
    public static String outPath = "F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\Output\\";
    public static String bookName = "Test";

//    public static String bookPath = "D:\\apurba\\projects\\ocr\\LayoutAnalyzer\\src\\main\\resources\\books\\";
//    public static String outPath = "D:\\apurba\\projects\\ocr\\LayoutAnalyzer\\src\\main\\resources\\books\\output\\";
//    public static String bookName = "Test";
    public static void main(String[] args) throws FileNotFoundException, IOException, XmlException {
        OpenCV.loadShared();

        PageLayoutDetection.processImage(bookPath, bookName, pageName);

//        rootPath += File.separator + directoryName + File.separator;
        File jsonFile = new File(bookPath + File.separator + bookName + File.separator + jsonPath);
//        File imageFile = new File(rootPath + imagePath);
        ObjectMapper objectMapper = new ObjectMapper();
        PageAnnotations pageAnnotations = objectMapper.readValue(jsonFile, PageAnnotations.class);

        System.out.println(pageAnnotations.getName());

        create(pageAnnotations, bookPath, bookName, pageName, outPath);
    }

    public static void create(PageAnnotations pageAnnotations, String bookPath, String bookName, String pageName, String outPath) throws IOException {
        DocxCreator.bookPath = bookPath;
        DocxCreator.bookName = bookName;
        DocxCreator.pageName = pageName;
        int bookID = bookName.hashCode();
        FileDatabase database = new FileDatabase(new File(bookPath));
        Page selectedPage = null;

        Book book = database.getBook(bookID);

        for (Page page : book.getPages()) {
            if (page.getName() != null && page.getName().equals(pageName)) {
                selectedPage = page;
                break;
            }
        }
        if (selectedPage == null) {
            System.out.println("Wrong file!!!");
            System.exit(0);
        }

        File imageFile = new File(bookPath + File.separator + selectedPage.getImages().get(0));

        ArrayList<Row> rows = PageLayoutDetection.detect(imageFile, pageAnnotations, bookPath, pageName);
        ComponentStyle componentStyle = new ComponentStyle();
        pageAnnotations = componentStyle.updateComponentAlignment(pageAnnotations, rows, bookPath + File.separator + selectedPage.getImages().get(0));
        for (Map.Entry<String, Region> entry : pageAnnotations.getSegments().entrySet()) {
            Region region = entry.getValue();
            if (region.getTextStyle() != null) {
                ParagraphAlignment paragraphAlignment = ParagraphAlignment.LEFT;
                String alignment = region.getTextStyle().getAlignment();
                if (alignment.equals("LEFT")) {
                    paragraphAlignment = ParagraphAlignment.LEFT;
                } else if (alignment.equals("RIGHT")) {
                    paragraphAlignment = ParagraphAlignment.RIGHT;
                } else if (alignment.equals("CENTER")) {
                    paragraphAlignment = ParagraphAlignment.CENTER;
                }
                regionIdToAlignment.put(region.getId(), paragraphAlignment);
            }
        }

        int lineGap = PageLayoutDetection.lineGap;
        try ( XWPFDocument document = new XWPFDocument()) {
            XWPFStyles styles = document.createStyles();
            CTFonts fonts = CTFonts.Factory.newInstance();
            styles.setDefaultFonts(fonts);
            for (int r = 0; r < rows.size(); r++) {
                XWPFParagraph paragraph = null;
                Row row = rows.get(r);
                int noOfColumns = row.getCols().size() > 1 ? row.getCols().size() : 1;
                if (row.getCols().size() > 1) {
                    for (int j = 0; j < row.getCols().size(); j++) {
                        Col col = row.getCols().get(j);
                        paragraph = document.createParagraph();
                        addColumn(paragraph, col.getRegions(), j);
                    }
                } else {
                    paragraph = document.createParagraph();
                    addColumn(paragraph, row.getRegions(), 0);
                }

                if (r + 1 < rows.size()) {
                    int height = row.getArea().y + row.getArea().height;
                    int height2 = rows.get(r + 1).getArea().y;
                    int interParaGap = height2 - height;
                    int noOfBlankLine = interParaGap / lineGap + (interParaGap % lineGap == 0 ? 0 : 1);
                    XWPFRun rn = paragraph.createRun();
                    for (int i = 0; i < (noOfBlankLine / 2) + 1; i++) {
                        rn.addBreak();
                    }
                }
                CTSectPr ctSectPr = paragraph.getCTP().addNewPPr().addNewSectPr();
                ctSectPr.addNewType().setVal(STSectionMark.CONTINUOUS);
                CTColumns ctColumns = ctSectPr.addNewCols();
                ctColumns.setNum(BigInteger.valueOf(noOfColumns));
                ctColumns.setEqualWidth(STOnOff.ON);
            }

            FileOutputStream out = new FileOutputStream(outPath + File.separator + pageName + ".docx");
            document.write(out);
            out.close();

            String sourcepath = outPath + File.separator + pageName + ".docx";;

            forverseTableCells(sourcepath, sourcepath);

            com.spire.doc.Document doc = new com.spire.doc.Document();
            doc.loadFromFile(outPath + File.separator + pageName + ".docx");

            PictureWatermark picture = new PictureWatermark();
//        picture.setPicture(outPath + File.separator + "choloman.png");
            picture.setPicture(Thread.currentThread().getContextClassLoader().getResourceAsStream("choloman.png"));
            picture.setScaling(50);

            picture.isWashout(false);
            doc.setWatermark(picture);

            doc.saveToFile(outPath + File.separator + pageName + ".docx", FileFormat.Docx);

            if (callBackURL != null && !callBackURL.isEmpty()) {
                File docxFile = new File(outPath + File.separator + pageName + ".docx");
                CloseableHttpClient httpClient = HttpClients.createDefault();
                HttpPost httpPost = new HttpPost(callBackURL);
                MultipartEntityBuilder builder = MultipartEntityBuilder.create();
                String fileName = docxFile.getName();
                if (callBackFileName != null && !callBackFileName.isEmpty()) {
                    fileName = callBackFileName;
                }
                builder.addBinaryBody(
                        "userFile",
                        new FileInputStream(docxFile),
                        ContentType.APPLICATION_OCTET_STREAM,
                        fileName
                );
                HttpEntity multipart = builder.build();
                httpPost.setEntity(multipart);
                CloseableHttpResponse response = httpClient.execute(httpPost);
                HttpEntity responseEntity = response.getEntity();
                String responseString = EntityUtils.toString(responseEntity, StandardCharsets.UTF_8);
                System.out.println(responseString);
            }
        }
    }

    public static void addColumn(XWPFParagraph paragraph, ArrayList<LayoutRegion> segments, int columnIndex) throws IOException {
        Collections.sort(segments, (t, t1) -> {
            return t.y - t1.y;
        });
        boolean columnBreakAdded = false;
        for (Region segment : segments) {
            XWPFRun run = paragraph.createRun();
            if (columnIndex > 0 && !columnBreakAdded) {
                run = paragraph.createRun();
                run.addBreak(BreakType.COLUMN);
                run = paragraph.createRun();
                columnBreakAdded = true;
            }
            run.setFontFamily("SolaimanLipi");
            run.setFontSize(10);

            if (regionIdToAlignment.containsKey(segment.getId())) {
                paragraph.setAlignment(regionIdToAlignment.get(segment.getId()));
            }
            if (segment.getType().equals(RegionType.ImageRegion.toString())) {
                try {
                    GeneralPath clip = new GeneralPath();
                    int index = 0;
                    for (Point point : segment.getPoints()) {
                        if (index <= 0) {
                            clip.moveTo(point.getX(), point.getY());
                        } else {
                            clip.lineTo(point.getX(), point.getY());
                        }

                        index++;
                    }
                    clip.closePath();
                    Rectangle bounds = clip.getBounds();

                    BufferedImage bufferedImage = ImageIO.read(new File(bookPath + File.separator + bookName + File.separator + pageName + ".png"));
                    int imgWidth = bufferedImage.getWidth();
                    int imgHeight = bufferedImage.getHeight();
                    int boundX = (int) bounds.getX() - 5 >= 0 ? (int) bounds.getX() - 5 : 0;
                    int boundY = (int) bounds.getY() - 5 >= 0 ? (int) bounds.getY() - 5 : 0;
                    int boundWidth = (boundX + (int) bounds.getWidth() + 10) < imgWidth ? (int) bounds.getWidth() + 10 : imgWidth - boundX;
                    int boundHeight = (boundY + (int) bounds.getHeight() + 10) < imgHeight ? (int) bounds.getHeight() + 10 : imgHeight - boundY;
                    BufferedImage resultBufferedImage = bufferedImage.getSubimage(boundX, boundY, boundWidth, boundHeight);
//                    BufferedImage resultBufferedImage = bufferedImage.getSubimage((int) bounds.getX() - 5, (int) bounds.getY() - 5, (int) bounds.getWidth() + 10, (int) bounds.getHeight() + 10);
                    ByteArrayOutputStream os = new ByteArrayOutputStream();
                    ImageIO.write(resultBufferedImage, "png", os);                          // Passing: ​(RenderedImage im, String formatName, OutputStream output)
                    InputStream is = new ByteArrayInputStream(os.toByteArray());
                    run.addBreak();
                    run.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, UUID.randomUUID().toString() + ".png", Units.toEMU(bounds.width / 3), Units.toEMU(bounds.height / 3));
                    run.addBreak();
//                    continue;
                } catch (IOException | InvalidFormatException ex) {
                    //Logger.getLogger(DocxCreator.class.getName()).log(Level.SEVERE, null, ex);
                    ex.printStackTrace();
                }

            } else if (segment.getType().equals(RegionType.TableRegion.toString())) {
                XWPFDocument document = paragraph.getDocument();

                TableCreator tableCreator = new TableCreator();
                tableCreator.processTable(segments, document);
            } else {
                int counter = 0;
                for (Map.Entry<String, TextLine> textLine : segment.getTextlines().entrySet()) {
                    for (Map.Entry<Integer, String> text : textLine.getValue().getText().entrySet()) {
                        if (!text.getValue().trim().isEmpty()) {
                            if (counter > 0) {
                                run.addBreak();
                            }
                            counter++;
                            run.setText(text.getValue().trim());
                        }
                    }
                }
            }
        }
    }

    /*public static void forverseTableCells(String sourceFile, String targetFile) throws FileNotFoundException, IOException {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(sourceFile));
        for (XWPFTable table : doc.getTables()) {//Table
            for (XWPFTableRow row : table.getRows()) {//row
                for (XWPFTableCell cell : row.getTableCells()) {//cell: Direct cell.setText() will only add text to the original and delete the text.
                    addBreakInCell(cell);
                }
            }
        }
        FileOutputStream fos = new FileOutputStream(targetFile);
        doc.write(fos);
        fos.close();
        System.out.println("end");
    }*/

    /*public static void addBreakInCell(XWPFTableCell cell) {
        //if (cell.getText() != null && cell.getText().contains("\n")) {
            for (XWPFParagraph p : cell.getParagraphs()) {
                for (XWPFRun run : p.getRuns()) {
                    run.setFontFamily("SolaimanLipi");
                    run.setFontSize(10);
                    if (run.getText(0) != null && run.getText(0).contains("\n")) {
                        String[] lines = run.getText(0).split("\n");
                        if (lines.length > 0) {
                            run.setText(lines[0], 0);
                            for (int i = 1; i < lines.length; i++) {
                                run.addBreak();
                                run.setText(lines[i]);
                            }
                        }
                    }
                }
            }
        //}
    }*/

}
