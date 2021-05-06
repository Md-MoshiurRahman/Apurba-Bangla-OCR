/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.uniwue.web;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.spire.doc.FileFormat;
import com.spire.doc.PictureWatermark;
import de.uniwue.algorithm.geometry.PointList;
import de.uniwue.algorithm.geometry.regions.type.RegionSubType;
import de.uniwue.algorithm.geometry.regions.type.RegionType;
import static de.uniwue.web.PageLayoutDetection.DEBUG;
import de.uniwue.web.io.FileDatabase;
import de.uniwue.web.io.ImageLoader;
import de.uniwue.web.model.Book;
import de.uniwue.web.model.Page;
import de.uniwue.web.model.PageAnnotations;
import de.uniwue.web.model.Point;
import de.uniwue.web.model.Region;
import de.uniwue.web.model.TextLine;
import java.awt.Rectangle;
import java.awt.geom.GeneralPath;
import java.awt.image.BufferedImage;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;
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
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.opencv.core.Mat;
import org.opencv.core.MatOfPoint;
import org.opencv.core.Rect;
import org.opencv.core.Scalar;
import org.opencv.imgcodecs.Imgcodecs;
import org.opencv.imgproc.Imgproc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColumns;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumbering;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import static org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff.Enum.table;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

/**
 *
 * @author alamgir
 */
public class List {

//    public static HashMap<String, ParagraphAlignment> regionIdToAlignment = new HashMap<>();    
    public static String callBackURL = "";
    public static String callBackFileName = "";
    public static String pageName = "79";
    public static String jsonPath = "79.xml";
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
        File jsonFile = new File(bookPath + File.separator + bookName + File.separator + jsonPath);
        ObjectMapper objectMapper = new ObjectMapper();
        PageAnnotations pageAnnotations = objectMapper.readValue(jsonFile, PageAnnotations.class);
        System.out.println(pageAnnotations.getName());

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
        double pageWidth = pageAnnotations.getWidth();

        XWPFDocument document = new XWPFDocument();
        FileOutputStream outDocx = new FileOutputStream(new File(bookPath + "List" + File.separator + "list_" + pageName + ".docx"));
        for (Region segment : pageAnnotations.getSegments().values()) {
            /*if (segment.getType().toString().equals("paragraph")) {
                System.out.println(segment.getType().toString());
            }*/

            if (segment.getType().toString().equals("paragraph")) {
                //Map<String, TextLine> textLines = segment.getTextlines();

                HashMap<Point[], String> textPoints = new HashMap<>();
                int count = 0;
                double avgWordWidth = 0;
                double avgWordHeight = 0;

                for (Map.Entry<String, TextLine> textLine : segment.getTextlines().entrySet()) {
                    boolean emptyLine = true;
                    for (Map.Entry<Integer, String> text : textLine.getValue().getText().entrySet()) {
                        if (!text.getValue().trim().isEmpty()) {
                            emptyLine = false;
                        }
                    }

                    if (!emptyLine) {
                        count++;
                        double minY = Long.MAX_VALUE;
                        double maxY = 0;

                        double minX = Long.MAX_VALUE;
                        double maxX = 0;
                        for (de.uniwue.web.model.Point point : textLine.getValue().getPoints()) {
                            if (minY > point.getY()) {
                                minY = point.getY();
                            }

                            if (maxY < point.getY()) {
                                maxY = point.getY();
                            }

                            if (minX > point.getX()) {
                                minX = point.getX();
                            }

                            if (maxX < point.getX()) {
                                maxX = point.getX();
                            }
                        }

                        for (Map.Entry<Integer, String> text : textLine.getValue().getText().entrySet()) {
                            if (!text.getValue().trim().isEmpty()) {
                                Point pointArray[] = new Point[2];
                                pointArray[0] = new Point(minX, minY);
                                pointArray[1] = new Point(maxX, maxY);

                                double width = maxX - minX;
                                double height = maxY - minY;
                                textPoints.put(pointArray, text.getValue().trim());
                                avgWordWidth += width;
                                avgWordHeight += height;
                                /*if (width >= 15 && width <= 300 && height <= 70.0 && height >= 15) {
                                    textPoints.put(pointArray, text.getValue().trim());
                                    avgWordWidth += width;
                                    avgWordHeight += height;
                                }*/
                                //textPoints.put(pointArray, text.getValue().trim());
                                //out.write("minnx " + minnY + " maxxX " + maxxY + " text " + text.getValue().trim() + "\n");
                            }
                        }
                    }
                }
                avgWordWidth = avgWordWidth / textPoints.size();
                avgWordHeight = avgWordHeight / textPoints.size();

                System.out.println("total count " + count + " | point count -> " + textPoints.size());

                //createTextFile(textPoints);
                createList(document, textPoints, pageWidth, avgWordWidth, avgWordHeight);
                //createWordRect(textPoints);

            }
        }
        document.write(outDocx);
        outDocx.close();
    }

    public static void createList(XWPFDocument document, HashMap<Point[], String> textPoints, double pageWidth, double avgWordWidth, double avgWordHeight) throws FileNotFoundException, IOException {
        ArrayList<String> texts = new ArrayList<>();
        for (Map.Entry textPoint : textPoints.entrySet()) {
            String text = (String) textPoint.getValue();
            texts.add(text);
        }

        ArrayList<Point[]> sortedKeyPoints = new ArrayList<Point[]>(textPoints.keySet());
        Collections.sort(sortedKeyPoints, new Comparator<Point[]>() {
            @Override
            public int compare(Point[] p1, Point[] p2) {
                return (int) (p1[0].getY() - p2[0].getY());
            }
        });

        ArrayList<Double> left = new ArrayList<>();
        ArrayList<Double> right = new ArrayList<>();

        for (Point[] point : sortedKeyPoints) {
            left.add(point[0].getX());
            right.add(point[0].getY());
        }

        double avgStartGap = 0;
        double min = Long.MAX_VALUE;
        for (int j = 0; j < left.size(); j++) {
            if (left.get(j) < min) {
                min = left.get(j);
            }
        }
        for (int j = 0; j < left.size(); j++) {
            avgStartGap += left.get(j) - min;
        }
        avgStartGap /= left.size();

        double avgEndGap = 0;
        double max = 0;
        for (int j = 0; j < right.size(); j++) {
            if (right.get(j) > max) {
                max = right.get(j);
            }
        }
        for (int j = 0; j < right.size(); j++) {
            avgEndGap += max - right.get(j);
        }
        avgEndGap /= right.size();

        System.out.println("AvgStartGap " + avgStartGap + "   AvgEndGap " + avgEndGap);

        XWPFParagraph paragraph;
        XWPFRun run;
        XWPFStyles styles = document.createStyles();
        CTFonts fonts = CTFonts.Factory.newInstance();
        fonts.setAscii("Tanmatra Internet");
        styles.setDefaultFonts(fonts);
        /*XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("SolaimanLipi");
        run.setFontSize(10);*/

 /*if (avgStartGap <= 2 && avgEndGap >= 100) {
            for (Point[] point : sortedKeyPoints) {
                run.setText(textPoints.get(point).trim());
                run.addBreak();
            }
        }*/

 /*XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("SolaimanLipi");
        run.setFontSize(10);
        run.setText("The list:");

        ArrayList<String> documentList = new ArrayList<String>(
                Arrays.asList(
                        new String[]{
                            "One",
                            "Two",
                            "Three"
                        }));*/
        CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
        //Next we set the AbstractNumId. This requires care. 
        //Since we are in a new document we can start numbering from 0. 
        //But if we have an existing document, we must determine the next free number first.
        cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));
        char[] banglaDigits = {'?', '?', '?', '?', '?', '?', '?', '?', '?', '?'};

        ///*Bullet list
        /*CTLvl cTLvl = cTAbstractNum.addNewLvl();
        cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
        cTLvl.addNewLvlText().setVal("•");*/
        //cTLvl.addNewLvlText().setVal("");
        ///* Decimal list
        CTLvl cTLvl = cTAbstractNum.addNewLvl();
        cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
        cTLvl.addNewLvlText().setVal("%1.");
        cTLvl.addNewStart().setVal(BigInteger.valueOf(1));

        XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
        XWPFNumbering numbering = document.createNumbering();
        BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
        BigInteger numID = numbering.addNum(abstractNumID);

        /*for (String string : texts) {
            paragraph = document.createParagraph();
            paragraph.setNumID(numID);
            run = paragraph.createRun();
            run.setFontFamily("SolaimanLipi");
            run.setFontSize(10);
            run.setText(string);
        }*/
        if (avgStartGap <= 170 && avgEndGap >= 200) {
            createWordRect(textPoints);
            Point[] prevPoint = sortedKeyPoints.get(0);
            String s = textPoints.get(prevPoint).trim();
            for (int i = 1; i < sortedKeyPoints.size(); i++) {
                Point[] point = sortedKeyPoints.get(i);
                if (point[0].getX() - min > 10 && point[0].getY() - prevPoint[0].getY() > 10) {
                    s += "\n";
                    s += "\t";
                    s += textPoints.get(point).trim();
                } else if (point[0].getX() - min > 10) {
                    s += " ";
                    s += textPoints.get(point).trim();
                } else {
                    paragraph = document.createParagraph();
                    run = paragraph.createRun();
                    run.setFontFamily("SolaimanLipi");
                    paragraph.setNumID(numID);
                    run.setFontSize(10);
                    run.setText(s);
                    if (s != null && s.contains("\n")) {
                        String[] lines = run.getText(0).split("\n");
                        if (lines.length > 0) {
                            run.setText(lines[0], 0);
                            for (int k = 1; k < lines.length; k++) {
                                run.addBreak();
                                run.setText(lines[k]);
                            }
                        }
                    }

                    s = textPoints.get(point).trim();
                }
                prevPoint = point;
            }
            paragraph = document.createParagraph();
            run = paragraph.createRun();
            run.setFontFamily("SolaimanLipi");
            paragraph.setNumID(numID);
            run.setFontSize(10);
            run.setText(s);
            if (s != null && s.contains("\n")) {
                String[] lines = run.getText(0).split("\n");
                if (lines.length > 0) {
                    run.setText(lines[0], 0);
                    for (int k = 1; k < lines.length; k++) {
                        run.addBreak();
                        run.setText(lines[k]);
                    }
                }
            }
        }
    }

    public static void createTextFile(HashMap<Point[], String> textPoints) throws FileNotFoundException, UnsupportedEncodingException, IOException {
        ArrayList<Point[]> sortedKeys = new ArrayList<Point[]>(textPoints.keySet());
        Collections.sort(sortedKeys, new Comparator<Point[]>() {
            @Override
            public int compare(Point[] p1, Point[] p2) {
                if ((int) Math.abs(p1[0].getY() - p2[0].getY()) > 7) {
                    return (int) (p1[0].getY() - p2[0].getY());
                } else {
                    return (int) (p1[0].getX() - p2[0].getX());
                }

            }
        });

        /*Collections.sort(sortedKeys, new Comparator<Point[]>() {
            @Override
            public int compare(Point[] p1, Point[] p2) {
                return (int) (p1[1].getX() - p2[1].getX());

            }
        });*/

 /*Collections.sort(sortedKeys, new Comparator<Point[]>() {
            @Override
            public int compare(Point[] p1, Point[] p2) {
                if ((int) Math.abs(p1[1].getY() - p2[0].getY()) > 10) {
                    return (int) (p1[0].getY() - p2[0].getY());
                } else {
                    if ((int) (p1[1].getX() - p2[0].getX()) > 0) {
                        return (int) (p1[0].getY() - p2[0].getY());
                    } else {
                        return (int) (p1[0].getX() - p2[0].getX());
                    }

                }

            }
        });*/
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\List\\filename.txt"), "UTF-8"));
        Point[] refPoint = sortedKeys.get(0);
        for (Point[] p : sortedKeys) {
            out.write("minnx " + p[0].getX() + " maxxX " + p[1].getX() + " minnY " + p[0].getY() + " maxxY " + p[1].getY() + " text " + textPoints.get(p).trim() + "\n");
            /*if (Math.abs(refPoint[0].getY() - p[0].getY()) <= 15) {
                out.write(textPoints.get(p).trim() + " ");
            } else {
                refPoint = p;
                out.write("\n" + textPoints.get(p).trim() + " ");
            }*/
        }
        out.close();

        Writer out2 = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\List\\new_file.txt"), "UTF-8"));
        Point[] prevPoint = sortedKeys.get(0);
        for (Point[] p : sortedKeys) {
            if (p[0].getY() - prevPoint[1].getY() > 0) {
                out2.write("prev-> " + textPoints.get(prevPoint).trim() + "   current-> " + textPoints.get(p).trim() + "   diff " + Math.abs(p[0].getY() - prevPoint[1].getY()) + "\n");
            }
            prevPoint = p;
            //out.write("minnx " + p[0].getX() + " maxxX " + p[1].getX() + " minnY " + p[0].getY() + " maxxY " + p[1].getY() + " text " + textPoints.get(p).trim() + "\n");
            /*if (Math.abs(refPoint[0].getY() - p[0].getY()) <= 15) {
                out.write(textPoints.get(p).trim() + " ");
            } else {
                refPoint = p;
                out.write("\n" + textPoints.get(p).trim() + " ");
            }*/
        }
        out2.close();

    }

    public static void createWordRect(HashMap<Point[], String> textPoints) {
        try {
            String wordOutput = bookPath + "List\\List_output" + File.separator + pageName + "-word.png";
            String originalImage = bookPath + "List\\List_output" + File.separator + pageName + ".png";

            //String wordOutput = bookPath + "Demo" + File.separator + pageName + "-word.png";
            //String originalImage = bookPath + "Demo" + File.separator + pageName + ".png";
            /**
             * Line height detection
             */
            Mat newimg = ImageLoader.readOriginal(new File(originalImage));

            ArrayList<Point[]> points = new ArrayList<Point[]>(textPoints.keySet());
            ArrayList<Rect> rectangles = new ArrayList<>();
            for (Point[] p : points) {
                int left = (int) p[0].getX();
                int top = (int) p[0].getY();
                int right = (int) p[1].getX();
                int bottom = (int) p[1].getY();

                org.opencv.core.Point tl = new org.opencv.core.Point(left, top);
                org.opencv.core.Point br = new org.opencv.core.Point(right, bottom);
                Rect rect = new Rect(tl, br);
                rectangles.add(rect);
            }

            ArrayList<Rect> temp = new ArrayList<>();
            for (Rect rectangle : rectangles) {
                Imgproc.rectangle(newimg, rectangle.tl(), rectangle.br(), new Scalar(0, 255, 0));
            }
            if (DEBUG) {
                Imgcodecs.imwrite(wordOutput, newimg);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

}
