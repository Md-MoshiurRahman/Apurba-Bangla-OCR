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
import org.apache.poi.xwpf.usermodel.XWPFDocument;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColumns;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import static org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff.Enum.table;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

/**
 *
 * @author alamgir
 */
public class Table {

//    public static HashMap<String, ParagraphAlignment> regionIdToAlignment = new HashMap<>();    
    public static String callBackURL = "";
    public static String callBackFileName = "";
    public static String pageName = "2020_06_08_01";
    public static String jsonPath = "2020_06_08_01.xml";
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

        for (Region segment : pageAnnotations.getSegments().values()) {

            if (segment.getType().equals(RegionType.TableRegion.toString())) {
                Map<String, TextLine> textLines = segment.getTextlines();

                double tableminX = Long.MAX_VALUE;
                double tablemaxX = 0;

                double tableminY = Long.MAX_VALUE;
                double tablemaxY = 0;

                //double width = 0;
                //double height = 0;
                System.out.println("###############In the table region!!#############");
                for (de.uniwue.web.model.Point point : segment.getPoints()) {
                    if (tableminX > point.getX()) {
                        tableminX = point.getX();
                    }
                    if (tablemaxX < point.getX()) {
                        tablemaxX = point.getX();
                    }
                    if (tableminY > point.getY()) {
                        tableminY = point.getY();
                    }
                    if (tablemaxY < point.getY()) {
                        tablemaxY = point.getY();
                    }
                }
                System.out.println("tableminx " + tableminX + " tablemaxX " + tablemaxX + " tableminy " + tableminY + " tablemaxy " + tablemaxY);

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
                        double minnY = Long.MAX_VALUE;
                        double maxxY = 0;

                        double minnX = Long.MAX_VALUE;
                        double maxxX = 0;
                        for (de.uniwue.web.model.Point point : textLine.getValue().getPoints()) {
                            if (minnY > point.getY()) {
                                minnY = point.getY();
                            }
                            if (minnY < tableminY) {
                                minnY = tableminY;
                            }

                            if (maxxY < point.getY()) {
                                maxxY = point.getY();
                            }
                            if (maxxY > tablemaxY) {
                                maxxY = tablemaxY;
                            }

                            if (minnX > point.getX()) {
                                minnX = point.getX();
                            }
                            if (minnX < tableminX) {
                                minnX = tableminX;
                            }

                            if (maxxX < point.getX()) {
                                maxxX = point.getX();
                            }
                            if (maxxX > tablemaxX) {
                                maxxX = tablemaxX;
                            }

                        }

                        for (Map.Entry<Integer, String> text : textLine.getValue().getText().entrySet()) {
                            if (!text.getValue().trim().isEmpty()) {
                                Point pointArray[] = new Point[2];
                                pointArray[0] = new Point(minnX, minnY);
                                pointArray[1] = new Point(maxxX, maxxY);

                                double width = maxxX - minnX;
                                double height = maxxY - minnY;
                                if (width >= 15 && width <= 300 && height <= 70.0 && height >= 15) {
                                    textPoints.put(pointArray, text.getValue().trim());
                                    avgWordWidth += width;
                                    avgWordHeight += height;
                                }
                                //textPoints.put(pointArray, text.getValue().trim());
                                //out.write("minnx " + minnY + " maxxX " + maxxY + " text " + text.getValue().trim() + "\n");
                            }
                        }
                    }
                }
                avgWordWidth = avgWordWidth / textPoints.size();
                avgWordHeight = avgWordHeight / textPoints.size();
                Point tablePoint[] = new Point[2];
                tablePoint[0] = new Point(tableminX, tableminY);
                tablePoint[1] = new Point(tablemaxX, tablemaxY);

                System.out.println("total count " + count + " | point count -> " + textPoints.size());

                createTextFile(textPoints);
                createTable(textPoints, pageWidth, tablePoint, avgWordWidth, avgWordHeight);
                createWordRect(textPoints);

            }
        }
    }

    public static void createTable(HashMap<Point[], String> textPoints, double pageWidth, Point tablePoint[], double avgWordWidth, double avgWordHeight) throws FileNotFoundException, IOException {
        double tableminX = tablePoint[0].getX();
        double tableminY = tablePoint[0].getY();
        double tablemaxX = tablePoint[1].getX();
        double tablemaxY = tablePoint[1].getY();

        System.out.println("avgWordWidth " + avgWordWidth + "  avgWordHeight " + avgWordHeight);

        ArrayList<Point[]> sortedKeyPoints = new ArrayList<Point[]>(textPoints.keySet());
        Collections.sort(sortedKeyPoints, new Comparator<Point[]>() {
            @Override
            public int compare(Point[] p1, Point[] p2) {
                if ((int) Math.abs(p1[0].getY() - p2[0].getY()) > avgWordHeight / 3) {
                    return (int) (p1[0].getY() - p2[0].getY());
                } else {
                    return (int) (p1[0].getX() - p2[0].getX());
                }

            }
        });

        double maxLineGap = 0;
        double minLineGap = Long.MAX_VALUE;
        double avgLineGap = 0;
        int lineGapCount = 0;
        double LineStart = tableminY - 1;
        double lineEnd = tableminY - 1;
        int flag3 = 1;
        for (double i = tableminY; i <= tablemaxY; i++) {
            int flag = 0;
            for (Map.Entry textPoint : textPoints.entrySet()) {
                Point[] point = (Point[]) textPoint.getKey();
                if (point[0].getY() <= i && i <= point[1].getY()) {
                    flag = 1;
                    break;
                }
            }

            if (flag == 1 && flag3 == 1) {
                System.out.println("lineEnd  " + lineEnd + "  lineStart " + LineStart);
                double lineGap = LineStart - lineEnd;
                if (lineEnd != tableminY - 1) {
                    if (lineGap > maxLineGap) {
                        maxLineGap = lineGap;
                    }
                    if (lineGap < minLineGap) {
                        minLineGap = lineGap;
                    }
                    avgLineGap += lineGap;
                    lineGapCount++;
                }

                flag3 = 0;
                lineEnd = 0;
            }
            if (flag == 0 && lineEnd == 0) {
                lineEnd = i;
                flag3 = 1;
            }

            if (flag == 0) {
                LineStart = i;
            }
        }
        avgLineGap = avgLineGap / lineGapCount;
        System.out.println("Max line gap " + maxLineGap);
        System.out.println("Min line gap " + minLineGap);
        System.out.println("Avg line gap " + avgLineGap);

        double maxWordGap = 0;
        double minWordGap = Long.MAX_VALUE;
        double avgWordGap = 0;
        int wordGapCount = 0;
        double wordStart = tableminX - 1;
        double wordEnd = tableminX - 1;
        flag3 = 1;
        for (double i = tableminX; i <= tablemaxX; i++) {
            int flag = 0;
            for (Map.Entry textPoint : textPoints.entrySet()) {
                Point[] point = (Point[]) textPoint.getKey();
                if (point[0].getX() <= i && i <= point[1].getX()) {
                    flag = 1;
                    break;
                }
            }

            if (flag == 1 && flag3 == 1) {
                System.out.println("wordEnd  " + wordEnd + "  wordStart " + wordStart);
                double wordGap = wordStart - wordEnd;
                if (wordEnd != tableminX - 1) {
                    if (wordGap > maxWordGap) {
                        maxWordGap = wordGap;
                    }
                    if (wordGap < minWordGap) {
                        minWordGap = wordGap;
                    }
                    avgWordGap += wordGap;
                    wordGapCount++;
                }

                flag3 = 0;
                wordEnd = 0;
            }
            if (flag == 0 && wordEnd == 0) {
                wordEnd = i;
                flag3 = 1;
            }

            if (flag == 0) {
                wordStart = i;
            }
        }
        avgWordGap = avgWordGap / wordGapCount;
        System.out.println("Max word gap " + maxWordGap);
        System.out.println("Min word gap " + minWordGap);
        System.out.println("Avg word gap " + avgWordGap);

        ArrayList<Double> colValues = new ArrayList<>();
        ArrayList<Double> rowValues = new ArrayList<>();

        double prevColvalue = tableminX - 1;
        //double wordEnd = -1;
        wordEnd = -1;
        boolean flag2 = true;
        for (double i = tableminX; i <= tablemaxX; i++) {
            int flag = 0;
            for (Map.Entry textPoint : textPoints.entrySet()) {
                Point[] point = (Point[]) textPoint.getKey();
                if (point[0].getX() <= i && i <= point[1].getX()) {
                    flag = 1;
                    break;
                }
            }

            if (flag == 1 && !colValues.contains(prevColvalue) && wordEnd != 0) {
                double wordGap = prevColvalue - wordEnd;
                System.out.println("wordEnd  " + wordEnd + "  prev " + prevColvalue);
                if (wordGap >= avgWordGap / 2 || wordGap > 10 || flag2) {
                    //if (wordGap >= 12 || flag2) {
                    flag2 = false;
                    System.out.println("### ### " + wordGap);
                    colValues.add((prevColvalue + wordEnd) / 2);
                }
                wordEnd = 0;
            }
            if (flag == 0 && wordEnd <= 0) {
                wordEnd = i;
                System.out.println("yesss");
            }

            if (flag == 0) {
                prevColvalue = i;
            }
        }

        colValues.add(tablemaxX + 1);

        /*double prevColValue = tableminX - 1;
        for (double i = tableminX; i <= tablemaxX; i++) {
            int flag = 0;
            for (Map.Entry textPoint : textPoints.entrySet()) {
                Point[] point = (Point[]) textPoint.getKey();
                if (point[0].getX() <= i && i <= point[1].getX()) {
                    flag = 1;
                    break;
                }
            }
            if (flag == 1 && !colValues.contains(prevColValue)) {
                colValues.add(prevColValue);
            }
            if (flag == 0) {
                prevColValue = i;
            }
        }
        colValues.add(tablemaxX + 1);*/
        System.out.print("colvalues ");
        for (double colval : colValues) {
            System.out.print(colval + " ");
        }
        System.out.println("\n");

        double prevRowvalue = tableminY - 1;
        //double lineEnd = -1;
        lineEnd = -1;
        //boolean flag2 = true;
        flag2 = true;
        for (double i = tableminY; i <= tablemaxY; i++) {
            int flag = 0;
            for (Map.Entry textPoint : textPoints.entrySet()) {
                Point[] point = (Point[]) textPoint.getKey();
                if (point[0].getY() <= i && i <= point[1].getY()) {
                    flag = 1;
                    break;
                }
            }

            if (flag == 1 && !rowValues.contains(prevRowvalue) && lineEnd != 0) {
                double lineGap = prevRowvalue - lineEnd;
                System.out.println("lineEnd  " + lineEnd + "  prev " + prevRowvalue);

                ArrayList<Double> top = new ArrayList<>();
                for (int j = 1; j < colValues.size(); j++) {
                    double min = Long.MAX_VALUE;
                    for (Point[] point : sortedKeyPoints) {
                        if (prevRowvalue < point[0].getY() && colValues.get(j - 1) < point[0].getX() && point[1].getX() < colValues.get(j)) {
                            if (point[0].getY() < min) {
                                min = point[0].getY();
                            }
                        }
                    }
                    /*if (min == Long.MAX_VALUE) {
                        top.add(prevRowvalue + 3 * avgWordHeight);
                    } else if(min - prevRowvalue < avgWordHeight * 2){
                        top.add(min);
                    }*/
                    if(min - prevRowvalue < avgWordHeight * 2){
                        top.add(min);
                    }
                }
                System.out.println("top " + top);
                double avgStartGap = 0;
                for (int j = 0; j < top.size(); j++) {
                    /*if (top.get(j) - prevRowvalue > avgWordHeight * 3) {
                        top.remove(j);
                    } else {
                        avgStartGap += top.get(j) - prevRowvalue;
                    }*/
                    avgStartGap += top.get(j) - prevRowvalue;
                }
                if(top.size() < Math.round(colValues.size()* 3.0/5.0)){
                    avgStartGap += Long.MAX_VALUE;
                } 
                avgStartGap /= top.size();
                System.out.println("avgStartGap " + avgStartGap);

                //if (lineGap >= maxLineGap * 3 / 8 || flag2) {
                if (lineGap >= avgLineGap * 2 / 3 || (lineGap >= avgLineGap /3 && avgStartGap <= avgWordHeight / 6) || flag2) {
                    flag2 = false;
                    System.out.println("### ### " + lineGap);
                    rowValues.add(prevRowvalue);
                }
                lineEnd = 0;
                //rowValues.add(prevRowvalue);
            }
            if (flag == 0 && lineEnd <= 0) {
                lineEnd = i;
                System.out.println("yesss");
            }

            if (flag == 0) {
                prevRowvalue = i;
            }
        }

        rowValues.add(tablemaxY + 1);
        /*double prevRowvalue = tableminY - 1;

        for (double i = tableminY; i <= tablemaxY; i++) {
            int flag = 0;
            for (Map.Entry textPoint : textPoints.entrySet()) {
                Point[] point = (Point[]) textPoint.getKey();
                if (point[0].getY() <= i && i <= point[1].getY()) {
                    flag = 1;
                    break;
                }
            }

            if (flag == 1 && !rowValues.contains(prevRowvalue)) {
                rowValues.add(prevRowvalue);
            }

            if (flag == 0) {
                prevRowvalue = i;
            }
        }

        rowValues.add(tablemaxY + 1);*/

        System.out.print("rowvalues ");
        for (double rowval : rowValues) {
            System.out.print(rowval + " ");
        }
        System.out.println("\n");

        ArrayList<Double> left = new ArrayList<>();
        ArrayList<Double> right = new ArrayList<>();
        ArrayList<String> colAllignment = new ArrayList<>();

        for (int i = 1; i < colValues.size(); i++) {
            double y = 0;
            for (Point[] point : sortedKeyPoints) {
                if (rowValues.get(1) < point[0].getY() && colValues.get(i - 1) < point[0].getX() && point[1].getX() < colValues.get(i)) {
                    if (point[0].getY() - y > avgWordHeight / 3) {
                        left.add(point[0].getX());
                        y = point[0].getY();
                    }
                }
            }
            System.out.println("left " + left);

            Point[] lastPoint = null;
            y = 0;
            //for (Point[] point : sortedKeyPoints) {
            for (int j = 0; j < sortedKeyPoints.size(); j++) {
                Point[] point = sortedKeyPoints.get(j);
                if (rowValues.get(1) < point[0].getY() && colValues.get(i - 1) < point[0].getX() && point[1].getX() < colValues.get(i)) {

                    if (point[0].getY() - y > avgWordHeight / 3 && y != 0) {
                        right.add(lastPoint[1].getX());
                    }
                    y = point[0].getY();
                    lastPoint = point;
                }
            }
            if (lastPoint != null) {
                right.add(lastPoint[1].getX());
            }
            System.out.println("right " + right);

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
            /*if(avgStartGap <= 10){
                colAllignment.add("LEFT");
            } else*/ if (Math.abs(avgStartGap - avgEndGap) <= avgWordHeight * 2 / 3) {
                colAllignment.add("CENTER");
            } else if (avgStartGap < avgWordHeight / 2) {
                colAllignment.add("LEFT");
            } else if (avgEndGap < avgWordHeight / 2) {
                colAllignment.add("RIGHT");
            } else {
                colAllignment.add("CENTER");
            }

            right.clear();
            left.clear();
        }
        //System.out.println("\n");

        ArrayList<String> headerAllignment = new ArrayList<>();
        for (int j = 1; j < colValues.size(); j++) {
            double y = 0;
            for (Point[] point : sortedKeyPoints) {
                if (rowValues.get(0) < point[0].getY() && point[1].getY() < rowValues.get(1) && colValues.get(j - 1) < point[0].getX() && point[1].getX() < colValues.get(j)) {
                    if (point[0].getY() - y > avgWordHeight / 3) {
                        left.add(point[0].getX());
                        y = point[0].getY();
                    }
                }
            }
            //System.out.println("left " + left);

            Point[] lastPoint = null;
            y = 0;
            //for (Point[] point : sortedKeyPoints) {
            for (int k = 0; k < sortedKeyPoints.size(); k++) {
                Point[] point = sortedKeyPoints.get(k);
                if (rowValues.get(0) < point[0].getY() && point[1].getY() < rowValues.get(1) && colValues.get(j - 1) < point[0].getX() && point[1].getX() < colValues.get(j)) {
                    if (point[0].getY() - y > avgWordHeight / 3 && y != 0) {
                        right.add(lastPoint[1].getX());
                    }
                    y = point[0].getY();
                    lastPoint = point;
                }
            }
            if (lastPoint != null) {
                right.add(lastPoint[1].getX());
            }
            //System.out.println("right " + right);

            double avgStartGap = 0;
            double min;
            if (left.size() >= 2) {
                min = Long.MAX_VALUE;
                for (int i = 0; i < left.size(); i++) {
                    if (left.get(i) < min) {
                        min = left.get(i);
                    }
                }
            } else {
                min = colValues.get(j - 1);
            }
            for (int k = 0; k < left.size(); k++) {
                avgStartGap += left.get(k) - min;
            }
            avgStartGap /= left.size();

            double avgEndGap = 0;
            double max;
            if (right.size() >= 2) {
                max = 0;
                for (int i = 0; i < right.size(); i++) {
                    if (right.get(i) > max) {
                        max = right.get(i);
                    }
                }
            } else {
                max = colValues.get(j);
            }
            for (int k = 0; k < right.size(); k++) {
                avgEndGap += max - right.get(k);
            }
            avgEndGap /= right.size();

            //System.out.println("AvgStartGap " + avgStartGap + "   AvgEndGap " + avgEndGap);
            if (Math.abs(avgStartGap - avgEndGap) <= avgWordHeight * 2 / 3) {
                headerAllignment.add("CENTER");
            } else if (avgStartGap < avgWordHeight / 2) {
                headerAllignment.add("LEFT");
            } else if (avgEndGap < avgWordHeight / 2) {
                headerAllignment.add("RIGHT");
            } else {
                headerAllignment.add("CENTER");
            }

            right.clear();
            left.clear();
        }

        /*ArrayList<ArrayList<String>> allignments = new ArrayList<>();
        for (int i = 1; i < rowValues.size(); i++) {
            ArrayList<Double> left = new ArrayList<>();
            ArrayList<Double> right = new ArrayList<>();
            ArrayList<String> cellAllignment = new ArrayList<>();
            for (int j = 1; j < colValues.size(); j++) {
                double y = 0;
                for (Point[] point : sortedKeyPoints) {
                    if (rowValues.get(i - 1) < point[0].getY() && point[1].getY() < rowValues.get(i) && colValues.get(j - 1) < point[0].getX() && point[1].getX() < colValues.get(j)) {
                        if (point[0].getY() - y > 10) {
                            left.add(point[0].getX());
                            y = point[0].getY();
                        }
                    }
                }
                System.out.println("left " + left);

                Point[] lastPoint = null;
                y = 0;
                //for (Point[] point : sortedKeyPoints) {
                for (int k = 0; k < sortedKeyPoints.size(); k++) {
                    Point[] point = sortedKeyPoints.get(k);
                    if (rowValues.get(i - 1) < point[0].getY() && point[1].getY() < rowValues.get(i) && colValues.get(j - 1) < point[0].getX() && point[1].getX() < colValues.get(j)) {
                        if (point[0].getY() - y > 10 && y != 0) {
                            right.add(lastPoint[1].getX());
                        }
                        y = point[0].getY();
                        lastPoint = point;
                    }
                }
                if (lastPoint != null) {
                    right.add(lastPoint[1].getX());
                }
                System.out.println("right " + right);

                double avgStartGap = 0;
                double min = colValues.get(j - 1);
                for (int k = 0; k < left.size(); k++) {
                    avgStartGap += left.get(k) - min;
                }
                avgStartGap /= left.size();

                double avgEndGap = 0;
                double max = colValues.get(j);
                for (int k = 0; k < right.size(); k++) {
                    avgEndGap += max - right.get(k);
                }
                avgEndGap /= right.size();

                System.out.println("AvgStartGap " + avgStartGap + "   AvgEndGap " + avgEndGap);

                if (Math.abs(avgStartGap - avgEndGap) < 15) {
                    cellAllignment.add("CENTER");
                } else if (avgStartGap < avgEndGap) {
                    cellAllignment.add("LEFT");
                } else {
                    cellAllignment.add("RIGHT");
                }

                right.clear();
                left.clear();
            }
            allignments.add(cellAllignment);
        }*/
        XWPFDocument document = new XWPFDocument();
        //FileOutputStream outDocx = new FileOutputStream(new File(bookPath + "Demo" + File.separator + "table_" + pageName + ".docx"));
        FileOutputStream outDocx = new FileOutputStream(new File(bookPath + File.separator + "table_" + pageName + ".docx"));
        XWPFTable table = document.createTable();
        int tableWidth = (int) Math.round((tablemaxX - tableminX) / pageWidth * 100);
        tableWidth = tableWidth * 3 / 2;
        if (tableWidth > 100) {
            tableWidth = 100;
        }
        table.setWidth(Integer.toString(tableWidth) + "%");
        //table.setWidth("100%");
        //System.out.println(tableWidth + " " + pageWidth);
        setTableAlign(table, ParagraphAlignment.CENTER);
        /*CTTblLayoutType type = table.getCTTbl().getTblPr().addNewTblLayout();
        type.setType(STTblLayoutType.FIXED);*/

 /*CTTblWidth width = table.getCTTbl().addNewTblPr().addNewTblW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(5 * (long) Math.abs(tablemaxX - tableminX)));*/
 /*XWPFTableRow tableRow = table.getRow(0);        
        String s = "";

        int counter = 0;
        for (int i = 1; i < rowValues.size(); i++) {
            for (int j = 1; j < colValues.size(); j++) {
                s = " ";
                Point[] prevPoint = null;
                for (Point[] point : sortedKeyPoints) {
                    if (rowValues.get(i - 1) < point[0].getY() && point[1].getY() < rowValues.get(i) && colValues.get(j - 1) < point[0].getX() && point[1].getX() < colValues.get(j)) {
                        if (prevPoint == null) {
                            s += textPoints.get(point).trim();
                            s += " ";
                            prevPoint = point;
                        } else {
                            if (Math.abs(prevPoint[0].getY() - point[0].getY()) > 10) {
                                s += "\n";
                            }
                            s += textPoints.get(point).trim();
                            s += " ";
                            prevPoint = point;
                        }
                        counter++;
                    }
                }
                //tableRow.getCell(j - 1).setVerticalAlignment(XWPFVertAlign.CENTER);
                tableRow.getCell(j - 1).setText(s);
                if (i == 1 && j < colValues.size() - 1) {
                    tableRow.addNewTableCell().setVerticalAlignment(XWPFVertAlign.CENTER);
                }
            }
            if (i < rowValues.size() - 1) {
                tableRow = table.createRow();
            }

        }
        System.out.println("\ncounter 2 --- " + counter);

        document.write(outDocx);
        outDocx.close();

        String spath = bookPath + File.separator + "table_" + pageName + ".docx";
        String tpath = "F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\table.docx";

        forverseTableCells(spath, spath);*/
        XWPFTableRow tableRow = table.getRow(0);
        //XWPFParagraph paragraph = tableRow.getCell(0).addParagraph();
        //XWPFRun run = paragraph.createRun();        
        String s = "";

        int counter = 0;
        for (int i = 1; i < rowValues.size(); i++) {
            for (int j = 1; j < colValues.size(); j++) {
                s = " ";
                Point[] prevPoint = null;
                for (Point[] point : sortedKeyPoints) {
                    if (rowValues.get(i - 1) < point[0].getY() && point[1].getY() < rowValues.get(i) && colValues.get(j - 1) < point[0].getX() && point[1].getX() < colValues.get(j)) {
                        if (prevPoint == null) {
                            s += textPoints.get(point).trim();
                            s += " ";
                            prevPoint = point;
                        } else {
                            if (Math.abs(prevPoint[0].getY() - point[0].getY()) > 10) {
                                s += "\n";
                            }
                            s += textPoints.get(point).trim();
                            s += " ";
                            prevPoint = point;
                        }
                        counter++;
                    }
                }
                tableRow.getCell(j - 1).removeParagraph(0);
                XWPFParagraph paragraph = tableRow.getCell(j - 1).addParagraph();
                if (i == 1) {
                    if (headerAllignment.get(j - 1).equals("CENTER")) {
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                    } else if (headerAllignment.get(j - 1).equals("LEFT")) {
                        paragraph.setAlignment(ParagraphAlignment.LEFT);
                    } else {
                        paragraph.setAlignment(ParagraphAlignment.RIGHT);
                    }
                } else if (colAllignment.get(j - 1).equals("CENTER")) {
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                } else if (colAllignment.get(j - 1).equals("LEFT")) {
                    paragraph.setAlignment(ParagraphAlignment.LEFT);
                } else {
                    paragraph.setAlignment(ParagraphAlignment.RIGHT);
                }
                XWPFRun run = paragraph.createRun();
                run.setFontFamily("SolaimanLipi");
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
                //tableRow.getCell(j - 1).setText(s);
                if (i == 1 && j < colValues.size() - 1) {
                    tableRow.addNewTableCell();
                }
            }
            if (i < rowValues.size() - 1) {
                tableRow = table.createRow();
            }

        }
        System.out.println("\ncounter 2 --- " + counter);

        document.write(outDocx);
        outDocx.close();

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
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\filename.txt"), "UTF-8"));
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

        Writer out2 = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\new_file.txt"), "UTF-8"));
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

    public static void forverseTableCells(String sourceFile, String targetFile) throws FileNotFoundException, IOException {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(sourceFile));
        for (XWPFTable table : doc.getTables()) {//Table
            int rowIndex = 0;
            for (XWPFTableRow row : table.getRows()) {//row
                int colIndex = 0;
                int totalCellNumber = row.getTableCells().size();
                for (XWPFTableCell cell : row.getTableCells()) {//cell: Direct cell.setText() will only add text to the original and delete the text.
                    addBreakInCell(cell, rowIndex, colIndex, totalCellNumber);
                    colIndex++;
                }
                rowIndex++;
            }
        }
        FileOutputStream fos = new FileOutputStream(targetFile);
        doc.write(fos);
        fos.close();
        System.out.println("end");
    }

    public static void addBreakInCell(XWPFTableCell cell, int rowIndex, int colIndex, int totalCellNumber) {
        // if (cell.getText() != null && cell.getText().contains("\n")) {
        for (XWPFParagraph p : cell.getParagraphs()) {
            if (rowIndex == 0) {
                p.setAlignment(ParagraphAlignment.CENTER);
            }
            for (XWPFRun run : p.getRuns()) {
                run.setFontFamily("SolaimanLipi");
                run.setFontSize(10);
                //run.setVerticalAlignment(ParagraphAlignment.CENTER.toString());
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
    }

    public static void createWordRect(HashMap<Point[], String> textPoints) {
        try {
            String wordOutput = bookPath + "Table_output" + File.separator + pageName + "-word.png";
            String originalImage = bookPath + "Table_output" + File.separator + pageName + ".png";

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

    public static void setTableAlign(XWPFTable table, ParagraphAlignment align) {
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        CTJc jc = (tblPr.isSetJc() ? tblPr.getJc() : tblPr.addNewJc());
        STJc.Enum en = STJc.Enum.forInt(align.getValue());
        jc.setVal(en);
    }

}
