package de.uniwue.web;

import com.fasterxml.jackson.databind.ObjectMapper;
import de.uniwue.algorithm.geometry.regions.type.RegionType;
import static de.uniwue.web.Table.bookName;
import static de.uniwue.web.Table.bookPath;
//import static de.uniwue.web.Table.createTable_2;
import static de.uniwue.web.Table.forverseTableCells;
import static de.uniwue.web.Table.jsonPath;
import static de.uniwue.web.Table.pageName;
import de.uniwue.web.io.FileDatabase;
import de.uniwue.web.model.Book;
import de.uniwue.web.model.Page;
import de.uniwue.web.model.PageAnnotations;
import de.uniwue.web.model.Point;
import de.uniwue.web.model.Region;
import de.uniwue.web.model.TextLine;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;
import nu.pattern.OpenCV;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

public class TableCreator {

    public static String callBackURL = "";
    public static String callBackFileName = "";
    public static String pageName = "95";
    public static String jsonPath = "95.xml";

    public static String bookPath = "F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\";
    public static String outPath = "F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\Output\\";
    public static String bookName = "Test";

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
        processTable(pageAnnotations);

    }

    public static void processTable(PageAnnotations pageAnnotations) throws IOException {
        for (Region segment : pageAnnotations.getSegments().values()) {

            if (segment.getType().equals(RegionType.TableRegion.toString())) {
                Map<String, TextLine> textLines = segment.getTextlines();

                double tableminX = Long.MAX_VALUE;
                double tablemaxX = 0;

                double tableminY = Long.MAX_VALUE;
                double tablemaxY = 0;

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
                                if (width <= 300 && height <= 70.0 && height >= 15) {
                                    textPoints.put(pointArray, text.getValue().trim());
                                }
                                //textPoints.put(pointArray, text.getValue().trim());
                                //out.write("minnx " + minnY + " maxxX " + maxxY + " text " + text.getValue().trim() + "\n");
                            }
                        }
                    }
                }
                XWPFDocument document = new XWPFDocument();
                //XWPFRun run = new
                //createTable_2(document, textPoints, tableminX, tableminY, tablemaxX, tablemaxY);
                System.out.println("total count " + count + " | point count -> " + textPoints.size());

                //createTextFile(textPoints);
                //createTable(textPoints, tableminX, tableminY, tablemaxX, tablemaxY);
                //createWordRect(textPoints);
            }
        }
    }

    public static void processTable(ArrayList<LayoutRegion> segments, XWPFDocument document) throws IOException {
        for (Region segment : segments) {
            if (segment.getType().equals(RegionType.TableRegion.toString())) {
                Map<String, TextLine> textLines = segment.getTextlines();

                double tableminX = Long.MAX_VALUE;
                double tablemaxX = 0;

                double tableminY = Long.MAX_VALUE;
                double tablemaxY = 0;

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
                                if (width <= 300 && height <= 70.0 && height >= 15) {
                                    textPoints.put(pointArray, text.getValue().trim());
                                }
                                //textPoints.put(pointArray, text.getValue().trim());
                                //out.write("minnx " + minnY + " maxxX " + maxxY + " text " + text.getValue().trim() + "\n");
                            }
                        }
                    }
                }

                System.out.println("total count " + count + " | point count -> " + textPoints.size());
                createTable(document, textPoints, tableminX, tableminY, tablemaxX, tablemaxY);
            }

        }
    }

    public static void createTable(XWPFDocument document, HashMap<Point[], String> textPoints, double tableminX, double tableminY, double tablemaxX, double tablemaxY) throws FileNotFoundException, IOException {
        ArrayList<Point[]> sortedKeyPoints = new ArrayList<Point[]>(textPoints.keySet());
        Collections.sort(sortedKeyPoints, new Comparator<Point[]>() {
            @Override
            public int compare(Point[] p1, Point[] p2) {
                if ((int) Math.abs(p1[0].getY() - p2[0].getY()) >= 6) {
                    return (int) (p1[0].getY() - p2[0].getY());
                } else {
                    return (int) (p1[0].getX() - p2[0].getX());
                }

            }
        });

        double maxLineGap = 0;
        double avgLineGap = 0;
        int lineGapCount = 0;
        if (sortedKeyPoints.size() != 0) {
            Point[] prevLinePoint = sortedKeyPoints.get(0);

            for (Point[] p : sortedKeyPoints) {
                double lineGap = p[0].getY() - prevLinePoint[1].getY();
                if (lineGap > 0) {
                    if (lineGap > maxLineGap) {
                        maxLineGap = lineGap;
                    }
                    avgLineGap += lineGap;
                    lineGapCount++;
                }
                prevLinePoint = p;
            }
            avgLineGap = avgLineGap / lineGapCount;
        }

        System.out.println("Max line gap " + maxLineGap);
        System.out.println("Avg line gap " + avgLineGap);

        ArrayList<Double> colValues = new ArrayList<>();
        ArrayList<Double> rowValues = new ArrayList<>();

        double prevColvalue = tableminX - 1;
        double wordEnd = -1;
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
                //if (lineGap >= maxLineGap * 3 / 8 || flag2) {
                if (wordGap >= 12 || flag2) {
                    flag2 = false;
                    System.out.println("### ### " + wordGap);
                    colValues.add(prevColvalue);
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
        System.out.print("colvalues ");
        for (double colval : colValues) {
            System.out.print(colval + " ");
        }
        System.out.println("\n");

        double prevRowvalue = tableminY - 1;
        double lineEnd = -1;
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
                //if (lineGap >= maxLineGap * 3 / 8 || flag2) {
                if (lineGap >= 5 || flag2) {
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
                if (colValues.get(i - 1) < point[0].getX() && point[1].getX() < colValues.get(i)) {
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
            for (int j = 0; j < sortedKeyPoints.size(); j++) {
                Point[] point = sortedKeyPoints.get(j);
                if (colValues.get(i - 1) < point[0].getX() && point[1].getX() < colValues.get(i)) {

                    if (point[0].getY() - y > 10 && y != 0) {
                        right.add(lastPoint[1].getX());
                    }
                    y = point[0].getY();
                    lastPoint = point;
                }
            }
            right.add(lastPoint[1].getX());
            System.out.println("right " + right);

            double avgStartGap = 0;
            double max = 0;
            for (int j = 1; j < left.size(); j++) {
                if (left.get(j) > max) {
                    max = left.get(j);
                }
            }
            for (int j = 1; j < left.size(); j++) {
                avgStartGap += max - left.get(j);
            }
            avgStartGap /= left.size() - 1;

            double avgEndGap = 0;
            double min = Long.MAX_VALUE;
            for (int j = 1; j < right.size(); j++) {
                if (right.get(j) < min) {
                    min = right.get(j);
                }
            }
            for (int j = 1; j < right.size(); j++) {
                avgEndGap += right.get(j) - min;
            }
            avgEndGap /= right.size() - 1;

            System.out.println("AvgStartGap " + avgStartGap + "   AvgEndGap " + avgEndGap);

            if (Math.abs(avgStartGap - avgEndGap) < 15) {
                colAllignment.add("CENTER");
            } else {
                colAllignment.add("LEFT");
            }

            right.clear();
            left.clear();
        }

        //XWPFDocument document = new XWPFDocument();
        //FileOutputStream outDocx = new FileOutputStream(new File("F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\create_table_2.docx"));
        FileOutputStream outDocx = new FileOutputStream(new File(bookPath + File.separator + "table_" + pageName + ".docx"));
        XWPFTable table = document.createTable();
        table.setWidth("100%");
        setTableAlign(table, ParagraphAlignment.CENTER);
        /*CTTblLayoutType type = table.getCTTbl().getTblPr().addNewTblLayout();
        type.setType(STTblLayoutType.FIXED);*/

        /*CTTblWidth width = table.getCTTbl().addNewTblPr().addNewTblW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(5 * (long) Math.abs(tablemaxX - tableminX)));*/
 
        XWPFTableRow tableRow = table.getRow(0);       
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
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                } else if (colAllignment.get(j - 1).equals("CENTER")) {
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                } else {
                    paragraph.setAlignment(ParagraphAlignment.LEFT);
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

    public static void setTableAlign(XWPFTable table, ParagraphAlignment align) {
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        CTJc jc = (tblPr.isSetJc() ? tblPr.getJc() : tblPr.addNewJc());
        STJc.Enum en = STJc.Enum.forInt(align.getValue());
        jc.setVal(en);
    }

}
