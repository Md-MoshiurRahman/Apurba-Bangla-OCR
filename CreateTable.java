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
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;

public class CreateTable {

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
        addTable(pageAnnotations);

    }

    public static void addTable(PageAnnotations pageAnnotations) throws IOException {
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
                            if (maxxY < point.getY()) {
                                maxxY = point.getY();
                            }

                            if (minnX > point.getX()) {
                                minnX = point.getX();
                            }
                            if (maxxX < point.getX()) {
                                maxxX = point.getX();
                            }
                        }

                        for (Map.Entry<Integer, String> text : textLine.getValue().getText().entrySet()) {
                            if (!text.getValue().trim().isEmpty()) {
                                Point pointArray[] = new Point[2];
                                pointArray[0] = new Point(minnX, minnY);
                                pointArray[1] = new Point(maxxX, maxxY);

                                double width = maxxX - minnX;
                                double height = maxxY - minnY;
                                if (width <= 550.0 && height <= 150.0) {
                                    textPoints.put(pointArray, text.getValue().trim());
                                }
                            }
                        }
                    }
                }
                System.out.println("Counter " + count);
                XWPFDocument document = new XWPFDocument();
                //XWPFRun run = new
                //createTable_2(document, textPoints, tableminX, tableminY, tablemaxX, tablemaxY);
            }
        }
    }

    public static void addTable(ArrayList<LayoutRegion> segments,  XWPFDocument document, XWPFRun run) throws IOException {
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
                            if (maxxY < point.getY()) {
                                maxxY = point.getY();
                            }

                            if (minnX > point.getX()) {
                                minnX = point.getX();
                            }
                            if (maxxX < point.getX()) {
                                maxxX = point.getX();
                            }
                        }

                        for (Map.Entry<Integer, String> text : textLine.getValue().getText().entrySet()) {
                            if (!text.getValue().trim().isEmpty()) {
                                Point pointArray[] = new Point[2];
                                pointArray[0] = new Point(minnX, minnY);
                                pointArray[1] = new Point(maxxX, maxxY);

                                double width = maxxX - minnX;
                                double height = maxxY - minnY;
                                if (width <= 550.0 && height <= 150.0) {
                                    textPoints.put(pointArray, text.getValue().trim());
                                }
                            }
                        }
                    }
                }
                System.out.println("Counter " + count);
                createTable_2(document, run, textPoints, tableminX, tableminY, tablemaxX, tablemaxY);
            }
        }
    }

    public static void createTable_2(XWPFDocument document, XWPFRun run, HashMap<Point[], String> textPoints, double tableminX, double tableminY, double tablemaxX, double tablemaxY) throws FileNotFoundException, IOException {
        ArrayList<Double> colValues = new ArrayList<>();
        ArrayList<Double> rowValues = new ArrayList<>();

        colValues.add(tableminX);
        double prevColValue = tableminX;

        for (double i = tableminX + 1; i <= tablemaxX; i++) {
            int flag = 0;
            for (Map.Entry textPoint : textPoints.entrySet()) {
                Point[] point = (Point[]) textPoint.getKey();
                if (point[0].getX() <= i && i <= point[1].getX()) {
                    flag = 1;
                    break;
                }
            }
            if (flag == 0) {
                if (i - prevColValue > 10) {
                    colValues.add(i);
                }
                prevColValue = i;
            }
        }
        System.out.print("colvalues ");
        for (double colval : colValues) {
            System.out.print(colval + " ");
        }

        rowValues.add(tableminY);
        double prevRowvalue = tableminY;

        for (double i = tableminY + 1; i <= tablemaxY; i++) {
            int flag = 0;
            for (Map.Entry textPoint : textPoints.entrySet()) {
                Point[] point = (Point[]) textPoint.getKey();
                if (point[0].getY() <= i && i <= point[1].getY()) {
                    flag = 1;
                    break;
                }
            }
            if (flag == 0) {
                if (i - prevRowvalue > 10) {
                    rowValues.add(i);
                }
                prevRowvalue = i;
            }
        }
        System.out.print("\nrowvalues ");
        for (double rowval : rowValues) {
            System.out.print(rowval + " ");
        }

        ArrayList<Point[]> sortedKeyPoints = new ArrayList<Point[]>(textPoints.keySet());
        Collections.sort(sortedKeyPoints, new Comparator<Point[]>() {
            @Override
            public int compare(Point[] p1, Point[] p2) {
                if ((int) Math.abs(p1[0].getY() - p2[0].getY()) > 5) {
                    return (int) (p1[0].getY() - p2[0].getY());
                } else {
                    return (int) (p1[0].getX() - p2[0].getX());
                }

            }
        });

        //XWPFDocument document = new XWPFDocument();
        //FileOutputStream outDocx = new FileOutputStream(new File("F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\create_table_2.docx"));
        FileOutputStream outDocx = new FileOutputStream(new File(bookPath + File.separator + "table_" + pageName + ".docx"));
        XWPFTable table = document.createTable();

        XWPFTableRow tableRow = table.getRow(0);
        String s = "";

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
                    }
                }
                
                
                /*tableRow.getCell(j-1).addParagraph().addRun(run);
                tableRow.getCell(j - 1).setText(s);
                for (XWPFParagraph p : tableRow.getCell(j - 1).getParagraphs()) {
                    
                    for (XWPFRun tableRun : p.getRuns()) {
                        
                        tableRun.setFontFamily("SolaimanLipi");
                        tableRun.setFontSize(10);
                        if(tableRun.getText(0) == null){
                            System.out.println("Nullllllllllllllll");
                        }
                        if(tableRun.getText(0).contains(" ")){
                            System.out.println("Containsssssssssssss ");
                        }
                        if (tableRun.getText(0) != null && tableRun.getText(0).contains("\n")) {
                            System.out.println("here in paragraph!!!!!!!!!");
                            String[] lines = run.getText(0).split("\n");
                            if (lines.length > 0) {
                                tableRun.setText(lines[0], 0);
                                for (int l = 1; l < lines.length; l++) {
                                    tableRun.addBreak();
                                    tableRun.setText(lines[l]);
                                }
                            }
                        }
                    }
                }*/
                tableRow.getCell(j - 1).setText(s);
                if (i == 1 && j < colValues.size() - 1) {
                    tableRow.addNewTableCell();
                }
            }
            if (i < rowValues.size() - 1) {
                tableRow = table.createRow();
            }

        }

        document.write(outDocx);
        outDocx.close();

        /*String spath = "F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\table_77.docx";
        String tpath = "F:\\Apurba\\LayoutAnalyzer\\src\\main\\resources\\books\\table.docx";

        forverseTableCells(spath, tpath);*/

    }

    public static void forverseTableCells(String sourceFile, String targetFile) throws FileNotFoundException, IOException {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(sourceFile));
        for (XWPFTable table : doc.getTables()) {//Table
            for (XWPFTableRow row : table.getRows()) {//row
                for (XWPFTableCell cell : row.getTableCells()) {//cell: Direct cell.setText() will only add text to the original and delete the text.
                    addBreakInCell(cell);
                    System.out.println("hereeeeee############");
                }
            }
        }
        FileOutputStream fos = new FileOutputStream(targetFile);
        doc.write(fos);
        fos.close();
        System.out.println("end");
    }

    public static void addBreakInCell(XWPFTableCell cell) {
        if (cell.getText() != null && cell.getText().contains("\n")) {
            System.out.println("##########yess########");
            for (XWPFParagraph p : cell.getParagraphs()) {
                for (XWPFRun run : p.getRuns()) {
                    //XWPFRun object defines a text area with a set of public properties
                    if (run.getText(0) != null && run.getText(0).contains("\n")) {
                        String[] lines = run.getText(0).split("\n");
                        if (lines.length > 0) {
                            run.setText(lines[0], 0); // set first line into XWPFRun
                            for (int i = 1; i < lines.length; i++) {
                                // add break and insert new text
                                run.addBreak();//interrupt
                                // run.addCarriageReturn();//carriage return, but does not work
                                run.setText(lines[i]);
                            }
                        }
                    }
                }
            }
        }
    }

}
