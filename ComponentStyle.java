package de.uniwue.web;

import de.uniwue.algorithm.geometry.regions.type.RegionType;
import de.uniwue.web.model.PageAnnotations;
import de.uniwue.web.model.Region;
import de.uniwue.web.model.TextLine;
import de.uniwue.web.model.TextStyle;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import javax.imageio.ImageIO;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

/**
 *
 * @author USER
 */
public class ComponentStyle {

    //final static Logger logger = LoggerFactory.getLogger(ComponentStyle.class);
    private int imageWidth, imageHeight, totalColumns, prevRowTotalColumns;
    private String prevRowImageSegmentId = null;
    private double textLineBoundaryStarts = Long.MAX_VALUE;
    private double textLineBoundaryEnds = 0;
    public HashMap<String, ParagraphAlignment> regionIdToAlignment = new HashMap<>();

    public PageAnnotations updateComponentAlignment(PageAnnotations pageAnnotations, ArrayList<Row> rows, String imgPath) {
        BufferedImage bimg;
        try {
            bimg = ImageIO.read(new File(imgPath));
        } catch (IOException ioEx) {
            //Logger.getLogger(ComponentStyle.class.getName()).log(Level.SEVERE, null, ioEx);
            return pageAnnotations;
        }
        imageWidth = bimg.getWidth();
        imageHeight = bimg.getHeight();
        for (Row row : rows) {
            if (row.getCols().size() > 1) {
                for (int i = 0; i < row.getCols().size(); i++) {
                    Col col = row.getCols().get(i);
                    extractDocumentLineBoundary(col.getRegions());
                }
            } else {
                extractDocumentLineBoundary(row.getRegions());
            }
        }
        System.out.println("textLineBoundaryStarts" + textLineBoundaryStarts);
        System.out.println("textLineBoundaryEnds" + textLineBoundaryEnds);

        for (int r = 0; r < rows.size(); r++) {
            Row row = rows.get(r);
            totalColumns = row.getCols().size() > 1 ? row.getCols().size() : 1;
            if (row.getCols().size() > 1) {
                for (int j = 0; j < row.getCols().size(); j++) {
                    Col col = row.getCols().get(j);
                    addColumn(col.getRegions(), j);
                }
            } else {
                addColumn(row.getRegions(), 0);
            }
            prevRowTotalColumns = totalColumns;
        }
        for (Map.Entry<String, Region> entry : pageAnnotations.getSegments().entrySet()) {
            Region region = entry.getValue();
            if (regionIdToAlignment.containsKey(region.getId())) {
                TextStyle textStyle = new TextStyle();
                textStyle.setAlignment(regionIdToAlignment.get(region.getId()).toString());
                entry.getValue().setTextStyle(textStyle);
            }
        }

        return pageAnnotations;
    }

    private void extractDocumentLineBoundary(ArrayList<LayoutRegion> segments) {
        double minX = Long.MAX_VALUE;
        double maxX = 0;
        for (Region segment : segments) {
            for (Map.Entry<String, TextLine> textLine : segment.getTextlines().entrySet()) {
                for (de.uniwue.web.model.Point point : textLine.getValue().getPoints()) {
                    if (minX > point.getX()) {
                        minX = point.getX();
                    }
                    if (maxX < point.getX()) {
                        maxX = point.getX();
                    }
                }
                if (minX > imageWidth - 1) {
                    minX = imageWidth - 1;
                }
                if (maxX > imageWidth - 1) {
                    maxX = imageWidth - 1;
                }
            }
        }
        if (textLineBoundaryStarts > minX) {
            textLineBoundaryStarts = minX;
        }
        if (textLineBoundaryEnds < maxX) {
            textLineBoundaryEnds = maxX;
        }
        //adding some thrseshold in right margin
        if (imageWidth - textLineBoundaryEnds > 20) {
            textLineBoundaryEnds = textLineBoundaryEnds + 20;
        }
    }

    private void addColumn(ArrayList<LayoutRegion> segments, int columnIndex) {
        int columnStart = 0;
        int columnEnd = 0;
        int boundaryStarts = (int) textLineBoundaryStarts;
        int boundaryEnds = (int) textLineBoundaryEnds;
        columnStart = boundaryStarts;
        if (totalColumns == 1) {
            columnEnd = boundaryEnds;
        } else {
            int textLineWidth = boundaryEnds - boundaryStarts;
            if (boundaryStarts == 0) {
                textLineWidth++;
            }
            int columnWidth = textLineWidth / totalColumns;
            columnStart = columnStart + (columnIndex * columnWidth);
            columnEnd = columnStart + columnWidth - 1;
        }
        Collections.sort(segments, (t, t1) -> {
            return t.y - t1.y;
        });
        String prevImageSegmentId = null;

        for (Region segment : segments) {
            int totalCentered = 0;
            int totalLeftAdjacent = 0;
            int totalRightAdjacent = 0;
            boolean isCentered = false;
            int prevLineStartX = 0;
            int prevLineEndX = 0;
            if (segment.getType().equals(RegionType.ImageRegion.toString()) || segment.getType().equals(RegionType.TableRegion.toString())) {

                double minX = Long.MAX_VALUE;
                double maxX = 0;
                for (de.uniwue.web.model.Point point : segment.getPoints()) {
                    if (minX > point.getX()) {
                        minX = point.getX();
                    }
                    if (maxX < point.getX()) {
                        maxX = point.getX();
                    }
                }
                if (minX > imageWidth - 1) {
                    minX = imageWidth - 1;
                }
                if (maxX > imageWidth - 1) {
                    maxX = imageWidth - 1;
                }
                //if there is miscalculation from ocr end then this is temporarily fixed, ocr team shoould fix it later
                if (minX < columnStart) {
                    minX = columnStart;
                }
                if (maxX > columnEnd) {
                    maxX = columnEnd;
                }
                int lineX = (int) minX;
                int lineWidth = (int) (maxX - minX);
                //checking single line center aligned line
                int lineStartGap = lineX - columnStart;
                int lineEndGap = columnEnd - (lineX + lineWidth);
                /*if (lineStartGap > 50 && lineEndGap > 50 && Math.abs(lineStartGap - lineEndGap) < 100) {
                    isCentered = true;
                } else if (lineStartGap <= lineEndGap) {
                    totalLeftAdjacent++;
                } else {
                    totalRightAdjacent++;
                }*/
                if (lineStartGap < lineEndGap - 50) {
                    totalLeftAdjacent++;
                } else if (lineStartGap - 50 > lineEndGap) {
                    totalRightAdjacent++;
                } else {
                    isCentered = true;
                }

                /*if (lineEndGap - lineStartGap >= 50) {
                    totalLeftAdjacent++;
                } else if (lineStartGap - lineEndGap >= 50) {
                    totalRightAdjacent++;
                } else {
                    totalCentered++;
                }*/
                if (isCentered) {
                    regionIdToAlignment.put(segment.getId(), ParagraphAlignment.CENTER);
                } else if (totalLeftAdjacent >= totalRightAdjacent) {
                    regionIdToAlignment.put(segment.getId(), ParagraphAlignment.LEFT);
                } else {
                    regionIdToAlignment.put(segment.getId(), ParagraphAlignment.RIGHT);
                }

                /*if (totalCentered > totalLeftAdjacent && totalCentered > totalRightAdjacent) {
                    regionIdToAlignment.put(segment.getId(), ParagraphAlignment.CENTER);
                    System.out.println("!!!!!!Center!!!!!! \n");
                } else if (totalRightAdjacent > totalCentered && totalRightAdjacent > totalLeftAdjacent) {
                    regionIdToAlignment.put(segment.getId(), ParagraphAlignment.RIGHT);
                    System.out.println("!!!!!!Right!!!!!! \n");
                } else {
                    regionIdToAlignment.put(segment.getId(), ParagraphAlignment.LEFT);
                    System.out.println("!!!!!!Left!!!!!! \n");
                }*/
                prevImageSegmentId = segment.getId();
                prevRowImageSegmentId = segment.getId();
            } /*else if (segment.getType().equals(RegionType.TableRegion.toString())) {
                regionIdToAlignment.put(segment.getId(), ParagraphAlignment.RIGHT);
            }*/ else {
                int totalLines = segment.getTextlines().size();

                for (Map.Entry<String, TextLine> textLine : segment.getTextlines().entrySet()) {
                    boolean emptyLine = true;
                    for (Map.Entry<Integer, String> text : textLine.getValue().getText().entrySet()) {
                        if (!text.getValue().trim().isEmpty()) {
                            emptyLine = false;
                        }
                    }
                    if (!emptyLine) {
                        double minX = Long.MAX_VALUE;
                        double maxX = 0;
                        for (de.uniwue.web.model.Point point : textLine.getValue().getPoints()) {
                            if (minX > point.getX()) {
                                minX = point.getX();
                            }
                            if (maxX < point.getX()) {
                                maxX = point.getX();
                            }

                        }
                        if (minX > imageWidth - 1) {
                            minX = imageWidth - 1;
                        }
                        if (maxX > imageWidth - 1) {
                            maxX = imageWidth - 1;
                        }
                        //if there is miscalculation from ocr end then this is temporarily fixed, ocr team shoould fix it later
                        if (minX < columnStart) {
                            minX = columnStart;
                        }
                        if (maxX > columnEnd) {
                            maxX = columnEnd;
                        }
                        int lineX = (int) minX;
                        int lineWidth = (int) (maxX - minX);
                        ///System.out.print(lineWidth + " ");
                        //checking single line center aligned line
                        int lineStartGap = lineX - columnStart;
                        int lineEndGap = columnEnd - (lineX + lineWidth);
                        //System.out.println(lineStartGap + " " + lineEndGap + "\n");
                        if (lineStartGap >= 60 && lineEndGap >= 60) {
                            totalCentered++;
                        } else if (lineStartGap - lineEndGap >= 50) {
                            totalRightAdjacent++;
                        } else if (lineEndGap - lineStartGap >= 50) {
                            totalLeftAdjacent++;
                        } else {
                            totalLeftAdjacent++;
                        }
                    }
                }
                ParagraphAlignment paragraphAlignment = ParagraphAlignment.LEFT;
                if (totalCentered > totalLeftAdjacent && totalCentered > totalRightAdjacent) {
                    regionIdToAlignment.put(segment.getId(), ParagraphAlignment.CENTER);
                    paragraphAlignment = ParagraphAlignment.CENTER;
                    //System.out.println("!!!!!!Center!!!!!! " + totalLines + "\n");
                } else if (totalRightAdjacent > totalCentered && totalRightAdjacent > totalLeftAdjacent) {
                    regionIdToAlignment.put(segment.getId(), ParagraphAlignment.RIGHT);
                    paragraphAlignment = ParagraphAlignment.RIGHT;
                    //System.out.println("!!!!!!Right!!!!!! " + totalLines + "\n");
                } else {
                    regionIdToAlignment.put(segment.getId(), ParagraphAlignment.LEFT);
                    paragraphAlignment = ParagraphAlignment.LEFT;
                    //System.out.println("!!!!!!Left!!!!!! " + totalLines + "\n");
                }

                if (prevImageSegmentId != null) {
                    regionIdToAlignment.put(prevImageSegmentId, paragraphAlignment);
                    prevImageSegmentId = null;
                }
                if (totalColumns == 1 && totalColumns == prevRowTotalColumns) {
                    if (prevRowImageSegmentId != null) {
                        regionIdToAlignment.put(prevRowImageSegmentId, paragraphAlignment);
                        prevRowImageSegmentId = null;
                    }
                } else {
                    prevRowImageSegmentId = null;
                }
            }
        }
    }
}
