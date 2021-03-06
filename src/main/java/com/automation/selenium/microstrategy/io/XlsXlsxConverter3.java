package com.automation.selenium.microstrategy.io;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;

import java.util.*;

public class XlsXlsxConverter3 {


    public static void run(String type, Workbook wb_1, Workbook wb_2, Workbook wb_final) {

        final XSSFSheet source_act = ((XSSFWorkbook) wb_1).getSheetAt(0);
        final XSSFSheet source_prev = ((XSSFWorkbook) wb_2).getSheetAt(0);

        XSSFSheet sMain = ((XSSFWorkbook) wb_final).getSheet("Actual Week");
        XSSFSheet sLastWeek = ((XSSFWorkbook) wb_final).getSheet("Last Week");

        copySheet(source_act, sMain);
        copySheet(source_prev, sLastWeek);

        populateSheet (type, sMain, sLastWeek);

    }

    public static void populateSheet(String type, XSSFSheet sMain, XSSFSheet sLastWeek) {

        //Insert formulas
        CellAddress first_tkt_address_sMain = null;
        CellAddress last_tkt_address_sMain  = null;

        CellAddress first_tkt_address_lw = null;
        CellAddress last_tkt_address_lw  = null;

        if(type.equals("SR")){
            first_tkt_address_sMain = getFirstValContainAddress(sMain, "SR1");
            last_tkt_address_sMain = getLastValContainAddress(sMain, "SR");

            first_tkt_address_lw = getFirstValContainAddress(sLastWeek, "SR1");
            last_tkt_address_lw = getLastValContainAddress(sLastWeek, "SR");
        }else
        if (type.equals("IN")){
            first_tkt_address_sMain = getFirstValContainAddress(sMain, "IN9");
            last_tkt_address_sMain = getLastValContainAddress(sMain, "IN");

            first_tkt_address_lw = getFirstValContainAddress(sLastWeek, "IN9");
            last_tkt_address_lw = getLastValContainAddress(sLastWeek, "IN");
        }

        //Header titles
        final Cell cell_header_sourc = sMain.getRow(2).getCell(15);
        Cell cell_header_dest = sMain.getRow(2).getCell(16);
        if (cell_header_dest == null)
            cell_header_dest = sMain.getRow(2).createCell(16);


        copyCell(cell_header_sourc, cell_header_dest, new ArrayList<CellStyle>());
        setCellValue(sMain, 2, 16, CellType.STRING, "Aging");


        //Set formulas
        final DataFormat format = sMain.getWorkbook().createDataFormat();

        for (int i = first_tkt_address_sMain.getRow(); i <= last_tkt_address_sMain.getRow(); i++){
            setFormulaCell(sMain,i, 16, format, CellType.FORMULA, "TODAY()-j" + (i+1));
            setFormulaCell(sMain, i, 17,
                    CellType.FORMULA,
                    getVlookUpFormulaOnPos(
                            "A" + (i+1),
                            sLastWeek.getSheetName(),
                            first_tkt_address_lw.toString() + ":"
                                    + "I" + (last_tkt_address_lw.getRow()+1),
                            7,
                            "FALSE"
                    ));

        }

    }

    private static class FormulaInfo {

        private String sheetName;
        private Integer rowIndex;
        private Integer cellIndex;
        private String formula;

        private FormulaInfo(String sheetName, Integer rowIndex, Integer cellIndex, String formula) {
            this.sheetName = sheetName;
            this.rowIndex = rowIndex;
            this.cellIndex = cellIndex;
            this.formula = formula;
        }

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public Integer getRowIndex() {
            return rowIndex;
        }

        public void setRowIndex(Integer rowIndex) {
            this.rowIndex = rowIndex;
        }

        public Integer getCellIndex() {
            return cellIndex;
        }

        public void setCellIndex(Integer cellIndex) {
            this.cellIndex = cellIndex;
        }

        public String getFormula() {
            return formula;
        }

        public void setFormula(String formula) {
            this.formula = formula;
        }
    }

    static List<FormulaInfo> formulaInfoList = new ArrayList<FormulaInfo>();


    public static void refreshFormula(XSSFWorkbook workbook) {
        for (FormulaInfo formulaInfo : formulaInfoList) {
            workbook.getSheet(formulaInfo.getSheetName()).getRow(formulaInfo.getRowIndex())
                    .getCell(formulaInfo.getCellIndex()).setCellFormula(formulaInfo.getFormula());
        }
        formulaInfoList.removeAll(formulaInfoList);
    }

    public static XSSFWorkbook convertWorkbookHSSFToXSSF(HSSFWorkbook source) {
        XSSFWorkbook retVal = new XSSFWorkbook();

        for (int i = 0; i < source.getNumberOfSheets(); i++) {

            HSSFSheet hssfsheet = source.getSheetAt(i);
            XSSFSheet xssfSheet = retVal.createSheet(hssfsheet.getSheetName());

            copySheetSettings(xssfSheet, hssfsheet);
            copySheet(hssfsheet, xssfSheet);
            copyPictures(xssfSheet, hssfsheet);
        }

        refreshFormula(retVal);

        return retVal;
    }

    private static void copySheet(XSSFSheet source, XSSFSheet destination) {
        copySheet(source, destination, true);
    }

    private static void copySheet(HSSFSheet source, XSSFSheet destination) {
        copySheet(source, destination, true);
    }

    private static void copySheet(XSSFSheet source, XSSFSheet destination, boolean copyStyle) {
        int maxColumnNum = 0;
        List<CellStyle> styleMap2 = (copyStyle) ? new ArrayList<CellStyle>() : null;
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            XSSFRow srcRow = source.getRow(i);
            XSSFRow destRow = destination.createRow(i);
            if (srcRow != null) {
                copyRow(source, destination, srcRow, destRow, styleMap2);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            destination.setColumnWidth(i, source.getColumnWidth(i));
        }
    }

    /**
     * @param source      the sheet to create from the copy.
     * @param destination sheet to copy.
     * @param copyStyle   true copy the style.
     */
    private static void copySheet(HSSFSheet source, XSSFSheet destination, boolean copyStyle) {
        int maxColumnNum = 0;
        List<CellStyle> styleMap2 = (copyStyle) ? new ArrayList<CellStyle>() : null;
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            HSSFRow srcRow = source.getRow(i);
            XSSFRow destRow = destination.createRow(i);
            if (srcRow != null) {
                copyRow(source, destination, srcRow, destRow, styleMap2);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            destination.setColumnWidth(i, source.getColumnWidth(i));
        }
    }

    /**
     * @param srcSheet  the sheet to copy.
     * @param destSheet the sheet to create.
     * @param srcRow    the row to copy.
     * @param destRow   the row to create.
     * @param styleMap  -
     */
    private static void copyRow(HSSFSheet srcSheet, XSSFSheet destSheet, HSSFRow srcRow, XSSFRow destRow,
                                List<CellStyle> styleMap) {
        // manage a list of merged zone in order to not insert two times a
        // merged zone
        Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
        short dh = srcSheet.getDefaultRowHeight();
        if (srcRow.getHeight() != dh) {
            destRow.setHeight(srcRow.getHeight());
        }


        int j = srcRow.getFirstCellNum();
        if (j < 0) {
            j = 0;
        }
        for (; j <= srcRow.getLastCellNum(); j++) {
            HSSFCell oldCell = srcRow.getCell(j); // ancienne cell
            XSSFCell newCell = destRow.getCell(j); // new cell
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                // copy chaque cell
                copyCell(oldCell, newCell, styleMap);
                // copy les informations de fusion entre les cellules
                // System.out.println("row num: " + srcRow.getRowNum() +
                // " , col: " + (short)oldCell.getColumnIndex());
                CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(),
                        (short) oldCell.getColumnIndex());

                if (mergedRegion != null) {
                    // System.out.println("Selected merged region: " +
                    // mergedRegion.toString());
                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                            mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                    // System.out.println("New merged region: " +
                    // newMergedRegion.toString());
                    CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
                    if (isNewMergedRegion(wrapper, mergedRegions)) {
                        mergedRegions.add(wrapper);
                        destSheet.addMergedRegion(wrapper.range);
                    }
                }
            }
        }

    }

    private static void copyRow(XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow, XSSFRow destRow,
                                List<CellStyle> styleMap) {
        // manage a list of merged zone in order to not insert two times a
        // merged zone
        Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
        short dh = srcSheet.getDefaultRowHeight();
        if (srcRow.getHeight() != dh) {
            destRow.setHeight(srcRow.getHeight());
        }


        int j = srcRow.getFirstCellNum();
        if (j < 0) {
            j = 0;
        }
        for (; j <= srcRow.getLastCellNum(); j++) {
            XSSFCell oldCell = srcRow.getCell(j); // ancienne cell
            XSSFCell newCell = destRow.getCell(j); // new cell
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }

                copyCell(oldCell, newCell, styleMap);
                CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(),
                        (short) oldCell.getColumnIndex());

                if (mergedRegion != null) {
                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                            mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                    CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
                    if (isNewMergedRegion(wrapper, mergedRegions)) {
                        mergedRegions.add(wrapper);
                        destSheet.addMergedRegion(wrapper.range);
                    }
                }
            }
        }

    }

    /**
     * Récupère les informations de fusion des cellules dans la sheet source_dir
     * pour les appliquer à la sheet destination... Récupère toutes les zones
     * merged dans la sheet source_dir et regarde pour chacune d'elle si elle se
     * trouve dans la current row que nous traitons. Si oui, retourne l'objet
     * CellRangeAddress.
     *
     * @param sheet   the sheet containing the data.
     * @param rowNum  the num of the row to copy.
     * @param cellNum the num of the cell to copy.
     * @return the CellRangeAddress created.
     */
    public static CellRangeAddress getMergedRegion(HSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }

    public static CellRangeAddress getMergedRegion(XSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }

    /**
     * Check that the merged region has been created in the destination sheet.
     *
     * @param newMergedRegion the merged region to copy or not in the destination sheet.
     * @param mergedRegions   the list containing all the merged region.
     * @return true if the merged region is already in the list or not.
     */
    private static boolean isNewMergedRegion(CellRangeAddressWrapper newMergedRegion,
                                             Set<CellRangeAddressWrapper> mergedRegions) {
        return !mergedRegions.contains(newMergedRegion);
    }

    private static void copyPictures(Sheet newSheet, Sheet sheet) {
        Drawing drawingOld = sheet.createDrawingPatriarch();
        Drawing drawingNew = newSheet.createDrawingPatriarch();
        CreationHelper helper = newSheet.getWorkbook().getCreationHelper();

        // if (drawingNew instanceof HSSFPatriarch) {
        if (drawingOld instanceof HSSFPatriarch) {
            List<HSSFShape> shapes = ((HSSFPatriarch) drawingOld).getChildren();
            for (int i = 0; i < shapes.size(); i++) {
                System.out.println(shapes.size());
                if (shapes.get(i) instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shapes.get(i);
                    HSSFPictureData picdata = pic.getPictureData();
                    int pictureIndex = newSheet.getWorkbook().addPicture(picdata.getData(), picdata.getFormat());
                    ClientAnchor anchor = null;
                    if (pic.getAnchor() != null) {
                        anchor = helper.createClientAnchor();
                        anchor.setDx1(((HSSFClientAnchor) pic.getAnchor()).getDx1());
                        anchor.setDx2(((HSSFClientAnchor) pic.getAnchor()).getDx2());
                        anchor.setDy1(((HSSFClientAnchor) pic.getAnchor()).getDy1());
                        anchor.setDy2(((HSSFClientAnchor) pic.getAnchor()).getDy2());
                        anchor.setCol1(((HSSFClientAnchor) pic.getAnchor()).getCol1());
                        anchor.setCol2(((HSSFClientAnchor) pic.getAnchor()).getCol2());
                        anchor.setRow1(((HSSFClientAnchor) pic.getAnchor()).getRow1());
                        anchor.setRow2(((HSSFClientAnchor) pic.getAnchor()).getRow2());
                        anchor.setAnchorType(((HSSFClientAnchor) pic.getAnchor()).getAnchorType());
                    }
                    drawingNew.createPicture(anchor, pictureIndex);
                }
            }
        } else {
            if (drawingNew instanceof XSSFDrawing) {
                List<XSSFShape> shapes = ((XSSFDrawing) drawingOld).getShapes();
                for (int i = 0; i < shapes.size(); i++) {
                    if (shapes.get(i) instanceof XSSFPicture) {
                        XSSFPicture pic = (XSSFPicture) shapes.get(i);
                        XSSFPictureData picdata = pic.getPictureData();
                        int pictureIndex = newSheet.getWorkbook().addPicture(picdata.getData(),
                                picdata.getPictureType());
                        XSSFClientAnchor anchor = null;
                        CTTwoCellAnchor oldAnchor = ((XSSFDrawing) drawingOld).getCTDrawing().getTwoCellAnchorArray(i);
                        if (oldAnchor != null) {
                            anchor = (XSSFClientAnchor) helper.createClientAnchor();
                            CTMarker markerFrom = oldAnchor.getFrom();
                            CTMarker markerTo = oldAnchor.getTo();
                            anchor.setDx1((int) markerFrom.getColOff());
                            anchor.setDx2((int) markerTo.getColOff());
                            anchor.setDy1((int) markerFrom.getRowOff());
                            anchor.setDy2((int) markerTo.getRowOff());
                            anchor.setCol1(markerFrom.getCol());
                            anchor.setCol2(markerTo.getCol());
                            anchor.setRow1(markerFrom.getRow());
                            anchor.setRow2(markerTo.getRow());
                        }
                        drawingNew.createPicture(anchor, pictureIndex);
                    }
                }
            }
        }
    }

    private static void copyCell(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
        if (styleList != null) {
            if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()) {
                newCell.setCellStyle(oldCell.getCellStyle());
            } else {
                DataFormat newDataFormat = newCell.getSheet().getWorkbook().createDataFormat();

                CellStyle newCellStyle = getSameCellStyle(oldCell, newCell, styleList);
                if (newCellStyle == null) {
                    // Create a new cell style
                    Font oldFont = oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());
                    // Find a existing font corresponding to avoid to create a
                    // new one
                    Font newFont = newCell
                            .getSheet()
                            .getWorkbook()
                            .findFont(oldFont.getBold(), oldFont.getColor(), oldFont.getFontHeight(),
                                    oldFont.getFontName(), oldFont.getItalic(), oldFont.getStrikeout(),
                                    oldFont.getTypeOffset(), oldFont.getUnderline());
                    if (newFont == null) {
                        newFont = newCell.getSheet().getWorkbook().createFont();
                        newFont.setColor(oldFont.getColor());
                        newFont.setFontHeight(oldFont.getFontHeight());
                        newFont.setFontName(oldFont.getFontName());
                        newFont.setItalic(oldFont.getItalic());
                        newFont.setStrikeout(oldFont.getStrikeout());
                        newFont.setTypeOffset(oldFont.getTypeOffset());
                        newFont.setUnderline(oldFont.getUnderline());
                        newFont.setCharSet(oldFont.getCharSet());
                    }

                    short newFormat = newDataFormat.getFormat(oldCell.getCellStyle().getDataFormatString());
                    newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
                    newCellStyle.setFont(newFont);
                    newCellStyle.setDataFormat(newFormat);

                    newCellStyle.setAlignment(oldCell.getCellStyle().getAlignmentEnum());
                    newCellStyle.setHidden(oldCell.getCellStyle().getHidden());
                    newCellStyle.setLocked(oldCell.getCellStyle().getLocked());
                    newCellStyle.setWrapText(oldCell.getCellStyle().getWrapText());
                    newCellStyle.setBorderBottom(oldCell.getCellStyle().getBorderBottomEnum());
                    newCellStyle.setBorderLeft(oldCell.getCellStyle().getBorderLeftEnum());
                    newCellStyle.setBorderRight(oldCell.getCellStyle().getBorderRightEnum());
                    newCellStyle.setBorderTop(oldCell.getCellStyle().getBorderTopEnum());
                    newCellStyle.setBottomBorderColor(oldCell.getCellStyle().getBottomBorderColor());
                    newCellStyle.setFillBackgroundColor(oldCell.getCellStyle().getFillBackgroundColor());
                    newCellStyle.setFillForegroundColor(oldCell.getCellStyle().getFillForegroundColor());
                    newCellStyle.setFillPattern(oldCell.getCellStyle().getFillPatternEnum());
                    newCellStyle.setIndention(oldCell.getCellStyle().getIndention());
                    newCellStyle.setLeftBorderColor(oldCell.getCellStyle().getLeftBorderColor());
                    newCellStyle.setRightBorderColor(oldCell.getCellStyle().getRightBorderColor());
                    newCellStyle.setRotation(oldCell.getCellStyle().getRotation());
                    newCellStyle.setTopBorderColor(oldCell.getCellStyle().getTopBorderColor());
                    newCellStyle.setVerticalAlignment(oldCell.getCellStyle().getVerticalAlignmentEnum());

                    styleList.add(newCellStyle);
                }
                newCell.setCellStyle(newCellStyle);
            }
        }
        switch (oldCell.getCellTypeEnum()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BLANK:
                newCell.setCellType(CellType.BLANK);
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                formulaInfoList.add(new FormulaInfo(oldCell.getSheet().getSheetName(), oldCell.getRowIndex(), oldCell
                        .getColumnIndex(), oldCell.getCellFormula()));
                break;
            default:
                break;
        }
    }

    private static CellStyle getSameCellStyle(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
        CellStyle styleToFind = oldCell.getCellStyle();
        CellStyle currentCellStyle = null;
        CellStyle returnCellStyle = null;
        Iterator<CellStyle> iterator = styleList.iterator();
        Font oldFont = null;
        Font newFont = null;
        while (iterator.hasNext() && returnCellStyle == null) {
            currentCellStyle = iterator.next();

            if (currentCellStyle.getAlignmentEnum() != styleToFind.getAlignmentEnum()) {
                continue;
            }
            if (currentCellStyle.getHidden() != styleToFind.getHidden()) {
                continue;
            }
            if (currentCellStyle.getLocked() != styleToFind.getLocked()) {
                continue;
            }
            if (currentCellStyle.getWrapText() != styleToFind.getWrapText()) {
                continue;
            }
            if (currentCellStyle.getBorderBottomEnum() != styleToFind.getBorderBottomEnum()) {
                continue;
            }
            if (currentCellStyle.getBorderLeftEnum() != styleToFind.getBorderLeftEnum()) {
                continue;
            }
            if (currentCellStyle.getBorderRightEnum() != styleToFind.getBorderRightEnum()) {
                continue;
            }
            if (currentCellStyle.getBorderTopEnum() != styleToFind.getBorderTopEnum()) {
                continue;
            }
            if (currentCellStyle.getBottomBorderColor() != styleToFind.getBottomBorderColor()) {
                continue;
            }
            if (currentCellStyle.getFillBackgroundColor() != styleToFind.getFillBackgroundColor()) {
                continue;
            }
            if (currentCellStyle.getFillForegroundColor() != styleToFind.getFillForegroundColor()) {
                continue;
            }
            if (currentCellStyle.getFillPatternEnum() != styleToFind.getFillPatternEnum()) {
                continue;
            }
            if (currentCellStyle.getIndention() != styleToFind.getIndention()) {
                continue;
            }
            if (currentCellStyle.getLeftBorderColor() != styleToFind.getLeftBorderColor()) {
                continue;
            }
            if (currentCellStyle.getRightBorderColor() != styleToFind.getRightBorderColor()) {
                continue;
            }
            if (currentCellStyle.getRotation() != styleToFind.getRotation()) {
                continue;
            }
            if (currentCellStyle.getTopBorderColor() != styleToFind.getTopBorderColor()) {
                continue;
            }
            if (currentCellStyle.getVerticalAlignmentEnum() != styleToFind.getVerticalAlignmentEnum()) {
                continue;
            }

            oldFont = oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());
            newFont = newCell.getSheet().getWorkbook().getFontAt(currentCellStyle.getFontIndex());

            if (newFont.getBold() == oldFont.getBold()) {
                continue;
            }
            if (newFont.getColor() == oldFont.getColor()) {
                continue;
            }
            if (newFont.getFontHeight() == oldFont.getFontHeight()) {
                continue;
            }
            if (newFont.getFontName() == oldFont.getFontName()) {
                continue;
            }
            if (newFont.getItalic() == oldFont.getItalic()) {
                continue;
            }
            if (newFont.getStrikeout() == oldFont.getStrikeout()) {
                continue;
            }
            if (newFont.getTypeOffset() == oldFont.getTypeOffset()) {
                continue;
            }
            if (newFont.getUnderline() == oldFont.getUnderline()) {
                continue;
            }
            if (newFont.getCharSet() == oldFont.getCharSet()) {
                continue;
            }
            if (oldCell.getCellStyle().getDataFormatString().equals(currentCellStyle.getDataFormatString())) {
                continue;
            }

            returnCellStyle = currentCellStyle;
        }
        return returnCellStyle;
    }

    private static void copySheetSettings(Sheet newSheet, Sheet sheetToCopy) {

        newSheet.setAutobreaks(sheetToCopy.getAutobreaks());
        newSheet.setDefaultColumnWidth(sheetToCopy.getDefaultColumnWidth());
        newSheet.setDefaultRowHeight(sheetToCopy.getDefaultRowHeight());
        newSheet.setDefaultRowHeightInPoints(sheetToCopy.getDefaultRowHeightInPoints());
        newSheet.setDisplayGuts(sheetToCopy.getDisplayGuts());
        newSheet.setFitToPage(sheetToCopy.getFitToPage());

        newSheet.setForceFormulaRecalculation(sheetToCopy.getForceFormulaRecalculation());

        PrintSetup sheetToCopyPrintSetup = sheetToCopy.getPrintSetup();
        PrintSetup newSheetPrintSetup = newSheet.getPrintSetup();

        newSheetPrintSetup.setPaperSize(sheetToCopyPrintSetup.getPaperSize());
        newSheetPrintSetup.setScale(sheetToCopyPrintSetup.getScale());
        newSheetPrintSetup.setPageStart(sheetToCopyPrintSetup.getPageStart());
        newSheetPrintSetup.setFitWidth(sheetToCopyPrintSetup.getFitWidth());
        newSheetPrintSetup.setFitHeight(sheetToCopyPrintSetup.getFitHeight());
        newSheetPrintSetup.setLeftToRight(sheetToCopyPrintSetup.getLeftToRight());
        newSheetPrintSetup.setLandscape(sheetToCopyPrintSetup.getLandscape());
        newSheetPrintSetup.setValidSettings(sheetToCopyPrintSetup.getValidSettings());
        newSheetPrintSetup.setNoColor(sheetToCopyPrintSetup.getNoColor());
        newSheetPrintSetup.setDraft(sheetToCopyPrintSetup.getDraft());
        newSheetPrintSetup.setNotes(sheetToCopyPrintSetup.getNotes());
        newSheetPrintSetup.setNoOrientation(sheetToCopyPrintSetup.getNoOrientation());
        newSheetPrintSetup.setUsePage(sheetToCopyPrintSetup.getUsePage());
        newSheetPrintSetup.setHResolution(sheetToCopyPrintSetup.getHResolution());
        newSheetPrintSetup.setVResolution(sheetToCopyPrintSetup.getVResolution());
        newSheetPrintSetup.setHeaderMargin(sheetToCopyPrintSetup.getHeaderMargin());
        newSheetPrintSetup.setFooterMargin(sheetToCopyPrintSetup.getFooterMargin());
        newSheetPrintSetup.setCopies(sheetToCopyPrintSetup.getCopies());

        Header sheetToCopyHeader = sheetToCopy.getHeader();
        Header newSheetHeader = newSheet.getHeader();
        newSheetHeader.setCenter(sheetToCopyHeader.getCenter());
        newSheetHeader.setLeft(sheetToCopyHeader.getLeft());
        newSheetHeader.setRight(sheetToCopyHeader.getRight());

        Footer sheetToCopyFooter = sheetToCopy.getFooter();
        Footer newSheetFooter = newSheet.getFooter();
        newSheetFooter.setCenter(sheetToCopyFooter.getCenter());
        newSheetFooter.setLeft(sheetToCopyFooter.getLeft());
        newSheetFooter.setRight(sheetToCopyFooter.getRight());

        newSheet.setHorizontallyCenter(sheetToCopy.getHorizontallyCenter());
        newSheet.setMargin(Sheet.LeftMargin, sheetToCopy.getMargin(Sheet.LeftMargin));
        newSheet.setMargin(Sheet.RightMargin, sheetToCopy.getMargin(Sheet.RightMargin));
        newSheet.setMargin(Sheet.TopMargin, sheetToCopy.getMargin(Sheet.TopMargin));
        newSheet.setMargin(Sheet.BottomMargin, sheetToCopy.getMargin(Sheet.BottomMargin));

        newSheet.setPrintGridlines(sheetToCopy.isPrintGridlines());
        newSheet.setRowSumsBelow(sheetToCopy.getRowSumsBelow());
        newSheet.setRowSumsRight(sheetToCopy.getRowSumsRight());
        newSheet.setVerticallyCenter(sheetToCopy.getVerticallyCenter());
        newSheet.setDisplayFormulas(sheetToCopy.isDisplayFormulas());
        newSheet.setDisplayGridlines(sheetToCopy.isDisplayGridlines());
        newSheet.setDisplayRowColHeadings(sheetToCopy.isDisplayRowColHeadings());
        newSheet.setDisplayZeros(sheetToCopy.isDisplayZeros());
        newSheet.setPrintGridlines(sheetToCopy.isPrintGridlines());
        newSheet.setRightToLeft(sheetToCopy.isRightToLeft());
        newSheet.setZoom(1);
        copyPrintTitle(newSheet, sheetToCopy);
    }

    private static void copyPrintTitle(Sheet newSheet, Sheet sheetToCopy) {
        int nbNames = sheetToCopy.getWorkbook().getNumberOfNames();
        Name name = null;
        String formula = null;

        String part1S = null;
        String part2S = null;
        String formS = null;
        String formF = null;
        String part1F = null;
        String part2F = null;
        int rowB = -1;
        int rowE = -1;
        int colB = -1;
        int colE = -1;

        for (int i = 0; i < nbNames; i++) {
            name = sheetToCopy.getWorkbook().getNameAt(i);
            if (name.getSheetIndex() == sheetToCopy.getWorkbook().getSheetIndex(sheetToCopy)) {
                if (name.getNameName().equals("Print_Titles")
                        || name.getNameName().equals(XSSFName.BUILTIN_PRINT_TITLE)) {
                    formula = name.getRefersToFormula();
                    int indexComma = formula.indexOf(",");
                    if (indexComma == -1) {
                        indexComma = formula.indexOf(";");
                    }
                    String firstPart = null;
                    ;
                    String secondPart = null;
                    if (indexComma == -1) {
                        firstPart = formula;
                    } else {
                        firstPart = formula.substring(0, indexComma);
                        secondPart = formula.substring(indexComma + 1);
                    }

                    formF = firstPart.substring(firstPart.indexOf("!") + 1);
                    part1F = formF.substring(0, formF.indexOf(":"));
                    part2F = formF.substring(formF.indexOf(":") + 1);

                    if (secondPart != null) {
                        formS = secondPart.substring(secondPart.indexOf("!") + 1);
                        part1S = formS.substring(0, formS.indexOf(":"));
                        part2S = formS.substring(formS.indexOf(":") + 1);
                    }

                    rowB = -1;
                    rowE = -1;
                    colB = -1;
                    colE = -1;
                    String rowBs, rowEs, colBs, colEs;
                    if (part1F.lastIndexOf("$") != part1F.indexOf("$")) {
                        rowBs = part1F.substring(part1F.lastIndexOf("$") + 1, part1F.length());
                        rowEs = part2F.substring(part2F.lastIndexOf("$") + 1, part2F.length());
                        rowB = Integer.parseInt(rowBs);
                        rowE = Integer.parseInt(rowEs);
                        if (secondPart != null) {
                            colBs = part1S.substring(part1S.lastIndexOf("$") + 1, part1S.length());
                            colEs = part2S.substring(part2S.lastIndexOf("$") + 1, part2S.length());
                            colB = CellReference.convertColStringToIndex(colBs);// CExportExcelHelperPoi.convertColumnLetterToInt(colBs);
                            colE = CellReference.convertColStringToIndex(colEs);// CExportExcelHelperPoi.convertColumnLetterToInt(colEs);
                        }
                    } else {
                        colBs = part1F.substring(part1F.lastIndexOf("$") + 1, part1F.length());
                        colEs = part2F.substring(part2F.lastIndexOf("$") + 1, part2F.length());
                        colB = CellReference.convertColStringToIndex(colBs);// CExportExcelHelperPoi.convertColumnLetterToInt(colBs);
                        colE = CellReference.convertColStringToIndex(colEs);// CExportExcelHelperPoi.convertColumnLetterToInt(colEs);

                        if (secondPart != null) {
                            rowBs = part1S.substring(part1S.lastIndexOf("$") + 1, part1S.length());
                            rowEs = part2S.substring(part2S.lastIndexOf("$") + 1, part2S.length());
                            rowB = Integer.parseInt(rowBs);
                            rowE = Integer.parseInt(rowEs);
                        }
                    }

//                    newSheet.getWorkbook().setRepeatingRowsAndColumns(newSheet.getWorkbook().getSheetIndex(newSheet),
                    //                           colB, colE, rowB - 1, rowE - 1);
                }
            }
        }
    }

    private static void setCellValue(Sheet sheet, int row, int cell, CellType type, Object value) {
        Row r = sheet.getRow(row);
        if (r == null)
            r = sheet.createRow(row);

        Cell c = r.getCell(cell);
        if (c == null)
            c = r.createCell(cell, type);

        if (value instanceof String)
            c.setCellValue(value.toString());
        else if (value instanceof Double)
            c.setCellValue((Double) value);
        else if (value instanceof Calendar)
            c.setCellValue((Calendar) value);
        else if (value instanceof Date)
            c.setCellValue((Date) value);
        else if (value instanceof Boolean)
            c.setCellValue((Boolean) value);
    }

    private static void setFormulaCell(Sheet sheet, int row, int cell, CellType formula, Object value) {
        Row r = sheet.getRow(row);
        if (r == null)
            r = sheet.createRow(row);

        Cell c = r.getCell(cell);
        if (c == null)
            c = r.createCell(cell, formula);
        c.setCellType(formula);
        c.setCellFormula(value.toString());

    }

    private static void setFormulaCell(Sheet sheet, int row, int cell, DataFormat format, CellType formula, Object value) {
        Row r = sheet.getRow(row);
        if (r == null)
            r = sheet.createRow(row);

        Cell c = r.getCell(cell);
        if (c == null)
            c = r.createCell(cell, formula);
        c.setCellType(formula);
        c.setCellFormula(value.toString());

        CellStyle cellstyle = c.getCellStyle();
        cellstyle.setDataFormat(format.getFormat("#"));

        c.setCellStyle(cellstyle);
        //region casting
/*
        if (value instanceof String)
            c.setFormulaCell(value.toString());
        else if (value instanceof Double)
            c.setFormulaCell((Double)value);
        else if(value instanceof Calendar)
            c.setFormulaCell((Calendar)value);
        else if(value instanceof Date)
            c.setFormulaCell((Date)value);
        else if(value instanceof Boolean)
            c.setFormulaCell((Boolean)value);
*/
        //endregion
    }

    private static void setCellNumericValue(Sheet sheet, int row, int cell, int number) {
        Row r = sheet.getRow(row);
        if (r == null)
            r = sheet.createRow(row);

        Cell c = r.getCell(cell);
        if (c == null)
            c = r.createCell(cell, CellType.NUMERIC);
        c.setCellType(CellType.NUMERIC);
        c.setCellValue(number);

    }

    private static CellAddress getFirstValAddress(Sheet s, String val) {

        Iterator<Row> iterator = s.iterator();
        CellAddress columnNumber = null;

        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellTypeEnum() == CellType.STRING) {
                    String text = cell.getStringCellValue();
                    if (val.equals(text)) {
                        columnNumber = cell.getAddress();
                        return columnNumber;
                    }
                }
            }
        }
        return columnNumber;
    }

    private static CellAddress getFirstValContainAddress(Sheet s, String val) {

        Iterator<Row> iterator = s.iterator();
        CellAddress columnNumber = null;

        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellTypeEnum() == CellType.STRING) {
                    String text = cell.getStringCellValue();
                    if (text.contains(val)) {
                        columnNumber = cell.getAddress();
                        return columnNumber;
                    }
                }
            }
        }
        return columnNumber;
    }

    private static CellAddress getLastValAddress(Sheet s, String val) {

        Iterator<Row> iterator = s.iterator();
        CellAddress columnNumber = null;

        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellTypeEnum() == CellType.STRING) {
                    String text = cell.getStringCellValue();
                    if (text.equals(val)) {
                        columnNumber = cell.getAddress();
                        break;
                    }
                }
            }
        }
        return columnNumber;
    }

    private static CellAddress getLastValContainAddress(Sheet s, String val) {

        Iterator<Row> iterator = s.iterator();
        CellAddress columnNumber = null;

        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellTypeEnum() == CellType.STRING) {
                    String text = cell.getStringCellValue();
                    if (text.contains(val)) {
                        columnNumber = cell.getAddress();
                        break;
                    }
                }
            }
        }
        return columnNumber;
    }

    /**
     * @param pos   first vlookup value
     * @param sheet sheet from value to compare
     * @param grid  section of the sheet
     * @param col   third vlookup value
     * @param bool  fourth vlookup value
     */
    private static String getVlookUpFormulaOnPos(String pos, String sheet, String grid, int col, String bool) {
        return "VLOOKUP("
                + pos + ","
                + "'" + sheet + "'!"
                + grid + ","
                + col + ","
                + bool + ")";
    }

    private static String getVlookUpFormulaOnPos(String pos, String sheet, String grid, int col) {
        return "VLOOKUP("
                + pos + ","
                + "'" + sheet + "'!"
                + grid + ","
                + col + ")";
    }

    private static void setHeader(String title, int col, CellStyle cs) {

    }

}
