package app.converter;

import com.microsoft.schemas.vml.CTGroup;
import com.microsoft.schemas.vml.CTLine;
import com.microsoft.schemas.vml.CTShape;
import com.microsoft.schemas.vml.CTTextbox;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTOrientation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.Map;

public class StyledDocument {

    public StyledDocument(InputStream inputStream, Map<String,String> documentAttr) throws IOException, InvalidFormatException {
            this.document = new XWPFDocument(OPCPackage.open(inputStream));
           /* CTPageSz pageSize = null;
            pageSize.setW(BigInteger.valueOf(595));
            pageSize.setH(BigInteger.valueOf(842));
            document.getDocument().getBody().getSectPr().setPgSz(pageSize);*/
        this.documentAttribute=documentAttr;
    }
    StyledDocument(String path, Map<String,String> documentAttr){
        try {
            FileInputStream fileInputStream = new FileInputStream(path);
            this.document = new XWPFDocument(OPCPackage.open(fileInputStream));
        } catch (IOException e) {
            System.out.println("Not able to load style template"+e.toString());
        } catch (InvalidFormatException e) {
            System.out.println("InvalidFormatException"+e.toString());
            throw new RuntimeException(e);
        }
        this.documentAttribute=documentAttr;
    }
    private final float mmInDxa = 144 / 2.54f;
    private int indenTbl = Math.round(12 * mmInDxa-8);
    private XWPFDocument document;

    private Map<String,String> documentAttribute;

    private static void setCellText(XWPFTableCell tableCell, String text) {
        tableCell.setText(text);
        tableCell.getParagraphArray(0).setAlignment(ParagraphAlignment.LEFT);
        tableCell.getParagraphArray(0).getRuns().get(0).setFontFamily("Arial");
        tableCell.getParagraphArray(0).getRuns().get(0).setFontSize(9);
        tableCell.getParagraphArray(0).getRuns().get(0).setItalic(true);
    }

    private static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTVMerge vmerge = CTVMerge.Factory.newInstance();
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                vmerge.setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                vmerge.setVal(STMerge.CONTINUE);
                // and the content should be removed
                for (int i = cell.getParagraphs().size(); i > 0; i--) {
                    cell.removeParagraph(0);
                }
                cell.addParagraph();
            }
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr == null) tcPr = cell.getCTTc().addNewTcPr();
            tcPr.setVMerge(vmerge);
        }
    }

    private static void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(colIndex);
            CTHMerge hmerge = CTHMerge.Factory.newInstance();
            if (colIndex == fromCol) {
                // The first merged cell is set with RESTART merge value
                hmerge.setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                hmerge.setVal(STMerge.CONTINUE);
                // and the content should be removed
                for (int i = cell.getParagraphs().size(); i > 0; i--) {
                    cell.removeParagraph(0);
                }
                cell.addParagraph();
            }
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr == null) tcPr = cell.getCTTc().addNewTcPr();
            tcPr.setHMerge(hmerge);
        }
    }

    private void setPageMargins() {
        int tabFromPageRightTopBottomBoarder = Math.round(5 * mmInDxa);
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(Math.round(8 * mmInDxa)));
        pageMar.setRight(BigInteger.valueOf(tabFromPageRightTopBottomBoarder));
        pageMar.setTop(BigInteger.valueOf(tabFromPageRightTopBottomBoarder));
        pageMar.setBottom(BigInteger.valueOf(12 * tabFromPageRightTopBottomBoarder));
        pageMar.setFooter(BigInteger.valueOf(tabFromPageRightTopBottomBoarder));
        pageMar.setHeader(BigInteger.valueOf(tabFromPageRightTopBottomBoarder));
        XWPFParagraph [] parArr=document.getParagraphs().toArray(new XWPFParagraph[0]);
        int pos = parArr.length - 1;
        while (pos >= 0) {
            parArr[pos].setIndentationLeft(Math.round(30 * mmInDxa));
            parArr[pos].setIndentationRight(Math.round(15 * mmInDxa));
            pos--;
        }
    }

    private void createNewFooter() {

        XWPFFooter footer = document.createFooter(HeaderFooterType.FIRST);
        int columnNumber = 11;
        int rowNumber = 8;
        XWPFTable table = footer.createTable(rowNumber, columnNumber);
        CTGroup gr=CTGroup.Factory.newInstance();
        CTJcTable ctJcTable = table.getCTTbl().getTblPr().addNewJc();
        ctJcTable.setVal(STJcTable.LEFT);
        CTTblWidth tableIndentation = table.getCTTbl().getTblPr().addNewTblInd();
        tableIndentation.setType(STTblWidth.DXA);
        tableIndentation.setW(BigInteger.valueOf(indenTbl));
        table.setCellMargins(0, 20, 0, 20);
        table.getCTTbl().getTblPr().addNewTblLayout().setType(STTblLayoutType.FIXED);
        CTTblGrid tblGrid = table.getCTTbl().addNewTblGrid();
        CTTblBorders ctBoarders = table.getCTTbl().getTblPr().getTblBorders();
        BigInteger tblBoarderWidth = BigInteger.valueOf(12);
        CTBorder[] Boarders = {ctBoarders.getTop(), ctBoarders.getBottom(), ctBoarders.getLeft(), ctBoarders.getRight(),
                ctBoarders.getTop(), ctBoarders.getInsideV(), ctBoarders.getInsideH()};
        for (CTBorder boarder : Boarders) {
            boarder.setVal(STBorder.SINGLE);
            boarder.setSz(tblBoarderWidth);
            boarder.setSpace(BigInteger.valueOf(0));
        }
        int[] columnWidth = {7, 10, 23, 15, 10, 70, 5, 5, 5, 15, 20};
        for (int i = 0; i < columnNumber; i++) {
            tblGrid.addNewGridCol().setW(BigInteger.valueOf(Math.round(columnWidth[i] * mmInDxa)));
        }
        for (XWPFTableRow row : table.getRows()) {
            row.setHeightRule(TableRowHeightRule.EXACT);
            row.setHeight(Math.round(5 * mmInDxa));
        }

        //Обозначение документа
        mergeCellHorizontally(table, rowNumber - 8, 5, 10);
        mergeCellHorizontally(table, rowNumber - 7, 5, 10);
        mergeCellHorizontally(table, rowNumber - 6, 5, 10);
        mergeCellVertically(table, 5, rowNumber - 8, rowNumber - 6);
        //Наименование документа
        mergeCellVertically(table, 5, rowNumber - 5, rowNumber - 1);
        //Лит
        mergeCellHorizontally(table, rowNumber - 5, 6, 8);
        //
        for (int i = rowNumber - 3; i < rowNumber; i++) {
            mergeCellHorizontally(table, i, 6, 10);
        }
        mergeCellVertically(table, 6, rowNumber - 3, rowNumber - 1);
        //
        for (int i = rowNumber - 5; i < rowNumber; i++) {
            mergeCellHorizontally(table, i, 0, 1);
        }
        //
        for (int j = 0; j < 5; j++) {
            for (int i = rowNumber - 1; i >= rowNumber - 5; i--) {
                CTBorder tblCTopBoarder = table.getRow(i).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewTop();
                CTBorder tblCBottomBoarder = table.getRow(i).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewBottom();
                tblCTopBoarder.setVal(STBorder.SINGLE);
                tblCTopBoarder.setSz(BigInteger.valueOf(4));
                if (i != rowNumber - 1) {
                    tblCBottomBoarder.setVal(STBorder.SINGLE);
                    tblCBottomBoarder.setSz(BigInteger.valueOf(4));
                }
            }
        }
        for (int i = 6; i < 9; i++) {
            CTBorder tblCRightBoarder = table.getRow(rowNumber - 4).getCell(i).getCTTc().addNewTcPr().addNewTcBorders().addNewRight();
            CTBorder tblCLeftBoarder = table.getRow(rowNumber - 4).getCell(i).getCTTc().addNewTcPr().addNewTcBorders().addNewLeft();
            if (i != 6) {
                tblCLeftBoarder.setVal(STBorder.SINGLE);
                tblCLeftBoarder.setSz(BigInteger.valueOf(4));
            }
            if (i != 8) {
                tblCRightBoarder.setVal(STBorder.SINGLE);
                tblCRightBoarder.setSz(BigInteger.valueOf(4));
            }
        }
        //Заполняем графы
        for (int j = rowNumber - 1; j >= 0; j--) {
            for (int i = 0; i < columnNumber; i++) {
                XWPFTableCell tblCell = table.getRow(j).getCell(i);
                tblCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                if ((j == rowNumber - 1) && (i == 0)) {
                    setCellText(tblCell, "Утв.");
                }
                if ((j == rowNumber - 1) && (i == 2)) {
                    setCellText(tblCell, documentAttribute.get("approver"));
                }
                if ((j == rowNumber - 2) && (i == 0)) {
                    setCellText(tblCell, "Н.контр");
                }
                if ((j == rowNumber - 4) && (i == 0)) {
                    setCellText(tblCell, "Пров.");
                }
                if ((j == rowNumber - 4) && (i == 2)) {
                    setCellText(tblCell, documentAttribute.get("checker"));
                }
                if ((j == rowNumber - 5) && (i == 0)) {
                    setCellText(tblCell, "Разраб.");
                }
                if ((j == rowNumber - 5) && (i == 2)) {
                    setCellText(tblCell, documentAttribute.get("developer"));
                }
                if ((j == rowNumber - 6) && (i == 0)) {
                    setCellText(tblCell, "Изм.");
                }
                if ((j == rowNumber - 6) && (i == 1)) {
                    setCellText(tblCell, "Лист");
                }
                if ((j == rowNumber - 6) && (i == 2)) {
                    setCellText(tblCell, "№ докум.");
                }
                if ((j == rowNumber - 6) && (i == 3)) {
                    setCellText(tblCell, "Подп.");
                }
                if ((j == rowNumber - 6) && (i == 4)) {
                    setCellText(tblCell, "Дата");
                }
                if ((j == rowNumber - 8) && (i == 5)) {
                    setCellText(tblCell, documentAttribute.get("documentCode"));
                    tblCell.getParagraphArray(0).getRuns().get(0).setFontSize(18);
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                }
                
                if ((j == rowNumber - 5) && (i == 5)) {
                    setCellText(tblCell, documentAttribute.get("productName"));
                    tblCell.addParagraph();
                    tblCell.getParagraphArray(1).createRun().setText(documentAttribute.get("documentName"));
                    tblCell.getParagraphArray(1).getRuns().get(0).setFontFamily("Arial");
                    tblCell.getParagraphArray(1).getRuns().get(0).setFontSize(9);
                    tblCell.getParagraphArray(1).getRuns().get(0).setItalic(true);
                    tblCell.getParagraphArray(0).getRuns().get(0).setFontSize(14);
                    tblCell.getParagraphArray(1).getRuns().get(0).setFontSize(14);
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                    tblCell.getParagraphArray(1).setAlignment(ParagraphAlignment.CENTER);
                }
                if ((j == rowNumber - 5) && (i == 6)) {
                    setCellText(tblCell, "Лит.");
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                }
                if ((j == rowNumber - 5) && (i == 9)) {
                    setCellText(tblCell, "Лист");
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                }
                if ((j == rowNumber - 4) && (i == 9)) {
                    setCellText(tblCell, "");
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.BEGIN);
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewInstrText().setStringValue(" PAGE ");
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.END);
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                }
                if ((j == rowNumber - 5) && (i == 10)) {
                    setCellText(tblCell, "Листов");
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);

                }
                if ((j == rowNumber - 4) && (i == 10)) {
                    setCellText(tblCell, "");
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.BEGIN);
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewInstrText().setStringValue(" NUMPAGES ");
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.END);
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);

                }

            }
        }

    }

    private void createNewShemeFooter() {

        XWPFFooter footer = document.createFooter(HeaderFooterType.FIRST);
        int columnNumber = 12;
        int rowNumber = 11;
        //int tableWidth=Math.round(185*mmInDxa);
        XWPFTable table = footer.createTable(rowNumber, columnNumber);
        CTJcTable ctJcTable = table.getCTTbl().getTblPr().addNewJc();
        ctJcTable.setVal(STJcTable.LEFT);
        CTTblWidth tableIndentation = table.getCTTbl().getTblPr().addNewTblInd();
        tableIndentation.setType(STTblWidth.DXA);
        tableIndentation.setW(BigInteger.valueOf(indenTbl));
        table.setCellMargins(0, 20, 0, 20);
        table.getCTTbl().getTblPr().addNewTblLayout().setType(STTblLayoutType.FIXED);
        CTTblGrid tblGrid = table.getCTTbl().addNewTblGrid();
        CTTblBorders ctBoarders = table.getCTTbl().getTblPr().getTblBorders();
        //XWPFTable.XWPFBorderType[] boarderType={table.getBottomBorderType(),};
        BigInteger tblBoarderWidth = BigInteger.valueOf(12);
        CTBorder[] Boarders = {ctBoarders.getTop(), ctBoarders.getBottom(), ctBoarders.getLeft(), ctBoarders.getRight(),
                ctBoarders.getTop(), ctBoarders.getInsideV(), ctBoarders.getInsideH()};
        for (CTBorder boarder : Boarders) {
            boarder.setVal(STBorder.SINGLE);
            boarder.setSz(tblBoarderWidth);
            boarder.setSpace(BigInteger.valueOf(0));
        }
        int[] columnWidth = {7, 10, 23, 15, 10, 70, 5, 5, 5, 5, 12, 18};
        for (int i = 0; i < columnNumber; i++) {
            tblGrid.addNewGridCol().setW(BigInteger.valueOf(Math.round(columnWidth[i] * mmInDxa)));
        }
        for (XWPFTableRow row : table.getRows()) {
            row.setHeightRule(TableRowHeightRule.EXACT);
            row.setHeight(Math.round(5 * mmInDxa));
        }

        //Обозначение документа
        mergeCellHorizontally(table, rowNumber - 11, 5, 11);
        mergeCellHorizontally(table, rowNumber - 10, 5, 11);
        mergeCellHorizontally(table, rowNumber - 9, 5, 11);
        mergeCellVertically(table, 5, rowNumber - 11, rowNumber - 9);
        //Наименование документа
        mergeCellHorizontally(table, rowNumber - 3, 6, 11);
        mergeCellHorizontally(table, rowNumber - 2, 6, 11);
        mergeCellHorizontally(table, rowNumber - 1, 6, 11);
        mergeCellVertically(table, 6, rowNumber - 3, rowNumber - 1);
        //
        mergeCellVertically(table, 5, rowNumber - 3, rowNumber - 1);
        //
        mergeCellVertically(table, 5, rowNumber - 8, rowNumber - 4);
        //
        for (int i = 6; i < 12; i++) {
            mergeCellVertically(table, i, rowNumber - 7, rowNumber - 5);
        }
        for (int i = rowNumber - 6; i < rowNumber; i++) {
            mergeCellHorizontally(table, i, 0, 1);
        }
        //
        for (int i = rowNumber - 6; i < rowNumber; i++) {
            mergeCellHorizontally(table, i, 0, 1);
        }
        //
        mergeCellHorizontally(table, rowNumber - 4, 6, 9);
        //
        mergeCellHorizontally(table, rowNumber - 4, 10, 11);
        //
        mergeCellHorizontally(table, rowNumber - 5, 9, 10);
        mergeCellHorizontally(table, rowNumber - 6, 9, 10);
        mergeCellHorizontally(table, rowNumber - 7, 9, 10);

        //
        mergeCellHorizontally(table, rowNumber - 8, 9, 10);
        //
        mergeCellHorizontally(table, rowNumber - 8, 6, 8);
        //
        for (int j = 0; j < 5; j++) {
            for (int i = rowNumber - 1; i >= rowNumber - 6; i--) {
                CTBorder tblCTopBoarder = table.getRow(i).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewTop();
                CTBorder tblCBottomBoarder = table.getRow(i).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewBottom();
                tblCTopBoarder.setVal(STBorder.SINGLE);
                tblCTopBoarder.setSz(BigInteger.valueOf(4));
                if (i != rowNumber - 1) {
                    tblCBottomBoarder.setVal(STBorder.SINGLE);
                    tblCBottomBoarder.setSz(BigInteger.valueOf(4));
                }
            }
        }
        for (int j = rowNumber - 5; j >= rowNumber - 7; j--) {
            for (int i = 6; i < 9; i++) {
                CTBorder tblCRightBoarder = table.getRow(j).getCell(i).getCTTc().addNewTcPr().addNewTcBorders().addNewRight();
                CTBorder tblCLeftBoarder = table.getRow(j).getCell(i).getCTTc().addNewTcPr().addNewTcBorders().addNewLeft();
                if (i != 6) {
                    tblCLeftBoarder.setVal(STBorder.SINGLE);
                    tblCLeftBoarder.setSz(BigInteger.valueOf(4));
                }
                if (i != 8) {
                    tblCRightBoarder.setVal(STBorder.SINGLE);
                    tblCRightBoarder.setSz(BigInteger.valueOf(4));
                }
            }
        }
        //Заполняем графы
        for (int j = rowNumber - 1; j >= 0; j--) {
            for (int i = 0; i < columnNumber; i++) {
                XWPFTableCell tblCell = table.getRow(j).getCell(i);
                tblCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                if ((j == rowNumber - 1) && (i == 0)) {
                    setCellText(tblCell, "Утв.");
                }
                if ((j == rowNumber - 2) && (i == 0)) {
                    setCellText(tblCell, "Н.контр");
                }
                if ((j == rowNumber - 4) && (i == 0)) {
                    setCellText(tblCell, "Т.контр");
                }

                if ((j == rowNumber - 5) && (i == 0)) {
                    setCellText(tblCell, "Пров.");
                }
                if ((j == rowNumber - 6) && (i == 0)) {
                    setCellText(tblCell, "Разраб.");
                }
                if ((j == rowNumber - 7) && (i == 0)) {
                    setCellText(tblCell, "Изм.");
                }
                if ((j == rowNumber - 7) && (i == 1)) {
                    setCellText(tblCell, "Лист");
                }
                if ((j == rowNumber - 7) && (i == 2)) {
                    setCellText(tblCell, "№ докум.");
                }
                if ((j == rowNumber - 7) && (i == 3)) {
                    setCellText(tblCell, "Подп.");
                }
                if ((j == rowNumber - 7) && (i == 4)) {
                    setCellText(tblCell, "Дата");
                }
                if ((j == rowNumber - 8) && (i == 6)) {
                    setCellText(tblCell, "Лит.");
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                }
                if ((j == rowNumber - 8) && (i == 9)) {
                    setCellText(tblCell, "Масса");
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                }
                if ((j == rowNumber - 8) && (i == 11)) {
                    setCellText(tblCell, "Масштаб");
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                }
                if ((j == rowNumber - 4) && (i == 6)) {
                    setCellText(tblCell, "Лист ");
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.BEGIN);
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewInstrText().setStringValue(" PAGE ");
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.END);
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                }
                if ((j == rowNumber - 4) && (i == 10)) {
                    setCellText(tblCell, "Листов ");
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.BEGIN);
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewInstrText().setStringValue(" NUMPAGES ");
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.END);
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);

                }

            }
        }

    }

    private void createNewDefaultFooter() {

        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        int columnNumber = 7;
        int rowNumber = 4;
        XWPFTable table = footer.createTable(rowNumber, columnNumber);
        CTJcTable ctJcTable = table.getCTTbl().getTblPr().addNewJc();
        ctJcTable.setVal(STJcTable.LEFT);
        CTTblWidth tableIndentation = table.getCTTbl().getTblPr().addNewTblInd();
        tableIndentation.setType(STTblWidth.DXA);
        tableIndentation.setW(BigInteger.valueOf(indenTbl));
        table.setCellMargins(0, 20, 0, 20);
        table.getCTTbl().getTblPr().addNewTblLayout().setType(STTblLayoutType.FIXED);
        CTTblGrid tblGrid = table.getCTTbl().addNewTblGrid();
        CTTblBorders ctBoarders = table.getCTTbl().getTblPr().getTblBorders();
        BigInteger tblBoarderWidth = BigInteger.valueOf(12);
        CTBorder[] Boarders = {ctBoarders.getTop(), ctBoarders.getBottom(), ctBoarders.getLeft(), ctBoarders.getRight(),
                ctBoarders.getTop(), ctBoarders.getInsideV(), ctBoarders.getInsideH()};
        for (CTBorder boarder : Boarders) {
            boarder.setVal(STBorder.SINGLE);
            boarder.setSz(tblBoarderWidth);
            boarder.setSpace(BigInteger.valueOf(0));
        }
        int[] columnWidth = {7, 10, 23, 15, 10, 110, 10};
        for (int i = 0; i < columnNumber; i++) {
            tblGrid.addNewGridCol().setW(BigInteger.valueOf(Math.round(columnWidth[i] * mmInDxa)));
        }
        int[] tblRowHeights = {5, 2, 3, 5};
        for (XWPFTableRow row : table.getRows()) {
            row.setHeightRule(TableRowHeightRule.EXACT);
            row.setHeight(Math.round(tblRowHeights[table.getRows().indexOf(row)] * mmInDxa));
        }

        //Обозначение документа
        mergeCellVertically(table, 5, 0, rowNumber - 1);
        //
        mergeCellVertically(table, 6, rowNumber - 2, rowNumber - 1);
        //
        mergeCellVertically(table, 6, 0, 1);
        //
        for (int j = 0; j < 5; j++) {
            mergeCellVertically(table, j, 1, 2);
            CTBorder tblCBottomBoarder = table.getRow(0).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewBottom();
            CTBorder tblCTopBoarder = table.getRow(1).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewTop();
            tblCTopBoarder.setVal(STBorder.SINGLE);
            tblCTopBoarder.setSz(BigInteger.valueOf(4));
            tblCBottomBoarder.setVal(STBorder.SINGLE);
            tblCBottomBoarder.setSz(BigInteger.valueOf(4));
        }
        //Заполняем графы
        for (int j = rowNumber - 1; j >= 0; j--) {
            for (int i = 0; i < columnNumber; i++) {
                XWPFTableCell tblCell = table.getRow(j).getCell(i);
                tblCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                if ((j == 0) && (i == 5)) {
                setCellText(tblCell, documentAttribute.get("documentCode"));
                tblCell.getParagraphArray(0).getRuns().get(0).setFontSize(18);
                tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                }
                if ((j == rowNumber - 1) && (i == 0)) {
                    setCellText(tblCell, "Изм.");
                }
                if ((j == rowNumber - 1) && (i == 1)) {
                    setCellText(tblCell, "Лист");
                }
                if ((j == rowNumber - 1) && (i == 2)) {
                    setCellText(tblCell, "№ докум.");
                }
                if ((j == rowNumber - 1) && (i == 3)) {
                    setCellText(tblCell, "Подп.");
                }
                if ((j == rowNumber - 1) && (i == 4)) {
                    setCellText(tblCell, "Дата");
                }
                if ((j == 0) && (i == 6)) {
                    setCellText(tblCell, "Лист ");
                }
                if ((j == rowNumber - 2) && (i == 6)) {
                    setCellText(tblCell, "");
                    tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.BEGIN);
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewInstrText().setStringValue(" PAGE ");
                    tblCell.getParagraphArray(0).getCTP().getRArray(0).addNewFldChar().setFldCharType(STFldCharType.END);
                }

            }
        }
    }

    private static void createLineInGroup(CTGroup group, int fromXmm, int fromYmm, int toXmm, int toYmm) {
        CTLine line = group.addNewLine();
        line.setStyle("position:absolute");
        float coeffConv = 3.778f;
        String sFrom = fromXmm * coeffConv + "," + fromYmm * coeffConv;
        line.setFrom(sFrom);
        String sTo = toXmm * coeffConv + "," + toYmm * coeffConv;
        line.setTo(sTo);
        line.setStrokeweight("1.5pt");
    }

    private void createNewHeader(HeaderFooterType type) {
        XWPFHeader header = document.createHeader(type);
        int columnNumber = 2;
        int rowNumber = 8;
        XWPFRun run = header.createParagraph().createRun();
        run.getCTR().addNewRPr().addNewNoProof();
        CTPicture pict = run.getCTR().addNewPict();
        CTGroup group1 = CTGroup.Factory.newInstance();
        //group1.setHralign(STHrAlign.RIGHT);
        createLineInGroup(group1, 12, 0, 197, 0);
        createLineInGroup(group1, 12, 0, 12, 287);
        createLineInGroup(group1, 197, 0, 197, 287);
        pict.set(group1);
        XWPFRun run2 = header.getParagraphArray(0).createRun();
        CTPicture pict2 = run2.getCTR().addNewPict();
        CTGroup group2 = CTGroup.Factory.newInstance();

        CTShape shape = group2.addNewShape();
        shape.setStyle("position:absolute;margin-left:-7.3pt;margin-top:-4.75pt;width:auto;height:1150;z-index:251662335");
        CTTextbox textBox = shape.addNewTextbox();
        CTTxbxContent txbxContent = textBox.addNewTxbxContent();
        XWPFTable table = new XWPFTable(txbxContent.addNewTbl(), run2.getParagraph().getBody(), rowNumber, columnNumber);
        CTJcTable ctJcTable = table.getCTTbl().getTblPr().addNewJc();
        ctJcTable.setVal(STJcTable.LEFT);
        table.setCellMargins(0, 30, 0, 30);
        table.getCTTbl().getTblPr().addNewTblLayout().setType(STTblLayoutType.FIXED);
        CTTblGrid tblGrid = table.getCTTbl().addNewTblGrid();
        CTTblBorders ctBoarders = table.getCTTbl().getTblPr().getTblBorders();
        BigInteger tblBoarderWidth = BigInteger.valueOf(12);
        CTBorder[] Boarders = {ctBoarders.getTop(), ctBoarders.getBottom(), ctBoarders.getLeft(), ctBoarders.getRight(),
                ctBoarders.getTop(), ctBoarders.getInsideV(), ctBoarders.getInsideH()};
        for (CTBorder boarder : Boarders) {
            boarder.setVal(STBorder.SINGLE);
            boarder.setSz(tblBoarderWidth);
            boarder.setSpace(BigInteger.valueOf(0));
        }
        int[] columnWidth = {5, 7};
        int[] columnHeight = {60, 60, 22, 35, 25, 25, 35, 25};
        for (int i = 0; i < columnNumber; i++) {
            tblGrid.addNewGridCol().setW(BigInteger.valueOf(Math.round(columnWidth[i] * mmInDxa)));
        }
        XWPFTableRow row = null;
        for (int i = 0; i < rowNumber; i++) {
            row = table.getRow(i);
            row.setHeightRule(TableRowHeightRule.EXACT);
            row.setHeight(Math.round(columnHeight[i] * mmInDxa - 2));
        }
        //
        mergeCellHorizontally(table, 2, 0, 1);
        //Заполняем графы
        String[] cellContents = {"Инв.№ подл.", "Подп.и дата", "Взам.инв.№", "Инв.№ дубл.", "Подп.и дата", "", "Справ.№", "Перв.примен."};
        for (int i = 0; i < rowNumber; i++) {
            if ((type == HeaderFooterType.DEFAULT) && (i < 2)) {
                cellContents[cellContents.length - 1 - i] = "";
                for (int j = 0; j < 2; j++) {
                    table.getRow(i).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewRight().setVal(STBorder.NIL);
                    table.getRow(i).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewLeft().setVal(STBorder.NIL);
                    table.getRow(i).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewTop().setVal(STBorder.NIL);
                    table.getRow(i).getCell(j).getCTTc().addNewTcPr().addNewTcBorders().addNewBottom().setVal(STBorder.NIL);
                }
            }
            if (i == 2) {
                table.getRow(i).getCell(0).getCTTc().addNewTcPr().addNewTcBorders().addNewLeft().setVal(STBorder.NIL);
                table.getRow(i).getCell(0).getCTTc().addNewTcPr().addNewTcBorders().addNewTop().setVal(STBorder.NIL);
            }
            table.getRow(i).getCell(1).getCTTc().addNewTcPr().addNewTcBorders().addNewRight().setVal(STBorder.NIL);
            XWPFTableCell tblCell = table.getRow(rowNumber - i - 1).getCell(0);
            tblCell.getCTTc().addNewTcPr().addNewTextDirection().setVal(STTextDirection.BT_LR);
            setCellText(tblCell, cellContents[i]);
            tblCell.getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
        }
        CTTbl[] tblArray = {table.getCTTbl()};
        txbxContent.setTblArray(tblArray);
        pict2.set(group2);
    }

    protected void createNewCleanHeaderFooter() {
        for (XWPFHeader header : document.getHeaderList()) {
            header.setHeaderFooter(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr.Factory.newInstance());
        }
        for (XWPFFooter footer : document.getFooterList()) {
            footer.setHeaderFooter(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr.Factory.newInstance());
        }
    }


    private void createFrames() {
        this.createNewCleanHeaderFooter();
        this.setPageMargins();
        this.createNewHeader(HeaderFooterType.FIRST);
        this.createNewHeader(HeaderFooterType.DEFAULT);
        this.createNewFooter();
        this.createNewDefaultFooter();
    }
    public void createFile(String outputPath) {
        this.createFrames();
        File outputFile = new File(outputPath);
        FileOutputStream os = null;
        try {
            os = new FileOutputStream(outputFile);
            this.document.write(os);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
    public void createStream(OutputStream os) throws IOException {
        this.createFrames();
        this.document.write(os);
    }

}