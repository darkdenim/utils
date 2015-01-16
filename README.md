# utils
utils

# begin
package utils;

import com.vaadin.data.Item;
import com.vaadin.server.FileDownloader;
import com.vaadin.server.StreamResource;
import com.vaadin.server.VaadinSession;
import com.vaadin.ui.Button;
import com.vaadin.ui.Table;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.*;

public class XLSExporter {

    private Workbook wb;

    public XLSExporter() {
        wb = new HSSFWorkbook();
    }

    private List<List<Object>> tableToList(final Table table, final List<String> columns) {

        List<List<Object>> rows = new ArrayList<>();
        for (Object itemId : table.getItemIds()) {
            List<Object> row = new ArrayList<>();
            for (String column : columns) {
                row.add(table.getItem(itemId).getItemProperty(column).getValue());
            }
            rows.add(row);
        }

        return rows;
    }

    public void export(final Table table, final Object[] columns, final Button button) throws Exception {

        Map<String, CellStyle> styles = createStyles(wb);

        Sheet sheet = wb.createSheet("Sheet1");
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);

        int headerRowIndex = 2;
        String title = table.getCaption();

        //title row
        /*if (title == null || title.isEmpty()) {
            Row titleRow = sheet.createRow(0);
            titleRow.setHeightInPoints(45);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue(table.getCaption());
            titleCell.setCellStyle(styles.get("title"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("$A$1:$L$1"));
        } else {
            headerRowIndex--;
        }*/

        String[] titles = Arrays.asList(columns).toArray(new String[columns.length]);
        List<String> columnList = Arrays.asList(titles);

        //header row
        /*Row headerRow = sheet.createRow(headerRowIndex);
        headerRow.setHeightInPoints(40);
        Cell headerCell;
        for (int i = 0; i < titles.length; i++) {
            if (columnList.contains(titles[i])) {
                headerCell = headerRow.createCell(i);
                headerCell.setCellValue(titles[i]);
                headerCell.setCellStyle(styles.get("header"));
            }
        }*/

        // title
        if (table.getCaption() != null && !table.getCaption().isEmpty()) {
            Row titleRow = sheet.createRow(0);
            sheet.createRow(1);
//            titleRow.setHeightInPoints(45);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue(table.getCaption());
            titleCell.setCellStyle(styles.get("title"));
//            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.length-1));
            sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, columns.length-1));
        } else {
            headerRowIndex = 0;
        }

        // write headings
        Row headerRow = sheet.createRow(headerRowIndex);
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i].toString());
            cell.setCellStyle(styles.get("header"));
        }
//        List<List<Object>> rows = new ArrayList<>();
//        for (Object itemId : table.getItemIds()) {
//            List<Object> row = new ArrayList<>();
//            Item item = table.getItem(itemId);
//            for (Object propertyId : item.getItemPropertyIds()) {
//                if (columnList.contains(propertyId.toString())) {
//                    row.add(item.getItemProperty(propertyId).getValue());
//                }
//            }
//            rows.add(row);
//        }

        List<List<Object>> rows = tableToList(table, columnList);
        for (int i = 0; i < rows.size(); i++) {
            int index = headerRowIndex+i;
            if (i == 0) {
                index++;
            }
            Row row = sheet.createRow(index);
            List<Object> l = rows.get(i);
            for (int j = 0; j < l.size(); j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j);
                }

                if (l.get(j) == null) {
                    continue;
                }
                if(l.get(j) instanceof String) {
                    cell.setCellValue((String) l.get(j));
                }else if(l.get(j) instanceof Integer) {
                    cell.setCellValue((Integer)l.get(j));
                } else {
                    cell.setCellValue((Double)l.get(j));
                }
            }
        }

        //finally set column widths, the width is measured in units of 1/256th of a character width
        sheet.setColumnWidth(0, 30*256); //30 characters wide
        for (int i = 2; i < 9; i++) {
            sheet.setColumnWidth(i, 6*256);  //6 characters wide
        }
        sheet.setColumnWidth(10, 10*256); //10 characters wide

        // Write the output to a file
//        final String file = "timesheet.xls";
//        if(wb instanceof XSSFWorkbook) file += "x";
        final File file = File.createTempFile("timesheet", "xls");
        final FileOutputStream out = new FileOutputStream(file);
        wb.write(out);
        out.close();

//        FileDownloader downloader = new FileDownloader(new StreamResource(new StreamResource.StreamSource() {
//            @Override
//            public InputStream getStream() {
//                InputStream is = null;
////                try {
//                    is =  new FileInputStream(file);
////                } catch (FileNotFoundException e) {
////                    e.printStackTrace();
////                }
//                return is;
//            }
//        }, file));
//        downloader.extend(button);
    }

    /**
     * Create a library of cell styles
     */
    private static Map<String, CellStyle> createStyles(Workbook wb){
        Map<String, CellStyle> styles = new HashMap<>();
        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)15);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(titleFont);
        styles.put("title", style);

        Font monthFont = wb.createFont();
//        monthFont.setFontHeightInPoints((short)11);
//        monthFont.setColor(IndexedColors.WHITE.getIndex());
        monthFont.setBold(true);
        style = wb.createCellStyle();
//        style.setAlignment(CellStyle.ALIGN_CENTER);
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(monthFont);
//        style.setWrapText(true);
        styles.put("header", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styles.put("cell", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula_2", style);

        return styles;
    }
}

# end
