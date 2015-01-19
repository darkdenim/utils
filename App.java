package test;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.text.ParseException;
import java.util.*;

public class App {
    public static void main(String[] args) {
        try {

            FileInputStream file = new FileInputStream(new File("F:\\Downloads\\Sample - Superstore Sales (Excel).xls"));

            //Get the workbook instance for XLS file
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            //Get first sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);

            List<ModelA> models = new ArrayList<>();
            List<List<String>> rows = new ArrayList<>();
            //Get iterator to all the rows in current sheet
            Iterator<Row> rowIterator = sheet.iterator();
            Row headerRow = sheet.getRow(0);
            while(rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (row == headerRow) {
                    continue;
                }
                models.add(toModelA(row));
            }

            System.out.println("Found " + models.size() + " models. Printing...");
            for (ModelA model : models) {
                System.out.println(model);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static ModelA toModelA(Row row) throws ParseException {
        List<Object> rowList = new ArrayList<>();
        Iterator<Cell> cellIterator = row.cellIterator();
        List<String> l = new ArrayList<>();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN: {
                    //rowList.add(Boolean.toString(cell.getBooleanCellValue()));
                    break;
                }
                case Cell.CELL_TYPE_ERROR: {
                    //rowList.add(cell.getErrorCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA: {
                    //rowList.add(cell.getCellFormula());
                    break;
                }
                case Cell.CELL_TYPE_NUMERIC: {
                    rowList.add(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_STRING: {
                    rowList.add(cell.getStringCellValue());
                    break;
                }
                default: {

//                    rowList.add(cell.getStringCellValue());
                }
            }
        }

        return new ModelA(rowList);
    }
}
