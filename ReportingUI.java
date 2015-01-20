package com.smallbizconsult.ui;

import com.smallbizconsult.model.ModelA;
import com.vaadin.annotations.Widgetset;
import com.vaadin.data.Item;
import com.vaadin.server.VaadinRequest;
import com.vaadin.ui.Component;
import com.vaadin.ui.Table;
import com.vaadin.ui.UI;
import com.vaadin.ui.VerticalLayout;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.dussan.vaadin.dcharts.DCharts;
import org.dussan.vaadin.dcharts.base.elements.XYaxis;
import org.dussan.vaadin.dcharts.data.DataSeries;
import org.dussan.vaadin.dcharts.data.Ticks;
import org.dussan.vaadin.dcharts.metadata.renderers.AxisRenderers;
import org.dussan.vaadin.dcharts.metadata.renderers.SeriesRenderers;
import org.dussan.vaadin.dcharts.options.Axes;
import org.dussan.vaadin.dcharts.options.Highlighter;
import org.dussan.vaadin.dcharts.options.Options;
import org.dussan.vaadin.dcharts.options.SeriesDefaults;
import org.vaadin.spring.VaadinUI;

import java.io.File;
import java.io.FileInputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

@VaadinUI
public class ReportingUI extends UI {

    public static final String fileLocation = "F:\\Downloads\\Sample - Superstore Sales (Excel).xls";

    @Override
    protected void init(VaadinRequest vaadinRequest) {
        Table table = new Table("Superstore Sales");
        table.addContainerProperty(ModelA.Headings.RowID.caption, Long.class, null);
        table.addContainerProperty(ModelA.Headings.OrderID.caption, Long.class, null);
        table.addContainerProperty(ModelA.Headings.OrderDate.caption, Date.class, null);
        table.addContainerProperty(ModelA.Headings.OrderQuantity.caption, Integer.class, null);
        table.addContainerProperty(ModelA.Headings.CustomerName.caption, String.class, null);
        
        for (ModelA model : buildModel()) {
            Object itemId = table.addItem();
            Item item = table.getItem(itemId);
            item.getItemProperty(ModelA.Headings.RowID.caption).setValue(model.getRowId());
            item.getItemProperty(ModelA.Headings.OrderID.caption).setValue(model.getOrderId());
            item.getItemProperty(ModelA.Headings.OrderDate.caption).setValue(model.getOrderDate());
            item.getItemProperty(ModelA.Headings.OrderQuantity.caption).setValue(model.getQuantity());
            item.getItemProperty(ModelA.Headings.CustomerName.caption).setValue(model.getCustomer());
        }

        VerticalLayout layout = new VerticalLayout();
        layout.addComponent(table);
        generateCharts(layout);
        setContent(layout);
    }

    private void generateCharts(VerticalLayout layout) {
        DataSeries dataSeries = new DataSeries()
                .add(2, 6, 7, 10);

        SeriesDefaults seriesDefaults = new SeriesDefaults()
                .setRenderer(SeriesRenderers.BAR);

        Axes axes = new Axes()
                .addAxis(
                        new XYaxis()
                                .setRenderer(AxisRenderers.CATEGORY)
                                .setTicks(
                                        new Ticks()
                                                .add("a", "b", "c", "d")));

        Highlighter highlighter = new Highlighter()
                .setShow(false);

        Options options = new Options()
                .setSeriesDefaults(seriesDefaults)
                .setAxes(axes)
                .setHighlighter(highlighter);

        DCharts chart = new DCharts()
                .setDataSeries(dataSeries)
                .setOptions(options);
//                .show();

        layout.addComponent(chart);
    }

    public List<ModelA> buildModel() {
        List<ModelA> models = new ArrayList<ModelA>();
        try {

            FileInputStream file = new FileInputStream(new File(fileLocation));

            //Get the workbook instance for XLS file
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            //Get first sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);

            List<List<String>> rows = new ArrayList<List<String>>();
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
        return models;
    }

    private static ModelA toModelA(Row row) throws ParseException {
        List<Object> rowList = new ArrayList<Object>();
        Iterator<Cell> cellIterator = row.cellIterator();
        List<String> l = new ArrayList<String>();
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
                    rowList.add(cell.getCellFormula());
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
                    rowList.add(cell.getStringCellValue());
                }
            }
        }

        return new ModelA(rowList);
    }

}
