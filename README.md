# utils
utils

Structure
src
-main
--java
--webapp
---WEB-INF\web.xml

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
#begin
package test;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class ModelA {

    private long rowId;
    private long orderId;
    private Date orderDate;
    private Priority priority;
    private int quantity;
    private double sales;
    private double discount;
    private ShipMode shipMode;
    private double profit;
    private double unitPrice;
    private double shipping;
    private String customer;
    private String province;
    private String region;
    private Segment segment;
    private String productCategory;
    private String productSubCategory;
    private String productName;
    private String productContainer;
    private double baseMargin;
    private Date shipDate;

    public long getRowId() {
        return rowId;
    }

    public void setRowId(long rowId) {
        this.rowId = rowId;
    }

    public void setRowId(List<String> row) {
        this.rowId = Long.parseLong(row.get(Headings.RowID.ordinal()));
    }

    public long getOrderId() {
        return orderId;
    }

    public void setOrderId(long orderId) {
        this.orderId = orderId;
    }

    public void setOrderId(List<String> row) {
        this.orderId = Long.parseLong(row.get(Headings.OrderID.ordinal()));
    }

    public Date getOrderDate() {
        return orderDate;
    }

    public void setOrderDate(Date orderDate) {
        this.orderDate = orderDate;
    }

    public Priority getPriority() {
        return priority;
    }

    public void setPriority(Priority priority) {
        this.priority = priority;
    }

    public int getQuantity() {
        return quantity;
    }

    public void setQuantity(int quantity) {
        this.quantity = quantity;
    }

    public double getSales() {
        return sales;
    }

    public void setSales(double sales) {
        this.sales = sales;
    }

    public double getDiscount() {
        return discount;
    }

    public void setDiscount(double discount) {
        this.discount = discount;
    }

    public ShipMode getShipMode() {
        return shipMode;
    }

    public void setShipMode(ShipMode shipMode) {
        this.shipMode = shipMode;
    }

    public double getProfit() {
        return profit;
    }

    public void setProfit(double profit) {
        this.profit = profit;
    }

    public double getUnitPrice() {
        return unitPrice;
    }

    public void setUnitPrice(double unitPrice) {
        this.unitPrice = unitPrice;
    }

    public double getShipping() {
        return shipping;
    }

    public void setShipping(double shipping) {
        this.shipping = shipping;
    }

    public String getCustomer() {
        return customer;
    }

    public void setCustomer(String customer) {
        this.customer = customer;
    }

    public String getProvince() {
        return province;
    }

    public void setProvince(String province) {
        this.province = province;
    }

    public String getRegion() {
        return region;
    }

    public void setRegion(String region) {
        this.region = region;
    }

    public Segment getSegment() {
        return segment;
    }

    public void setSegment(Segment segment) {
        this.segment = segment;
    }

    public String getProductCategory() {
        return productCategory;
    }

    public void setProductCategory(String productCategory) {
        this.productCategory = productCategory;
    }

    public String getProductSubCategory() {
        return productSubCategory;
    }

    public void setProductSubCategory(String productSubCategory) {
        this.productSubCategory = productSubCategory;
    }

    public String getProductName() {
        return productName;
    }

    public void setProductName(String productName) {
        this.productName = productName;
    }

    public String getProductContainer() {
        return productContainer;
    }

    public void setProductContainer(String productContainer) {
        this.productContainer = productContainer;
    }

    public double getBaseMargin() {
        return baseMargin;
    }

    public void setBaseMargin(double baseMargin) {
        this.baseMargin = baseMargin;
    }

    public Date getShipDate() {
        return shipDate;
    }

    public void setShipDate(Date shipDate) {
        this.shipDate = shipDate;
    }

    public enum Priority {
        Low, Medium, High, Critical;

        public static Priority getPriority(String caption) {
            for (Priority priority : Priority.values()) {
                if (priority.equals(caption)) {
                    return priority;
                }
            }
            return null;
        }
    }

    public enum Headings {
        RowID("Row ID"), OrderID("Order ID"), OrderDate("Order Date"), OrderPriority("Order Priority"),
        OrderQuantity("Order Quantity"), Sales("Sales"), SalesDiscount("Sales	Discount"), ShipMode("Ship Mode"),
        Profit("Profit"), UnitPrice("Unit Price"), ShippingCost("Shipping Cost"), CustomerName("Customer Name"),
        Province("Province"), Region("Region"), CustomerSegment("Customer Segment"), ProductCategory("Product Category"),
        ProductSubCategory("Product Sub-Category"), ProductName("Product Name"),
        ProductContainer("Product Container"), ProductBaseMargin("Product Base Margin"), ShipDate("Ship Date");

        String caption;

        Headings(String caption) {
            this.caption = caption;
        }

        public static Headings getHeading(String caption) {
            for (Headings heading : Headings.values()) {
                if (heading.equals(caption)) {
                    return heading;
                }
            }
            return null;
        }
    }

    public enum ShipMode {
        DeliveryTruck("Delivery Truck"), RegularAir("Regular Air"), ExpressAir("Express Air");

        String caption;

        ShipMode(String caption) {
            this.caption = caption;
        }

        public static ShipMode getShipMode(String caption) {
            for (ShipMode mode : ShipMode.values()) {
                if (mode.caption.equals(caption)) {
                    return mode;
                }
            }
            return null;
        }
    }

    public enum Segment {
        Consumer("Consumer"), Corporate("Corporate"), SmallBusiness("Small Business"), HomeOffice("Home Office"),;

        String caption;

        Segment(String caption) {
            this.caption = caption;
        }

        public static Segment getSegment(String caption) {
            for (Segment segment : Segment.values()) {
                if (segment.caption.equals(caption)) {
                    return segment;
                }
            }
            return null;
        }
    }

    public ModelA(List<Object> row) throws ParseException {
        this.rowId = ((Double)row.get(Headings.RowID.ordinal())).longValue();
        this.orderId = ((Double)row.get(Headings.OrderID.ordinal())).longValue();
        this.orderDate = toDate(row.get(Headings.OrderDate.ordinal()));
        this.priority = Priority.getPriority(row.get(Headings.OrderPriority.ordinal()).toString());
        this.quantity = ((Double)row.get(Headings.OrderQuantity.ordinal())).intValue();
        this.sales = (double)row.get(Headings.Sales.ordinal());
        try {
            this.discount = (double) row.get(Headings.SalesDiscount.ordinal());
        } catch (ClassCastException e) {
            this.discount = 0;
        }
        this.shipMode = ShipMode.getShipMode(row.get(Headings.ShipMode.ordinal()).toString());
        this.profit = (double)row.get(Headings.Profit.ordinal());
        this.unitPrice = (double)row.get(Headings.UnitPrice.ordinal());
        this.shipping = (double)row.get(Headings.ShippingCost.ordinal());
        this.customer = row.get(Headings.CustomerName.ordinal()).toString();
        this.province = row.get(Headings.Province.ordinal()).toString();
        this.region = row.get(Headings.Region.ordinal()).toString();
        this.segment = Segment.getSegment(row.get(Headings.CustomerSegment.ordinal()).toString());
        this.productCategory = row.get(Headings.ProductCategory.ordinal()).toString();
        this.productSubCategory = row.get(Headings.ProductSubCategory.ordinal()).toString();
        this.productName = row.get(Headings.ProductName.ordinal()).toString();
        this.productContainer = row.get(Headings.ProductContainer.ordinal()).toString();
        this.baseMargin = (double)row.get(Headings.ProductBaseMargin.ordinal());

        String formula = row.get(Headings.ShipDate.ordinal()).toString();
        int x = Integer.parseInt(formula.substring(2));
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(this.orderDate);
        if (formula.startsWith("+")) {
            calendar.add(Calendar.DAY_OF_MONTH, x);
        } else {
            calendar.add(Calendar.DAY_OF_MONTH, x*-1);
        }
        this.shipDate = calendar.getTime();
    }

    private Date toDate(Object obj) throws ParseException {
        String cellValue = obj.toString();
        cellValue = cellValue.replace("DATE(", "").replace(")", ",");
        String[] tokens = cellValue.split(",");
        if (tokens.length < 4) {
            return new Date();
        }
        String date = tokens[0] + "/" + tokens[1] + "/" + tokens[2];
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
        Date dt = sdf.parse(date);
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(dt);

        String days = tokens[3];
        if (days.startsWith("-")) {
            calendar.add(Calendar.DAY_OF_MONTH, Integer.parseInt(days.substring(1)) * -1);
        } else {
            calendar.add(Calendar.DAY_OF_MONTH, Integer.parseInt(days.substring(1)));
        }

        return calendar.getTime();
    }

    @Override
    public String toString() {
        return "ModelA{" +
                "rowId=" + rowId +
                ", orderId=" + orderId +
                ", orderDate=" + orderDate +
                ", priority=" + priority +
                ", quantity=" + quantity +
                ", sales=" + sales +
                ", discount=" + discount +
                ", shipMode=" + shipMode +
                ", profit=" + profit +
                ", unitPrice=" + unitPrice +
                ", shipping=" + shipping +
                ", customer='" + customer + '\'' +
                ", province='" + province + '\'' +
                ", region='" + region + '\'' +
                ", segment=" + segment +
                ", productCategory='" + productCategory + '\'' +
                ", productSubCategory='" + productSubCategory + '\'' +
                ", productName='" + productName + '\'' +
                ", productContainer='" + productContainer + '\'' +
                ", baseMargin=" + baseMargin +
                ", shipDate=" + shipDate +
                '}';
    }
}

#end
