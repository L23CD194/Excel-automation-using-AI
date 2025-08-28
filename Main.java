    import java.io.*;
    import java.text.*;
    import java.util.*;
    import org.apache.poi.ss.usermodel.*;
    import org.apache.poi.xssf.usermodel.*;
    import org.apache.poi.ss.util.CellRangeAddress;
    import org.apache.poi.xddf.usermodel.chart.*;
    import org.apache.poi.xddf.usermodel.*;
    import org.apache.poi.ss.usermodel.Font;

    public class Main {

        // === Config ===
        private static final String INPUT_FILE = "shop_inventory.xlsx";
        private static final String OUTPUT_FILE = "processed_inventory.xlsx";

        // Input columns (0-based)
        private static final int COL_ITEM = 1;
        private static final int COL_QTY = 2;
        private static final int COL_COST = 3;
        private static final int COL_SELL = 4;
        private static final int COL_PROFIT = 5;
        private static final int COL_CATEGORY = 6;
        private static final int COL_EXPIRY_INPUT = 7;

        // Output computed columns (in Data sheet)
        private static final int OUT_COL_PROFIT = 5;
        private static final int OUT_COL_EXPIRY_STATUS = 8;
        private static final int OUT_COL_DUPLICATE = 9;
        private static final int OUT_COL_RECOMM = 10;

        // Date formats to try for string expiry values
        private static final String[] DATE_FORMATS = {
            "d/M/yyyy", "dd/MM/yyyy", "dd-MM-yyyy", "d-M-yyyy", "d MMM yyyy", "dd MMM yyyy",
            "yyyy-MM-dd", "yyyy/MM/dd", "MM/dd/yyyy", "MMM d, yyyy", "MMMM d, yyyy",
            "MMM d yyyy", "MMMM d yyyy", "d-MMM-yyyy", "dd MMMM yyyy"
        };

        public static void main(String[] args) {
            System.out.println("▶ Processing: " + INPUT_FILE);
            try (
                FileInputStream fis = new FileInputStream(INPUT_FILE);
                Workbook inWb = new XSSFWorkbook(fis);
                Workbook outWb = new XSSFWorkbook()
            ) {
                Sheet inSheet = inWb.getSheetAt(0);

                // ====== Create output sheets ======
                XSSFSheet dataSheet = (XSSFSheet) outWb.createSheet("Data");
                XSSFSheet dashSheet = (XSSFSheet) outWb.createSheet("Dashboard");

                // ====== Basic styles ======
                CellStyle headerStyle = outWb.createCellStyle();
                Font headerFont = outWb.createFont();
                headerFont.setBold(true);
                headerStyle.setFont(headerFont);
                headerStyle.setAlignment(HorizontalAlignment.CENTER);
                headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                headerStyle.setBorderBottom(BorderStyle.THIN);

                CellStyle normalStyle = outWb.createCellStyle();
                normalStyle.setVerticalAlignment(VerticalAlignment.CENTER);

                // ====== Counters for dashboard ======
                int expiredCount = 0, nearCount = 0, validCount = 0, noExpiryCount = 0;

                // ====== Duplicate detection ======
                Set<String> seen = new HashSet<>();
                Set<String> duplicatesSet = new LinkedHashSet<>();

                // ====== Copy header to Data sheet + add new headers ======
                Row inHeader = inSheet.getRow(0);
                Row outHeader = dataSheet.createRow(0);
                for (int c = 0; c <= 7; c++) {
                    Cell src = inHeader.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Cell dst = outHeader.createCell(c);
                    dst.setCellValue(cellToString(src));
                    dst.setCellStyle(headerStyle);
                }

                setHeader(outHeader, OUT_COL_PROFIT, "Profit", headerStyle);
                setHeader(outHeader, OUT_COL_EXPIRY_STATUS, "Expiry Status", headerStyle);
                setHeader(outHeader, OUT_COL_DUPLICATE, "Duplicate?", headerStyle);
                setHeader(outHeader, OUT_COL_RECOMM, "Recommendation", headerStyle);

                int outRowIdx = 1;

                // ====== Process each input row ======
                for (int r = 1; r <= inSheet.getLastRowNum(); r++) {
                    Row inRow = inSheet.getRow(r);
                    if (inRow == null) continue;

                    String itemName = cellToString(inRow.getCell(COL_ITEM)).trim();
                    if (itemName.isEmpty()) continue;

                    String quantity = cellToString(inRow.getCell(COL_QTY)).trim();
                    String category = cellToString(inRow.getCell(COL_CATEGORY)).trim();

                    double cost = safeGetNumeric(inRow.getCell(COL_COST));
                    double sell = safeGetNumeric(inRow.getCell(COL_SELL));
                    double profit = sell - cost;

                    Cell expiryCell = inRow.getCell(COL_EXPIRY_INPUT, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String expiryStatus;
                    Date expiryDate = parseDateFromCell(expiryCell);

                    if (expiryCell == null || isBlankCell(expiryCell)) {
                        expiryStatus = "No Expiry";
                        noExpiryCount++;
                    } else if (expiryDate == null) {
                        String raw = cellToString(expiryCell).trim();
                        if (raw.equalsIgnoreCase("no expiry") ||
                            raw.equalsIgnoreCase("none") ||
                            raw.equalsIgnoreCase("n/a") ||
                            raw.equalsIgnoreCase("na")) {
                            expiryStatus = "No Expiry";
                            noExpiryCount++;
                        } else {
                            expiryStatus = "Invalid Date";
                        }
                    } else {
                        expiryStatus = classifyExpiry(expiryDate);
                        switch (expiryStatus) {
                            case "❌ Expired": expiredCount++; break;
                            case "⚠️ Near Expiry": nearCount++; break;
                            case "✅ Valid": validCount++; break;
                        }
                    }

                    String key = itemName.toLowerCase();
                    boolean isDup = !seen.add(key);
                    if (isDup) duplicatesSet.add(itemName);

                    String recommendation = "-";
                    if ("❌ Expired".equals(expiryStatus)) {
                        recommendation = "🛒 Reorder: " + itemName;
                    } else if ("⚠️ Near Expiry".equals(expiryStatus)) {
                        recommendation = "🔖 Discount / Move front";
                    } else if (profit < 5) {
                        recommendation = "Low profit — review price";
                    } else {
                        recommendation = "Stock OK";
                    }

                    Row outRow = dataSheet.createRow(outRowIdx++);
                    for (int c = 0; c <= 7; c++) {
                        Cell src = inRow.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        Cell dst = outRow.createCell(c);
                        if (c == COL_EXPIRY_INPUT) {
                            Date norm = parseDateFromCell(src);
                            if (norm != null) {
                                dst.setCellValue(new SimpleDateFormat("yyyy-MM-dd").format(norm));
                            } else {
                                dst.setCellValue(cellToString(src));
                            }
                        } else if (src.getCellType() == CellType.NUMERIC) {
                            dst.setCellValue(src.getNumericCellValue());
                        } else if (src.getCellType() == CellType.BOOLEAN) {
                            dst.setCellValue(src.getBooleanCellValue());
                        } else {
                            dst.setCellValue(cellToString(src));
                        }
                        dst.setCellStyle(normalStyle);
                    }

                    outRow.createCell(OUT_COL_PROFIT).setCellValue(profit);
                    outRow.createCell(OUT_COL_EXPIRY_STATUS).setCellValue(expiryStatus);
                    outRow.createCell(OUT_COL_DUPLICATE).setCellValue(isDup ? "Yes" : "No");
                    outRow.createCell(OUT_COL_RECOMM).setCellValue(recommendation);
                }

                for (int c = 0; c <= OUT_COL_RECOMM; c++) {
                    dataSheet.autoSizeColumn(c);
                }

                buildDashboard(dashSheet, outWb, expiredCount, nearCount, validCount, noExpiryCount);

                try (FileOutputStream fos = new FileOutputStream(OUTPUT_FILE)) {
                    outWb.write(fos);
                }

                System.out.println("✅ Done. Created: " + OUTPUT_FILE);
                System.out.println(" Status counts -> Expired: " + expiredCount + ", Near: " + nearCount +
                                   ", Valid: " + validCount + ", No Expiry: " + noExpiryCount);
                System.out.println(" Duplicates: " + duplicatesSet.size() + " -> " + duplicatesSet);

            } catch (FileNotFoundException fnf) {
                System.err.println("❌ Input file not found: " + INPUT_FILE);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        // ---------- Dashboard builder ----------
        private static void buildDashboard(XSSFSheet dashSheet, Workbook wb,
                                           int expired, int near, int valid, int noExpiry) {
            Row title = dashSheet.createRow(0);
            Cell t = title.createCell(0);
            t.setCellValue("Inventory Dashboard");

            CellStyle titleStyle = wb.createCellStyle();
            Font f = wb.createFont();
            f.setBold(true);
            f.setFontHeightInPoints((short) 14);
            titleStyle.setFont(f);
            t.setCellStyle(titleStyle);

            int startRow = 2;
            Row h = dashSheet.createRow(startRow);
            h.createCell(0).setCellValue("Expiry Status");
            h.createCell(1).setCellValue("Count");

            Object[][] rows = {
                { "❌ Expired", expired },
                { "⚠️ Near Expiry", near },
                { "✅ Valid", valid },
                { "No Expiry", noExpiry }
            };

            for (int i = 0; i < rows.length; i++) {
                Row r = dashSheet.createRow(startRow + 1 + i);
                r.createCell(0).setCellValue((String) rows[i][0]);
                r.createCell(1).setCellValue((Integer) rows[i][1]);
            }

            dashSheet.autoSizeColumn(0);
            dashSheet.autoSizeColumn(1);

            createBarChart(dashSheet, startRow, wb);
        }

        // ---------- Chart ----------
        private static void createBarChart(Sheet sheet, int startRow, Workbook wb) {
            try {
                XSSFDrawing drawing = (XSSFDrawing) ((XSSFSheet) sheet).createDrawingPatriarch();
                XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, startRow, 10, startRow + 16);

                XSSFChart chart = drawing.createChart(anchor);
                chart.setTitleText("Expiry Status Summary");
                chart.setTitleOverlay(false);

                XDDFChartLegend legend = chart.getOrAddLegend();
                legend.setPosition(LegendPosition.RIGHT);

                XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
                bottomAxis.setTitle("Status");
                XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
                leftAxis.setTitle("Count");

                String sheetName = sheet.getSheetName();
                String catRange = sheetName + "!$A$" + (startRow + 2) + ":$A$" + (startRow + 5);
                String valRange = sheetName + "!$B$" + (startRow + 2) + ":$B$" + (startRow + 5);

                XDDFDataSource<String> categories =
                    XDDFDataSourcesFactory.fromStringCellRange((XSSFSheet) sheet, CellRangeAddress.valueOf(catRange));
                XDDFNumericalDataSource<Double> values =
                    XDDFDataSourcesFactory.fromNumericCellRange((XSSFSheet) sheet, CellRangeAddress.valueOf(valRange));

                XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
                data.setVaryColors(true);
                XDDFChartData.Series s = data.addSeries(categories, values);
                s.setTitle("Items", null);
                chart.plot(data);
            } catch (Exception e) {
                System.err.println("Chart creation failed: " + e.getMessage());
            }
        }

        // ---------- Helpers ----------
        private static void setHeader(Row header, int col, String text, CellStyle style) {
            Cell c = header.createCell(col);
            c.setCellValue(text);
            c.setCellStyle(style);
        }

        private static boolean isBlankCell(Cell c) {
            if (c == null) return true;
            if (c.getCellType() == CellType.BLANK) return true;
            if (c.getCellType() == CellType.STRING && c.getStringCellValue().trim().isEmpty()) return true;
            return false;
        }

        private static String cellToString(Cell cell) {
            if (cell == null) return "";
            try {
                switch (cell.getCellType()) {
                    case STRING: return cell.getStringCellValue();
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            Date d = cell.getDateCellValue();
                            return new SimpleDateFormat("yyyy-MM-dd").format(d);
                        } else {
                            double v = cell.getNumericCellValue();
                            if (v == (long) v) return Long.toString((long) v);
                            return Double.toString(v);
                        }
                    case BOOLEAN: return Boolean.toString(cell.getBooleanCellValue());
                    case FORMULA:
                        try {
                            FormulaEvaluator ev = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                            CellValue cv = ev.evaluate(cell);
                            if (cv.getCellType() == CellType.STRING) return cv.getStringValue();
                            if (cv.getCellType() == CellType.NUMERIC) return Double.toString(cv.getNumberValue());
                            if (cv.getCellType() == CellType.BOOLEAN) return Boolean.toString(cv.getBooleanValue());
                        } catch (Exception ignored) { }
                        return cell.getCellFormula();
                    default: return cell.toString();
                }
            } catch (Exception e) {
                return "";
            }
        }

        private static double safeGetNumeric(Cell c) {
            if (c == null) return 0;
            try {
                if (c.getCellType() == CellType.NUMERIC) return c.getNumericCellValue();
                if (c.getCellType() == CellType.STRING) {
                    String s = c.getStringCellValue().trim().replaceAll("[^0-9.\\-]", "");
                    if (s.isEmpty()) return 0;
                    return Double.parseDouble(s);
                }
                if (c.getCellType() == CellType.FORMULA) {
                    FormulaEvaluator ev = c.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue cv = ev.evaluate(c);
                    if (cv.getCellType() == CellType.NUMERIC) return cv.getNumberValue();
                }
            } catch (Exception ignored) { }
            return 0;
        }

        private static Date parseDateFromCell(Cell c) {
            if (c == null) return null;
            try {
                if (c.getCellType() == CellType.NUMERIC) {
                    double n = c.getNumericCellValue();
                    if (DateUtil.isCellDateFormatted(c) || DateUtil.isValidExcelDate(n)) {
                        return DateUtil.getJavaDate(n);
                    }
                } else {
                    String s = cellToString(c).trim();
                    if (s.isEmpty()) return null;
                    Date parsed = tryParseStringDate(s);
                    if (parsed != null) return parsed;
                }
            } catch (Exception ignored) { }
            return null;
        }

        private static Date tryParseStringDate(String s) {
            if (s == null || s.trim().isEmpty()) return null;
            s = s.trim();
            for (String fmt : DATE_FORMATS) {
                try {
                    SimpleDateFormat sdf = new SimpleDateFormat(fmt, Locale.ENGLISH);
                    sdf.setLenient(false);
                    return sdf.parse(s);
                } catch (Exception ignored) { }
            }
            String noComma = s.replace(",", "");
            if (!noComma.equals(s)) {
                for (String fmt : DATE_FORMATS) {
                    try {
                        SimpleDateFormat sdf = new SimpleDateFormat(fmt, Locale.ENGLISH);
                        sdf.setLenient(false);
                        return sdf.parse(noComma);
                    } catch (Exception ignored) { }
                }
            }
            return null;
        }

        private static String classifyExpiry(Date expiryDate) {
            Date now = new Date();
            long diffMs = expiryDate.getTime() - now.getTime();
            long days = diffMs / (1000L * 60 * 60 * 24);
            if (days < 0) return "❌ Expired";
            if (days <= 30) return "⚠️ Near Expiry";
            return "✅ Valid";
        }
    }