package com.FirstProject.FirstProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class FirstProjectApplication {

    public static void main(String[] args) {
        SpringApplication.run(FirstProjectApplication.class, args);
        
        if (args.length == 0) {
            System.out.println("Usage: java -jar XlToIncConverter.jar <ExcelFilePath>");
            return;
        }

        try {
            convert(args[0]);
            System.out.println("Conversion complete.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void convert(String excelPath) throws Exception {
        FileInputStream fis = new FileInputStream(new File(excelPath));
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        String outputFileName1 = excelPath.replace("XlInput", "IncOutput").replace(".xlsx", ".sql");
        String outputFileName2 = excelPath.replace("XlInput", "UdeOutput").replace(".xlsx", "_ude.sql");

        int totalRows = 0;
        int udeRowsWritten = 0;
        int skippedBecauseNoInterest = 0;

        try (
                OutputStreamWriter writer1 = new OutputStreamWriter(new FileOutputStream(outputFileName1), StandardCharsets.UTF_8);
                OutputStreamWriter writer2 = new OutputStreamWriter(new FileOutputStream(outputFileName2), StandardCharsets.UTF_8)
        ) {
            DataFormatter df = new DataFormatter();
            SimpleDateFormat inputDateFormat = new SimpleDateFormat("dd-MM-yyyy");
            SimpleDateFormat outputDateFormat = new SimpleDateFormat("dd-MM-yyyy");

            for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
                XSSFSheet sheet = workbook.getSheetAt(s);
                Row headerRow = sheet.getRow(0);
                if (headerRow == null) continue;

                Map<String, Integer> colIndex = new HashMap<>();
                for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                    Cell hCell = headerRow.getCell(c);
                    if (hCell == null) continue;
                    String h = new DataFormatter().formatCellValue(hCell).trim();
                    if (h.isEmpty()) continue;
                    colIndex.put(normalizeHeader(h), c);
                }

                String kCreditCasa = normalizeHeader("CREDIT CASA ACCOUNT");
                String kStaffNo = normalizeHeader("STAFF NO.");
                String kProductCode = normalizeHeader("LOAN PRODUCT CODE");
                String kAmountFin = normalizeHeader("AMOUNT FINANCED");
                String kValueDate = normalizeHeader("VALUE DATE");
                String kMaturityDate = normalizeHeader("MATURITY DATE");
                String kInterest = normalizeHeader("INTEREST RATE");

                if (!colIndex.containsKey(kInterest)) {
                    System.out.println("WARNING: Could not find header for INTEREST RATE.");
                    writer2.write("-- No INTEREST RATE column found. UDE inserts not generated.\n");
                }

                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null || isRowEmpty(row)) continue;
                    totalRows++;

                    String creditCasa = getCellAsString(row, colIndex.get(kCreditCasa), df, evaluator);
                    String staffNo = getCellAsString(row, colIndex.get(kStaffNo), df, evaluator);
                    String productCode = getCellAsString(row, colIndex.get(kProductCode), df, evaluator);
                    String amountFinanced = getCellAsString(row, colIndex.get(kAmountFin), df, evaluator).replace(",", "");
                    String valueDate = getCellAsString(row, colIndex.get(kValueDate), df, evaluator);
                    String maturityDate = getCellAsString(row, colIndex.get(kMaturityDate), df, evaluator);

                    if (creditCasa == null || creditCasa.isEmpty()) continue;
                    String branchCode = creditCasa.substring(0, Math.min(3, creditCasa.length()));

                    valueDate = formatDate(valueDate, inputDateFormat, outputDateFormat);
                    maturityDate = formatDate(maturityDate, inputDateFormat, outputDateFormat);

                    // Main insert
                    String insertSQL1 = String.format(
                            "INSERT INTO cltb_account_upload " +
                                    "(SOURCE_CODE, BRANCH_CODE, ALT_ACC_NO, CUSTOMER_ID, PRODUCT_CODE, PRODUCT_CATEGORY, BOOK_DATE, " +
                                    "VALUE_DATE, MATURITY_DATE, AMOUNT_FINANCED, CURRENCY, CR_PAYMENT_MODE, CR_PROD_AC, " +
                                    "CR_ACC_BRN, USGT_STATUS, UPLOAD_STATUS, JOB_ID, PROCESS_CODE, SEQUENCE_NO, PROP_HANDOVER, HANDOVER_CONF, STATUS_CHANGE_MODE)\n" +
                                    "VALUES ('FCAT', '%s', '%s', (SELECT cust_no FROM sttm_cust_account WHERE cust_ac_no ='%s'), '%s', " +
                                    "(SELECT product_category FROM cltm_product WHERE product_code = '%s'), " +
                                    "(SELECT today FROM sttm_dates WHERE branch_code = '%s'), " +
                                    "TO_DATE('%s','DD-MM-YYYY'), TO_DATE('%s','DD-MM-YYYY'), %s, 'GHS', 'ACC', '%s', " +
                                    "(SELECT branch_code FROM sttm_cust_account WHERE cust_ac_no ='%s'), 'N', 'U', 1, 'BOOK', 1, 'N', 'N', 'A');\n\n",
                            branchCode, staffNo, creditCasa, productCode, productCode, branchCode,
                            valueDate, maturityDate, amountFinanced, creditCasa, creditCasa
                    );
                    writer1.write(insertSQL1);

                    // UDE insert
                    Double interestRate = null;
                    if (colIndex.containsKey(kInterest)) {
                        int idx = colIndex.get(kInterest);
                        interestRate = getCellAsDouble(row, idx, df, evaluator);
                    }

                    if (interestRate != null) {
                        // Scale percentage cells
                        Cell cell = row.getCell(colIndex.get(kInterest));
                        if (cell != null && cell.getCellType() == CellType.NUMERIC &&
                                cell.getCellStyle().getDataFormatString().contains("%")) {
                            interestRate *= 100;
                        }

                        String udeInsert = String.format(
                                "INSERT INTO cltbs_ac_ude_upload " +
                                        "(BRANCH_CODE, EFFECTIVE_DATE, UDE_ID, UDE_VALUE, SOURCE_CODE, EXT_REF_NO, RESOLVED_VALUE, MAINT_RSLV_FLAG, PROCESS_CODE, SEQUENCE_NO)\n" +
                                        "VALUES ('%s', TO_DATE('%s','DD-MM-YYYY'), 'INTEREST_RATE', %s, 'FCAT', '%s', %s, 'M', 'BOOK', 1);\n\n",
                                branchCode, valueDate, formatDoubleForSql(interestRate), staffNo, formatDoubleForSql(interestRate)
                        );
                        writer2.write(udeInsert);
                        udeRowsWritten++;
                    } else {
                        skippedBecauseNoInterest++;
                        writer2.write(String.format("-- Row %d skipped: no numeric INTEREST_RATE (creditCasa=%s, staff=%s)\n",
                                r + 1, creditCasa, staffNo == null ? "" : staffNo));
                    }
                }
            }

            writer2.write(String.format("\n-- Summary: total rows processed=%d, ude inserts written=%d, skipped=%d\n",
                    totalRows, udeRowsWritten, skippedBecauseNoInterest));

            System.out.println("Total rows processed: " + totalRows);
            System.out.println("UDE inserts written: " + udeRowsWritten);
            System.out.println("Rows skipped (no numeric interest): " + skippedBecauseNoInterest);

        } finally {
            workbook.close();
            fis.close();
        }
    }

    private static String normalizeHeader(String h) {
        if (h == null) return "";
        return h.trim().toLowerCase().replaceAll("\\s+", " ").replaceAll("[ %()]", "");
    }

    private static boolean isRowEmpty(Row row) {
        short last = row.getLastCellNum();
        if (last < 0) return true;
        for (int i = 0; i < last; i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK && !new DataFormatter().formatCellValue(cell).trim().isEmpty())
                return false;
        }
        return true;
    }

    private static String getCellAsString(Row row, Integer colIdx, DataFormatter df, FormulaEvaluator evaluator) {
        if (colIdx == null) return "";
        Cell cell = row.getCell(colIdx);
        if (cell == null) return "";
        if (cell.getCellType() == CellType.FORMULA) {
            evaluator.evaluateInCell(cell);
        }
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            Date d = cell.getDateCellValue();
            return new SimpleDateFormat("dd-MM-yyyy").format(d);
        }
        return df.formatCellValue(cell).trim();
    }

    private static Double getCellAsDouble(Row row, int colIdx, DataFormatter df, FormulaEvaluator evaluator) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) return null;
        if (cell.getCellType() == CellType.FORMULA) {
            evaluator.evaluateInCell(cell);
        }
        if (cell.getCellType() == CellType.NUMERIC && !DateUtil.isCellDateFormatted(cell)) {
            return cell.getNumericCellValue();
        }
        String s = df.formatCellValue(cell).trim().replace("%", "").replace(",", "");
        try {
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return null;
        }
    }
    private static String formatDoubleForSql(Double d) {
        if (d == null) return "NULL";
        if (d == Math.rint(d)) {
            return String.format("%.0f", d);
        } else {
            return d.toString();
        }
    }

    private static String formatDate(String input, SimpleDateFormat inFmt, SimpleDateFormat outFmt) {
        try {
            if (input == null || input.trim().isEmpty()) return "";
            Date d = inFmt.parse(input);
            return outFmt.format(d);
        } catch (ParseException e) {
            return input;
        }
    }
}
