package com.example.etoll;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// working properly and giving current output files

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * EtollVoucherGenerator - standalone version that reads dsr_report.xlsx from project root,
 * uses RUN_NUMBER=1 and OUTPUT_ROOT=E-tollAcquiringSettlement/Processing.
 *
 * Produces an XLSX with two sheets:
 *   - Voucher
 *   - Upload
 *
 * Usage:
 *   mvn package
 *   java -jar target/etoll-1.0.0.jar
 */
public class EtollVoucherGenerator {

    // ---------------- CONFIG (fixed as requested) ----------------
    private static final String DSR_FILE = "dsr_report.xlsx";
    private static final int RUN_NUMBER = 1;
    private static final Path OUTPUT_ROOT = Paths.get("E-tollAcquiringSettlement", "Processing");

    // Column names (must match Excel)
    private static final String COL_SETTLEMENT_DATE   = "Settlement Date";
    private static final String COL_TRANSACTION_CYCLE = "Transaction Cycle";
    private static final String COL_TRANSACTION_TYPE  = "Transaction Type";
    private static final String COL_CHANNEL           = "Channel";
    private static final String COL_SETAMTDR          = "SETAMTDR";
    private static final String COL_SETAMTCR          = "SETAMTCR";
    private static final String COL_SERVICE_FEE_DR    = "Service Fee Amt Dr";
    private static final String COL_SERVICE_FEE_CR    = "Service Fee Amt Cr";
    private static final String COL_FINAL_NET_AMT     = "Final Net Amt";
    private static final String COL_INWARD_OUTWARD    = "Inward/Outward";

    // TEMPLATE (unchanged)
    private static final List<TemplateRow> TEMPLATE = List.of(
            new TemplateRow("0103SLRGTSRC", "NPCIR5{yyyymmdd} {ddmmyy}_{cycle} ETCAC", "Final Net Amt"),
            new TemplateRow("", "", ""),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy}_{cycle}", "NETC Settled Transaction"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} Dr.Adj_{cycle}", "Debit Adjustment"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} GF Accp_{cycle}", "Good Faith Acceptance Credit"),
            new TemplateRow("", "", ""),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} Cr.Adj_{cycle}", "Credit Adjustment"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} Chbk_{cycle}", "Chargeback Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} GF Accp_{cycle}", "Good Faith Acceptance Debit"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} PrArbtAc_{cycle}", "Pre-Arbitration Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} DrPrAbAc_{cycle}", "Pre-Arbitration Deemed Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} DrChbAc_{cycle}", "Debit chargeback deemed Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} ArbtAc_{cycle}", "Arbitration Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {dd_mm_yy} ArbtVer_{cycle}", "Arbitration Vedict"),
            new TemplateRow("", "", ""),
            new TemplateRow("0103CNETCACQ", "Etoll acq {dd_mm_yy}_{cycle}", "Income Debit"),
            new TemplateRow("0103SLPPCIGT", "Etoll acq {dd_mm_yy}_{cycle}", "GST Debit"),
            new TemplateRow("0103CNETCACQ", "Etoll acq {dd_mm_yy}_{cycle}", "Income Credit"),
            new TemplateRow("0103SLPPCIGT", "Etoll acq {dd_mm_yy}_{cycle}", "GST Credit")
    );

    // RULES (unchanged)
    private static final Map<String, Rule> RULES = initRules();
    private static Map<String, Rule> initRules() {
        Map<String, Rule> m = new HashMap<>();
        m.put("NETC Settled Transaction", new Rule(List.of("netc settled transaction"), COL_SETAMTCR, "credit", null));
        m.put("Debit Adjustment", new Rule(List.of("debitadjustment", "debit adjustment"), COL_SETAMTCR, "credit", null));
        m.put("Good Faith Acceptance Credit", new Rule(List.of("good faith acceptance"), COL_SETAMTCR, "credit", null));
        m.put("Credit Adjustment", new Rule(List.of("credit adjustment"), COL_SETAMTDR, "debit", null));
        m.put("Chargeback Acceptance", new Rule(List.of("chargeback acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Good Faith Acceptance Debit", new Rule(List.of("good faith acceptance"), null, "goodfaith", null));
        m.put("Pre-Arbitration Acceptance", new Rule(List.of("pre-arbitration acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Pre-Arbitration Deemed Acceptance", new Rule(List.of("pre-arbitration deemed acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Debit chargeback deemed Acceptance", new Rule(List.of("debit chargeback deemed acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Arbitration Acceptance", new Rule(List.of("arbitration acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Arbitration Vedict", new Rule(List.of("arbitration vedict"), COL_SETAMTDR, "debit", null));
        m.put("Income Debit", new Rule(null, null, null, "inward_dr"));
        m.put("GST Debit", new Rule(null, null, null, "inward_dr"));
        m.put("Income Credit", new Rule(null, null, null, "inward_cr"));
        m.put("GST Credit", new Rule(null, null, null, "inward_cr"));
        m.put("Final Net Amt", new Rule(null, null, null, "final"));
        return m;
    }

    public static void main(String[] args) {
        System.out.println("Starting EtollVoucherGenerator...");
        try {
            Path dsr = Paths.get(DSR_FILE);
            if (!Files.exists(dsr)) {
                System.err.println("Error: " + DSR_FILE + " not found in current directory: " + Paths.get("").toAbsolutePath());
                System.exit(1);
            }
            Map<String,Object> result = generateVoucher(dsr);
            System.out.println("Result: " + result);
            if ("ok".equals(result.get("status"))) {
                System.out.println("Done. Output saved to: " + result.get("path"));
                System.exit(0);
            } else {
                System.err.println("Completed with error. Details: " + result);
                System.exit(2);
            }
        } catch (Throwable t) {
            t.printStackTrace();
            System.exit(3);
        }
    }

    public static Map<String,Object> generateVoucher(Path dsrPath) throws Exception {
        // load workbook
        Workbook wb;
        try (InputStream is = Files.newInputStream(dsrPath, StandardOpenOption.READ)) {
            wb = WorkbookFactory.create(is);
        }

        Sheet sheet = wb.getSheetAt(0); // first sheet assumed
        List<Map<String,String>> rows = readSheetToMaps(sheet);

        // forward-fill TC & TT
        forwardFill(rows, COL_TRANSACTION_CYCLE);
        forwardFill(rows, COL_TRANSACTION_TYPE);

        // find settlement date
        LocalDate settlement = findSettlementDate(rows);
        if (settlement == null) settlement = LocalDate.now(ZoneId.systemDefault());

        String yyyymmdd = settlement.format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        String ddmmyy   = settlement.format(DateTimeFormatter.ofPattern("ddMMyy"));
        String dd_mm_yy = settlement.format(DateTimeFormatter.ofPattern("dd.MM.yy"));
        String cycle = RUN_NUMBER + "C";

        // normalize helpers
        for (Map<String,String> r : rows) {
            r.put("TC", safeTrimLower(r.getOrDefault(COL_TRANSACTION_CYCLE,"")));
            r.put("TT", safeTrimLower(r.getOrDefault(COL_TRANSACTION_TYPE,"")));
            r.put("CH", safeTrimLower(r.getOrDefault(COL_CHANNEL,"")));
        }

        // Final net amt: last non-empty
        BigDecimal totalFinal = BigDecimal.ZERO;
        for (int i = rows.size()-1; i >=0; --i) {
            String v = rows.get(i).getOrDefault(COL_FINAL_NET_AMT,"").trim();
            if (!v.isEmpty()) {
                totalFinal = toDecimal(v);
                totalFinal = round2(totalFinal);
                break;
            }
        }
        System.out.println("Final Net Amt (Rightmost+Lowest) = " + totalFinal);

        // Inward GST detection
        BigDecimal incomeDebit = BigDecimal.ZERO, incomeCredit = BigDecimal.ZERO, gstDebit = BigDecimal.ZERO, gstCredit = BigDecimal.ZERO;
        for (int i=0;i<rows.size();i++) {
            String io = rows.get(i).getOrDefault(COL_INWARD_OUTWARD,"");
            if ("INWARD GST".equalsIgnoreCase(io.trim())) {
                if (i>0) {
                    Map<String,String> ra = rows.get(i-1);
                    incomeDebit = round2(toDecimal(ra.getOrDefault(COL_SERVICE_FEE_DR,"0")));
                    incomeCredit = round2(toDecimal(ra.getOrDefault(COL_SERVICE_FEE_CR,"0")));
                }
                Map<String,String> rg = rows.get(i);
                gstDebit = round2(toDecimal(rg.getOrDefault(COL_SERVICE_FEE_DR,"0")));
                gstCredit = round2(toDecimal(rg.getOrDefault(COL_SERVICE_FEE_CR,"0")));
                break;
            }
        }
        System.out.println("Derived INWARD values -> Income Debit: " + incomeDebit + ", GST Debit: " + gstDebit
                + ", Income Credit: " + incomeCredit + ", GST Credit: " + gstCredit);

        // Build voucher list
        List<VoucherRow> voucher = new ArrayList<>();
        for (TemplateRow t : TEMPLATE) {
            String acct = t.accountNo;
            String tmpl = t.template;
            String desc = t.description;

            String narration = tmpl.replace("{yyyymmdd}", yyyymmdd)
                    .replace("{ddmmyy}", ddmmyy)
                    .replace("{dd_mm_yy}", dd_mm_yy)
                    .replace("{cycle}", cycle);

            if (acct.isEmpty() && desc.isEmpty()) {
                voucher.add(new VoucherRow("", null, null, narration, desc));
                continue;
            }

            Rule rule = RULES.getOrDefault(desc, null);

            if (rule != null && "final".equals(rule.special)) {
                voucher.add(new VoucherRow(acct, totalFinal.equals(BigDecimal.ZERO) ? null : totalFinal, null, narration, desc));
                continue;
            }

            if (rule != null && "inward_dr".equals(rule.special)) {
                BigDecimal amt = "Income Debit".equals(desc) ? incomeDebit : gstDebit;
                voucher.add(new VoucherRow(acct, amt.equals(BigDecimal.ZERO) ? null : amt, null, narration, desc));
                continue;
            }

            if (rule != null && "inward_cr".equals(rule.special)) {
                BigDecimal amt = "Income Credit".equals(desc) ? incomeCredit : gstCredit;
                voucher.add(new VoucherRow(acct, null, amt.equals(BigDecimal.ZERO) ? null : amt, narration, desc));
                continue;
            }

            if ("Arbitration Vedict".equals(desc)) {
                BigDecimal amt = BigDecimal.ZERO;
                for (Map<String,String> r : rows) {
                    if ("arbitration vedict".equals(r.getOrDefault("TC","")) &&
                            List.of("debit","non_fin").contains(r.getOrDefault("TT","").toLowerCase()) &&
                            r.getOrDefault(COL_CHANNEL,"").trim().length()>0) {
                        amt = amt.add(toDecimal(r.getOrDefault(COL_SETAMTDR,"0")));
                    }
                }
                amt = round2(amt);
                voucher.add(new VoucherRow(acct, amt.equals(BigDecimal.ZERO) ? null : amt, null, narration, desc));
                System.out.println("Arbitration Vedict: summed SETAMTDR = " + amt);
                continue;
            }

            List<String> cycles = rule != null && rule.cycles != null ? rule.cycles : List.of();
            List<Map<String,String>> sel = new ArrayList<>();
            for (Map<String,String> r : rows) {
                if (cycles.contains(r.getOrDefault("TC",""))) sel.add(r);
            }

            String sumCol = rule != null ? rule.sumCol : null;
            String side = rule != null ? rule.side : "debit";

            if ("goodfaith".equals(side)) {
                BigDecimal dr = BigDecimal.ZERO, cr = BigDecimal.ZERO;
                for (Map<String,String> r : sel) {
                    dr = dr.add(toDecimal(r.getOrDefault(COL_SETAMTDR,"0")));
                    cr = cr.add(toDecimal(r.getOrDefault(COL_SETAMTCR,"0")));
                }
                dr = round2(dr); cr = round2(cr);
                if (dr.compareTo(BigDecimal.ZERO) != 0) voucher.add(new VoucherRow(acct, dr, null, narration, desc));
                else if (cr.compareTo(BigDecimal.ZERO) != 0) voucher.add(new VoucherRow(acct, null, cr, narration, desc));
                else voucher.add(new VoucherRow(acct, null, null, narration, desc));
                continue;
            }

            BigDecimal amt = BigDecimal.ZERO;
            if (sumCol != null) {
                for (Map<String,String> r : sel) {
                    amt = amt.add(toDecimal(r.getOrDefault(sumCol,"0")));
                }
                amt = round2(amt);
            }

            if ("credit".equals(side)) voucher.add(new VoucherRow(acct, null, amt.equals(BigDecimal.ZERO) ? null : amt, narration, desc));
            else voucher.add(new VoucherRow(acct, amt.equals(BigDecimal.ZERO) ? null : amt, null, narration, desc));
        }

        // Build upload rows
        List<List<Object>> uploadRows = new ArrayList<>();
        uploadRows.add(List.of("Account No","C/D","Amount","Narration"));
        for (VoucherRow vr : voucher) {
            BigDecimal d = vr.debit;
            BigDecimal c = vr.credit;
            String cd;
            BigDecimal amount;
            if (d != null && d.compareTo(BigDecimal.ZERO) != 0) { cd = "D"; amount = d; }
            else if (c != null && c.compareTo(BigDecimal.ZERO) != 0) { cd = "C"; amount = c; }
            else { cd = ""; amount = null; }
            if (amount != null) uploadRows.add(List.of(vr.accountNo, cd, amount.doubleValue(), vr.narration));
        }

        // TALLY
        BigDecimal dTotal = BigDecimal.ZERO, cTotal = BigDecimal.ZERO;
        for (VoucherRow vr : voucher) {
            if (vr.debit != null) dTotal = dTotal.add(vr.debit);
            if (vr.credit != null) cTotal = cTotal.add(vr.credit);
        }
        dTotal = round2(dTotal); cTotal = round2(cTotal);
        System.out.println("Voucher totals -> Debit: " + dTotal + " Credit: " + cTotal);

        // create output folder
        Path folder = OUTPUT_ROOT.resolve(String.valueOf(settlement.getYear()))
                .resolve(String.format("%02d", settlement.getMonthValue()))
                .resolve(String.format("%02d", settlement.getDayOfMonth()));
        Files.createDirectories(folder);

        String ddmmyyStr = ddmmyy;
        Path file = folder.resolve("ETOLL_ACQUIRING_VOUCHER_" + ddmmyyStr + "_N" + RUN_NUMBER + ".xlsx");
        Path errFile = folder.resolve("ERROR_ETOLL_ACQUIRING_VOUCHER_" + ddmmyyStr + "_N" + RUN_NUMBER + ".xlsx");

        if (dTotal.compareTo(cTotal) != 0) {
            // write error workbook
            try (Workbook out = new XSSFWorkbook()) {
                writeVoucherSheet(out, voucher);
                writeUploadSheet(out, uploadRows);
                try (OutputStream os = Files.newOutputStream(errFile, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
                    out.write(os);
                }
            }
            return Map.of("status","error","message","Debit and credit not tallied","path",errFile.toString(),"debit",dTotal,"credit",cTotal);
        }

        try (Workbook out = new XSSFWorkbook()) {
            writeVoucherSheet(out, voucher);
            writeUploadSheet(out, uploadRows);
            try (OutputStream os = Files.newOutputStream(file, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
                out.write(os);
            }
        }

        return Map.of("status","ok","path",file.toString(),"debit",dTotal,"credit",cTotal);
    }

    // ---------- helpers ----------
    private static List<Map<String,String>> readSheetToMaps(Sheet sheet) {
        List<Map<String,String>> rows = new ArrayList<>();
        Iterator<Row> it = sheet.iterator();
        if (!it.hasNext()) return rows;
        Row header = it.next();
        List<String> headers = new ArrayList<>();
        for (Cell c : header) {
            headers.add(c.getStringCellValue().trim());
        }
        while (it.hasNext()) {
            Row r = it.next();
            Map<String,String> map = new HashMap<>();
            for (int i=0;i<headers.size();i++) {
                Cell cell = r.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String val = cellToString(cell);
                map.put(headers.get(i), val);
            }
            rows.add(map);
        }
        return rows;
    }

    private static String cellToString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toLocalDate().toString();
                } else {
                    return BigDecimal.valueOf(cell.getNumericCellValue()).stripTrailingZeros().toPlainString();
                }
            case BOOLEAN: return Boolean.toString(cell.getBooleanCellValue());
            case BLANK: return "";
            case FORMULA:
                try {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getLocalDateTimeCellValue().toLocalDate().toString();
                    }
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            default: return cell.toString();
        }
    }

    private static void forwardFill(List<Map<String,String>> rows, String column) {
        String last = "";
        for (Map<String,String> r : rows) {
            String v = r.getOrDefault(column,"");
            if (v == null) v = "";
            v = v.trim();
            if (!v.isEmpty()) last = v;
            else r.put(column, last);
        }
    }

    private static LocalDate findSettlementDate(List<Map<String,String>> rows) {
        for (Map<String,String> r : rows) {
            String v = r.getOrDefault(COL_SETTLEMENT_DATE,"").trim();
            if (v.isEmpty()) continue;
            try {
                if (v.matches("\\d{4}-\\d{2}-\\d{2}")) return LocalDate.parse(v, DateTimeFormatter.ISO_LOCAL_DATE);
                if (v.matches("\\d{2}-\\d{2}-\\d{4}")) return LocalDate.parse(v, DateTimeFormatter.ofPattern("dd-MM-yyyy"));
                if (v.matches("\\d{2}/\\d{2}/\\d{4}")) return LocalDate.parse(v, DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                return LocalDate.parse(v);
            } catch (Exception ignored) {}
        }
        return null;
    }

    private static BigDecimal toDecimal(String s) {
        if (s == null) return BigDecimal.ZERO;
        s = s.trim().replace(",","");
        if (s.isEmpty() || "nan".equalsIgnoreCase(s)) return BigDecimal.ZERO;
        try {
            return new BigDecimal(s);
        } catch (Exception ex) {
            try {
                double d = Double.parseDouble(s);
                return BigDecimal.valueOf(d);
            } catch (Exception ex2) {
                return BigDecimal.ZERO;
            }
        }
    }

    private static BigDecimal round2(BigDecimal d) {
        return d.setScale(2, RoundingMode.HALF_UP);
    }

    private static String safeTrimLower(String s) {
        if (s == null) return "";
        return s.trim().toLowerCase();
    }

    private static void writeVoucherSheet(Workbook wb, List<VoucherRow> voucher) {
        Sheet sh = wb.createSheet("Voucher");
        Row header = sh.createRow(0);
        header.createCell(0).setCellValue("Account No");
        header.createCell(1).setCellValue("Debit");
        header.createCell(2).setCellValue("Credit");
        header.createCell(3).setCellValue("Narration");
        header.createCell(4).setCellValue("Description");

        int r = 1;
        for (VoucherRow vr : voucher) {
            Row row = sh.createRow(r++);
            row.createCell(0).setCellValue(vr.accountNo);
            if (vr.debit != null) row.createCell(1).setCellValue(vr.debit.doubleValue());
            if (vr.credit != null) row.createCell(2).setCellValue(vr.credit.doubleValue());
            row.createCell(3).setCellValue(vr.narration);
            row.createCell(4).setCellValue(vr.description);
        }
        autosizeColumns(sh, 5);
    }

    private static void writeUploadSheet(Workbook wb, List<List<Object>> uploadRows) {
        Sheet sh = wb.createSheet("Upload");
        int r = 0;
        for (List<Object> rowData : uploadRows) {
            Row row = sh.createRow(r++);
            for (int c=0;c<rowData.size();c++) {
                Object o = rowData.get(c);
                Cell cell = row.createCell(c);
                if (o == null) cell.setBlank();
                else if (o instanceof Number) cell.setCellValue(((Number) o).doubleValue());
                else cell.setCellValue(o.toString());
            }
        }
        autosizeColumns(sh, uploadRows.isEmpty() ? 4 : uploadRows.get(0).size());
    }

    private static void autosizeColumns(Sheet sh, int n) {
        for (int i=0;i<n;i++) sh.autoSizeColumn(i);
    }

    // inner helper classes
    private static class TemplateRow {
        final String accountNo;
        final String template;
        final String description;
        TemplateRow(String accountNo, String template, String description) {
            this.accountNo = accountNo; this.template = template; this.description = description;
        }
    }
    private static class Rule {
        final List<String> cycles;
        final String sumCol;
        final String side;
        final String special;
        Rule(List<String> cycles, String sumCol, String side, String special) {
            this.cycles = cycles; this.sumCol = sumCol; this.side = side; this.special = special;
        }
    }
    private static class VoucherRow {
        final String accountNo;
        final BigDecimal debit;
        final BigDecimal credit;
        final String narration;
        final String description;
        VoucherRow(String accountNo, BigDecimal debit, BigDecimal credit, String narration, String description) {
            this.accountNo = accountNo; this.debit = debit; this.credit = credit; this.narration = narration; this.description = description;
        }
    }
}
