package com.example.etoll;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


// kafka working properly + guidelines added in the text file


import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * EtollVoucherGenerator - final single-file generator.
 *
 * - Use dsr_reports/<folder>/dsr_report.xlsx as input (processes all folders)
 * - Fixed RUN_NUMBER=1 and output root E-tollAcquiringSettlement/Processing
 * - Minimal logging to logs/etoll_log_<ts>.txt
 * - Public API: generateVoucher(Path dsrPath)
 */
public class EtollVoucherGenerator {

    // ---------------- CONFIG ----------------
    private static final Path DSR_ROOT = Paths.get("dsr_reports");
    private static final int RUN_NUMBER = 1;
    private static final Path OUTPUT_ROOT = Paths.get("E-tollAcquiringSettlement", "Processing");

    private static final Path LOG_FOLDER = Paths.get("logs");
    private static PrintWriter LOG_WRITER = null;

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
        m.put("NETC Settled Transaction",         new Rule(List.of("netc settled transaction"), COL_SETAMTCR, "credit", null));
        m.put("Debit Adjustment",                 new Rule(List.of("debitadjustment", "debit adjustment"), COL_SETAMTCR, "credit", null));
        m.put("Good Faith Acceptance Credit",     new Rule(List.of("good faith acceptance"), COL_SETAMTCR, "credit", null));
        m.put("Credit Adjustment",                new Rule(List.of("credit adjustment"), COL_SETAMTDR, "debit", null));
        m.put("Chargeback Acceptance",            new Rule(List.of("chargeback acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Good Faith Acceptance Debit",      new Rule(List.of("good faith acceptance"), null, "goodfaith", null));
        m.put("Pre-Arbitration Acceptance",       new Rule(List.of("pre-arbitration acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Pre-Arbitration Deemed Acceptance",new Rule(List.of("pre-arbitration deemed acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Debit chargeback deemed Acceptance",new Rule(List.of("debit chargeback deemed acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Arbitration Acceptance",           new Rule(List.of("arbitration acceptance"), COL_SETAMTDR, "debit", null));
        m.put("Arbitration Vedict",               new Rule(List.of("arbitration vedict"), COL_SETAMTDR, "debit", null));
        m.put("Income Debit",                     new Rule(null, null, null, "inward_dr"));
        m.put("GST Debit",                        new Rule(null, null, null, "inward_dr"));
        m.put("Income Credit",                    new Rule(null, null, null, "inward_cr"));
        m.put("GST Credit",                       new Rule(null, null, null, "inward_cr"));
        m.put("Final Net Amt",                    new Rule(null, null, null, "final"));
        return m;
    }

    // ---------------- MAIN — process existing folders ----------------
    public static void main(String[] args) {
        try {
            setupLogging();
            log("Starting batch processing (EtollVoucherGenerator) ...");

            if (!Files.exists(DSR_ROOT) || !Files.isDirectory(DSR_ROOT)) {
                log("ERROR: dsr_reports folder not found in project root (" + Paths.get("").toAbsolutePath() + ")");
                return;
            }

            try (DirectoryStream<Path> ds = Files.newDirectoryStream(DSR_ROOT)) {
                for (Path folder : ds) {
                    if (Files.isDirectory(folder)) {
                        processFolder(folder);
                    }
                }
            }

            log("Batch processing completed.");
        } catch (Exception e) {
            e.printStackTrace();
            log("Fatal: " + e.getMessage());
        } finally {
            if (LOG_WRITER != null) LOG_WRITER.close();
        }
    }

    private static void processFolder(Path folder) {
        try {
            Path dsr = folder.resolve("dsr_report.xlsx");
            if (!Files.exists(dsr)) {
                log("[SKIP] No dsr_report.xlsx in " + folder.getFileName());
                return;
            }

            log("Processing folder: " + folder.getFileName());
            Map<String,Object> result = generateVoucher(dsr);
            log("Result: " + result);
        } catch (Exception e) {
            log("[ERROR] processing " + folder.getFileName() + " : " + e.getMessage());
        }
    }

    /**
     * Public generator method used by consumers: accepts full path to dsr_report.xlsx.
     * Returns map {status, path, debit, credit, message?}
     */
    public static Map<String,Object> generateVoucher(Path dsrPath) throws Exception {
        Workbook wb;
        try (InputStream is = Files.newInputStream(dsrPath, StandardOpenOption.READ)) {
            wb = WorkbookFactory.create(is);
        }

        Sheet sheet = wb.getSheetAt(0);
        List<Map<String,String>> rows = readSheetToMaps(sheet);

        // forward-fill TC & TT only
        forwardFill(rows, COL_TRANSACTION_CYCLE);
        forwardFill(rows, COL_TRANSACTION_TYPE);

        // Settlement from Excel (Option B)
        LocalDate settlement = findSettlementDate(rows);
        if (settlement == null) settlement = LocalDate.now();

        log("Settlement date inside Excel = " + settlement);

        String yyyymmdd = settlement.format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        String ddmmyy   = settlement.format(DateTimeFormatter.ofPattern("ddMMyy"));
        String dd_mm_yy = settlement.format(DateTimeFormatter.ofPattern("dd.MM.yy"));
        String cycle = RUN_NUMBER + "C";

        // normalize helpers
        for (Map<String,String> r : rows) {
            r.put("TC", safeLower(r.getOrDefault(COL_TRANSACTION_CYCLE,"")));
            r.put("TT", safeLower(r.getOrDefault(COL_TRANSACTION_TYPE,"")));
            r.put("CH", safeLower(r.getOrDefault(COL_CHANNEL,"")));
        }

        // FINAL NET: last non-empty
        BigDecimal totalFinal = BigDecimal.ZERO;
        for (int i = rows.size()-1; i >=0; --i) {
            String v = rows.get(i).getOrDefault(COL_FINAL_NET_AMT,"").trim();
            if (!v.isEmpty()) {
                totalFinal = round2(toDecimal(v));
                break;
            }
        }
        log("Final Net Amt (Rightmost+Lowest) = " + totalFinal);

        // INWARD GST detection
        BigDecimal income_debit = BigDecimal.ZERO, income_credit = BigDecimal.ZERO, gst_debit = BigDecimal.ZERO, gst_credit = BigDecimal.ZERO;
        for (int i=0;i<rows.size();i++) {
            String io = rows.get(i).getOrDefault(COL_INWARD_OUTWARD,"");
            if ("INWARD GST".equalsIgnoreCase(io.trim())) {
                if (i>0) {
                    Map<String,String> ra = rows.get(i-1);
                    income_debit = round2(toDecimal(ra.getOrDefault(COL_SERVICE_FEE_DR, "0")));
                    income_credit = round2(toDecimal(ra.getOrDefault(COL_SERVICE_FEE_CR, "0")));
                }
                Map<String,String> rg = rows.get(i);
                gst_debit = round2(toDecimal(rg.getOrDefault(COL_SERVICE_FEE_DR, "0")));
                gst_credit = round2(toDecimal(rg.getOrDefault(COL_SERVICE_FEE_CR, "0")));
                break;
            }
        }
        log("Derived INWARD values -> Income Debit: " + income_debit + ", GST Debit: " + gst_debit + ", Income Credit: " + income_credit + ", GST Credit: " + gst_credit);

        // ---------------- BUILD VOUCHER ----------------
        List<VoucherRow> voucher = new ArrayList<>();

        for (TemplateRow t : TEMPLATE) {
            String acct = t.accountNo;
            String tmpl = t.template;
            String desc = t.description;

            String narration = tmpl.replace("{yyyymmdd}", yyyymmdd)
                    .replace("{ddmmyy}", ddmmyy)
                    .replace("{dd_mm_yy}", dd_mm_yy)
                    .replace("{cycle}", cycle);

            // spacer (include blank rows)
            if (acct.isEmpty() && desc.isEmpty()) {
                voucher.add(new VoucherRow("", null, null, narration, desc));
                continue;
            }

            Rule rule = RULES.get(desc);

            // final net special
            if (rule != null && "final".equals(rule.special)) {
                voucher.add(new VoucherRow(acct, totalFinal.equals(BigDecimal.ZERO) ? null : totalFinal, null, narration, desc));
                continue;
            }

            // inward dr
            if (rule != null && "inward_dr".equals(rule.special)) {
                BigDecimal amt = "Income Debit".equals(desc) ? income_debit : gst_debit;
                voucher.add(new VoucherRow(acct, amt.equals(BigDecimal.ZERO) ? null : amt, null, narration, desc));
                continue;
            }

            // inward cr
            if (rule != null && "inward_cr".equals(rule.special)) {
                BigDecimal amt = "Income Credit".equals(desc) ? income_credit : gst_credit;
                voucher.add(new VoucherRow(acct, null, amt.equals(BigDecimal.ZERO) ? null : amt, narration, desc));
                continue;
            }

            // arbitration vedict special
            if ("Arbitration Vedict".equals(desc)) {
                BigDecimal amt = BigDecimal.ZERO;
                for (Map<String,String> r : rows) {
                    if ("arbitration vedict".equals(r.get("TC")) &&
                            List.of("debit","non_fin").contains(r.get("TT")) &&
                            r.getOrDefault(COL_CHANNEL,"").trim().length() > 0) {
                        amt = amt.add(toDecimal(r.getOrDefault(COL_SETAMTDR, "0")));
                    }
                }
                amt = round2(amt);
                voucher.add(new VoucherRow(acct, amt.equals(BigDecimal.ZERO) ? null : amt, null, narration, desc));
                log("Arbitration Vedict: summed SETAMTDR = " + amt);
                continue;
            }

            // normal rules
            List<String> cycles = rule != null && rule.cycles != null ? rule.cycles : List.of();
            BigDecimal amt = BigDecimal.ZERO;
            if (rule != null && rule.sumCol != null) {
                for (Map<String,String> r : rows) {
                    if (cycles.contains(r.get("TC"))) {
                        amt = amt.add(toDecimal(r.getOrDefault(rule.sumCol, "0")));
                    }
                }
            }
            amt = round2(amt);

            if (rule != null && "credit".equals(rule.side)) {
                voucher.add(new VoucherRow(acct, null, amt.equals(BigDecimal.ZERO) ? null : amt, narration, desc));
            } else {
                voucher.add(new VoucherRow(acct, amt.equals(BigDecimal.ZERO) ? null : amt, null, narration, desc));
            }
        }

        // ---------------- UPLOAD SHEET ----------------
        List<List<Object>> uploadRows = new ArrayList<>();
        uploadRows.add(List.of("Account No","C/D","Amount","Narration"));

        for (VoucherRow vr : voucher) {
            BigDecimal d = vr.debit;
            BigDecimal c = vr.credit;
            if (d != null && d.compareTo(BigDecimal.ZERO) != 0) uploadRows.add(List.of(vr.accountNo, "D", d.doubleValue(), vr.narration));
            else if (c != null && c.compareTo(BigDecimal.ZERO) != 0) uploadRows.add(List.of(vr.accountNo, "C", c.doubleValue(), vr.narration));
        }

        // TALLY
        BigDecimal dTotal = BigDecimal.ZERO, cTotal = BigDecimal.ZERO;
        for (VoucherRow vr : voucher) {
            if (vr.debit != null) dTotal = dTotal.add(vr.debit);
            if (vr.credit != null) cTotal = cTotal.add(vr.credit);
        }
        dTotal = round2(dTotal);
        cTotal = round2(cTotal);
        log("Voucher totals -> Debit: " + dTotal + " Credit: " + cTotal);

        // create output folder and write
        Path folder = OUTPUT_ROOT.resolve(String.valueOf(settlement.getYear()))
                .resolve(String.format("%02d", settlement.getMonthValue()))
                .resolve(String.format("%02d", settlement.getDayOfMonth()));
        Files.createDirectories(folder);

        Path okFile = folder.resolve("ETOLL_ACQUIRING_VOUCHER_" + ddmmyy + "_N" + RUN_NUMBER + ".xlsx");
        Path errFile = folder.resolve("ERROR_ETOLL_ACQUIRING_VOUCHER_" + ddmmyy + "_N" + RUN_NUMBER + ".xlsx");

        boolean ok = dTotal.compareTo(cTotal) == 0;

        try (Workbook out = new XSSFWorkbook()) {
            writeVoucherSheet(out, voucher);
            writeUploadSheet(out, uploadRows);

            Path writeTo = ok ? okFile : errFile;

            // write workbook to disk (create/overwrite)
            try (OutputStream os = Files.newOutputStream(writeTo, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
                out.write(os);
            }
            log("Voucher written to: " + writeTo.toAbsolutePath());
        }

        Map<String,Object> result = new HashMap<>();
        result.put("status", ok ? "ok" : "error");
        result.put("path", (ok ? okFile : errFile).toString());
        result.put("debit", dTotal);
        result.put("credit", cTotal);
        if (!ok) result.put("message", "Debit and credit not tallied");

        return result;
    }

    // ---------------- logging ----------------
    private static synchronized void setupLogging() throws IOException {
        Files.createDirectories(LOG_FOLDER);
        String name = "etoll_log_" + System.currentTimeMillis() + ".txt";
        LOG_WRITER = new PrintWriter(Files.newBufferedWriter(LOG_FOLDER.resolve(name), java.nio.charset.StandardCharsets.UTF_8, StandardOpenOption.CREATE, StandardOpenOption.APPEND));
    }

    private static synchronized void log(String msg) {
        String line = "[" + java.time.ZonedDateTime.now() + "] " + msg;
        System.out.println(line);
        if (LOG_WRITER != null) {
            LOG_WRITER.println(line);
            LOG_WRITER.flush();
        }
    }

    // ---------------- helpers ----------------
    private static String safeLower(String s) {
        return s == null ? "" : s.trim().toLowerCase();
    }

    private static BigDecimal toDecimal(String s) {
        if (s == null) return BigDecimal.ZERO;
        s = s.trim().replace(",","");
        if (s.isEmpty() || s.equalsIgnoreCase("nan")) return BigDecimal.ZERO;
        try { return new BigDecimal(s); }
        catch (Exception e) {
            try { return BigDecimal.valueOf(Double.parseDouble(s)); }
            catch (Exception ex) { return BigDecimal.ZERO; }
        }
    }

    private static BigDecimal round2(BigDecimal d) {
        return d.setScale(2, RoundingMode.HALF_UP);
    }

    private static List<Map<String,String>> readSheetToMaps(Sheet sheet) {
        List<Map<String,String>> rows = new ArrayList<>();
        Iterator<Row> it = sheet.iterator();
        if (!it.hasNext()) return rows;
        Row header = it.next();
        List<String> headers = new ArrayList<>();
        for (Cell c : header) headers.add(c.getStringCellValue().trim());
        while (it.hasNext()) {
            Row r = it.next();
            Map<String,String> map = new HashMap<>();
            for (int i=0;i<headers.size();i++) {
                Cell cell = r.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                map.put(headers.get(i), cellToString(cell));
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
                if (DateUtil.isCellDateFormatted(cell)) return cell.getLocalDateTimeCellValue().toLocalDate().toString();
                else return BigDecimal.valueOf(cell.getNumericCellValue()).stripTrailingZeros().toPlainString();
            case BOOLEAN: return Boolean.toString(cell.getBooleanCellValue());
            case BLANK: return "";
            case FORMULA:
                try {
                    if (DateUtil.isCellDateFormatted(cell)) return cell.getLocalDateTimeCellValue().toLocalDate().toString();
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
        for (int c=0;c<5;c++) sh.autoSizeColumn(c);
    }

    private static void writeUploadSheet(Workbook wb, List<List<Object>> rows) {
        Sheet sh = wb.createSheet("Upload");
        int r = 0;
        for (List<Object> rowData : rows) {
            Row row = sh.createRow(r++);
            for (int c=0;c<rowData.size();c++) {
                Object o = rowData.get(c);
                Cell cell = row.createCell(c);
                if (o == null) cell.setBlank();
                else if (o instanceof Number) cell.setCellValue(((Number)o).doubleValue());
                else cell.setCellValue(o.toString());
            }
        }
        for (int c=0;c<4;c++) sh.autoSizeColumn(c);
    }

    // ---------- inner classes ----------
    private static class TemplateRow {
        final String accountNo;
        final String template;
        final String description;
        TemplateRow(String a, String b, String c) { accountNo=a; template=b; description=c; }
    }

    private static class Rule {
        final List<String> cycles;
        final String sumCol;
        final String side;
        final String special;
        Rule(List<String> cycles, String sumCol, String side, String special) { this.cycles=cycles; this.sumCol=sumCol; this.side=side; this.special=special; }
    }

    private static class VoucherRow {
        final String accountNo;
        final BigDecimal debit;
        final BigDecimal credit;
        final String narration;
        final String description;
        VoucherRow(String a, BigDecimal d, BigDecimal c, String n, String desc) { accountNo=a; debit=d; credit=c; narration=n; description=desc; }
    }
}
