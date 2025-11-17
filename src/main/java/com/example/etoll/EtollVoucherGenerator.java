package com.example.etoll;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * EtollVoucherGenerator
 * - Reads dsr_report.xlsx (xlsx)
 * - Implements the mapping & aggregation logic (Option A)
 * - Produces voucher and diagnostic Excel files
 *
 * Usage:
 * 1) Place dsr_report.xlsx in project root (or change DSR_PATH)
 * 2) mvn package
 * 3) java -jar target/etoll-voucher-1.0.0-jar-with-dependencies.jar
 */
public class EtollVoucherGenerator {

    // ---------- CONFIG ----------
    // Path to DSR file (relative to project root)
    private static final Path DSR_PATH = Paths.get("dsr_report.xlsx");

    // Output root
    private static final Path OUTPUT_ROOT = Paths.get("E-tollAcquiringSettlement/Processing");

    // Run number (1 => 1C)
    private static final int RUN_NUMBER = 1;

    // Column names (exact as in your DSR)
    private static final String COL_SETTLEMENT_DATE = "Settlement Date";
    private static final String COL_INWARD_OUTWARD = "Inward/Outward";
    private static final String COL_TRANSACTION_CYCLE = "Transaction Cycle";
    private static final String COL_CHANNEL = "Channel";
    private static final String COL_SETAMTDR = "SETAMTDR";
    private static final String COL_SETAMTCR = "SETAMTCR";
    private static final String COL_SERVICE_FEE_DR = "Service Fee Amt Dr";
    private static final String COL_SERVICE_FEE_CR = "Service Fee Amt Cr";
    private static final String COL_FINAL_NET_AMT = "Final Net Amt";

    // Template rows
    private static final List<TemplateRow> TEMPLATE = Arrays.asList(
            new TemplateRow("0103SLRTGSTRC", "NPCIR5{date_yyyymmdd} {date_ddmmyy}_{cycle} ETCAC", "Final Net Amt"),
            new TemplateRow("", "", ""),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy}_{cycle}", "NETC Settled Transaction Credit"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} Dr.Adj_{cycle}", "DebitAdjustment"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} GF Accp_{cycle}", "Good Faith Acceptance Credit"),
            new TemplateRow("", "", ""),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} Cr.Adj_{cycle}", "Credit Adjustment"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} Chbk_{cycle}", "Chargeback Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} GF Acp_{cycle}", "Good Faith Acceptance Debit"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} PrArbtAc_{cycle}", "Pre-Arbitration Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} Dr PrArAc_{cycle}", "Pre-Arbitration Deemed Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} DrArbtAc_{cycle}", "Debit chargeback deemed Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} ArbtAc_{cycle}", "Arbitration Acceptance"),
            new TemplateRow("0103SLETCACQ", "Etoll acq {date_dd_mm_yy} ArbVer_{cycle}", "Arbitration Vedit"),
            new TemplateRow("", "", ""),
            new TemplateRow("0103CNETCACQ", "Etoll acq {date_dd_mm_yy}_{cycle}", "Income Debit"),
            new TemplateRow("0103SLPPCIGT", "Etoll acq {date_dd_mm_yy}_{cycle}", "GST Debit"),
            new TemplateRow("0103CNETCACQ", "Etoll acq {date_dd_mm_yy}_{cycle}", "Income Credit"),
            new TemplateRow("0103SLPPCIGT", "Etoll acq {date_dd_mm_yy}_{cycle}", "GST Credit")
    );

    // mapping of description to default rule (source column + side)
    private static final Map<String, Rule> ROW_RULES = new LinkedHashMap<>();
    static {
        ROW_RULES.put("NETC Settled Transaction Credit", new Rule(COL_SETAMTCR, "credit"));
        ROW_RULES.put("DebitAdjustment", new Rule(COL_SETAMTCR, "credit"));
        ROW_RULES.put("Credit Adjustment", new Rule(COL_SETAMTDR, "debit"));
        ROW_RULES.put("Chargeback Acceptance", new Rule(COL_SETAMTDR, "debit"));
        ROW_RULES.put("Good Faith Acceptance Credit", new Rule(COL_SETAMTCR, "credit"));
        ROW_RULES.put("Good Faith Acceptance Debit", new Rule(COL_SETAMTDR, "debit"));
        ROW_RULES.put("Pre-Arbitration Acceptance", new Rule(COL_SETAMTDR, "debit"));
        ROW_RULES.put("Pre-Arbitration Deemed Acceptance", new Rule(COL_SETAMTDR, "debit"));
        ROW_RULES.put("Debit chargeback deemed Acceptance", new Rule(COL_SETAMTDR, "debit"));
        ROW_RULES.put("Arbitration Acceptance", new Rule(COL_SETAMTDR, "debit"));
        ROW_RULES.put("Arbitration Vedit", new Rule(COL_SETAMTDR, "debit"));
        ROW_RULES.put("Income Debit", new Rule(COL_SERVICE_FEE_DR, "debit", true));
        ROW_RULES.put("GST Debit", new Rule(COL_SERVICE_FEE_DR, "debit", true));
        ROW_RULES.put("Income Credit", new Rule(COL_SERVICE_FEE_CR, "credit", true));
        ROW_RULES.put("GST Credit", new Rule(COL_SERVICE_FEE_CR, "credit", true));
        ROW_RULES.put("Final Net Amt", new Rule(COL_FINAL_NET_AMT, "debit_or_credit"));
    }

    // mapping description -> cycles in DSR to check
    private static final Map<String, List<String>> DESC_TO_CYCLES = new HashMap<>();
    static {
        DESC_TO_CYCLES.put("NETC Settled Transaction Credit", Arrays.asList("NETC Settled Transaction"));
        DESC_TO_CYCLES.put("DebitAdjustment", Arrays.asList("DebitAdjustment", "Debit Adjustment"));
        DESC_TO_CYCLES.put("Credit Adjustment", Arrays.asList("Credit Adjustment", "CreditAdjustment"));
        DESC_TO_CYCLES.put("Chargeback Acceptance", Arrays.asList("Chargeback Acceptance"));
        DESC_TO_CYCLES.put("Good Faith Acceptance Credit", Arrays.asList("Good Faith Acceptance"));
        DESC_TO_CYCLES.put("Good Faith Acceptance Debit", Arrays.asList("Good Faith Acceptance"));
        DESC_TO_CYCLES.put("Pre-Arbitration Acceptance", Arrays.asList("Pre-Arbitration Acceptance"));
        DESC_TO_CYCLES.put("Pre-Arbitration Deemed Acceptance", Arrays.asList("Pre-Arbitration Deemed Acceptance"));
        DESC_TO_CYCLES.put("Debit chargeback deemed Acceptance", Arrays.asList("Debit chargeback deemed Acceptance"));
        DESC_TO_CYCLES.put("Arbitration Acceptance", Arrays.asList("Arbitration Acceptance"));
        DESC_TO_CYCLES.put("Arbitration Vedit", Arrays.asList("Arbitration Vedit", "Arbitration Vedict", "Arbitration Verdict"));
    }

    // ---------- Main ----------
    public static void main(String[] args) {
        try {
            new EtollVoucherGenerator().run();
        } catch (Exception e) {
            System.err.println("Fatal error: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }

    private void run() throws Exception {
        if (!Files.exists(DSR_PATH)) {
            throw new FileNotFoundException("DSR file not found at: " + DSR_PATH.toAbsolutePath());
        }

        // read DSR into list of maps (rows)
        List<Map<String, String>> rows = readXlsxToRowMaps(DSR_PATH);

        // detect settlement date
        LocalDate settlementDate = detectSettlementDate(rows);
        Map<String,String> dates = buildDatesMap(settlementDate);
        String cycleSuffix = RUN_NUMBER + "C";

        System.out.println("Detected settlement date: " + settlementDate);
        System.out.println("Cycle suffix: " + cycleSuffix);

        // compute per-description candidates
        List<Map<String,Object>> voucherRows = new ArrayList<>();
        List<Map<String,Object>> diagRows = new ArrayList<>();

        for (TemplateRow tr : TEMPLATE) {
            String desc = tr.description == null ? "" : tr.description;
            String acct = tr.account == null ? "" : tr.account;
            String narrT = tr.narration == null ? "" : tr.narration;

            Candidate c = computeCandidates(rows, desc);
            // pick final debit/credit based on rule
            Pick pick = pickAmount(c, desc, rows);

            String narration = "";
            if (!narrT.isEmpty()) {
                narration = narrT.replace("{date_ddmmyy}", dates.get("date_ddmmyy"))
                                  .replace("{date_dd_mm_yy}", dates.get("date_dd_mm_yy"))
                                  .replace("{date_yyyymmdd}", dates.get("date_yyyymmdd"))
                                  .replace("{cycle}", cycleSuffix);
            }

            Map<String,Object> vrow = new LinkedHashMap<>();
            vrow.put("Account No", acct);
            vrow.put("Debit", pick.debit);
            vrow.put("Credit", pick.credit);
            vrow.put("Narration", narration);
            vrow.put("Description", desc);
            voucherRows.add(vrow);

            // diag
            Map<String,Object> d = new LinkedHashMap<>();
            d.put("Account No", acct);
            d.put("Description", desc);
            d.put("SETAMTDR", c.setamtDr);
            d.put("SETAMTCR", c.setamtCr);
            d.put("FinalNet", c.finalNet);
            d.put("SvcDr", c.svcDr);
            d.put("SvcCr", c.svcCr);
            d.put("ChosenDebit", pick.debit);
            d.put("ChosenCredit", pick.credit);
            d.put("Why", pick.reason);
            diagRows.add(d);
        }

        // if Final Net Amt negative -> terminate (per spec)
        BigDecimal totalFinalNet = computeTotalFinalNet(rows);
        if (totalFinalNet.compareTo(BigDecimal.ZERO) < 0) {
            System.err.println("Final Net Amt negative (" + totalFinalNet + "). Terminating and notifying process owner.");
            System.exit(2);
        }

        // build output folder
        Path outFolder = OUTPUT_ROOT.resolve(settlementDate.format(DateTimeFormatter.ofPattern("yyyy")))
                .resolve(settlementDate.format(DateTimeFormatter.ofPattern("MM")))
                .resolve(settlementDate.format(DateTimeFormatter.ofPattern("dd")));
        Files.createDirectories(outFolder);

        // write voucher and diagnostic
        String dtStr = settlementDate.format(DateTimeFormatter.ofPattern("ddMMyy"));
        Path voucherFile = outFolder.resolve("ETOLL ACQUIRING VOUCHER_" + dtStr + "_N" + RUN_NUMBER + ".xlsx");
        Path diagFile = outFolder.resolve("voucher_diagnostic_" + dtStr + "_N" + RUN_NUMBER + ".xlsx");

        writeRowsToXlsx(voucherRows, Arrays.asList("Account No","Debit","Credit","Narration","Description"), voucherFile);
        writeRowsToXlsx(diagRows, Arrays.asList("Account No","Description","SETAMTDR","SETAMTCR","FinalNet","SvcDr","SvcCr","ChosenDebit","ChosenCredit","Why"), diagFile);

        System.out.println("Voucher saved to: " + voucherFile.toAbsolutePath());
        System.out.println("Diagnostic saved to: " + diagFile.toAbsolutePath());
    }

    // ---------- helpers ----------

    private List<Map<String,String>> readXlsxToRowMaps(Path xlsx) throws IOException {
        List<Map<String,String>> rows = new ArrayList<>();
        try (InputStream in = Files.newInputStream(xlsx);
             Workbook wb = new XSSFWorkbook(in)) {
            Sheet sheet = wb.getSheetAt(0);

            Iterator<Row> it = sheet.iterator();
            if (!it.hasNext()) return rows;
            Row header = it.next();
            List<String> headers = new ArrayList<>();
            for (Cell c : header) {
                headers.add(getCellString(c).trim());
            }

            while (it.hasNext()) {
                Row r = it.next();
                Map<String,String> map = new LinkedHashMap<>();
                for (int i=0;i<headers.size();i++) {
                    Cell cell = r.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String val = cell == null ? "" : getCellString(cell);
                    map.put(headers.get(i).trim(), val);
                }
                rows.add(map);
            }
        }
        return rows;
    }

    private String getCellString(Cell c) {
        if (c == null) return "";
        switch (c.getCellType()) {
            case STRING: return c.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(c)) {
                    return c.getLocalDateTimeCellValue().toLocalDate().toString();
                } else {
                    double d = c.getNumericCellValue();
                    // remove trailing .0 if integer-like
                    if (d == Math.floor(d)) return String.valueOf((long)d);
                    return String.valueOf(d);
                }
            case BOOLEAN: return String.valueOf(c.getBooleanCellValue());
            case FORMULA:
                try {
                    return c.getStringCellValue();
                } catch (Exception e) {
                    double d = c.getNumericCellValue();
                    if (d == Math.floor(d)) return String.valueOf((long)d);
                    return String.valueOf(d);
                }
            default: return "";
        }
    }

    private LocalDate detectSettlementDate(List<Map<String,String>> rows) {
        // Try to find from first non-empty Settlement Date column
        for (Map<String,String> r : rows) {
            if (r.containsKey(COL_SETTLEMENT_DATE)) {
                String val = r.get(COL_SETTLEMENT_DATE);
                if (val != null && !val.trim().isEmpty()) {
                    // try known patterns: dd-MM-yyyy or variants
                    try {
                        DateTimeFormatter f = DateTimeFormatter.ofPattern("dd-MM-yyyy");
                        return LocalDate.parse(val.trim(), f);
                    } catch (Exception ignored) {}
                    try {
                        DateTimeFormatter f2 = DateTimeFormatter.ofPattern("dd.MM.yyyy");
                        return LocalDate.parse(val.trim(), f2);
                    } catch (Exception ignored) {}
                    // fallback: parse by splitting
                    try {
                        // support dd-MM-yy etc
                        String s = val.trim().replace(".", "-");
                        String[] parts = s.split("-");
                        if (parts.length >= 3) {
                            int day = Integer.parseInt(parts[0]);
                            int month = Integer.parseInt(parts[1]);
                            int year = Integer.parseInt(parts[2].length() == 2 ? ("20"+parts[2]) : parts[2]);
                            return LocalDate.of(year, month, day);
                        }
                    } catch (Exception ignored) {}
                }
            }
        }
        // fallback: search first 6 columns of first row for date-like value
        if (!rows.isEmpty()) {
            Map<String,String> first = rows.get(0);
            int i=0;
            for (String key : first.keySet()) {
                if (i++ > 5) break;
                String v = first.get(key);
                if (v != null && !v.trim().isEmpty()) {
                    try {
                        String s = v.trim().replace(".", "-");
                        String[] parts = s.split("-");
                        if (parts.length >= 3) {
                            int day = Integer.parseInt(parts[0]);
                            int month = Integer.parseInt(parts[1]);
                            int year = Integer.parseInt(parts[2].length() == 2 ? ("20"+parts[2]) : parts[2]);
                            return LocalDate.of(year, month, day);
                        }
                    } catch (Exception ignored) {}
                }
            }
        }
        // if not found, use today
        return LocalDate.now();
    }

    private Map<String,String> buildDatesMap(LocalDate d) {
        Map<String,String> m = new HashMap<>();
        m.put("date_ddmmyy", d.format(DateTimeFormatter.ofPattern("ddMMyy")));
        m.put("date_dd_mm_yy", d.format(DateTimeFormatter.ofPattern("dd.MM.yy")));
        m.put("date_yyyymmdd", d.format(DateTimeFormatter.ofPattern("yyyyMMdd")));
        return m;
    }

    private Candidate computeCandidates(List<Map<String,String>> rows, String desc) {
        Candidate c = new Candidate();
        if (desc == null || desc.trim().isEmpty()) {
            return c;
        }
        List<String> cycles = DESC_TO_CYCLES.getOrDefault(desc, Collections.emptyList());
        BigDecimal setDr = BigDecimal.ZERO, setCr = BigDecimal.ZERO, fn = BigDecimal.ZERO, sdr = BigDecimal.ZERO, scr = BigDecimal.ZERO;

        for (String cyc : cycles) {
            // sum only rows where Channel is non-empty and transaction cycle matches (case-insensitive)
            for (Map<String,String> r : rows) {
                String cycleVal = r.getOrDefault(COL_TRANSACTION_CYCLE, "");
                String channelVal = r.getOrDefault(COL_CHANNEL, "");
                if (cycleVal == null) cycleVal = "";
                if (channelVal == null) channelVal = "";
                if (!channelVal.trim().isEmpty() && cycleVal.trim().equalsIgnoreCase(cyc.trim())) {
                    setDr = setDr.add(parseNumberSafe(r.getOrDefault(COL_SETAMTDR, "0")));
                    setCr = setCr.add(parseNumberSafe(r.getOrDefault(COL_SETAMTCR, "0")));
                    fn = fn.add(parseNumberSafe(r.getOrDefault(COL_FINAL_NET_AMT, "0")));
                    sdr = sdr.add(parseNumberSafe(r.getOrDefault(COL_SERVICE_FEE_DR, "0")));
                    scr = scr.add(parseNumberSafe(r.getOrDefault(COL_SERVICE_FEE_CR, "0")));
                }
            }
        }

        // special for Income/GST: sum across rows where Inward/Outward contains INWARD
        if (desc.equals("Income Debit") || desc.equals("GST Debit") || desc.equals("Income Credit") || desc.equals("GST Credit")) {
            BigDecimal sdr_i = BigDecimal.ZERO, scr_i = BigDecimal.ZERO;
            for (Map<String,String> r : rows) {
                String inout = r.getOrDefault(COL_INWARD_OUTWARD, "");
                if (inout != null && inout.toUpperCase().contains("INWARD")) {
                    sdr_i = sdr_i.add(parseNumberSafe(r.getOrDefault(COL_SERVICE_FEE_DR, "0")));
                    scr_i = scr_i.add(parseNumberSafe(r.getOrDefault(COL_SERVICE_FEE_CR, "0")));
                }
            }
            sdr = sdr_i;
            scr = scr_i;
        }

        c.setamtDr = setDr;
        c.setamtCr = setCr;
        c.finalNet = fn;
        c.svcDr = sdr;
        c.svcCr = scr;
        return c;
    }

    private Pick pickAmount(Candidate c, String desc, List<Map<String,String>> rows) {
        if (desc == null || desc.trim().isEmpty()) {
            return new Pick(BigDecimal.ZERO, BigDecimal.ZERO, "empty_desc");
        }
        Rule r = ROW_RULES.get(desc);
        if (r == null) {
            return new Pick(BigDecimal.ZERO, BigDecimal.ZERO, "no_rule");
        }

        if (desc.equals("Final Net Amt")) {
            BigDecimal totalFn = computeTotalFinalNet(rows);
            if (totalFn.compareTo(BigDecimal.ZERO) < 0) {
                throw new RuntimeException("Final Net Amt negative: " + totalFn);
            }
            return new Pick(totalFn.setScale(2, BigDecimal.ROUND_HALF_UP), BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), "FinalNet_used");
        }

        if (r.inwardOnly) {
            if ("debit".equalsIgnoreCase(r.side)) {
                return new Pick(c.svcDr.setScale(2, BigDecimal.ROUND_HALF_UP), BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), "SvcDr_inward");
            } else {
                return new Pick(BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), c.svcCr.setScale(2, BigDecimal.ROUND_HALF_UP), "SvcCr_inward");
            }
        }

        if (COL_SETAMTDR.equals(r.amountColumn)) {
            if ("debit".equalsIgnoreCase(r.side)) {
                return new Pick(c.setamtDr.setScale(2, BigDecimal.ROUND_HALF_UP), BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), "SETAMTDR_used");
            } else {
                return new Pick(BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), c.setamtDr.setScale(2, BigDecimal.ROUND_HALF_UP), "SETAMTDR_used_as_credit");
            }
        } else if (COL_SETAMTCR.equals(r.amountColumn)) {
            if ("credit".equalsIgnoreCase(r.side)) {
                return new Pick(BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), c.setamtCr.setScale(2, BigDecimal.ROUND_HALF_UP), "SETAMTCR_used");
            } else {
                return new Pick(c.setamtCr.setScale(2, BigDecimal.ROUND_HALF_UP), BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), "SETAMTCR_used_as_debit");
            }
        } else {
            // fallback use amount column if present in candidate
            if (COL_SERVICE_FEE_DR.equals(r.amountColumn)) {
                if ("debit".equalsIgnoreCase(r.side)) {
                    return new Pick(c.svcDr.setScale(2, BigDecimal.ROUND_HALF_UP), BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), "SvcDr_used");
                } else {
                    return new Pick(BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), c.svcDr.setScale(2, BigDecimal.ROUND_HALF_UP), "SvcDr_used_as_credit");
                }
            } else if (COL_SERVICE_FEE_CR.equals(r.amountColumn)) {
                if ("credit".equalsIgnoreCase(r.side)) {
                    return new Pick(BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), c.svcCr.setScale(2, BigDecimal.ROUND_HALF_UP), "SvcCr_used");
                } else {
                    return new Pick(c.svcCr.setScale(2, BigDecimal.ROUND_HALF_UP), BigDecimal.ZERO.setScale(2, BigDecimal.ROUND_HALF_UP), "SvcCr_used_as_debit");
                }
            }
        }
        return new Pick(BigDecimal.ZERO, BigDecimal.ZERO, "fallback");
    }

    private BigDecimal computeTotalFinalNet(List<Map<String,String>> rows) {
        BigDecimal total = BigDecimal.ZERO;
        for (Map<String,String> r : rows) {
            total = total.add(parseNumberSafe(r.getOrDefault(COL_FINAL_NET_AMT, "0")));
        }
        return total;
    }

    private BigDecimal parseNumberSafe(String s) {
        if (s == null) return BigDecimal.ZERO;
        String cleaned = s.trim().replace(",", "");
        if (cleaned.isEmpty()) return BigDecimal.ZERO;
        try {
            return new BigDecimal(cleaned);
        } catch (Exception e) {
            try {
                Double d = Double.parseDouble(cleaned);
                return BigDecimal.valueOf(d);
            } catch (Exception ex) {
                return BigDecimal.ZERO;
            }
        }
    }

    private void writeRowsToXlsx(List<Map<String,Object>> rows, List<String> headers, Path out) throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("Sheet1");
            Row h = s.createRow(0);
            for (int i=0;i<headers.size();i++) {
                Cell c = h.createCell(i);
                c.setCellValue(headers.get(i));
            }
            int rnum = 1;
            for (Map<String,Object> rmap : rows) {
                Row r = s.createRow(rnum++);
                for (int i=0;i<headers.size();i++) {
                    Object v = rmap.getOrDefault(headers.get(i), "");
                    Cell c = r.createCell(i);
                    if (v instanceof Number) {
                        c.setCellValue(((Number) v).doubleValue());
                    } else {
                        c.setCellValue(String.valueOf(v));
                    }
                }
            }
            // autosize columns modestly
            for (int i=0;i<headers.size();i++) s.autoSizeColumn(i);
            try (OutputStream outStream = Files.newOutputStream(out)) {
                wb.write(outStream);
            }
        }
    }

    // ---------- helper data classes ----------
    private static class TemplateRow {
        String account;
        String narration;
        String description;
        TemplateRow(String account, String narration, String description) {
            this.account = account;
            this.narration = narration;
            this.description = description;
        }
    }

    private static class Rule {
        String amountColumn;
        String side;
        boolean inwardOnly;
        Rule(String amountColumn, String side) { this(amountColumn, side, false); }
        Rule(String amountColumn, String side, boolean inwardOnly) {
            this.amountColumn = amountColumn;
            this.side = side;
            this.inwardOnly = inwardOnly;
        }
    }

    private static class Candidate {
        BigDecimal setamtDr = BigDecimal.ZERO;
        BigDecimal setamtCr = BigDecimal.ZERO;
        BigDecimal finalNet = BigDecimal.ZERO;
        BigDecimal svcDr = BigDecimal.ZERO;
        BigDecimal svcCr = BigDecimal.ZERO;
    }

    private static class Pick {
        BigDecimal debit;
        BigDecimal credit;
        String reason;
        Pick(BigDecimal debit, BigDecimal credit, String reason) {
            this.debit = debit;
            this.credit = credit;
            this.reason = reason;
        }
    }
}
