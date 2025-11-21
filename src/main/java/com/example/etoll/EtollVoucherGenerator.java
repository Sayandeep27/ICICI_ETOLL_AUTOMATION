package com.example.etoll;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

// proper input and output folder structure added

public class EtollVoucherGenerator {

    private static final Path DSR_ROOT = Paths.get("dsr_reports");
    private static final int RUN_NUMBER = 1;
    private static final Path OUTPUT_ROOT = Paths.get("E-tollAcquiringSettlement", "Processing");

    private static final Path LOG_FOLDER = Paths.get("logs");
    private static PrintWriter LOG;

    private static final String COL_SETTLEMENT_DATE = "Settlement Date";
    private static final String COL_TRANSACTION_CYCLE = "Transaction Cycle";
    private static final String COL_TRANSACTION_TYPE = "Transaction Type";
    private static final String COL_CHANNEL = "Channel";
    private static final String COL_SETAMTDR = "SETAMTDR";
    private static final String COL_SETAMTCR = "SETAMTCR";
    private static final String COL_SERVICE_FEE_DR = "Service Fee Amt Dr";
    private static final String COL_SERVICE_FEE_CR = "Service Fee Amt Cr";
    private static final String COL_FINAL_NET_AMT = "Final Net Amt";
    private static final String COL_INWARD_OUTWARD = "Inward/Outward";

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
        try {
            setupLogging();
            println("Starting batch processing...");

            if (!Files.exists(DSR_ROOT)) {
                println("ERROR: dsr_reports folder not found.");
                return;
            }

            try (DirectoryStream<Path> stream = Files.newDirectoryStream(DSR_ROOT)) {
                for (Path folder : stream) {
                    if (Files.isDirectory(folder)) processFolder(folder);
                }
            }

            println("Batch processing completed.");
        } catch (Exception e) {
            println("Fatal error: " + e.getMessage());
        } finally {
            if (LOG != null) LOG.close();
        }
    }

    private static void processFolder(Path folder) {
        try {
            Path dsr = folder.resolve("dsr_report.xlsx");
            if (!Files.exists(dsr)) {
                println("[SKIP] No dsr_report.xlsx in " + folder.getFileName());
                return;
            }

            println("Processing: " + folder.getFileName());
            Map<String, Object> res = generateVoucher(dsr);
            println("Result: " + res.get("status") + " | file: " + res.get("path"));

        } catch (Exception e) {
            println("[ERROR] Failed processing folder " + folder.getFileName() + ": " + e.getMessage());
        }
    }

    public static Map<String, Object> generateVoucher(Path dsrPath) throws Exception {

        Workbook wb;
        try (InputStream is = Files.newInputStream(dsrPath)) {
            wb = WorkbookFactory.create(is);
        }

        Sheet sheet = wb.getSheetAt(0);
        List<Map<String, String>> rows = readSheetToMaps(sheet);

        forwardFill(rows, COL_TRANSACTION_CYCLE);
        forwardFill(rows, COL_TRANSACTION_TYPE);

        LocalDate settlement = findSettlementDate(rows);
        if (settlement == null) settlement = LocalDate.now();

        println("Settlement date inside Excel = " + settlement);

        String yyyymmdd = settlement.format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        String ddmmyy = settlement.format(DateTimeFormatter.ofPattern("ddMMyy"));
        String dd_mm_yy = settlement.format(DateTimeFormatter.ofPattern("dd.MM.yy"));
        String cycle = RUN_NUMBER + "C";

        for (Map<String, String> r : rows) {
            r.put("TC", safeLower(r.get(COL_TRANSACTION_CYCLE)));
            r.put("TT", safeLower(r.get(COL_TRANSACTION_TYPE)));
            r.put("CH", safeLower(r.get(COL_CHANNEL)));
        }

        BigDecimal totalFinal = BigDecimal.ZERO;
        for (int i = rows.size() - 1; i >= 0; i--) {
            String v = rows.get(i).getOrDefault(COL_FINAL_NET_AMT, "").trim();
            if (!v.isEmpty()) {
                totalFinal = round2(toDecimal(v));
                break;
            }
        }

        println("Final Net Amount = " + totalFinal);

        BigDecimal incomeDebit = BigDecimal.ZERO, incomeCredit = BigDecimal.ZERO;
        BigDecimal gstDebit = BigDecimal.ZERO, gstCredit = BigDecimal.ZERO;

        for (int i = 0; i < rows.size(); i++) {
            String io = rows.get(i).getOrDefault(COL_INWARD_OUTWARD, "");
            if ("INWARD GST".equalsIgnoreCase(io.trim())) {

                if (i > 0) {
                    Map<String, String> prev = rows.get(i - 1);
                    incomeDebit = round2(toDecimal(prev.getOrDefault(COL_SERVICE_FEE_DR, "0")));
                    incomeCredit = round2(toDecimal(prev.getOrDefault(COL_SERVICE_FEE_CR, "0")));
                }

                Map<String, String> rg = rows.get(i);
                gstDebit = round2(toDecimal(rg.getOrDefault(COL_SERVICE_FEE_DR, "0")));
                gstCredit = round2(toDecimal(rg.getOrDefault(COL_SERVICE_FEE_CR, "0")));
                break;
            }
        }

        println("INWARD → incomeDr=" + incomeDebit + " gstDr=" + gstDebit +
                " incomeCr=" + incomeCredit + " gstCr=" + gstCredit);

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

            Rule rule = RULES.get(desc);

            if (rule != null && "final".equals(rule.special)) {
                voucher.add(new VoucherRow(acct,
                        totalFinal.equals(BigDecimal.ZERO) ? null : totalFinal,
                        null, narration, desc));
                continue;
            }

            if (rule != null && "inward_dr".equals(rule.special)) {
                BigDecimal amt = desc.equals("Income Debit") ? incomeDebit : gstDebit;
                voucher.add(new VoucherRow(acct, amt.equals(BigDecimal.ZERO) ? null : amt, null, narration, desc));
                continue;
            }

            if (rule != null && "inward_cr".equals(rule.special)) {
                BigDecimal amt = desc.equals("Income Credit") ? incomeCredit : gstCredit;
                voucher.add(new VoucherRow(acct, null, amt.equals(BigDecimal.ZERO) ? null : amt, narration, desc));
                continue;
            }

            if ("Arbitration Vedict".equals(desc)) {
                BigDecimal amt = BigDecimal.ZERO;
                for (Map<String, String> r : rows) {
                    if ("arbitration vedict".equals(r.get("TC")) &&
                            List.of("debit", "non_fin").contains(r.get("TT")) &&
                            r.getOrDefault(COL_CHANNEL, "").trim().length() > 0) {
                        amt = amt.add(toDecimal(r.getOrDefault(COL_SETAMTDR, "0")));
                    }
                }
                amt = round2(amt);
                voucher.add(new VoucherRow(acct, amt.equals(BigDecimal.ZERO) ? null : amt, null, narration, desc));
                continue;
            }

            List<String> cycles = rule != null && rule.cycles != null ? rule.cycles : List.of();
            BigDecimal amt = BigDecimal.ZERO;

            for (Map<String, String> r : rows) {
                if (cycles.contains(r.get("TC"))) {
                    amt = amt.add(toDecimal(r.getOrDefault(rule.sumCol, "0")));
                }
            }

            amt = round2(amt);

            if ("credit".equals(rule.side)) {
                voucher.add(new VoucherRow(acct, null, amt.equals(BigDecimal.ZERO) ? null : amt, narration, desc));
            } else {
                voucher.add(new VoucherRow(acct, amt.equals(BigDecimal.ZERO) ? null : amt, null, narration, desc));
            }
        }

        List<List<Object>> uploadRows = new ArrayList<>();
        uploadRows.add(List.of("Account No", "C/D", "Amount", "Narration"));

        for (VoucherRow vr : voucher) {
            BigDecimal d = vr.debit;
            BigDecimal c = vr.credit;

            if (d != null && d.compareTo(BigDecimal.ZERO) != 0)
                uploadRows.add(List.of(vr.accountNo, "D", d.doubleValue(), vr.narration));
            else if (c != null && c.compareTo(BigDecimal.ZERO) != 0)
                uploadRows.add(List.of(vr.accountNo, "C", c.doubleValue(), vr.narration));
        }

        BigDecimal dTotal = BigDecimal.ZERO, cTotal = BigDecimal.ZERO;
        for (VoucherRow vr : voucher) {
            if (vr.debit != null) dTotal = dTotal.add(vr.debit);
            if (vr.credit != null) cTotal = cTotal.add(vr.credit);
        }

        dTotal = round2(dTotal);
        cTotal = round2(cTotal);

        println("Debit=" + dTotal + "  Credit=" + cTotal);

        Path folder = OUTPUT_ROOT.resolve("" + settlement.getYear())
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

            try (OutputStream os = Files.newOutputStream(writeTo)) {
                out.write(os);
            }
        }

        return Map.of(
                "status", ok ? "ok" : "error",
                "path", (ok ? okFile : errFile).toString(),
                "debit", dTotal,
                "credit", cTotal
        );
    }

    private static void setupLogging() throws IOException {
        Files.createDirectories(LOG_FOLDER);
        String name = "etoll_log_" + System.currentTimeMillis() + ".txt";
        LOG = new PrintWriter(Files.newBufferedWriter(LOG_FOLDER.resolve(name)));
    }

    private static void println(String msg) {
        System.out.println(msg);
        if (LOG != null) {
            LOG.println(msg);
            LOG.flush();
        }
    }

    private static String safeLower(String s) {
        return s == null ? "" : s.trim().toLowerCase();
    }

    private static BigDecimal toDecimal(String s) {
        if (s == null) return BigDecimal.ZERO;
        s = s.trim().replace(",", "");
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

    private static List<Map<String, String>> readSheetToMaps(Sheet sheet) {
        List<Map<String, String>> rows = new ArrayList<>();
        Iterator<Row> it = sheet.iterator();
        if (!it.hasNext()) return rows;

        Row header = it.next();
        List<String> headers = new ArrayList<>();
        for (Cell c : header) headers.add(c.getStringCellValue().trim());

        while (it.hasNext()) {
            Row r = it.next();
            Map<String, String> row = new HashMap<>();
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = r.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                row.put(headers.get(i), cellToString(cell));
            }
            rows.add(row);
        }
        return rows;
    }

    private static String cellToString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell))
                    return cell.getLocalDateTimeCellValue().toLocalDate().toString();
                return BigDecimal.valueOf(cell.getNumericCellValue()).toPlainString();
            case BOOLEAN: return Boolean.toString(cell.getBooleanCellValue());
            case BLANK: return "";
            case FORMULA:
                try {
                    if (DateUtil.isCellDateFormatted(cell))
                        return cell.getLocalDateTimeCellValue().toLocalDate().toString();
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            default: return cell.toString();
        }
    }

    private static void forwardFill(List<Map<String, String>> rows, String col) {
        String last = "";
        for (Map<String, String> r : rows) {
            String v = r.getOrDefault(col, "").trim();
            if (!v.isEmpty()) last = v;
            else r.put(col, last);
        }
    }

    private static LocalDate findSettlementDate(List<Map<String, String>> rows) {
        for (Map<String, String> r : rows) {
            String v = r.getOrDefault(COL_SETTLEMENT_DATE, "").trim();
            if (v.isEmpty()) continue;

            try {
                if (v.matches("\\d{4}-\\d{2}-\\d{2}"))
                    return LocalDate.parse(v);
                if (v.matches("\\d{2}-\\d{2}-\\d{4}"))
                    return LocalDate.parse(v, DateTimeFormatter.ofPattern("dd-MM-yyyy"));
                if (v.matches("\\d{2}/\\d{2}/\\d{4}"))
                    return LocalDate.parse(v, DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                return LocalDate.parse(v);
            } catch (Exception ignored) {}
        }
        return null;
    }

    private static void writeVoucherSheet(Workbook wb, List<VoucherRow> voucher) {
        Sheet sh = wb.createSheet("Voucher");
        Row head = sh.createRow(0);

        head.createCell(0).setCellValue("Account No");
        head.createCell(1).setCellValue("Debit");
        head.createCell(2).setCellValue("Credit");
        head.createCell(3).setCellValue("Narration");
        head.createCell(4).setCellValue("Description");

        int r = 1;
        for (VoucherRow vr : voucher) {
            Row row = sh.createRow(r++);
            row.createCell(0).setCellValue(vr.accountNo);
            if (vr.debit != null) row.createCell(1).setCellValue(vr.debit.doubleValue());
            if (vr.credit != null) row.createCell(2).setCellValue(vr.credit.doubleValue());
            row.createCell(3).setCellValue(vr.narration);
            row.createCell(4).setCellValue(vr.description);
        }

        for (int c = 0; c < 5; c++) sh.autoSizeColumn(c);
    }

    private static void writeUploadSheet(Workbook wb, List<List<Object>> rows) {
        Sheet sh = wb.createSheet("Upload");
        int r = 0;

        for (List<Object> list : rows) {
            Row row = sh.createRow(r++);
            for (int c = 0; c < list.size(); c++) {
                Object o = list.get(c);
                Cell cell = row.createCell(c);
                if (o == null) cell.setBlank();
                else if (o instanceof Number) cell.setCellValue(((Number) o).doubleValue());
                else cell.setCellValue(o.toString());
            }
        }

        for (int c = 0; c < 4; c++) sh.autoSizeColumn(c);
    }

    private static class TemplateRow {
        final String accountNo, template, description;
        TemplateRow(String a, String b, String c) {
            accountNo = a; template = b; description = c;
        }
    }

    private static class Rule {
        final List<String> cycles;
        final String sumCol;
        final String side;
        final String special;

        Rule(List<String> cycles, String sumCol, String side, String special) {
            this.cycles = cycles;
            this.sumCol = sumCol;
            this.side = side;
            this.special = special;
        }
    }

    private static class VoucherRow {
        final String accountNo;
        final BigDecimal debit, credit;
        final String narration, description;

        VoucherRow(String a, BigDecimal d, BigDecimal c, String n, String desc) {
            accountNo = a;
            debit = d;
            credit = c;
            narration = n;
            description = desc;
        }
    }
}
