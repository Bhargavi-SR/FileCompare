package com.hcl.comparefiles;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;
import javax.net.ssl.*;
import java.security.cert.X509Certificate;
import org.apache.log4j.Logger;

/* Volt MX Middleware */
import com.hcl.voltmx.middleware.common.JavaService2;
import com.hcl.voltmx.middleware.controller.DataControllerRequest;
import com.hcl.voltmx.middleware.controller.DataControllerResponse;
import com.hcl.voltmx.middleware.dataobject.*;

/* Apache POI & PDFBox */
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class SideBySideFileCompareService implements JavaService2 {

    private static final Logger logger = Logger.getLogger(SideBySideFileCompareService.class);
    private static final double SIMILARITY_THRESHOLD = 0.65; 

    @Override
    public Object invoke(String methodID, Object[] inputArray, DataControllerRequest request, DataControllerResponse response) {
        Result result = new Result();
        Dataset ds = new Dataset("diffResults");

        try {
            String urlA = request.getParameter("urlA");
            String urlB = request.getParameter("urlB");
            String fileNameA = request.getParameter("fileNameA");
            String fileNameB = request.getParameter("fileNameB");

            disableSSL();

            List<String> leftLines = extractTextFromUrl(urlA, fileNameA);
            List<String> rightLines = extractTextFromUrl(urlB, fileNameB);

            if (leftLines.equals(rightLines) && !leftLines.isEmpty()) {
                result.addParam(new Param("isIdentical", "true"));
                result.addParam(new Param("status", "SUCCESS"));
                return result;
            }

            List<DiffRow> finalRows = performPriorityAlignment(leftLines, rightLines);

            int leftLineCounter = 1;
            int rightLineCounter = 1;

            for (DiffRow row : finalRows) {
                Record rec = new Record();
                
                rec.addParam(new Param("leftLineNo", !row.type.equals("ADDED") ? String.valueOf(leftLineCounter++) : " "));
                rec.addParam(new Param("rightLineNo", !row.type.equals("REMOVED") ? String.valueOf(rightLineCounter++) : " "));

                if ("MODIFIED".equals(row.type)) {
                    rec.addParam(new Param("leftText", getInlineDiff(row.left, row.right, true)));
                    rec.addParam(new Param("rightText", getInlineDiff(row.left, row.right, false)));
                } else {
                    rec.addParam(new Param("leftText", escapeHtml(row.left)));
                    rec.addParam(new Param("rightText", escapeHtml(row.right)));
                }
                
                rec.addParam(new Param("diffType", row.type));
                ds.addRecord(rec);
            }

            result.addDataset(ds);
            result.addParam(new Param("isIdentical", "false"));
            result.addParam(new Param("status", "SUCCESS"));

        } catch (Exception e) {
            logger.error("Error in SideBySideFileCompareService", e);
            result.addParam(new Param("status", "FAILED"));
            result.addParam(new Param("errorMessage", e.getMessage()));
        }
        return result;
    }

    private String getInlineDiff(String left, String right, boolean isOldSide) {
        String delimiter = (left.contains("|")) ? "\\|" : "\\s+";
        String joiner = (left.contains("|")) ? " | " : " ";
        String[] leftWords = left.split(delimiter);
        String[] rightWords = right.split(delimiter);
        
        Set<String> otherSide = new HashSet<>();
        for (String w : (isOldSide ? rightWords : leftWords)) otherSide.add(w.trim().toLowerCase());

        StringBuilder sb = new StringBuilder();
        String[] currentSide = isOldSide ? leftWords : rightWords;
        String highlightColor = isOldSide ? "#C0392B" : "#27AE60"; 

        for (String word : currentSide) {
            String cleanWord = word.trim();
            if (cleanWord.isEmpty()) continue;
            if (!otherSide.contains(cleanWord.toLowerCase())) {
                sb.append("<b><font color='").append(highlightColor).append("'>").append(escapeHtml(cleanWord)).append("</font></b>").append(joiner);
            } else {
                sb.append(escapeHtml(cleanWord)).append(joiner);
            }
        }
        return sb.toString().trim();
    }

    private String escapeHtml(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }

    private List<DiffRow> performPriorityAlignment(List<String> a, List<String> b) {
        int m = a.size(), n = b.size();
        int[][] lcs = new int[m + 1][n + 1];
        for (int i = 1; i <= m; i++) {
            for (int j = 1; j <= n; j++) {
                if (a.get(i - 1).equals(b.get(j - 1))) lcs[i][j] = lcs[i - 1][j - 1] + 1;
                else lcs[i][j] = Math.max(lcs[i - 1][j], lcs[i][j - 1]);
            }
        }
        List<DiffRow> result = new ArrayList<>();
        int i = m, j = n;
        while (i > 0 || j > 0) {
            if (i > 0 && j > 0 && a.get(i - 1).equals(b.get(j - 1))) {
                result.add(new DiffRow(a.get(i - 1), b.get(j - 1), "UNCHANGED"));
                i--; j--;
            } else if (i > 0 && j > 0 && getSimilarity(a.get(i - 1), b.get(j - 1)) >= SIMILARITY_THRESHOLD) {
                result.add(new DiffRow(a.get(i - 1), b.get(j - 1), "MODIFIED"));
                i--; j--;
            } else if (i > 0 && (j == 0 || lcs[i - 1][j] >= lcs[i][j - 1])) {
                result.add(new DiffRow(a.get(i - 1), "", "REMOVED"));
                i--;
            } else {
                result.add(new DiffRow("", b.get(j - 1), "ADDED"));
                j--;
            }
        }
        Collections.reverse(result);
        return result;
    }

    private double getSimilarity(String s1, String s2) {
        if (s1.isEmpty() || s2.isEmpty()) return 0;
        // Table Alignment Boost: If it's a table row and the first column matches, boost similarity
        if (s1.contains("|") && s2.contains("|")) {
            String p1 = s1.split("\\|")[0].trim();
            String p2 = s2.split("\\|")[0].trim();
            if (p1.equalsIgnoreCase(p2) && !p1.isEmpty()) return 0.8; 
        }
        int distance = levenshteinDistance(s1, s2);
        return 1.0 - ((double) distance / Math.max(s1.length(), s2.length()));
    }

    private int levenshteinDistance(String s1, String s2) {
        int[] costs = new int[s2.length() + 1];
        for (int j = 0; j <= s2.length(); j++) costs[j] = j;
        for (int i = 1; i <= s1.length(); i++) {
            costs[0] = i;
            int nw = i - 1;
            for (int j = 1; j <= s2.length(); j++) {
                int cj = Math.min(1 + Math.min(costs[j], costs[j - 1]), 
                         s1.charAt(i - 1) == s2.charAt(j - 1) ? nw : nw + 1);
                nw = costs[j];
                costs[j] = cj;
            }
        }
        return costs[s2.length()];
    }

    private List<String> extractTextFromUrl(String urlStr, String fileName) throws Exception {
        File file = downloadFile(urlStr, fileName);
        String ext = fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase();
        List<String> lines = new ArrayList<>();
        try (InputStream fis = new FileInputStream(file)) {
            if (ext.equals("pdf")) {
                try (PDDocument doc = Loader.loadPDF(file)) {
                    PDFTextStripper stripper = new PDFTextStripper();
                    for (String s : stripper.getText(doc).split("\\r?\\n")) if (!s.trim().isEmpty()) lines.add(s.trim());
                }
            } else if (ext.startsWith("doc")) {
                try (XWPFDocument doc = new XWPFDocument(fis)) {
                    for (IBodyElement el : doc.getBodyElements()) {
                        if (el instanceof XWPFParagraph) {
                            String t = ((XWPFParagraph) el).getText().replaceAll("\\s+", " ").trim();
                            if (!t.isEmpty()) lines.add(t);
                        } else if (el instanceof XWPFTable) {
                            for (XWPFTableRow row : ((XWPFTable) el).getRows()) {
                                StringBuilder sb = new StringBuilder();
                                List<XWPFTableCell> cells = row.getTableCells();
                                for (int k = 0; k < cells.size(); k++) {
                                    sb.append(cells.get(k).getText().replaceAll("\\s+", " ").trim());
                                    if (k < cells.size() - 1) sb.append(" | ");
                                }
                                lines.add(sb.toString().trim());
                            }
                        }
                    }
                }
            } else if (ext.startsWith("xls")) {
                try (Workbook wb = WorkbookFactory.create(fis)) {
                    DataFormatter formatter = new DataFormatter();
                    Sheet s = wb.getSheetAt(0);
                    for (Row r : s) {
                        StringBuilder sb = new StringBuilder();
                        int lastCol = r.getLastCellNum();
                        for (int cn = 0; cn < lastCol; cn++) {
                            Cell c = r.getCell(cn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            sb.append(formatter.formatCellValue(c).replaceAll("\\s+", " ").trim());
                            if (cn < lastCol - 1) sb.append(" | ");
                        }
                        if (sb.length() > 0) lines.add(sb.toString().trim());
                    }
                }
            }
        } finally { if (file != null) file.delete(); }
        return lines;
    }

    private File downloadFile(String urlStr, String name) throws Exception {
        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");
        File f = File.createTempFile("cmp_", "_" + name);
        try (InputStream in = conn.getInputStream(); FileOutputStream out = new FileOutputStream(f)) {
            byte[] buf = new byte[4096];
            int len;
            while ((len = in.read(buf)) != -1) out.write(buf, 0, len);
        }
        return f;
    }

    private void disableSSL() throws Exception {
        TrustManager[] trustAll = new TrustManager[]{
            new X509TrustManager() {
                public X509Certificate[] getAcceptedIssuers() { return null; }
                public void checkClientTrusted(X509Certificate[] c, String a) {}
                public void checkServerTrusted(X509Certificate[] c, String a) {}
            }
        };
        SSLContext sc = SSLContext.getInstance("TLS");
        sc.init(null, trustAll, new java.security.SecureRandom());
        HttpsURLConnection.setDefaultSSLSocketFactory(sc.getSocketFactory());
        HttpsURLConnection.setDefaultHostnameVerifier((h, s) -> true);
    }

    static class DiffRow {
        String left, right, type;
        DiffRow(String l, String r, String t) { 
            this.left = l == null ? "" : l; 
            this.right = r == null ? "" : r; 
            this.type = t; 
        }
    }
}