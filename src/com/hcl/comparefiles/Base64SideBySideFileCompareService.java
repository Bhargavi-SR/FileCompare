package com.hcl.comparefiles;

import java.io.*;
import java.util.*;
import java.util.Base64;
import org.apache.log4j.Logger;

/* Volt MX Middleware */
import com.hcl.voltmx.middleware.common.JavaService2;
import com.hcl.voltmx.middleware.controller.DataControllerRequest;
import com.hcl.voltmx.middleware.controller.DataControllerResponse;
import com.hcl.voltmx.middleware.dataobject.*;

/* Apache POI & PDFBox */
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class Base64SideBySideFileCompareService implements JavaService2 {

    private static final Logger logger = Logger.getLogger(Base64SideBySideFileCompareService.class);
    private static final double SIMILARITY_THRESHOLD = 0.70; 

    @Override
    public Object invoke(String methodID, Object[] inputArray, DataControllerRequest request, DataControllerResponse response) {
        Result result = new Result();
        Dataset ds = new Dataset("diffResults");

        try {
            String base64A = request.getParameter("fileA_base64");
            String base64B = request.getParameter("fileB_base64");
            String extA = request.getParameter("extA").toLowerCase();
            String extB = request.getParameter("extB").toLowerCase();

            List<String> leftLines = extractText(base64A, extA);
            List<String> rightLines = extractText(base64B, extB);

            List<DiffRow> finalRows = performAlignment(leftLines, rightLines);

            for (DiffRow row : finalRows) {
                Record rec = new Record();
                
                // INTEGRATED: Generate inline word-level diff for RichText if MODIFIED
                if ("MODIFIED".equals(row.type)) {
                    rec.addParam(new Param("leftText", getInlineDiff(row.left, row.right, true)));
                    rec.addParam(new Param("rightText", getInlineDiff(row.left, row.right, false)));
                } else {
                    rec.addParam(new Param("leftText", row.left));
                    rec.addParam(new Param("rightText", row.right));
                }
                
                rec.addParam(new Param("diffType", row.type));
                ds.addRecord(rec);
            }

            result.addDataset(ds);
            result.addParam(new Param("status", "SUCCESS"));
        } catch (Exception e) {
            logger.error("Error comparing files", e);
            result.addParam(new Param("status", "FAILED"));
            result.addParam(new Param("errorMessage", e.getMessage()));
        }
        return result;
    }

    /* ================= WORD-LEVEL HIGHLIGHTING FOR RICHTEXT ================= */
    
    private String getInlineDiff(String left, String right, boolean isOldSide) {
        // Detect if table row or paragraph to use correct delimiter
        String delimiter = (left.contains("|")) ? "\\|" : "\\s+";
        String joiner = (left.contains("|")) ? " | " : " ";
        
        String[] leftWords = left.split(delimiter);
        String[] rightWords = right.split(delimiter);
        
        Set<String> otherSide = new HashSet<>();
        for (String w : (isOldSide ? rightWords : leftWords)) {
            otherSide.add(w.trim());
        }

        StringBuilder sb = new StringBuilder();
        String[] currentSide = isOldSide ? leftWords : rightWords;
        
        // Highlight Color: Red for deletions (Old), Green for additions (New)
        String highlightColor = isOldSide ? "#C0392B" : "#27AE60"; 

        for (String word : currentSide) {
            String cleanWord = word.trim();
            if (cleanWord.isEmpty()) continue;

            if (!otherSide.contains(cleanWord)) {
                // Wrap only the modified words in Font tags for the Iris RichText widget
                sb.append("<b><font color='").append(highlightColor).append("'>")
                  .append(cleanWord).append("</font></b>").append(joiner);
            } else {
                sb.append(cleanWord).append(joiner);
            }
        }
        return sb.toString().trim();
    }

    /* ================= ALIGNMENT & SIMILARITY ================= */

    private List<DiffRow> performAlignment(List<String> a, List<String> b) {
        int m = a.size(), n = b.size();
        int[][] lcs = new int[m + 1][n + 1];

        for (int i = 1; i <= m; i++) {
            for (int j = 1; j <= n; j++) {
                if (a.get(i - 1).equals(b.get(j - 1))) {
                    lcs[i][j] = lcs[i - 1][j - 1] + 1;
                } else {
                    lcs[i][j] = Math.max(lcs[i - 1][j], lcs[i][j - 1]);
                }
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
        if (s1.contains("|") && s2.contains("|")) {
            String prefix1 = s1.split("\\|")[0].trim();
            String prefix2 = s2.split("\\|")[0].trim();
            if (!prefix1.equals(prefix2)) return 0.0;
        } else if (s1.contains("|") != s2.contains("|")) {
            return 0.0;
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

    /* ================= EXTRACTION LOGIC ================= */

    private List<String> extractText(String base64, String ext) throws Exception {
        String clean = base64.contains(",") ? base64.split(",")[1] : base64;
        byte[] bytes = Base64.getDecoder().decode(clean.replaceAll("\\s+", ""));
        List<String> lines = new ArrayList<>();

        try (InputStream is = new ByteArrayInputStream(bytes)) {
            if (ext.endsWith("docx")) {
                try (XWPFDocument doc = new XWPFDocument(is)) {
                    for (IBodyElement element : doc.getBodyElements()) {
                        if (element instanceof XWPFParagraph) {
                            String text = ((XWPFParagraph) element).getText();
                            if (!text.trim().isEmpty()) lines.add(text.trim());
                        } else if (element instanceof XWPFTable) {
                            for (XWPFTableRow row : ((XWPFTable) element).getRows()) {
                                StringBuilder sb = new StringBuilder();
                                for (XWPFTableCell cell : row.getTableCells()) {
                                    sb.append(cell.getText().trim()).append(" | ");
                                }
                                lines.add(sb.toString().trim());
                            }
                        }
                    }
                }
            } else if (ext.endsWith("xlsx")) {
                try (Workbook wb = WorkbookFactory.create(is)) {
                    Sheet sheet = wb.getSheetAt(0);
                    for (Row row : sheet) {
                        StringBuilder sb = new StringBuilder();
                        for (Cell cell : row) {
                            sb.append(cell.toString().trim()).append(" | ");
                        }
                        if (sb.length() > 0) lines.add(sb.toString().trim());
                    }
                }
            } else if (ext.endsWith("pdf")) {
                try (PDDocument pdf = Loader.loadPDF(bytes)) {
                    PDFTextStripper stripper = new PDFTextStripper();
                    String text = stripper.getText(pdf);
                    for (String line : text.split("\\r?\\n")) {
                        if (!line.trim().isEmpty()) lines.add(line.trim());
                    }
                }
            }
        }
        return lines;
    }

    static class DiffRow {
        String left, right, type;
        DiffRow(String l, String r, String t) {
            this.left = (l == null) ? "" : l;
            this.right = (r == null) ? "" : r;
            this.type = t;
        }
    }
}