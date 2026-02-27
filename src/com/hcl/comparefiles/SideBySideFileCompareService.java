package com.hcl.comparefiles;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;
import java.util.Base64;
import javax.net.ssl.*;
import java.security.cert.X509Certificate;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import javax.imageio.ImageIO;
import org.apache.log4j.Logger;

/* Volt MX Middleware */
import com.hcl.voltmx.middleware.common.JavaService2;
import com.hcl.voltmx.middleware.controller.DataControllerRequest;
import com.hcl.voltmx.middleware.controller.DataControllerResponse;
import com.hcl.voltmx.middleware.dataobject.*;

/* Apache POI & PDFBox */
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class SideBySideFileCompareService implements JavaService2 {

    private static final Logger logger = Logger.getLogger(SideBySideFileCompareService.class);
    private static final double SIMILARITY_THRESHOLD = 0.50;
    private static final int MAX_IMAGE_WIDTH = 400;

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

            List<DiffRow> finalRows = performPriorityAlignment(leftLines, rightLines);

            int leftLineCounter = 1;
            int rightLineCounter = 1;

            for (DiffRow row : finalRows) {
                Record rec = new Record();
                
                // IMAGE DETECTION & FORMATTING
                if (row.left.startsWith("IMG_DATA:") || row.right.startsWith("IMG_DATA:")) {
                    rec.addParam(new Param("isImage", "true"));
                    rec.addParam(new Param("leftText", row.left.replace("IMG_DATA:", "")));
                    rec.addParam(new Param("rightText", row.right.replace("IMG_DATA:", "")));
                } else {
                    rec.addParam(new Param("isImage", "false"));
                    if ("MODIFIED".equals(row.type)) {
                        rec.addParam(new Param("leftText", getInlineDiff(row.left, row.right, true)));
                        rec.addParam(new Param("rightText", getInlineDiff(row.left, row.right, false)));
                    } else {
                        rec.addParam(new Param("leftText", escapeHtml(row.left)));
                        rec.addParam(new Param("rightText", escapeHtml(row.right)));
                    }
                }

                rec.addParam(new Param("leftLineNo", !row.type.equals("ADDED") ? String.valueOf(leftLineCounter++) : " "));
                rec.addParam(new Param("rightLineNo", !row.type.equals("REMOVED") ? String.valueOf(rightLineCounter++) : " "));
                rec.addParam(new Param("diffType", row.type));
                ds.addRecord(rec);
            }

            result.addDataset(ds);
            result.addParam(new Param("isIdentical", String.valueOf(leftLines.equals(rightLines))));
            result.addParam(new Param("status", "SUCCESS"));

        } catch (Exception e) {
            logger.error("Error in SideBySideFileCompareService", e);
            result.addParam(new Param("status", "FAILED"));
            result.addParam(new Param("errorMessage", e.getMessage()));
        }
        return result;
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
            } else if (ext.equals("pptx")) {
                try (XMLSlideShow ppt = new XMLSlideShow(fis)) {
                    for (XSLFSlide slide : ppt.getSlides()) {
                        for (XSLFShape shape : slide.getShapes()) {
                            if (shape instanceof XSLFTextShape) {
                                for (XSLFTextParagraph para : ((XSLFTextShape) shape).getTextParagraphs()) {
                                    String bullet = para.isBullet() ? "• " : "";
                                    String txt = para.getText().trim();
                                    if (!txt.isEmpty()) lines.add(bullet + txt);
                                }
                            } else if (shape instanceof XSLFPictureShape) {
                                try {
                                    XSLFPictureData data = ((XSLFPictureShape) shape).getPictureData();
                                    byte[] resized = resizeImage(data.getData(), MAX_IMAGE_WIDTH);
                                    lines.add("IMG_DATA:image/jpeg;base64," + Base64.getEncoder().encodeToString(resized));
                                } catch (Exception imgEx) {
                                    logger.warn("Image skip: " + imgEx.getMessage());
                                }
                            } else if (shape instanceof XSLFTable) {
                                for (XSLFTableRow row : ((XSLFTable) shape).getRows()) {
                                    StringBuilder rowSb = new StringBuilder();
                                    for (XSLFTableCell cell : row.getCells()) {
                                        rowSb.append(cell.getText().replace("\n", " ").trim()).append(" | ");
                                    }
                                    lines.add(rowSb.toString().trim());
                                }
                            }
                        }
                    }
                }
            } else if (ext.startsWith("doc")) {
                try (XWPFDocument doc = new XWPFDocument(fis)) {
                    for (IBodyElement el : doc.getBodyElements()) {
                        if (el instanceof XWPFParagraph) {
                            String t = ((XWPFParagraph) el).getText().trim();
                            if (!t.isEmpty()) lines.add(t);
                        } else if (el instanceof XWPFTable) {
                            for (XWPFTableRow row : ((XWPFTable) el).getRows()) {
                                StringBuilder sb = new StringBuilder();
                                for (XWPFTableCell cell : row.getTableCells()) {
                                    sb.append(cell.getText().replace("\n", " ").trim()).append(" | ");
                                }
                                lines.add(sb.toString().trim());
                            }
                        }
                    }
                }
            } else if (ext.startsWith("xls")) {
                try (Workbook wb = WorkbookFactory.create(fis)) {
                                DataFormatter formatter = new DataFormatter();
                                FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                    Sheet s = wb.getSheetAt(0);
                              // DataFormatter df = new DataFormatter();
                    for (Row r : s) {
                        StringBuilder sb = new StringBuilder();
                                    int lastCol = r.getLastCellNum();
                                    for (int cn = 0; cn < lastCol; cn++) {
                                        Cell c = r.getCell(cn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        String val = formatter.formatCellValue(c, evaluator).replaceAll("\\s+", " ").trim();
                                        sb.append(val);
                                        if (cn < lastCol - 1) sb.append(" | ");
                                    }
                                  // for (Cell c : r) sb.append(df.formatCellValue(c)).append(" | ");
                        if (sb.length() > 0) lines.add(sb.toString().trim());
                    }
                }
            }
        } finally { 
            if (file != null && file.exists()) file.delete(); 
        }
        return lines;
    }

    private byte[] resizeImage(byte[] originalData, int maxWidth) throws Exception {
        try (InputStream in = new ByteArrayInputStream(originalData)) {
            BufferedImage originalImage = ImageIO.read(in);
            if (originalImage == null) return originalData;
            int width = originalImage.getWidth();
            int height = originalImage.getHeight();
            if (width <= maxWidth) return originalData;
            int newHeight = (maxWidth * height) / width;
            BufferedImage resized = new BufferedImage(maxWidth, newHeight, BufferedImage.TYPE_INT_RGB);
            Graphics2D g = resized.createGraphics();
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            g.drawImage(originalImage, 0, 0, maxWidth, newHeight, null);
            g.dispose();
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(resized, "jpg", baos);
            return baos.toByteArray();
        }
    }

    private String getInlineDiff(String left, String right, boolean isOldSide) {
        String delimiter = left.contains("|") ? "\\|" : "\\s+";
        String joiner = left.contains("|") ? " | " : " ";
        String[] lWords = left.split(delimiter, -1);
        String[] rWords = right.split(delimiter, -1);
        Set<String> otherSide = new HashSet<>();
        for (String w : (isOldSide ? rWords : lWords)) otherSide.add(w.trim().toLowerCase());

        StringBuilder sb = new StringBuilder();
        String[] currentSide = isOldSide ? lWords : rWords;
        String color = isOldSide ? "#C0392B" : "#27AE60"; 

        for (String word : currentSide) {
            if (word.trim().isEmpty()) { sb.append(joiner); continue; }
            if (!otherSide.contains(word.trim().toLowerCase())) {
                sb.append("<b><font color='").append(color).append("'>").append(escapeHtml(word.trim())).append("</font></b>").append(joiner);
            } else {
                sb.append(escapeHtml(word.trim())).append(joiner);
            }
        }
        return sb.toString().trim();
    }

    private String escapeHtml(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;").replace("'", "&apos;");
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
        if (s1.isEmpty() || s2.isEmpty() || s1.startsWith("IMG_DATA") || s2.startsWith("IMG_DATA")) return 0;
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