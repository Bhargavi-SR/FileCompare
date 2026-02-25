package com.hcl.comparefiles;

import java.io.*;
import java.util.*;
import java.util.Base64;
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
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class Base64SideBySideFileCompareService implements JavaService2 {

    private static final Logger logger = Logger.getLogger(Base64SideBySideFileCompareService.class);
    private static final double SIMILARITY_THRESHOLD = 0.65;
    private static final int MAX_IMAGE_WIDTH = 400; // Optimal for mobile display

    @Override
    public Object invoke(String methodID, Object[] inputArray, DataControllerRequest request, DataControllerResponse response) {
        Result result = new Result();
        Dataset ds = new Dataset("diffResults");

        try {
            String base64A = request.getParameter("fileA_base64");
            String base64B = request.getParameter("fileB_base64");
            String extA = request.getParameter("extA");
            String extB = request.getParameter("extB");

            if (base64A == null || base64B == null) throw new Exception("Input base64 missing");

            List<String> leftLines = extractText(base64A, extA);
            List<String> rightLines = extractText(base64B, extB);

            List<DiffRow> finalRows = performAlignment(leftLines, rightLines);

            boolean hasChanges = false;
            for (DiffRow row : finalRows) {
                if (!"UNCHANGED".equals(row.type)) hasChanges = true;
                Record rec = new Record();
                
                // IMAGE DETECTION LOGIC
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

                rec.addParam(new Param("leftLineNo", row.leftNo > 0 ? String.valueOf(row.leftNo) : ""));
                rec.addParam(new Param("rightLineNo", row.rightNo > 0 ? String.valueOf(row.rightNo) : ""));
                rec.addParam(new Param("diffType", row.type));
                ds.addRecord(rec);
            }

            result.addDataset(ds);
            result.addParam(new Param("status", "SUCCESS"));
            result.addParam(new Param("isIdentical", String.valueOf(!hasChanges)));

        } catch (Exception e) {
            logger.error("Service Critical Failure", e);
            result.addParam(new Param("status", "FAILED"));
            result.addParam(new Param("errorMessage", e.getMessage()));
        }
        return result;
    }

    private List<String> extractText(String base64, String ext) throws Exception {
        String cleanB64 = base64.contains(",") ? base64.split(",")[1] : base64;
        byte[] bytes = Base64.getDecoder().decode(cleanB64.replaceAll("\\s+", ""));
        List<String> lines = new ArrayList<>();
        String nExt = (ext == null) ? "" : ext.toLowerCase().trim();
        boolean isZip = (bytes.length > 2 && bytes[0] == 0x50 && bytes[1] == 0x4B);

        try (InputStream is = new ByteArrayInputStream(bytes)) {
        	if (ext.endsWith("docx")) {
                try (XWPFDocument doc = new XWPFDocument(is)) {
                    for (IBodyElement element : doc.getBodyElements()) {
                        if (element instanceof XWPFParagraph) {
                            String text = ((XWPFParagraph) element).getText();
                            if (text != null && !text.trim().isEmpty()) lines.add(text.trim());
                        } else if (element instanceof XWPFTable) {
                            for (XWPFTableRow row : ((XWPFTable) element).getRows()) {
                                StringBuilder sb = new StringBuilder();
                                for (XWPFTableCell cell : row.getTableCells()) {
                                    sb.append(cell.getText().replace("\n", " ").trim()).append(" | ");
                                }
                                lines.add(sb.toString().trim());
                            }
                        }
                    }
                }
            } 
            // FIXED: Added legacy .doc handler to prevent binary gibberish
	        else if (ext.equals("doc")) {
	            try (HWPFDocument doc = new HWPFDocument(is); 
	                 WordExtractor extractor = new WordExtractor(doc)) {
	                for (String p : extractor.getParagraphText()) {
	                    if (p != null && !p.trim().isEmpty()) lines.add(p.trim());
	                }
	            }
	        } 
	        // PPTX Extraction with Image Handling
	        else if (nExt.contains("pptx") || (isZip && nExt.isEmpty())) {
	            try (XMLSlideShow ppt = new XMLSlideShow(is)) {
	                for (XSLFSlide slide : ppt.getSlides()) {
	                    for (XSLFShape shape : slide.getShapes()) {
	                        if (shape instanceof XSLFTextShape) {
	                            for (XSLFTextParagraph p : ((XSLFTextShape) shape).getTextParagraphs()) {
	                                String bullet = p.isBullet() ? "• " : "";
	                                if (!p.getText().trim().isEmpty()) lines.add(bullet + p.getText().trim());
	                            }
	                        } else if (shape instanceof XSLFPictureShape) {
	                            try {
	                                XSLFPictureData data = ((XSLFPictureShape) shape).getPictureData();
	                                byte[] resized = resizeImage(data.getData(), MAX_IMAGE_WIDTH);
	                                lines.add("IMG_DATA:image/jpeg;base64," + Base64.getEncoder().encodeToString(resized));
	                            } catch (Exception imgEx) {
	                                logger.warn("Image resize failed", imgEx);
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
	        }else if (ext.contains("xls")) {
	                    try (Workbook wb = WorkbookFactory.create(is)) {
	                        Sheet sheet = wb.getSheetAt(0);
	                        DataFormatter formatter = new DataFormatter();
	                        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
	                        for (Row row : sheet) {
	                            StringBuilder sb = new StringBuilder();
	                            for (Cell cell : row) {
	                                sb.append(formatter.formatCellValue(cell, evaluator)).append(" | ");
	                            }
	                            if (sb.length() > 0) lines.add(sb.toString().trim());
	                        }
	                    }
	                } else if (ext.endsWith("pdf")) {
	                    try (PDDocument pdf = Loader.loadPDF(bytes)) {
	                        String text = new PDFTextStripper().getText(pdf);
	                        for (String line : text.split("\\r?\\n")) {
	                            if (!line.trim().isEmpty()) lines.add(line.trim());
	                        }
	                    }
	                } else {
	                    try (BufferedReader reader = new BufferedReader(new InputStreamReader(is, "UTF-8"))) {
	                        String line;
	                        while ((line = reader.readLine()) != null) {
	                            if (line.trim().isEmpty()) continue;
	                            if (ext.contains("csv")) {
	                                String[] parts = line.split(",", -1);
	                                StringBuilder sb = new StringBuilder();
	                                for (String p : parts) sb.append(p.trim().isEmpty() ? " " : p.trim()).append(" | ");
	                                lines.add(sb.toString().trim());
	                            } else lines.add(line.trim());
	                        }
	                    }
	                }
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
            ImageIO.write(resized, "jpg", baos); // JPEG is much smaller for Base64
            return baos.toByteArray();
        }
    }

    // --- LCS Alignment Algorithm ---
    private List<DiffRow> performAlignment(List<String> a, List<String> b) {
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
                result.add(new DiffRow(a.get(i - 1), b.get(j - 1), "UNCHANGED", i, j));
                i--; j--;
            } else if (i > 0 && j > 0 && getSimilarity(a.get(i - 1), b.get(j - 1)) >= SIMILARITY_THRESHOLD) {
                result.add(new DiffRow(a.get(i - 1), b.get(j - 1), "MODIFIED", i, j));
                i--; j--;
            } else if (i > 0 && (j == 0 || lcs[i - 1][j] >= lcs[i][j - 1])) {
                result.add(new DiffRow(a.get(i - 1), "", "REMOVED", i, 0));
                i--;
            } else {
                result.add(new DiffRow("", b.get(j - 1), "ADDED", 0, j));
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

    static class DiffRow {
        String left, right, type;
        int leftNo, rightNo;
        DiffRow(String l, String r, String t, int lNo, int rNo) {
            this.left = (l == null) ? "" : l;
            this.right = (r == null) ? "" : r;
            this.type = t;
            this.leftNo = lNo;
            this.rightNo = rNo;
        }
    }
}