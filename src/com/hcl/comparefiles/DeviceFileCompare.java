package com.hcl.comparefiles;

import com.hcl.voltmx.middleware.common.JavaService2;
import com.hcl.voltmx.middleware.controller.DataControllerRequest;
import com.hcl.voltmx.middleware.controller.DataControllerResponse;
import com.hcl.voltmx.middleware.dataobject.Param;
import com.hcl.voltmx.middleware.dataobject.Result;
import com.hcl.voltmx.middleware.dataobject.Dataset;
import com.hcl.voltmx.middleware.dataobject.Record;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.log4j.Logger;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.*;
import java.util.*;

public class DeviceFileCompare implements JavaService2 {
    private Logger logger = Logger.getLogger(DeviceFileCompare.class);

    @Override
    public Object invoke(String methodID, Object[] inputArray, DataControllerRequest request, 
                         DataControllerResponse response) throws Exception {
        
        Result result = new Result();
        Dataset dsResults = new Dataset("comparisonResults");

        try {
        	logger.error("got into devicefilecompare java service");
            String base64A = (String) request.getParameter("fileA_base64");
            String base64B = (String) request.getParameter("fileB_base64");
            String extA = (String) request.getParameter("extA");
            String extB = (String) request.getParameter("extB");

            if (base64A == null || base64B == null) {
                throw new Exception("Input files are missing.");
            }

            List<String> contentA = extractFromBase64(base64A, extA);
            List<String> contentB = extractFromBase64(base64B, extB);

            compareContent(contentA, contentB, dsResults);

            result.addDataset(dsResults);
            result.addParam(new Param("status", "success", "string"));

        } catch (Exception e) {
            result.addParam(new Param("status", "error", "string"));
            result.addParam(new Param("message", e.getMessage(), "string"));
        }

        return result;
    }

    private List<String> extractFromBase64(String base64, String ext) throws Exception {
    	logger.error("before base64--->"+base64);
        String sanitizedBase64 = base64.replace("%2F", "/").replace("%2B", "+").replace("%3D", "=");
        logger.error("after base64--->"+sanitizedBase64);
        byte[] bytes = Base64.getDecoder().decode(sanitizedBase64);
        List<String> lines = new ArrayList<>();
        ext = ext.toLowerCase();

        try (InputStream is = new ByteArrayInputStream(bytes)) {
            if (ext.endsWith(".pdf")) {
                try (PDDocument pdf = Loader.loadPDF(bytes)) {
                    PDFTextStripper stripper = new PDFTextStripper();
                    lines.addAll(Arrays.asList(stripper.getText(pdf).split("\\r?\\n")));
                }
            } else if (ext.endsWith(".docx") || ext.endsWith(".doc")) {
                try {
                    XWPFDocument docx = new XWPFDocument(is);
                    // Updated to use getBodyElements to capture Tables
                    for (IBodyElement element : docx.getBodyElements()) {
                        if (element instanceof XWPFParagraph) {
                            XWPFParagraph p = (XWPFParagraph) element;
                            if (!p.getText().trim().isEmpty()) {
                                lines.add(p.getText());
                            }
                        } else if (element instanceof XWPFTable) {
                            XWPFTable table = (XWPFTable) element;
                            for (XWPFTableRow row : table.getRows()) {
                                StringBuilder rowData = new StringBuilder("[TABLE ROW]: ");
                                for (XWPFTableCell cell : row.getTableCells()) {
                                    rowData.append(cell.getText().trim()).append(" | ");
                                }
                                lines.add(rowData.toString());
                            }
                        }
                    }
                    docx.close();
                } catch (Exception e) {
                    try (InputStream is2 = new ByteArrayInputStream(bytes);
                         WordExtractor extractor = new WordExtractor(is2)) {
                        lines.addAll(Arrays.asList(extractor.getParagraphText()));
                    }
                }
            } else if (ext.endsWith(".xlsx") || ext.endsWith(".xls")) {
                try (Workbook wb = WorkbookFactory.create(is)) {
                    Sheet sheet = wb.getSheetAt(0);
                    for (Row row : sheet) {
                        StringBuilder sb = new StringBuilder();
                        for (Cell cell : row) sb.append(cell.toString()).append(" | ");
                        lines.add(sb.toString());
                    }
                }
            }
        }
        return lines;
    }

    private void compareContent(List<String> list1, List<String> list2, Dataset ds) {
        int max = Math.max(list1.size(), list2.size());
        for (int i = 0; i < max; i++) {
            String s1 = i < list1.size() ? list1.get(i).trim() : "[End of File A]";
            String s2 = i < list2.size() ? list2.get(i).trim() : "[End of File B]";

            if (!s1.equals(s2)) {
                Record record = new Record();
                record.addParam(new Param("index", String.valueOf(i + 1), "string"));
                record.addParam(new Param("lineA", s1, "string"));
                record.addParam(new Param("lineB", s2, "string"));
                record.addParam(new Param("diff", getWordDiff(s1, s2), "string"));
                ds.addRecord(record);
            }
        }
    }

    private String getWordDiff(String oldT, String newT) {
        StringBuilder sb = new StringBuilder();
        // Detect if we are comparing a table row to use cell-based splitting
        boolean isTable = oldT.startsWith("[TABLE ROW]:") || newT.startsWith("[TABLE ROW]:");
        String delimiter = isTable ? "\\|" : "\\s+";

        List<String> oldWords = Arrays.asList(oldT.split(delimiter));
        String[] newWords = newT.split(delimiter);

        for (String word : newWords) {
            String cleanWord = word.trim();
            if (cleanWord.equals("[TABLE ROW]:")) continue;

            if (!oldWords.contains(word) && !cleanWord.isEmpty()) {
                sb.append("[+").append(cleanWord).append("+] ");
            } else if (!cleanWord.isEmpty()) {
                sb.append(cleanWord).append(" ");
            }
        }
        return sb.toString().trim();
    }
}