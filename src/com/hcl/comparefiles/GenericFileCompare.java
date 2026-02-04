package com.hcl.comparefiles;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.*;
import java.util.*;

public class GenericFileCompare {

    public static void main(String[] args) {
        try {
            // Replace these with your test file paths
            String fileA = "input/DocTable1.docx";
            String fileB = "input/DocTable2.docx";

            System.out.println("Processing Comparison...");
            List<String> content1 = extractText(fileA);
            List<String> content2 = extractText(fileB);

            compareLines(content1, content2);

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Extracts text from various formats into a List of Strings (lines/paragraphs)
     */
    public static List<String> extractText(String filePath) throws Exception {
        List<String> lines = new ArrayList<>();
        File file = new File(filePath);
        String ext = filePath.substring(filePath.lastIndexOf(".")).toLowerCase();

        switch (ext) {
            case ".txt":
            	System.out.println("xls file");
                try (BufferedReader br = new BufferedReader(new FileReader(file))) {
                    String line;
                    while ((line = br.readLine()) != null) lines.add(line);
                }
                break;

            case ".xlsx":
            case ".xls":
            	System.out.println("xls/xlsx file");
                try (Workbook wb = WorkbookFactory.create(file)) {
                    Sheet sheet = wb.getSheetAt(0);
                    for (Row row : sheet) {
                        StringBuilder rowText = new StringBuilder();
                        for (Cell cell : row) {
                            rowText.append(cell.toString()).append(" | ");
                        }
                        lines.add(rowText.toString());
                    }
                }
                break;

            case ".docx":
            case ".doc":
                System.out.println("Processing doc/docx file with tables...");
                try (InputStream is = new FileInputStream(file)) {
                    try {
                        XWPFDocument docx = new XWPFDocument(is);
                        
                        // Use getBodyElements to maintain the order of paragraphs and tables
                        for (IBodyElement element : docx.getBodyElements()) {
                            if (element instanceof XWPFParagraph) {
                                XWPFParagraph p = (XWPFParagraph) element;
                                if (!p.getText().trim().isEmpty()) {
                                    lines.add(p.getText());
                                }
                            } else if (element instanceof XWPFTable) {
                                XWPFTable table = (XWPFTable) element;
                                System.out.println("Table found, extracting rows...");
                                
                                // Iterate through table rows and cells
                                for (XWPFTableRow row : table.getRows()) {
                                    StringBuilder rowData = new StringBuilder("[TABLE ROW]: ");
                                    for (XWPFTableCell cell : row.getTableCells()) {
                                        // cell.getText() aggregates all text within that cell
                                        rowData.append(cell.getText().trim()).append(" | ");
                                    }
                                    lines.add(rowData.toString());
                                }
                            }
                        }
                        docx.close();
                        System.out.println("Parsed as .docx successfully.");
                    } catch (org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException | org.apache.poi.poifs.filesystem.OfficeXmlFileException e) {
                        // Fallback logic for legacy .doc binary files
                        System.out.println("Falling back to HWPF for binary .doc...");
                        try (InputStream is2 = new FileInputStream(file)) {
                            WordExtractor extractor = new WordExtractor(is2);
                            lines.addAll(Arrays.asList(extractor.getParagraphText()));
                            extractor.close();
                        }
                    }
                }
                break;
            case ".pdf":
            	System.out.println("pdf file");
                try (PDDocument pdf = Loader.loadPDF(file)) {
                    PDFTextStripper stripper = new PDFTextStripper();
                    String text = stripper.getText(pdf);
                    lines.addAll(Arrays.asList(text.split("\\r?\\n")));
                }
                break;

            default:
                throw new IllegalArgumentException("Unsupported file format: " + ext);
        }
        return lines;
    }

    /**
     * Logic to compare two lists of strings
     */
    public static void compareLines(List<String> list1, List<String> list2) {
        int max = Math.max(list1.size(), list2.size());
        boolean foundDifference = false;

        for (int i = 0; i < max; i++) {
            String s1 = i < list1.size() ? list1.get(i).trim() : "";
            String s2 = i < list2.size() ? list2.get(i).trim() : "";

            if (!s1.equals(s2)) {
                foundDifference = true;
                System.out.println("\n--------------------------------------");
                System.out.println("Difference at Line/Paragraph " + (i + 1) + ":");
                
                // Call the word-level diff helper
                printWordDiff(s1, s2);
            }
        }

        if (!foundDifference) {
            System.out.println("\nFiles are identical.");
        }
    }
    /**
     * Compares two strings word-by-word to show exact modifications
     */
    private static void printWordDiff(String oldText, String newText) {
        // Check if we are comparing a table row or standard paragraph
        boolean isTable = oldText.startsWith("[TABLE ROW]:") || newText.startsWith("[TABLE ROW]:");
        
        // Use | as a delimiter for tables, or whitespace for standard text
        String delimiter = isTable ? "\\|" : "\\s+";
        
        String[] oldParts = oldText.isEmpty() ? new String[0] : oldText.split(delimiter);
        String[] newParts = newText.isEmpty() ? new String[0] : newText.split(delimiter);

        List<String> oldList = Arrays.asList(oldParts);
        List<String> newList = Arrays.asList(newParts);

        System.out.print(isTable ? "TABLE CELL CHANGES: " : "TEXT CHANGES: ");
        
        // Highlight deletions
        for (String part : oldParts) {
            String cleanPart = part.trim();
            if (!newList.contains(part) && !cleanPart.isEmpty() && !cleanPart.equals("[TABLE ROW]:")) {
                System.out.print("[-" + cleanPart + "-] ");
            }
        }

        // Highlight additions and keep context
        for (String part : newParts) {
            String cleanPart = part.trim();
            if (cleanPart.equals("[TABLE ROW]:")) continue;

            if (!oldList.contains(part) && !cleanPart.isEmpty()) {
                System.out.print("[+" + cleanPart + "+] ");
            } else if (!cleanPart.isEmpty()) {
                System.out.print(cleanPart + " ");
            }
        }
        System.out.println();
    }
    
    
}