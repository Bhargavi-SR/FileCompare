package com.hcl.comparefiles;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

import javax.net.ssl.*;
import java.security.cert.X509Certificate;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

/**
 * Unified page -> paragraph diff for
 * PDF / DOCX / XLSX / TXT
 */
public class ParagraphDiffData {

    /* ========================= MAIN ========================= */

    public static void main(String[] args) throws Exception {
        disableSSLValidation();
        System.out.print(org.apache.pdfbox.util.Version.getVersion());
        String urlA = "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=lhxglimh513pigye1vu994qfv5mi1w2e&file_id=f_2096927530914";
        String urlB = "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=mwtmfxccb4o99tm2w5px25htspi0vaf6&file_id=f_2096924194868";

        File oldFile = downloadFile(urlA, "oldFile.pdf");
        File newFile = downloadFile(urlB, "newFile.pdf");

        Map<Integer, List<String>> oldPages = readByPage(oldFile);
        Map<Integer, List<String>> newPages = readByPage(newFile);

        List<PageParagraphDiff> diffs = comparePages(oldPages, newPages);
        printSideBySide(diffs);
    }

    /* ===================== DATA MODELS ===================== */

    static class DiffSegment {
        String text;
        String type; // UNCHANGED / ADDED / REMOVED

        DiffSegment(String text, String type) {
            this.text = text;
            this.type = type;
        }
    }

    static class PageParagraphDiff {
        int pageNo;
        int paragraphNo;
        List<DiffSegment> left;
        List<DiffSegment> right;

        PageParagraphDiff(int pageNo, int paragraphNo,
                          List<DiffSegment> left,
                          List<DiffSegment> right) {
            this.pageNo = pageNo;
            this.paragraphNo = paragraphNo;
            this.left = left;
            this.right = right;
        }
    }

    /* ====================== SSL BYPASS ====================== */

    static void disableSSLValidation() throws Exception {

        TrustManager[] trustAll = new TrustManager[]{
            new X509TrustManager() {
                public X509Certificate[] getAcceptedIssuers() { return new X509Certificate[0]; }
                public void checkClientTrusted(X509Certificate[] c, String a) {}
                public void checkServerTrusted(X509Certificate[] c, String a) {}
            }
        };

        SSLContext sc = SSLContext.getInstance("TLS");
        sc.init(null, trustAll, new java.security.SecureRandom());
        HttpsURLConnection.setDefaultSSLSocketFactory(sc.getSocketFactory());
        HttpsURLConnection.setDefaultHostnameVerifier((h, s) -> true);
    }

    /* ====================== DOWNLOAD ====================== */

    static File downloadFile(String urlStr, String name) throws Exception {

        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");

        File file = File.createTempFile("diff_", "_" + name);

        try (InputStream in = conn.getInputStream();
             FileOutputStream out = new FileOutputStream(file)) {

            byte[] buf = new byte[4096];
            int len;
            while ((len = in.read(buf)) != -1)
                out.write(buf, 0, len);
        }
        return file;
    }

    /* ====================== DISPATCHER ====================== */

    static Map<Integer, List<String>> readByPage(File file) throws Exception {
        String ext = file.getName()
                .substring(file.getName().lastIndexOf('.') + 1)
                .toLowerCase();
        switch (ext) {
            case "pdf":
                return readPdf(file);
            case "docx":
            case "doc": 
                return readDocx(file);
            case "xlsx":
            case "xls": 
                return readXlsx(file);
            case "txt":
                return readTxt(file);
            default:
                throw new RuntimeException(
                    "Unsupported file type: " + ext);
        }
    }

    /* ====================== READERS ====================== */

    static Map<Integer, List<String>> readPdf(File file) throws Exception {

        Map<Integer, List<String>> pages = new LinkedHashMap<>();

        try (PDDocument doc = Loader.loadPDF(file)) {
            PDFTextStripper stripper = new PDFTextStripper();
            int total = doc.getNumberOfPages();

            for (int p = 1; p <= total; p++) {
                stripper.setStartPage(p);
                stripper.setEndPage(p);

                String text = stripper.getText(doc);
                List<String> paras = splitParagraphs(text);
                pages.put(p, paras);
            }
        }
        return pages;
    }

    static Map<Integer, List<String>> readDocx(File file) throws Exception {

        Map<Integer, List<String>> pages = new LinkedHashMap<>();
        List<String> paras = new ArrayList<>();

        XWPFDocument doc = new XWPFDocument(new FileInputStream(file));
        for (XWPFParagraph p : doc.getParagraphs()) {
            if (!p.getText().trim().isEmpty())
                paras.add(p.getText().trim());
        }
        doc.close();

        pages.put(1, paras); // pseudo page
        return pages;
    }

    static Map<Integer, List<String>> readXlsx(File file) throws Exception {

        Map<Integer, List<String>> pages = new LinkedHashMap<>();
        Workbook wb = new XSSFWorkbook(new FileInputStream(file));

        int page = 1;
        for (Sheet sheet : wb) {
            List<String> rows = new ArrayList<>();
            for (Row r : sheet) {
                StringBuilder sb = new StringBuilder();
                for (Cell c : r)
                    sb.append(c.toString()).append(" ");
                if (sb.length() > 0)
                    rows.add(sb.toString().trim());
            }
            pages.put(page++, rows);
        }
        wb.close();
        return pages;
    }

    static Map<Integer, List<String>> readTxt(File file) throws Exception {

        Map<Integer, List<String>> pages = new LinkedHashMap<>();
        List<String> paras = new ArrayList<>();

        BufferedReader br = new BufferedReader(new FileReader(file));
        String line;
        StringBuilder sb = new StringBuilder();

        while ((line = br.readLine()) != null) {
            if (line.trim().isEmpty()) {
                if (sb.length() > 0) {
                    paras.add(sb.toString().trim());
                    sb.setLength(0);
                }
            } else {
                sb.append(line).append(" ");
            }
        }
        if (sb.length() > 0)
            paras.add(sb.toString().trim());

        br.close();
        pages.put(1, paras);
        return pages;
    }

    static List<String> splitParagraphs(String text) {

        List<String> list = new ArrayList<>();
        for (String p : text.split("\\n\\s*\\n")) {
            String clean = p.replaceAll("\\s+", " ").trim();
            if (!clean.isEmpty())
                list.add(clean);
        }
        return list;
    }

    /* ====================== COMPARE ====================== */

    static List<PageParagraphDiff> comparePages(
            Map<Integer, List<String>> oldPages,
            Map<Integer, List<String>> newPages) {

        List<PageParagraphDiff> diffs = new ArrayList<>();
        int maxPages = Math.max(oldPages.size(), newPages.size());

        for (int p = 1; p <= maxPages; p++) {

            List<String> oldParas =
                    oldPages.getOrDefault(p, Collections.emptyList());
            List<String> newParas =
                    newPages.getOrDefault(p, Collections.emptyList());

            int maxParas = Math.max(oldParas.size(), newParas.size());

            for (int i = 0; i < maxParas; i++) {

                String oldText = i < oldParas.size() ? oldParas.get(i) : "";
                String newText = i < newParas.size() ? newParas.get(i) : "";

                if (!oldText.equals(newText)) {

                    diffs.add(new PageParagraphDiff(
                            p,
                            i + 1,
                            diffWords(oldText, newText, true),
                            diffWords(oldText, newText, false)
                    ));
                }
            }
        }
        return diffs;
    }

    /* ====================== WORD DIFF ====================== */

    static List<DiffSegment> diffWords(
            String oldText,
            String newText,
            boolean left) {

        String[] a = oldText.split("\\s+");
        String[] b = newText.split("\\s+");

        int[][] lcs = new int[a.length + 1][b.length + 1];

        for (int i = a.length - 1; i >= 0; i--)
            for (int j = b.length - 1; j >= 0; j--)
                lcs[i][j] = a[i].equals(b[j])
                        ? lcs[i + 1][j + 1] + 1
                        : Math.max(lcs[i + 1][j], lcs[i][j + 1]);

        List<DiffSegment> res = new ArrayList<>();
        int i = 0, j = 0;

        while (i < a.length && j < b.length) {
            if (a[i].equals(b[j])) {
                res.add(new DiffSegment(a[i] + " ", "UNCHANGED"));
                i++; j++;
            } else if (lcs[i + 1][j] >= lcs[i][j + 1]) {
                if (left)
                    res.add(new DiffSegment(a[i] + " ", "REMOVED"));
                i++;
            } else {
                if (!left)
                    res.add(new DiffSegment(b[j] + " ", "ADDED"));
                j++;
            }
        }

        while (i < a.length) {
            if (left)
                res.add(new DiffSegment(a[i] + " ", "REMOVED"));
            i++;
        }

        while (j < b.length) {
            if (!left)
                res.add(new DiffSegment(b[j] + " ", "ADDED"));
            j++;
        }

        return res;
    }

    /* ====================== OUTPUT ====================== */

    static void printSideBySide(List<PageParagraphDiff> diffs) {

        for (PageParagraphDiff d : diffs) {

            System.out.println("\nPage " + d.pageNo +
                               " | Paragraph " + d.paragraphNo);

            System.out.print("OLD: ");
            for (DiffSegment s : d.left)
                System.out.print("[" + s.type + "]" + s.text);

            System.out.print("\nNEW: ");
            for (DiffSegment s : d.right)
                System.out.print("[" + s.type + "]" + s.text);

            System.out.println();
        }
    }
}
