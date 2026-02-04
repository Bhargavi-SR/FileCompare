package com.hcl.comparefiles;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

import javax.net.ssl.*;
import java.security.cert.X509Certificate;

/* Apache POI */
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

/* PDFBox */
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class SideBySideDiffDemo {

    /* ======================= MODEL ======================= */

    static class DiffRow {
        int page;
        int paragraph;
        int leftLineNo;
        int rightLineNo;
        String leftText;
        String rightText;
        String diffType; // UNCHANGED / ADDED / REMOVED / MODIFIED
    }

    /* ======================= MAIN ======================= */

    public static void main(String[] args) throws Exception {

        // 1️⃣ IMPORTANT: bypass SSL FIRST
        disableSSLValidation();

        // 2️⃣ Box URLs (use ?download=1)
        String urlA =
            "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=qw82nisgd9saa2dzju0lhm3pyjgp4fqm&file_id=f_2096743062735";
        String urlB =
            "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=xpy1xa5z8mokuc27rqhdfomapkzxn76h&file_id=f_2096746970182";

        File fileA = downloadFile(urlA, "old.docx");
        File fileB = downloadFile(urlB, "new.docx");

        // 3️⃣ Read page → paragraph → lines
        Map<Integer, List<List<String>>> left =
                readByPageAndParagraph(fileA);
        Map<Integer, List<List<String>>> right =
                readByPageAndParagraph(fileB);

        // 4️⃣ Diff
        List<DiffRow> diffs = diff(left, right);

        // 5️⃣ Print result (console demo)
        for (DiffRow d : diffs) {
            System.out.println(
                "PAGE " + d.page +
                " | PARA " + d.paragraph +
                " | L" + d.leftLineNo +
                " | R" + d.rightLineNo +
                " | " + d.diffType
            );
            System.out.println("LEFT : " + d.leftText);
            System.out.println("RIGHT: " + d.rightText);
            System.out.println("------------------------------------");
        }
    }

    /* ======================= SSL ======================= */

    static void disableSSLValidation() throws Exception {

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

    /* ======================= DOWNLOAD ======================= */

    static File downloadFile(String urlStr, String name) throws Exception {

        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");
        conn.setInstanceFollowRedirects(true);

        File f = File.createTempFile("diff_", "_" + name);

        try (InputStream in = conn.getInputStream();
             FileOutputStream out = new FileOutputStream(f)) {

            byte[] buf = new byte[4096];
            int len;
            while ((len = in.read(buf)) != -1)
                out.write(buf, 0, len);
        }
        return f;
    }

    /* ======================= READERS ======================= */

    static Map<Integer, List<List<String>>> readByPageAndParagraph(File file)
            throws Exception {

        String ext = file.getName()
                .substring(file.getName().lastIndexOf('.') + 1)
                .toLowerCase();

        switch (ext) {
            case "pdf":  return readPdf(file);
            case "docx": return readDocx(file);
            case "xlsx": return readXlsx(file);
            case "txt":  return readTxt(file);
            default:
                throw new RuntimeException("Unsupported file type: " + ext);
        }
    }

    /* -------- PDF -------- */

    static Map<Integer, List<List<String>>> readPdf(File file)
            throws Exception {

        if (!isValidPdf(file))
            throw new RuntimeException("Invalid PDF downloaded");

        Map<Integer, List<List<String>>> pages = new LinkedHashMap<>();

        try (PDDocument doc = Loader.loadPDF(file)) {
            PDFTextStripper stripper = new PDFTextStripper();

            for (int p = 1; p <= doc.getNumberOfPages(); p++) {

                stripper.setStartPage(p);
                stripper.setEndPage(p);

                List<List<String>> paragraphs = new ArrayList<>();
                List<String> current = new ArrayList<>();

                for (String line : stripper.getText(doc).split("\\r?\\n")) {

                    if (line.trim().isEmpty()) {
                        if (!current.isEmpty()) {
                            paragraphs.add(new ArrayList<>(current));
                            current.clear();
                        }
                    } else {
                        current.add(line.trim());
                    }
                }
                if (!current.isEmpty())
                    paragraphs.add(current);

                pages.put(p, paragraphs);
            }
        }
        return pages;
    }

    static boolean isValidPdf(File f) throws Exception {
        try (FileInputStream fis = new FileInputStream(f)) {
            byte[] h = new byte[4];
            fis.read(h);
            return "%PDF".equals(new String(h));
        }
    }

    /* -------- DOCX -------- */

    static Map<Integer, List<List<String>>> readDocx(File file)
            throws Exception {

        Map<Integer, List<List<String>>> pages = new LinkedHashMap<>();
        List<List<String>> paras = new ArrayList<>();

        XWPFDocument doc = new XWPFDocument(new FileInputStream(file));

        for (XWPFParagraph p : doc.getParagraphs()) {
            if (!p.getText().trim().isEmpty()) {
                paras.add(
                    Arrays.asList(p.getText().trim().split("\\n"))
                );
            }
        }
        doc.close();
        pages.put(1, paras);
        return pages;
    }

    /* -------- XLSX -------- */

    static Map<Integer, List<List<String>>> readXlsx(File file)
            throws Exception {

        Map<Integer, List<List<String>>> pages = new LinkedHashMap<>();
        Workbook wb = WorkbookFactory.create(new FileInputStream(file));

        int page = 1;
        for (Sheet s : wb) {
            List<List<String>> paras = new ArrayList<>();
            for (Row r : s) {
                List<String> line = new ArrayList<>();
                for (Cell c : r)
                    line.add(c.toString());
                paras.add(line);
            }
            pages.put(page++, paras);
        }
        wb.close();
        return pages;
    }

    /* -------- TXT -------- */

    static Map<Integer, List<List<String>>> readTxt(File file)
            throws Exception {

        Map<Integer, List<List<String>>> pages = new LinkedHashMap<>();
        List<List<String>> paras = new ArrayList<>();

        BufferedReader br = new BufferedReader(new FileReader(file));
        String line;
        while ((line = br.readLine()) != null)
            if (!line.trim().isEmpty())
                paras.add(Collections.singletonList(line.trim()));
        br.close();

        pages.put(1, paras);
        return pages;
    }

    /* ======================= DIFF ENGINE ======================= */

    static List<DiffRow> diff(
            Map<Integer, List<List<String>>> left,
            Map<Integer, List<List<String>>> right) {

        List<DiffRow> rows = new ArrayList<>();
        int maxPage = Math.max(left.size(), right.size());

        for (int p = 1; p <= maxPage; p++) {

            List<List<String>> lp = left.getOrDefault(p, List.of());
            List<List<String>> rp = right.getOrDefault(p, List.of());

            int maxPara = Math.max(lp.size(), rp.size());

            for (int para = 0; para < maxPara; para++) {

                List<String> l = para < lp.size() ? lp.get(para) : List.of();
                List<String> r = para < rp.size() ? rp.get(para) : List.of();

                rows.addAll(diffLines(p, para + 1, l, r));
            }
        }
        return rows;
    }
    static List<DiffRow> diffLines(
            int page, int para,
            List<String> left,
            List<String> right) {

        List<DiffRow> result = new ArrayList<>();

        int i = 0, j = 0;

        while (i < left.size() || j < right.size()) {

            String l = i < left.size() ? left.get(i) : null;
            String r = j < right.size() ? right.get(j) : null;

            DiffRow d = new DiffRow();
            d.page = page;
            d.paragraph = para;

            if (l != null && r != null && l.equals(r)) {
                d.diffType = "UNCHANGED";
                d.leftLineNo = i + 1;
                d.rightLineNo = j + 1;
                d.leftText = l;
                d.rightText = r;
                i++; j++;
            }
            else if (r != null && !left.contains(r)) {
                d.diffType = "ADDED";
                d.leftLineNo = -1;
                d.rightLineNo = j + 1;
                d.leftText = "";
                d.rightText = r;
                j++;
            }
            else if (l != null && !right.contains(l)) {
                d.diffType = "REMOVED";
                d.leftLineNo = i + 1;
                d.rightLineNo = -1;
                d.leftText = l;
                d.rightText = "";
                i++;
            }
            else {
                d.diffType = "MODIFIED";
                d.leftLineNo = i + 1;
                d.rightLineNo = j + 1;
                d.leftText = l != null ? l : "";
                d.rightText = r != null ? r : "";
                i++; j++;
            }

            result.add(d);
        }
        return result;
    }

}
