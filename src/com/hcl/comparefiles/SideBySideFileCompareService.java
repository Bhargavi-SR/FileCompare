package com.hcl.comparefiles;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

import javax.net.ssl.*;
import java.security.cert.X509Certificate;

import org.apache.log4j.Logger;

/* Volt MX */
import com.hcl.voltmx.middleware.common.JavaService2;
import com.hcl.voltmx.middleware.controller.DataControllerRequest;
import com.hcl.voltmx.middleware.controller.DataControllerResponse;
import com.hcl.voltmx.middleware.dataobject.Dataset;
import com.hcl.voltmx.middleware.dataobject.Param;
import com.hcl.voltmx.middleware.dataobject.Record;
import com.hcl.voltmx.middleware.dataobject.Result;

/* Apache POI */
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

/* PDFBox */
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class SideBySideFileCompareService implements JavaService2 {

    private static final Logger logger =
            Logger.getLogger(SideBySideFileCompareService.class);

    /* ================= ENTRY ================= */

    @Override
    public Object invoke(String methodID, Object[] inputArray,
                         DataControllerRequest request,
                         DataControllerResponse response) {

        Result result = new Result();
        Dataset ds = new Dataset("diffResults");

        try {
            logger.info("Entered SideBySideFileCompareService");

            String urlA = request.getParameter("urlA");
            String urlB = request.getParameter("urlB");
            String fileNameA = request.getParameter("fileNameA");
            String fileNameB = request.getParameter("fileNameB");

            disableSSL();

            File leftFile = downloadFile(urlA, fileNameA);
            File rightFile = downloadFile(urlB, fileNameB);

            Map<Integer, List<List<String>>> left =
                    readByPageAndParagraph(leftFile);
            Map<Integer, List<List<String>>> right =
                    readByPageAndParagraph(rightFile);

            buildDiff(left, right, ds);

            result.addDataset(ds);
            result.addParam(new Param("status", "SUCCESS"));
            result.addParam(new Param("differenceCount",
                    String.valueOf(ds.getAllRecords().size())));

        } catch (Exception e) {
            logger.error("Error in SideBySideFileCompareService", e);
            result.addParam(new Param("status", "FAILED"));
            result.addParam(new Param("errorMessage", e.getMessage()));
        }
        return result;
    }

    /* ================= SSL ================= */

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

    /* ================= DOWNLOAD ================= */

    private File downloadFile(String urlStr, String name) throws Exception {

        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");
        conn.setInstanceFollowRedirects(true);

        File f = File.createTempFile("cmp_", "_" + name);

        try (InputStream in = conn.getInputStream();
             FileOutputStream out = new FileOutputStream(f)) {

            byte[] buf = new byte[4096];
            int len;
            while ((len = in.read(buf)) != -1)
                out.write(buf, 0, len);
        }
        return f;
    }

    /* ================= FILE DISPATCH ================= */

    private Map<Integer, List<List<String>>> readByPageAndParagraph(File file)
            throws Exception {

        String ext = file.getName()
                .substring(file.getName().lastIndexOf('.') + 1)
                .toLowerCase();

        switch (ext) {
            case "pdf":  return readPdf(file);
            case "doc":
            case "docx": return readDocx(file);
            case "xls":
            case "xlsx": return readXlsx(file);
            case "txt":  return readTxt(file);
            default:
                throw new RuntimeException("Unsupported file type: " + ext);
        }
    }

    /* ================= READERS ================= */

    private Map<Integer, List<List<String>>> readPdf(File file)
            throws Exception {

        if (!isValidPdf(file))
            throw new RuntimeException("Invalid PDF downloaded");

        Map<Integer, List<List<String>>> pages = new LinkedHashMap<>();

        try (PDDocument doc = Loader.loadPDF(file)) {
            PDFTextStripper stripper = new PDFTextStripper();

            for (int p = 1; p <= doc.getNumberOfPages(); p++) {

                stripper.setStartPage(p);
                stripper.setEndPage(p);

                List<List<String>> paras = new ArrayList<>();
                List<String> current = new ArrayList<>();

                for (String line : stripper.getText(doc).split("\\r?\\n")) {
                    if (line.trim().isEmpty()) {
                        if (!current.isEmpty()) {
                            paras.add(new ArrayList<>(current));
                            current.clear();
                        }
                    } else {
                        current.add(line.trim());
                    }
                }
                if (!current.isEmpty())
                    paras.add(current);

                pages.put(p, paras);
            }
        }
        return pages;
    }

    private boolean isValidPdf(File f) throws Exception {
        try (FileInputStream fis = new FileInputStream(f)) {
            byte[] h = new byte[4];
            fis.read(h);
            return "%PDF".equals(new String(h));
        }
    }

    private Map<Integer, List<List<String>>> readDocx(File file)
            throws Exception {

        Map<Integer, List<List<String>>> pages = new LinkedHashMap<>();
        List<List<String>> paras = new ArrayList<>();

        XWPFDocument doc = new XWPFDocument(new FileInputStream(file));
        for (XWPFParagraph p : doc.getParagraphs()) {
            if (!p.getText().trim().isEmpty())
                paras.add(Collections.singletonList(p.getText().trim()));
        }
        doc.close();
        pages.put(1, paras);
        return pages;
    }

    private Map<Integer, List<List<String>>> readXlsx(File file)
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

    private Map<Integer, List<List<String>>> readTxt(File file)
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

    /* ================= DIFF ENGINE ================= */

    private void buildDiff(
            Map<Integer, List<List<String>>> left,
            Map<Integer, List<List<String>>> right,
            Dataset ds) {

        int maxPage = Math.max(left.size(), right.size());

        for (int p = 1; p <= maxPage; p++) {

            List<List<String>> lp =
                    left.getOrDefault(p, Collections.emptyList());
            List<List<String>> rp =
                    right.getOrDefault(p, Collections.emptyList());

            int li = 0, ri = 0;
            int paraNo = 1;

            while (li < lp.size() || ri < rp.size()) {

                List<String> lpara =
                        li < lp.size() ? lp.get(li) : null;
                List<String> rpara =
                        ri < rp.size() ? rp.get(ri) : null;

                // Both paragraphs exist
                if (lpara != null && rpara != null) {

                    if (lpara.equals(rpara)) {
                        addParagraph(ds, p, paraNo, lpara, rpara, "UNCHANGED");
                        li++; ri++;
                    } else {
                        addParagraph(ds, p, paraNo, lpara, rpara, "MODIFIED");
                        li++; ri++;
                    }

                }
                // Paragraph REMOVED
                else if (lpara != null) {
                    addParagraph(ds, p, paraNo, lpara,
                            Collections.emptyList(), "REMOVED");
                    li++;
                }
                // Paragraph ADDED
                else {
                    addParagraph(ds, p, paraNo,
                            Collections.emptyList(), rpara, "ADDED");
                    ri++;
                }

                paraNo++;
            }
        }
    }

    private void addParagraph(
            Dataset ds,
            int page,
            int paraNo,
            List<String> leftLines,
            List<String> rightLines,
            String diffType) {

        int maxLine = Math.max(leftLines.size(), rightLines.size());

        for (int i = 0; i < maxLine; i++) {

            String l = i < leftLines.size() ? leftLines.get(i) : "";
            String r = i < rightLines.size() ? rightLines.get(i) : "";

            Record rec = new Record();
            rec.addParam(new Param("page", String.valueOf(page)));
            rec.addParam(new Param("paragraph", String.valueOf(paraNo)));
            rec.addParam(new Param("leftLineNo",
                    l.isEmpty() ? "" : String.valueOf(i + 1)));
            rec.addParam(new Param("rightLineNo",
                    r.isEmpty() ? "" : String.valueOf(i + 1)));
            rec.addParam(new Param("leftText", l));
            rec.addParam(new Param("rightText", r));
            rec.addParam(new Param("diffType", diffType));

            ds.addRecord(rec);
        }
    }

}
