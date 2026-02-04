package com.hcl.comparefiles;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

import javax.net.ssl.*;
import java.security.cert.X509Certificate;

import org.apache.log4j.Logger;

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

public class PageParaCompare implements JavaService2 {

    private static final Logger logger =
            Logger.getLogger(PageParaCompare.class);

    /* ================= ENTRY ================= */

    @Override
    public Object invoke(String methodID, Object[] inputArray,
                         DataControllerRequest request,
                         DataControllerResponse response) {

        Result result = new Result();
        Dataset ds = new Dataset("diffResults");

        try {
            logger.info("Entered FileComparePageParagraph service");

            String urlA = request.getParameter("urlA");
            String urlB = request.getParameter("urlB");
            String fileNameA = request.getParameter("fileNameA");
            String fileNameB = request.getParameter("fileNameB");

            disableSSL();

            File fileA = downloadFile(urlA, fileNameA);
            File fileB = downloadFile(urlB, fileNameB);

            Map<Integer, List<String>> pagesA = readByPage(fileA);
            Map<Integer, List<String>> pagesB = readByPage(fileB);

            buildDiff(pagesA, pagesB, ds);

            result.addDataset(ds);
            result.addParam(new Param("status", "SUCCESS"));
            result.addParam(new Param("totalDifferences",
                    String.valueOf(ds.getAllRecords().size())));

        } catch (Exception e) {
            logger.error("Error in FileComparePageParagraph", e);
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

    private File downloadFile(String urlStr, String fileName) throws Exception {

        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");
        conn.setInstanceFollowRedirects(true);

        File file = File.createTempFile("cmp_", "_" + fileName);

        try (InputStream in = conn.getInputStream();
             FileOutputStream out = new FileOutputStream(file)) {

            byte[] buf = new byte[4096];
            int len;
            while ((len = in.read(buf)) != -1)
                out.write(buf, 0, len);
        }
        return file;
    }

    /* ================= FILE ROUTER ================= */

    private Map<Integer, List<String>> readByPage(File file) throws Exception {

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

    private Map<Integer, List<String>> readPdf(File file) throws Exception {

        Map<Integer, List<String>> pages = new LinkedHashMap<>();

        try (PDDocument doc = Loader.loadPDF(file)) {
            PDFTextStripper stripper = new PDFTextStripper();

            for (int i = 1; i <= doc.getNumberOfPages(); i++) {
                stripper.setStartPage(i);
                stripper.setEndPage(i);

                List<String> paras = new ArrayList<>();
                for (String line : stripper.getText(doc).split("\\r?\\n"))
                    if (!line.trim().isEmpty())
                        paras.add(line.trim());

                pages.put(i, paras);
            }
        }
        return pages;
    }

    private Map<Integer, List<String>> readDocx(File file) throws Exception {

        Map<Integer, List<String>> pages = new LinkedHashMap<>();
        List<String> paras = new ArrayList<>();

        XWPFDocument doc = new XWPFDocument(new FileInputStream(file));

        for (XWPFParagraph p : doc.getParagraphs())
            if (!p.getText().trim().isEmpty())
                paras.add(p.getText().trim());

        pages.put(1, paras); // DOC has logical page
        doc.close();
        return pages;
    }

    private Map<Integer, List<String>> readXlsx(File file) throws Exception {

        Map<Integer, List<String>> pages = new LinkedHashMap<>();
        Workbook wb = WorkbookFactory.create(new FileInputStream(file));

        int page = 1;
        for (Sheet s : wb) {
            List<String> rows = new ArrayList<>();
            for (Row r : s) {
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

    private Map<Integer, List<String>> readTxt(File file) throws Exception {

        Map<Integer, List<String>> pages = new LinkedHashMap<>();
        List<String> lines = new ArrayList<>();

        BufferedReader br = new BufferedReader(new FileReader(file));
        String line;
        while ((line = br.readLine()) != null)
            if (!line.trim().isEmpty())
                lines.add(line.trim());
        br.close();

        pages.put(1, lines);
        return pages;
    }

    /* ================= DIFF LOGIC ================= */

    private void buildDiff(Map<Integer, List<String>> a,
                           Map<Integer, List<String>> b,
                           Dataset ds) {

        int maxPage = Math.max(a.size(), b.size());

        for (int p = 1; p <= maxPage; p++) {

            List<String> pa = a.getOrDefault(p, new ArrayList<>());
            List<String> pb = b.getOrDefault(p, new ArrayList<>());

            int maxPara = Math.max(pa.size(), pb.size());

            for (int i = 0; i < maxPara; i++) {

                String left = i < pa.size() ? pa.get(i) : "";
                String right = i < pb.size() ? pb.get(i) : "";

                String status;
                if (left.equals(right)) status = "UNCHANGED";
                else if (left.isEmpty()) status = "ADDED";
                else if (right.isEmpty()) status = "REMOVED";
                else status = "MODIFIED";

                Record r = new Record();
                r.addParam(new Param("page", String.valueOf(p)));
                r.addParam(new Param("paragraph", String.valueOf(i + 1)));
                r.addParam(new Param("leftText", left));
                r.addParam(new Param("rightText", right));
                r.addParam(new Param("diffType", status));

                ds.addRecord(r);
            }
        }
    }
}
