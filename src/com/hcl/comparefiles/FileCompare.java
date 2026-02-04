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
import com.hcl.voltmx.middleware.dataobject.Param;
import com.hcl.voltmx.middleware.dataobject.Result;
import com.hcl.voltmx.middleware.dataobject.Dataset;
import com.hcl.voltmx.middleware.dataobject.Record;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class FileCompare implements JavaService2 {

    private static final Logger logger = Logger.getLogger(FileCompare.class);
    @Override
    public Object invoke(String methodID, Object[] inputArray,
                         DataControllerRequest request,
                         DataControllerResponse response) throws Exception {
        Result result = new Result();
        Dataset dsResults = new Dataset("comparisonResults");

        try {
            logger.error("Entered FileCompare Java Service");
            logger.error("pdf version----"+org.apache.pdfbox.util.Version.getVersion());
            String urlA = request.getParameter("urlA");
            String urlB = request.getParameter("urlB");
            String fileNameA = request.getParameter("fileNameA");
            String fileNameB = request.getParameter("fileNameB");

            disableSSLValidation();

            File fileA = downloadFile(urlA, fileNameA);
            File fileB = downloadFile(urlB, fileNameB);

            List<String> contentA = readFileByType(fileA);
            List<String> contentB = readFileByType(fileB);
           
            compareFiles(contentA, contentB, dsResults);

            result.addDataset(dsResults);
            result.addParam(new Param("status", "SUCCESS"));
            result.addParam(new Param(
                    "differenceCount",
                    String.valueOf(dsResults.getAllRecords().size())
            ));

        } catch (Exception e) {
            logger.error("Error in FileCompare Java Service", e);
            result.addParam(new Param("status", "FAILED"));
            result.addParam(new Param("errorMessage", e.getMessage()));
        }

        return result;
    }

    /* ================= SSL ================= */

    private void disableSSLValidation() throws Exception {

        TrustManager[] trustAllCerts = new TrustManager[]{
            new X509TrustManager() {
                public X509Certificate[] getAcceptedIssuers() { return null; }
                public void checkClientTrusted(X509Certificate[] c, String a) {}
                public void checkServerTrusted(X509Certificate[] c, String a) {}
            }
        };

        SSLContext sc = SSLContext.getInstance("TLS");
        sc.init(null, trustAllCerts, new java.security.SecureRandom());
        HttpsURLConnection.setDefaultSSLSocketFactory(sc.getSocketFactory());
        HttpsURLConnection.setDefaultHostnameVerifier((h, s) -> true);
    }

    /* ================= DOWNLOAD ================= */

    private File downloadFile(String fileUrl, String fileName) throws Exception {

        URL url = new URL(fileUrl);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();

        conn.setRequestProperty("User-Agent", "Mozilla/5.0");
        conn.setRequestProperty("Accept", "*/*");
        conn.setInstanceFollowRedirects(true);

        File file = File.createTempFile("compare_", "_" + fileName);

        try (InputStream in = conn.getInputStream();
             FileOutputStream out = new FileOutputStream(file)) {

            byte[] buffer = new byte[4096];
            int len;
            while ((len = in.read(buffer)) != -1) {
                out.write(buffer, 0, len);
            }
        }
        return file;
    }

    /* ================= FILE TYPE ================= */

    private String getExtension(String fileName) {
        return fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase();
    }

    private List<String> readFileByType(File file) throws Exception {

        String ext = getExtension(file.getName());

        switch (ext) {
            case "docx": return readDocx(file);
            case "xlsx": return readXlsx(file);
            case "pdf":  return readPdf(file);
            case "txt":  return readTxt(file);
            default:
                throw new RuntimeException("Unsupported file type: " + ext);
        }
    }

    /* ================= READERS ================= */

    private List<String> readDocx(File file) throws Exception {

        List<String> lines = new ArrayList<>();
        XWPFDocument doc = new XWPFDocument(new FileInputStream(file));

        for (XWPFParagraph p : doc.getParagraphs()) {
            if (!p.getText().trim().isEmpty()) {
                lines.add(p.getText().trim());
            }
        }

        for (XWPFTable table : doc.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    lines.add(cell.getText().trim());
                }
            }
        }

        doc.close();
        return lines;
    }

    private List<String> readXlsx(File file) throws Exception {

        List<String> lines = new ArrayList<>();
        Workbook wb = new XSSFWorkbook(new FileInputStream(file));

        for (Sheet sheet : wb) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    lines.add(cell.toString().trim());
                }
            }
        }

        wb.close();
        return lines;
    }

    private List<String> readPdf(File file) throws Exception {
        List<String> lines = new ArrayList<>();
        try (PDDocument doc = Loader.loadPDF(file)) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(doc);
            for (String line : text.split("\\r?\\n")) {
                if (!line.trim().isEmpty()) {
                    lines.add(line.trim());
                }
            }
        }
        return lines;
    }

    private List<String> readTxt(File file) throws Exception {

        List<String> lines = new ArrayList<>();
        BufferedReader br = new BufferedReader(new FileReader(file));

        String line;
        while ((line = br.readLine()) != null) {
            if (!line.trim().isEmpty()) {
                lines.add(line.trim());
            }
        }
        br.close();
        return lines;
    }

    /* ================= COMPARE ================= */

    private void compareFiles(List<String> a,
                              List<String> b,
                              Dataset dsResults) {
        int max = Math.max(a.size(), b.size());
        for (int i = 0; i < max; i++) {
            String lineA = i < a.size() ? a.get(i) : "<NO LINE>";
            String lineB = i < b.size() ? b.get(i) : "<NO LINE>";
            if (!lineA.equals(lineB)) {
                Record record = new Record();
                record.addParam(new Param("line", String.valueOf(i + 1)));
                record.addParam(new Param("fileA", lineA));
                record.addParam(new Param("fileB", lineB));
                /*if ("<NO LINE>".equals(lineA)) {
                    record.addParam(new Param("modifiedText", lineB)); // Added
                } else if ("<NO LINE>".equals(lineB)) {
                    record.addParam(new Param("modifiedText", "<REMOVED>")); // Removed
                } else {
                    record.addParam(new Param("modifiedText", lineB)); // Modified
                }*/
                dsResults.addRecord(record);
            }
        }
    }
}
