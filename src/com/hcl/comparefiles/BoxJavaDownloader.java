package com.hcl.comparefiles;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import javax.net.ssl.*;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.security.cert.X509Certificate;
import java.util.*;

public class BoxJavaDownloader {

    // IMPORTANT: Generate this in Box Developer Console. It expires every 60 mins.
    private static final String DEVELOPER_TOKEN = "j9H8eaPcq04GHyoWGEmfW953TLMrS53g";

    static {
        try {
            TrustManager[] trustAllCerts = new TrustManager[]{
                new X509TrustManager() {
                    public X509Certificate[] getAcceptedIssuers() { return null; }
                    public void checkClientTrusted(X509Certificate[] certs, String authType) {}
                    public void checkServerTrusted(X509Certificate[] certs, String authType) {}
                }
            };
            SSLContext sc = SSLContext.getInstance("SSL");
            sc.init(null, trustAllCerts, new java.security.SecureRandom());
            HttpsURLConnection.setDefaultSSLSocketFactory(sc.getSocketFactory());
            HttpsURLConnection.setDefaultHostnameVerifier((hostname, session) -> true);
        } catch (Exception e) { e.printStackTrace(); }
    }

    public static void main(String[] args) {
        String urlA = "https://app.box.com/s/pgu5np17dw0x56c6wkvpfm3kztulwpqa";
        String urlB = "https://app.box.com/s/16dzf51mjlkmeuz1s14xle3bpdifzpg0";

        try {
            BoxJavaDownloader downloader = new BoxJavaDownloader();
            List<Map<String, String>> results = downloader.processBoxFiles(urlA, urlB);
            results.forEach(diff -> System.out.println("Line " + diff.get("index") + ": " + diff.get("diff")));
        } catch (Exception e) { e.printStackTrace(); }
    }

    public List<Map<String, String>> processBoxFiles(String linkA, String linkB) throws Exception {
        return performComparison(extractText(downloadFile(linkA)), extractText(downloadFile(linkB)));
    }

    private byte[] downloadFile(String sharedLink) throws Exception {
        // STEP 1: Get the File ID from the Shared Link
        URL resolveUrl = new URL("https://api.box.com/2.0/shared_items");
        HttpURLConnection resolveConn = (HttpURLConnection) resolveUrl.openConnection();
        resolveConn.setRequestMethod("GET");
        resolveConn.setRequestProperty("Authorization", "Bearer " + DEVELOPER_TOKEN);
        resolveConn.setRequestProperty("BoxApi", "shared_link=" + sharedLink);

        String fileId = "";
        if (resolveConn.getResponseCode() == 200) {
            try (Scanner s = new Scanner(resolveConn.getInputStream())) {
                String response = s.useDelimiter("\\A").next();
                // Simple parsing to find "id":"XXXXX"
                int idIndex = response.indexOf("\"id\":\"") + 6;
                fileId = response.substring(idIndex, response.indexOf("\"", idIndex));
            }
        } else {
            throw new Exception("Step 1 Failed: Could not resolve shared link. Code: " + resolveConn.getResponseCode());
        }

        // STEP 2: Use that File ID to get the actual content
        // We MUST use the /files/{id}/content endpoint
        URL contentUrl = new URL("https://api.box.com/2.0/files/" + fileId + "/content");
        HttpURLConnection contentConn = (HttpURLConnection) contentUrl.openConnection();
        contentConn.setRequestMethod("GET");
        contentConn.setRequestProperty("Authorization", "Bearer " + DEVELOPER_TOKEN);
        // Even here, you must pass the BoxApi header to prove access via the link
        contentConn.setRequestProperty("BoxApi", "shared_link=" + sharedLink);
        contentConn.setInstanceFollowRedirects(true);

        int status = contentConn.getResponseCode();
        if (status == 200) {
            try (InputStream in = new BufferedInputStream(contentConn.getInputStream());
                 ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                byte[] buffer = new byte[8192];
                int n;
                while ((n = in.read(buffer)) != -1) out.write(buffer, 0, n);
                return out.toByteArray();
            }
        } else {
            throw new Exception("Step 2 Failed: Download error. Code: " + status);
        }
    }

    private List<String> extractText(byte[] data) throws Exception {
        List<String> lines = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(data);
             XWPFDocument doc = new XWPFDocument(is)) {
            for (XWPFParagraph p : doc.getParagraphs()) lines.add(p.getText());
        }
        return lines;
    }

    private List<Map<String, String>> performComparison(List<String> list1, List<String> list2) {
        List<Map<String, String>> diffs = new ArrayList<>();
        int max = Math.max(list1.size(), list2.size());
        for (int i = 0; i < max; i++) {
            String s1 = i < list1.size() ? list1.get(i).trim() : "[End]";
            String s2 = i < list2.size() ? list2.get(i).trim() : "[End]";
            if (!s1.equals(s2)) {
                Map<String, String> map = new HashMap<>();
                map.put("index", String.valueOf(i + 1));
                map.put("diff", "A: " + s1 + " | B: " + s2);
                diffs.add(map);
            }
        }
        return diffs;
    }
}