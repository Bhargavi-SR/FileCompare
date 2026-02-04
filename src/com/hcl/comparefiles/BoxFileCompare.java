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
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.Loader;


public class BoxFileCompare {
	
	public static String getFileExtension(String fileName) {
	    return fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
	}
	
	public static void disableSSLValidation() throws Exception {

	    TrustManager[] trustAllCerts = new TrustManager[]{
	        new X509TrustManager() {
	            public X509Certificate[] getAcceptedIssuers() {
	                return null;
	            }
	            public void checkClientTrusted(X509Certificate[] certs, String authType) {
	            }
	            public void checkServerTrusted(X509Certificate[] certs, String authType) {
	            }
	        }
	    };

	    SSLContext sc = SSLContext.getInstance("TLS");
	    sc.init(null, trustAllCerts, new java.security.SecureRandom());
	    HttpsURLConnection.setDefaultSSLSocketFactory(sc.getSocketFactory());

	    HttpsURLConnection.setDefaultHostnameVerifier((hostname, session) -> true);
	}
	
	public static List<String> readDocx(File file) throws Exception {

	    List<String> content = new ArrayList<>();

	    FileInputStream fis = new FileInputStream(file);
	    XWPFDocument document = new XWPFDocument(fis);

	    // Read paragraphs
	    for (XWPFParagraph para : document.getParagraphs()) {
	        if (!para.getText().trim().isEmpty()) {
	            content.add(para.getText().trim());
	        }
	    }

	    // Read table content
	    for (XWPFTable table : document.getTables()) {
	        for (XWPFTableRow row : table.getRows()) {
	            for (XWPFTableCell cell : row.getTableCells()) {
	                content.add(cell.getText().trim());
	            }
	        }
	    }

	    document.close();
	    fis.close();

	    return content;
	}
	public static List<String> readXlsx(File file) throws Exception {

	    List<String> content = new ArrayList<>();
	    FileInputStream fis = new FileInputStream(file);
	    Workbook workbook = new XSSFWorkbook(fis);

	    for (Sheet sheet : workbook) {
	        for (Row row : sheet) {
	            for (Cell cell : row) {
	                content.add(cell.toString().trim());
	            }
	        }
	    }

	    workbook.close();
	    fis.close();
	    return content;
	}
	
	public static List<String> readPdf(File file) throws Exception {

	    List<String> lines = new ArrayList<>();

	    try (PDDocument document = Loader.loadPDF(file)) {

	        PDFTextStripper stripper = new PDFTextStripper();
	        String text = stripper.getText(document);

	        for (String line : text.split("\\r?\\n")) {
	            if (!line.trim().isEmpty()) {
	                lines.add(line.trim());
	            }
	        }
	    }

	    return lines;
	}
	
	public static List<String> readTxt(File file) throws Exception {
	    List<String> lines = new ArrayList<>();
	    BufferedReader reader = new BufferedReader(new FileReader(file));

	    String line;
	    while ((line = reader.readLine()) != null) {
	        if (!line.trim().isEmpty()) {
	            lines.add(line.trim());
	        }
	    }
	    reader.close();
	    return lines;
	}


    // STEP 1: Download file from Box URL
	public static File downloadFile(String fileUrl, String fileName) throws Exception {

	    URL url = new URL(fileUrl);
	    HttpURLConnection conn = (HttpURLConnection) url.openConnection();

	    // Pretend to be a browser (VERY IMPORTANT)
	    conn.setRequestProperty("User-Agent", "Mozilla/5.0");
	    conn.setRequestProperty("Accept", "*/*");
	    conn.setRequestProperty("Connection", "keep-alive");
	    conn.setInstanceFollowRedirects(true);

	    InputStream input = conn.getInputStream();
	    FileOutputStream output = new FileOutputStream(fileName);

	    byte[] buffer = new byte[4096];
	    int bytesRead;
	    while ((bytesRead = input.read(buffer)) != -1) {
	        output.write(buffer, 0, bytesRead);
	    }

	    input.close();
	    output.close();

	    return new File(fileName);
	}



    // STEP 2: Read file content line by line
	public static List<String> readFileByType(File file) throws Exception {

	    String ext = getFileExtension(file.getName());

	    switch (ext) {
	        case "docx":
	            return readDocx(file);

	        case "xlsx":
	            return readXlsx(file);

	        case "pdf":
	            return readPdf(file);

	        case "txt":
	            return readTxt(file);

	        default:
	            throw new RuntimeException("Unsupported file type: " + ext);
	    }
	}

    // STEP 3: Compare two files
    public static void compareFiles(List<String> fileA, List<String> fileB) {

        int maxLines = Math.max(fileA.size(), fileB.size());

        for (int i = 0; i < maxLines; i++) {

            String lineA = (i < fileA.size()) ? fileA.get(i) : "<NO LINE>";
            String lineB = (i < fileB.size()) ? fileB.get(i) : "<NO LINE>";

            if (!lineA.equals(lineB)) {
                System.out.println("Difference at line " + (i + 1));
                System.out.println("File A: " + lineA);
                System.out.println("File B: " + lineB);
                System.out.println("-------------------------");
            }
        }
    }

    // STEP 4: Main method (program starts here)
    public static void main(String[] args) throws Exception {
    	disableSSLValidation();
    	
    	/*---- Xslx Files -------*/
        /*String urlA = "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=uqjwk09zz9gj031wav6f3n5gk51m3i17&file_id=f_2095403071019";
        String urlB = "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=73hnszvwlnngjqrjudn7se2px19r8qy6&file_id=f_2095396319754";
        File fileA = downloadFile(urlA, "fileA.xlsx");
        File fileB = downloadFile(urlB, "fileB.xlsx");*/
        /* ----- Docx Files------- */
        /*String urlA = "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=gqjtvdptusfpzj313puiwez1wb6or7im&file_id=f_2095452969360";
        String urlB = "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=hepyav2t6wqvdnj3kchj5i0wwukmtqzw&file_id=f_2095419549116";
        File fileA = downloadFile(urlA, "fileA.docx");
    	File fileB = downloadFile(urlB, "fileB.docx");*/
        
        /*------ PDF Files-------*/
    	String urlA = "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=rwrda3cdhttlw13pmhzelpvm71h15b3e&file_id=f_2095457773355";
    	String urlB = "https://app.box.com/index.php?rm=box_download_shared_file&shared_name=id0e6fr2zxt2iky0q12yp145wv4x60lt&file_id=f_2095431733227";
    	File fileA = downloadFile(urlA, "fileA.pdf");
    	File fileB = downloadFile(urlB, "fileB.pdf");
    	

    	
    	List<String> contentA = readFileByType(fileA);
    	List<String> contentB = readFileByType(fileB);

        compareFiles(contentA, contentB);
    }
}
