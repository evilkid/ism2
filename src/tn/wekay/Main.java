package tn.wekay;

import org.apache.commons.cli.*;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import javax.net.ssl.*;
import java.io.*;
import java.security.GeneralSecurityException;
import java.security.SecureRandom;
import java.security.cert.X509Certificate;
import java.util.HashSet;
import java.util.Set;
import java.util.concurrent.CopyOnWriteArraySet;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

/**
 * @author Ouerghi Yassine
 */
public class Main {

    public static String EXCEL_LOCATION = "D:/items.xls";

    public static int MAX_PRODUCT_PER_FILE = 5000;

    private static ExecutorService eventExecutor = Executors.newFixedThreadPool(2);

    public static Set<Event> allEvents = new CopyOnWriteArraySet<>();

    static int documentCount = 1;

    public final static String BASE_URL = "https://www.ism-cologne.com";

    static {
        TrustManager[] trustAllCertificates = {new X509TrustManager() {
            public X509Certificate[] getAcceptedIssuers() {
                return null;
            }


            public void checkClientTrusted(X509Certificate[] certs, String authType) {
            }


            public void checkServerTrusted(X509Certificate[] certs, String authType) {
            }
        }};
        HostnameVerifier trustAllHostnames = (hostname, session) -> true;


        try {
            SSLContext sc = SSLContext.getInstance("SSL");
            sc.init(null, trustAllCertificates, new SecureRandom());
            HttpsURLConnection.setDefaultSSLSocketFactory(sc.getSocketFactory());
            HttpsURLConnection.setDefaultHostnameVerifier(trustAllHostnames);
        } catch (GeneralSecurityException e) {
            throw new ExceptionInInitializerError(e);
        }
    }

    public static void main(String[] args) throws IOException, InterruptedException {
        CommandLine cmd;
        Options options = new Options();

        Option output = new Option("o", "output", true, "Output file");
        output.setRequired(false);
        options.addOption(output);

        Option threads = new Option("t", "threads", true, "Number of threads");
        threads.setRequired(false);
        options.addOption(threads);

        Option productPerFile = new Option("m", "max", true, "Maximum product per file");
        productPerFile.setRequired(false);
        options.addOption(productPerFile);

        Option countryOption = new Option("c", "country", true, "Country");
        countryOption.setRequired(false);
        options.addOption(countryOption);

        DefaultParser defaultParser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();

        try {
            cmd = defaultParser.parse(options, args);
        } catch (ParseException e) {
            System.out.println(e.getMessage());
            formatter.printHelp("utility-name", options);

            System.exit(1);

            return;
        }
        EXCEL_LOCATION = cmd.getOptionValue("output", EXCEL_LOCATION);
        MAX_PRODUCT_PER_FILE = Integer.parseInt(cmd.getOptionValue("max", String.valueOf(MAX_PRODUCT_PER_FILE)));


        String url = BASE_URL + "/exhibitors-and-products/exhibitor-index/exhibitor-index-9.php?start=%d";


        String country = cmd.getOptionValue("country");
        if (country != null) {
            country = country.toUpperCase();
            url += "&" + "paginatevalues={\"country2\":\"" + country + "\",\"origcountry2\":\"" + country + "\"}";
        }


        Document document = Jsoup.connect(String.format(url, 0))
                .userAgent("Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36")
                .header("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
                .timeout(100000)
                .method(Connection.Method.GET)
                .get();

        Elements select = document.select(".notmobile").first().select("a");

        int maxpages = 0;
        if (select.size() > 0) {
            String maxcount = StringUtils.substringBetween((select.get(select.size() - 2)).attr("href"), "start=", "&");

            maxpages = 1 + Integer.parseInt(maxcount) / 20;

        }

        eventExecutor.submit(new EventGrabberThread(getEventList(String.format(url, 0)), 1));

        for (int i = 1; i < maxpages; i++) {
            eventExecutor.submit(new EventGrabberThread(getEventList(String.format(url, i * 20)), i + 1));
        }

        eventExecutor.shutdown();

        while (!eventExecutor.isTerminated()) ;

        saveListToExcel(allEvents);
    }

    public static Set<Event> getEventList(String url) throws IOException {
        Set<Event> events = new HashSet<>();

        Document document = Jsoup.connect(url)
                .userAgent("Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36")
                .header("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
                .timeout(100000)
                .method(Connection.Method.GET)
                .get();

        int pageAttemps = 15;

        while (pageAttemps > 0 && document.select("#ausform > .search-results").isEmpty()) {
            try {
                Thread.sleep(3000L);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            pageAttemps--;

            document = Jsoup.connect(url)
                    .userAgent("Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36")
                    .header("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
                    .timeout(100000)
                    .method(Connection.Method.GET)
                    .get();
        }

        if (pageAttemps <= 0) {
            System.out.println("Error getting page list (15 attemps exhausted): " + url);
            return events;
        }

        Element table = document.select("#ausform > .search-results").first();

        for (Element item : table.select(".item")) {

            Event event = new Event();

            event.company = item.select(".col1ergebnis > a").attr("title").trim();
            event.url = BASE_URL + item.select(".col1ergebnis > a").attr("href");
            event.country = item.select(".col1ergebnis > p").text().trim();
            event.hall = item.select(".col3ergebnis > p").first().text().split("\\|")[0].replaceAll("Hall", "").trim();
            event.stand = item.select(".col3ergebnis > p").first().text().split("\\|")[1].replaceAll("Stand", "").trim();

            events.add(event);

        }

        return events;
    }

    public static void saveListToExcel(Set<Event> events) {
        try {
            System.out.println("writing to excel...");
            EXCEL_LOCATION = FilenameUtils.removeExtension(EXCEL_LOCATION);

            HSSFWorkbook workbook;

            if ((new File(EXCEL_LOCATION + documentCount + ".xls")).exists()) {
                try (InputStream inp = new FileInputStream(EXCEL_LOCATION + documentCount + ".xls")) {
                    workbook = new HSSFWorkbook(inp);
                }
            } else {
                workbook = new HSSFWorkbook();
            }


            HSSFSheet spreadsheet = workbook.getSheet("ism");

            if (spreadsheet == null) {
                spreadsheet = workbook.createSheet("ism");
            }

            if (spreadsheet.getPhysicalNumberOfRows() == 0) {
                writeHeaders(spreadsheet);
            }


            HSSFCellStyle hSSFCellStyle = workbook.createCellStyle();
            hSSFCellStyle.setAlignment(HorizontalAlignment.LEFT);
            hSSFCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            spreadsheet.setColumnWidth(0, 5000);
            spreadsheet.setColumnWidth(1, 5000);
            spreadsheet.setColumnWidth(2, 3000);
            spreadsheet.setColumnWidth(3, 3000);
            spreadsheet.setColumnWidth(4, 15000);
            spreadsheet.setColumnWidth(5, 4000);
            spreadsheet.setColumnWidth(6, 6000);
            spreadsheet.setColumnWidth(7, 10000);

            int rowId = spreadsheet.getPhysicalNumberOfRows();
            for (Event event : events) {
                HSSFRow row = spreadsheet.createRow(rowId++);

                int cellId = 0;

                HSSFCell hSSFCell = row.createCell(cellId++);
                hSSFCell.setCellValue(event.company);
                hSSFCell.setCellStyle((CellStyle) hSSFCellStyle);

                hSSFCell = row.createCell(cellId++);
                hSSFCell.setCellValue(event.country);
                hSSFCell.setCellStyle((CellStyle) hSSFCellStyle);

                hSSFCell = row.createCell(cellId++);
                hSSFCell.setCellValue(event.hall);
                hSSFCell.setCellStyle((CellStyle) hSSFCellStyle);

                hSSFCell = row.createCell(cellId++);
                hSSFCell.setCellValue(event.stand);
                hSSFCell.setCellStyle((CellStyle) hSSFCellStyle);

                hSSFCell = row.createCell(cellId++);
                hSSFCell.setCellValue(event.address);
                hSSFCell.setCellStyle((CellStyle) hSSFCellStyle);

                hSSFCell = row.createCell(cellId++);
                hSSFCell.setCellValue(event.telephone);
                hSSFCell.setCellStyle((CellStyle) hSSFCellStyle);

                hSSFCell = row.createCell(cellId++);
                hSSFCell.setCellValue(event.email);
                hSSFCell.setCellStyle((CellStyle) hSSFCellStyle);


                hSSFCell = row.createCell(cellId);
                hSSFCell.setCellValue(event.website);
                hSSFCell.setCellStyle((CellStyle) hSSFCellStyle);

                event = null;

                if (rowId % 1000 == 0) {
                    System.out.println("Writing progress " + rowId + "/" + events.size());
                }

                if (rowId % MAX_PRODUCT_PER_FILE == 0) {
                    try (FileOutputStream out = new FileOutputStream(new File(EXCEL_LOCATION + documentCount++ + ".xls"))) {

                        workbook.write(out);
                    }

                    workbook = new HSSFWorkbook();

                    spreadsheet = workbook.createSheet("ism");
                    writeHeaders(spreadsheet);

                    rowId = spreadsheet.getPhysicalNumberOfRows();

                    hSSFCellStyle = workbook.createCellStyle();
                    hSSFCellStyle.setAlignment(HorizontalAlignment.LEFT);
                    hSSFCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

                    spreadsheet.setColumnWidth(0, 5000);
                    spreadsheet.setColumnWidth(1, 5000);
                    spreadsheet.setColumnWidth(2, 3000);
                    spreadsheet.setColumnWidth(3, 3000);
                    spreadsheet.setColumnWidth(4, 15000);
                    spreadsheet.setColumnWidth(5, 4000);
                    spreadsheet.setColumnWidth(6, 6000);
                    spreadsheet.setColumnWidth(7, 10000);
                }
            }


            System.out.println("Done writing images");


            try (FileOutputStream out = new FileOutputStream(new File(EXCEL_LOCATION + documentCount + ".xls"))) {

                workbook.write(out);
            }

            System.out.println("Excel written successfully under: " + EXCEL_LOCATION + " ...");


            Cell cell = null;
            hSSFCellStyle = null;
            HSSFRow row = null;
            spreadsheet = null;
            workbook.close();
            workbook = null;
        } catch (Exception e) {
            System.out.println("Error: " + e.getMessage());
        }
    }


    private static void writeHeaders(HSSFSheet spreadsheet) {
        HSSFRow headerRow = spreadsheet.createRow(0);
        int col = 0;

        HSSFCell hSSFCell = headerRow.createCell(col++);
        hSSFCell.setCellValue("Company");

        hSSFCell = headerRow.createCell(col++);
        hSSFCell.setCellValue("Country");

        hSSFCell = headerRow.createCell(col++);
        hSSFCell.setCellValue("Hall");

        hSSFCell = headerRow.createCell(col++);
        hSSFCell.setCellValue("Stand");

        hSSFCell = headerRow.createCell(col++);
        hSSFCell.setCellValue("Address");

        hSSFCell = headerRow.createCell(col++);
        hSSFCell.setCellValue("Telephone");

        hSSFCell = headerRow.createCell(col++);
        hSSFCell.setCellValue("Email");

        hSSFCell = headerRow.createCell(col);
        hSSFCell.setCellValue("Website");
    }
}


class EventGrabberThread implements Runnable {
    private Set<Event> events;
    private int page;
    private int attempts;

    public EventGrabberThread(Set<Event> events, int page) {
        this.attempts = 0;

        this.events = events;
        this.page = page;
    }


    public void run() {
        try {
            System.out.println("Executing page " + this.page + "...");
            for (Event event : this.events) {
                Document document = Jsoup.connect(event.url).get();

                int pageAttemps = 15;

                while (pageAttemps > 0 && document.select(".cont").isEmpty()) {
                    System.out.println("Retrying");
                    try {
                        Thread.sleep(3000L);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }

                    pageAttemps--;
                    document = Jsoup.connect(event.url).get();
                }

                if (pageAttemps <= 0) {
                    System.out.println("Error getting page (15 attemps exhausted): " + event.url);

                    continue;
                }

                event.address = document.select(".cont").text().trim();
                event.telephone = document.select(".sico.ico_phone").text().trim();
                event.email = document.select(".sico.ico_email").text().trim();
                event.website = document.select(".sico.ico_link").text().trim();

                Main.allEvents.add(event);
            }

        } catch (IOException e) {
            this.attempts++;
            if (this.attempts < 5) {
                run();
            } else {
                System.out.println("Error getting event");
            }
        }
        System.out.println("Done page " + this.page + ".");
    }
}

class Event {
    String url;
    String company;
    String country;
    String hall;
    String stand;
    String address;
    String telephone;
    String email;
    String website;

    public String toString() {
        return "Event{url='" + this.url + '\'' + ", company='" + this.company + '\'' + ", country='" + this.country + '\'' + ", hall='" + this.hall + '\'' + ", stand='" + this.stand + '\'' + ", address='" + this.address + '\'' + ", telephone='" + this.telephone + '\'' + ", email='" + this.email + '\'' + ", website='" + this.website + '\'' + '}';
    }
}
