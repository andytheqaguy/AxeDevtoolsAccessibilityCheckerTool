package com.andytheqaguy;

import com.deque.axe.AXE;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.*;
import java.net.URL;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Properties;

public class Main {
    private static final String propertiesFileName = "config"; // Properties file name to be used for the tests
    private static final URL scriptUrl = Main.class.getResource("/axe.min.js");
    static ChromeOptions options = new ChromeOptions();
    static WebDriver driver; // Initializes the driver
    private static int rowNumberFirst = 0;
    private static int rowNumberLast = 1;
    private static int rowNumber = 2;
    private static final int firstColumn = 0;
    private static final int secondColumn = 1;
    private static String fileName = "Accessibility Report "+ getProperty("project.name") + " "; // File name to be used for the report, final file name will be: Accessibility Report <Project> dd.MM.yyyy
    private static final Logger log = LoggerFactory.getLogger(Main.class);
    private final String[] header = {"User Type", "URL", "Name", "Impact", "Count", "HTML Target"}; // Header columns
    private final XSSFWorkbook workbook = new XSSFWorkbook();
    private final Sheet sheet = workbook.createSheet("Accessibility Report"); // Creates the sheet named "Accessibility report"
    private final Row headerRow = sheet.createRow(0);
    private final String env = getProperty("env");
    private final String lang = getProperty("lang");
    private final String defaultUrl = getProperty("url."+env) + getProperty("lang") + "/";
    private final String loginPage = getProperty("login.page"); // Path (optional) to be used in case tests need to be run with a logged-in user
    private final String logoutPage = getProperty("logout.page");
    private boolean writeFile = false;
    private int writeCount = 0;
    private String typeOfUser = "";

    public static void main(String[] args) {
        Main tool = new Main();
        log.info("--------------------");
        log.info("SCRIPT STARTED!");
        tool.startScript();
        log.info("SCRIPT ENDED!");
        log.info("--------------------");
    }

    public static void startDriver() { // Adds arguments into the driver
        //"--headless"
        options.addArguments();
        driver = new ChromeDriver(options);
    }

    public static void closeDriver() { // Closes the driver after execution
        if (driver != null) {
            driver.quit();
        }
    }

    public static String getProperty(String propertyName) { // Method to retrieve properties from .properties file
        try {
            Properties properties = new Properties();
            String fileName = "src/main/resources/" + propertiesFileName + ".properties";
            FileInputStream fileInputStream = new FileInputStream(fileName);
            properties.load(fileInputStream);

            return properties.getProperty(propertyName);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }

    public void startScript() {
        startDriver();

        log.info("--------------------");
        log.info("Environment used is: " + env.toUpperCase());
        log.info("Lang used is: " + lang.toUpperCase());
        log.info("Default URL used is: " + defaultUrl);

        String[] userTypeList = (getProperty("usertype")).split("\\s*,\\s*");

        for (String userType : userTypeList) {
            String username = getProperty("username." + userType);
            String password = getProperty("password." + userType);

            typeOfUser = userType;

            log.info("--------------------");
            log.info("User type used is: " + typeOfUser.toUpperCase());

            if (getProperty("login." + userType).equals("true")) {
                if (!username.isEmpty() && !password.isEmpty()) {
                    userLogin(username, password);
                    writeFile = true;
                    analyzePaths();
                    userLogout();
                } else {
                    log.info("Username or password for user type '" + userType + "' in the config file is empty");
                    log.info("Please correct the information and run the script again");
                }
            } else {
                writeFile = true;
                analyzePaths();
            }
        }

        if (writeFile && writeCount > 0) { // It writes the Excel file only if at least 1 path was tested
            writeExcelFile();

            log.info("--------------------");
            log.info("SUCCESS!");
            log.info("Excel file was created with name '" + fileName + "'");
            log.info("--------------------");
        }
        else {
            log.info("--------------------");
            log.info("ERROR!");
            log.info("Excel file was not created");
            log.info("--------------------");
        }

        closeDriver();
    }

    public void userLogin(String username, String password) { // Login method in case pages need to be tested from a logged-in user perspective as well
        driver.navigate().to(defaultUrl + loginPage);
        WebElement usernameField = driver.findElement(By.id("edit-name"));
        WebElement passwordField = driver.findElement(By.id("edit-pass"));
        usernameField.sendKeys(username);
        passwordField.sendKeys(password + Keys.ENTER);
    }

    private void userLogout() { // Logout method
        driver.navigate().to(defaultUrl + logoutPage);
    }

    public void analyzePaths() { // Method to test all the paths from paths.userType
        if (writeFile) {
            createExcelSheetAndHeader();

            try {
                String[] pathsList = (getProperty("paths." + typeOfUser)).split("\\s*,\\s*");
                for (String path : pathsList) {  // Iterates through the list of paths
                    String url = defaultUrl + validatePath(path); // Creates the url with path to navigate to
                    log.info("URL used is: " + url);

                    writeCount++;
                    driver.navigate().to(url);

                    JSONObject response = new AXE.Builder(driver, scriptUrl).analyze(); // Returns the analyzed web page as a JSONObject response
                    JSONArray violations = response.getJSONArray("violations"); // Returns only the violations from the response

                    if (!violations.isEmpty()) { // Checks if the number of violations is greater than 0
                        rowNumberFirst = rowNumberLast + 1; // Creates the first row number for the violations
                        rowNumberLast = rowNumberLast + violations.length(); // Creates the last row number for the violations

                        analyzeViolationsAndCreateExcelRow(violations); // Analyzes violations and creates Excel rows for each violation

                        mergeURLCells(rowNumberFirst, rowNumberLast, firstColumn); // Merges cells for the current "violation" node when URL is the same
                        mergeURLCells(rowNumberFirst, rowNumberLast, secondColumn); // Merges cells for the current "violation" node when User Type is the same
                    }
                }
            } catch(Exception ignored){
            }
        }
    }

    public String validatePath(String path) { // Removes the '/' from the path's first character if it exists
        if (path.startsWith("/"))
            return path.substring(1);
         else
            return path;
    }

    public void analyzeViolationsAndCreateExcelRow(JSONArray violations) {
        for (int i = 0; i < violations.length(); i++) {
            JSONObject violation = violations.getJSONObject(i);

            String URL = driver.getCurrentUrl(); // Tested URL
            String name = violation.getString("help"); // Name of the violation

            JSONArray parentNode = violation.getJSONArray("nodes");
            JSONObject childNode = parentNode.getJSONObject(0);

            String impact = childNode.getString("impact"); // Impact level of the violation

            int count = 0;

            String[] htmlTargets = new String[50]; // Creates a String array, size set to 50
            StringBuilder htmlTarget = new StringBuilder();

            for (int j = 0; j < parentNode.length(); j++) { // Iterates through the nodes of "nodes" node
                JSONObject node = parentNode.getJSONObject(j);
                if (node.has("html")){ // Checks if the node has a "html" key
                    htmlTargets[j] = node.getString("html"); // Creates the StringBuilder that contains the html element
                    count+=1; // Counter of the keys "html" inside "nodes" mode
                    if (count == parentNode.length()-1) { // Adds "linebreakhere" only if the count is not equal to the first to penultimate StringBuilder variable
                        htmlTarget.append(htmlTargets[j]).append(" linebreakhere ");
                    }
                    else
                        htmlTarget.append(htmlTargets[j]);
                }
            }

            writeExcelRow(typeOfUser, rowNumber, URL, name, impact, count, htmlTarget); // Writes a single Excel row with the gathered information
            rowNumber++;
        }
    }

    public void createExcelSheetAndHeader() {
        // Creates the header style, with bold text
        CellStyle headerStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        short fontSize = 14;
        font.setFontHeightInPoints(fontSize);
        headerStyle.setFont(font);
        headerStyle.setBorderBottom(BorderStyle.MEDIUM);

        // Creates the header with the header style
        for (int i = 0; i < header.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(header[i]);
        }
    }

    public void writeExcelRow(String userType, int rowNumber, String URL, String name, String impact, int count, StringBuilder target) {
        // Creates data rows
        Row dataRow = sheet.createRow(rowNumber-1);
        Cell userTypeCell = dataRow.createCell(0);
        Cell URLCell = dataRow.createCell(1);
        Cell nameCell = dataRow.createCell(2);
        Cell impactCell = dataRow.createCell(3);
        Cell countCell = dataRow.createCell(4);
        Cell targetCell = dataRow.createCell(5);

        dataRow.setRowStyle(createAlignCenterStyle());

        userTypeCell.setCellStyle(createAlignCenterStyle());
        userTypeCell.setCellValue(userType);

        URLCell.setCellStyle(createAlignCenterStyle());
        URLCell.setCellValue(URL);

        nameCell.setCellStyle(createAlignCenterStyle());
        nameCell.setCellValue(name);

        impactCell.setCellStyle(createAlignCenterStyle());
        impactCell.setCellValue(impact);

        countCell.setCellStyle(createAlignCenterStyle());
        countCell.setCellValue(count);

        // Creates a String variable from StringBuilder, with each String on a separate cell row
        String targetInputString = target.toString();
        String[] substrings = targetInputString.split("linebreakhere");
        String targetToBeWritten = String.join("\n", substrings);

        targetCell.setCellStyle(createAlignCenterStyle());
        targetCell.setCellValue(targetToBeWritten);

        // Accesses the setBorders method only if it's the last row of the violations found on the particular page
        if (rowNumber == rowNumberLast) {
            setBorders(userTypeCell, URLCell, nameCell, impactCell, countCell, targetCell);
        }
    }

    public void setBorders(Cell userTypeCell, Cell URLCell, Cell nameCell, Cell impactCell, Cell countCell, Cell targetCell) {
        // Sets thin borders after the last cell of the violations found on the particular page
        userTypeCell.setCellStyle(createBottomThinBorderStyle());
        URLCell.setCellStyle(createBottomThinBorderStyle());
        nameCell.setCellStyle(createBottomThinBorderStyle());
        impactCell.setCellStyle(createBottomThinBorderStyle());
        countCell.setCellStyle(createBottomThinBorderStyle());
        targetCell.setCellStyle(createBottomThinBorderStyle());
    }

    public void mergeURLCells(int firstRow, int lastRow, int column) {
        // Merges a range of cells
        CellRangeAddress cellMerge = new CellRangeAddress(firstRow-1, lastRow-1, column, column);
        sheet.addMergedRegion(cellMerge);
    }

    private CellStyle createAlignCenterStyle() {
        // Creates a CellStyle with style vertical alignment
        CellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private CellStyle createBottomThinBorderStyle() {
        // Creates a CellStyle with style bottom thin border
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        return style;
    }

    public void writeExcelFile() {
        // Creates a CellRangeAddress for column 1 to column 6
        CellRangeAddress cellFilter = new CellRangeAddress(0, 0, 0, header.length-1);

        // Sets a filter for the above-mentioned ranges
        sheet.setAutoFilter(cellFilter);

        // Creates freeze pane on row 1
        sheet.createFreezePane(0, 0);

        // Sets the column width for all columns
        sheet.setColumnWidth(0, 11*256);
        sheet.setColumnWidth(1, 101*256);
        sheet.setColumnWidth(2, 61*256);
        sheet.setColumnWidth(3, 9*256);
        sheet.setColumnWidth(4, 8*256);
        sheet.setColumnWidth(5, 255*256);

        // Sets file name with the current date
        LocalDate today = LocalDate.now();
        String formattedDate = today.format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
        fileName = fileName + formattedDate;

        // Writes Excel file
        try {
            FileOutputStream out = new FileOutputStream("./" + fileName +".xlsx");
            workbook.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
