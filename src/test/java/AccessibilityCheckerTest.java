import com.deque.axe.AXE;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
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

public class AccessibilityCheckerTest {
    private static final URL scriptUrl = AccessibilityCheckerTest.class.getResource("/axe.min.js");
    private static int rowNumberFirst = 0;
    private static int rowNumberLast = 1;
    private static int rowNumber = 2;
    private final String[] header = {"URL", "Name", "Impact", "Count", "HTML Target"}; // Header columns
    private final XSSFWorkbook workbook = new XSSFWorkbook();
    private final Sheet sheet = workbook.createSheet("Accessibility report"); // Creates the sheet named "Accessibility report"
    private final Row headerRow = sheet.createRow(0);
    private final String defaultUrl = getProperty("url") + getProperty("lang"); // Creates the URL, taking into consideration the lang as well
    private final String loginPagePath = "/user/login"; // Path (optional) to be used in case tests need to be run with a logged-in user
    static ChromeOptions options = new ChromeOptions();
    private static String fileName = "Accessibility Report <Project> "; // File name to be used for the report, final file name will be: Accessibility Report <Project> dd.MM.yyyy
    private static String propertiesFileName = "accessibility"; // Properties file name to be used for the tests

    @BeforeAll
    public static void setupDriver() { // Adds arguments into the driver
        //"--headless"
        options.addArguments();
        driver = new ChromeDriver(options);
    }

    static WebDriver driver; // Initializes the driver
    @AfterAll
    public static void quitDriver() { // Closes the driver after execution
        if (driver != null) {
            driver.quit();
        }
    }

    @Test
    public void startScript() {
        String[] userTypeList = (getProperty("usertype")).split("\\s*,\\s*");

        String username = getProperty("username");
        String password = getProperty("password");

        createExcelSheetAndHeader();

        for (String userType : userTypeList) { // Iterates through the list of user types
            switch (userType) { // Different steps for different user types
                case "userType1":
                    testPath(userType);
                    break;
                case "userType2":
                    userLogin(username, password);
                    testPath(userType);
                    break;
            }
        }
        writeExcelFile();
    }

    public void userLogin(String username, String password) { // Login method in case pages need to be tested from a logged-in user perspective as well
        driver.navigate().to(defaultUrl + loginPagePath);
        WebElement usernameField = driver.findElement(By.id("edit-name"));
        WebElement passwordField = driver.findElement(By.id("edit-pass"));
        usernameField.sendKeys(username);
        passwordField.sendKeys(password + Keys.ENTER);
    }

    public static String getProperty(String propertyName) { // Method to retrieve properties from .properties file
        try {
            Properties properties = new Properties();
            propertiesFileName = "src/test/resources/" + propertiesFileName + ".properties";

            FileInputStream fileInputStream = new FileInputStream(propertiesFileName);
            properties.load(fileInputStream);

            return properties.getProperty(propertyName);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public void testPath (String typeOfUser) { // Method to test all the paths from paths.userType
        try {
            String[] pathsList = (getProperty("paths." + typeOfUser)).split("\\s*,\\s*");
            for (String path : pathsList){ // Iterates through the list of paths
                String url = defaultUrl + path; // Creates the url with path to navigate to
                driver.navigate().to(url);
                JSONObject response = new AXE.Builder(driver, scriptUrl).analyze(); // Returns the analyzed web page as a JSONObject response
                JSONArray violations = response.getJSONArray("violations"); // Returns only the violations from the response
                if (violations.length() > 0) { // Checks if the number of violations is greater than 0
                    rowNumberFirst = rowNumberLast+1; // Creates the first row number for the violations
                    rowNumberLast = rowNumberLast + violations.length(); // Creates the last row number for the creations
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
                                count+=1; // Counter of the keys "html" inside "nodes" mpde
                                if (count == parentNode.length()-1) { // Adds "linebreakhere" only if the count is not equal to the first to penultimate StringBuilder variable
                                    htmlTarget.append(htmlTargets[j]).append(" linebreakhere ");
                                }
                                else
                                    htmlTarget.append(htmlTargets[j]);
                            }
                        }
                        writeExcelRow(rowNumber, URL, name, impact, count, htmlTarget); // Writes a single Excel row with the gathered informations
                        rowNumber++;
                    }
                    mergeURLCells(rowNumberFirst, rowNumberLast); // Merges cells for the current "violation" node when URL is the same
                }
            }
        } catch (Exception ignored) {
            }
    }

    public void createExcelSheetAndHeader() {
        // Creates the header style, with bold text
        CellStyle headerStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        headerStyle.setFont(font);
        headerStyle.setBorderBottom(BorderStyle.MEDIUM);

        // Creates the header with the header style
        for (int i = 0; i < header.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(header[i]);
        }
    }

    public void writeExcelRow(int rowNumber, String URL, String name, String impact, int count, StringBuilder target) {
        // Creates data rows
        Row dataRow = sheet.createRow(rowNumber-1);

        Cell URLCell = dataRow.createCell(0);
        Cell nameCell = dataRow.createCell(1);
        Cell impactCell = dataRow.createCell(2);
        Cell countCell = dataRow.createCell(3);
        Cell targetCell = dataRow.createCell(4);

        dataRow.setRowStyle(createAlignCenterStyle());

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
            setBorders(URLCell, nameCell, impactCell, countCell, targetCell);
        }
    }

    public void setBorders(Cell URLCell, Cell nameCell, Cell impactCell, Cell countCell, Cell targetCell) {
        // Sets thin borders after the last cell of the violations found on the particular page
        URLCell.setCellStyle(createBottomThinBorderStyle());
        nameCell.setCellStyle(createBottomThinBorderStyle());
        impactCell.setCellStyle(createBottomThinBorderStyle());
        countCell.setCellStyle(createBottomThinBorderStyle());
        targetCell.setCellStyle(createBottomThinBorderStyle());
    }

    public void mergeURLCells(int firstRow, int lastRow) {
        // Merges a range of cells
        CellRangeAddress cellMerge = new CellRangeAddress(firstRow-1, lastRow-1, 0, 0);
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
        // Auto-size all columns, commented because it doesn't work as expected
        /*for (int i = 0; i < header.length; i++){
            System.out.println("Column: " + i + " has been resized");
            sheet.autoSizeColumn(i);
        }*/

        // Sets file name with the current date
        LocalDate today = LocalDate.now();
        String formattedDate = today.format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
        fileName = fileName + formattedDate;

        // Writes Excel file
        try {
            FileOutputStream out = new FileOutputStream("./" + fileName +".xlsx");
            workbook.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
