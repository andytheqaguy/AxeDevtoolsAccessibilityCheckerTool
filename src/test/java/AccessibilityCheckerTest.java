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
    private final String[] header = {"URL", "Name", "Impact", "Count", "HTML Target"};
    private final XSSFWorkbook workbook = new XSSFWorkbook();
    private final Sheet sheet = workbook.createSheet("Accessibility report");
    private final Row headerRow = sheet.createRow(0);
    private final String defaultUrl = getProperty("url") + getProperty("lang");
    private final String loginPagePath = "/user/login";
    static ChromeOptions options = new ChromeOptions();
    private String fileName = "Accessibility Report <Project> ";

    @BeforeAll
    public static void setupDriver() {
        //"--headless"
        options.addArguments();
        driver = new ChromeDriver(options);
    }

    static WebDriver driver;
    @AfterAll
    public static void quitDriver() {
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

        for (String userType : userTypeList){
            switch (userType){
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

    public void userLogin(String username, String password) {
        driver.navigate().to(defaultUrl + loginPagePath);
        WebElement usernameField = driver.findElement(By.id("edit-name"));
        WebElement passwordField = driver.findElement(By.id("edit-pass"));
        usernameField.sendKeys(username);
        passwordField.sendKeys(password + Keys.ENTER);
    }

    public static String getProperty(String propertyName) {
        try {
            Properties properties = new Properties();
            String filePath = "src/test/resources/accessibility.properties";

            FileInputStream fileInputStream = new FileInputStream(filePath);
            properties.load(fileInputStream);

            return properties.getProperty(propertyName);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public void testPath (String typeOfUser) {
        try {
            String[] pathsList = (getProperty("paths." + typeOfUser)).split("\\s*,\\s*");
            for (String path : pathsList){
                String url = defaultUrl + path;
                driver.navigate().to(url);
                JSONObject response = new AXE.Builder(driver, scriptUrl).analyze();
                JSONArray violations = response.getJSONArray("violations");
                if (violations.length() > 0) {
                    rowNumberFirst = rowNumberLast+1;
                    rowNumberLast = rowNumberLast + violations.length();
                    for (int i = 0; i < violations.length(); i++) {

                        JSONObject violation = violations.getJSONObject(i);
                        String URL = driver.getCurrentUrl();
                        String name = violation.getString("help");

                        JSONArray parentNode = violation.getJSONArray("nodes");
                        JSONObject childNode = parentNode.getJSONObject(0);
                        String impact = childNode.getString("impact");
                        int count = 0;
                        String[] htmlTargets = new String[50];
                        StringBuilder htmlTarget = new StringBuilder();
                        for (int j = 0; j < parentNode.length(); j++) {
                            JSONObject node = parentNode.getJSONObject(j);
                            if (node.has("html")){
                                htmlTargets[j] = node.getString("html");
                                count+=1;
                                if (count == parentNode.length()-1) {
                                    htmlTarget.append(htmlTargets[j]).append(" linebreakhere ");
                                }
                                else
                                    htmlTarget.append(htmlTargets[j]);
                            }
                        }
                        writeExcelRow(rowNumber, URL, name, impact, count, htmlTarget);
                        rowNumber++;
                    }
                    mergeURLCells(rowNumberFirst, rowNumberLast);
                }
            }
        } catch (Exception ignored) {
            }
    }

    public void createExcelSheetAndHeader() {
        //Creating the header
        CellStyle headerStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        headerStyle.setFont(font);
        headerStyle.setBorderBottom(BorderStyle.MEDIUM);

        for (int i = 0; i < header.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(header[i]);
        }
    }

    public void writeExcelRow(int rowNumber, String URL, String name, String impact, int count, StringBuilder target) {
        //Creating data rows
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


        String targetInputString = target.toString();
        String[] substrings = targetInputString.split("linebreakhere");
        String targetToBeWritten = String.join("\n", substrings);

        targetCell.setCellStyle(createAlignCenterStyle());
        targetCell.setCellValue(targetToBeWritten);

        if (rowNumber == rowNumberLast) {
            setBorders(URLCell, nameCell, impactCell, countCell, targetCell);
        }
    }

    public void setBorders(Cell URLCell, Cell nameCell, Cell impactCell, Cell countCell, Cell targetCell) {
        URLCell.setCellStyle(createBottomThinBorderStyle());
        nameCell.setCellStyle(createBottomThinBorderStyle());
        impactCell.setCellStyle(createBottomThinBorderStyle());
        countCell.setCellStyle(createBottomThinBorderStyle());
        targetCell.setCellStyle(createBottomThinBorderStyle());
    }

    public void mergeURLCells(int firstRow, int lastRow) {
        CellRangeAddress cellMerge = new CellRangeAddress(firstRow-1, lastRow-1, 0, 0);
        sheet.addMergedRegion(cellMerge);
    }

    private CellStyle createAlignCenterStyle() {
        CellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private CellStyle createBottomThinBorderStyle() {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        return style;
    }
    
    public void writeExcelFile() {
        //Auto-size all columns
        /*for (int i = 0; i < header.length; i++){
            System.out.println("Column: " + i + " has been resized");
            sheet.autoSizeColumn(i);
        }*/

        //Set file name
        LocalDate today = LocalDate.now();
        String formattedDate = today.format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
        fileName = fileName + formattedDate;

        //Write Excel file
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
