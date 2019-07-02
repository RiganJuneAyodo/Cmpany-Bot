package sample;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Map;

public class Controller {

    private static String generalDir;

    public static final String ANSI_RESET = "\u001B[0m";
    public static final String ANSI_BLACK = "\u001B[30m";
    public static final String ANSI_RED = "\u001B[31m";
    public static final String ANSI_GREEN = "\u001B[32m";
    public static final String ANSI_YELLOW = "\u001B[33m";
    public static final String ANSI_BLUE = "\u001B[34m";
    public static final String ANSI_PURPLE = "\u001B[35m";
    public static final String ANSI_CYAN = "\u001B[36m";
    public static final String ANSI_WHITE = "\u001B[37m";


    public static ArrayList<String> columns = new ArrayList<>();


    public static void main(String[] args) throws InterruptedException {

        //Get all the columns
        try {

            generalDir = new JFileChooser().getFileSystemView().getDefaultDirectory().toString() + File.separator + "CompanyBot";

            File docDirectory = new File(generalDir);

            if (!docDirectory.exists()) {

                docDirectory.mkdirs();

            }

            //Add Columns

            String Company_Number = "Company Number".toUpperCase();
            String Status = "Status".toUpperCase();
            String Incorporation_Date = "Incorporation Date".toUpperCase();
            String MAILING_ADDRESS = "MAILING ADDRESS";
            String TELEPHONE_NUMBER = "TELEPHONE NUMBER";

            columns.add(Status.trim().toUpperCase());
            columns.add(Incorporation_Date.trim().toUpperCase());
            columns.add(MAILING_ADDRESS.trim().toUpperCase());
            columns.add(Company_Number.trim().toUpperCase());
            columns.add(TELEPHONE_NUMBER.trim().toUpperCase());


            //Get all the columns
            ChromeOptions options = new ChromeOptions();

//            options.addArguments("--headless"); //, "--disable-gpu", "--window-size=1920,1200","--ignore-certificate-errors", "--silent");

            options.addArguments("--disable-popup-blocking");

            Map<String, Object> preferences = new Hashtable<String, Object>();

            options.setExperimentalOption("prefs", preferences);

            System.setProperty("webdriver.chrome.driver","C:\\chromedriver.exe");

            System.out.println("Launching Bot ");

            ChromeDriver driver = new ChromeDriver(options);

            System.out.println("Opening Chrome Browser ");


            driver.get("https://opencorporates.com/companies?jurisdiction_code=&q=&utf8=%E2%9C%93");

            System.out.println("Connection established to the site ");

            File[] allFiles = new File(generalDir).listFiles();


            System.out.println("Scanning Documents ");

            //Check if there is a file
            if (allFiles != null && allFiles.length <= 0 ) {

                JOptionPane.showMessageDialog(null,"No files in the Directory (" + docDirectory.getPath() + ")", "Files", JOptionPane.ERROR_MESSAGE);

                return;

            }

            assert allFiles != null;

            for (File docFile: allFiles) {

                if (driver.getSessionId() == null) {

                    System.exit(0);

                }




                //Check if the document has been processed
                if (!FilenameUtils.getBaseName(docFile.getName()).contains("Done") || !FilenameUtils.getBaseName(docFile.getName()).contains("InProgress")) {

                    //Check if the document is of the correct format
                    if (FilenameUtils.getExtension(docFile.getName()).equalsIgnoreCase("xls") ||
                            FilenameUtils.getExtension(docFile.getName()).equalsIgnoreCase("xlsx") ) {

                        docFile = renamedFile(docFile, "InProgress");

//                        System.out.println(ANSI_BLUE + "Processing Document: " + docFile.getName());

                        if (docFile != null ) {

                            System.out.println(ANSI_BLUE + "Processing Document: " + docFile.getName());

                            try {

                                FileInputStream adfile = new FileInputStream(docFile);

                                XSSFWorkbook workbook = new XSSFWorkbook(adfile);

                                XSSFSheet tmpSheet = workbook.getSheetAt(0);

                                setUpColumns(workbook, docFile);
                                
                                String companyName, status = null, incooperationDate = null, mailingAddress  = null, company_number = null, telephone_numbers = null;

                                for (int rowIndex = 1; rowIndex <= tmpSheet.getLastRowNum(); rowIndex++) {

                                    Map<String, String>  companyData = new HashMap<>();

                                    if (driver.getSessionId() == null) {

                                        System.exit(0);

                                    }

                                    try {

                                        XSSFRow  row = tmpSheet.getRow(rowIndex);

                                        if (row != null) {

                                            DataFormatter formatter = new DataFormatter();

                                            companyName = formatter.formatCellValue(row.getCell(0));

                                            if (companyName != null && !companyName.isEmpty()) {

                                                Row firstRow = tmpSheet.getRow(0);

                                                //Search for value at the Status column and check if its done or not
                                                for (int cn=0; cn< firstRow.getLastCellNum(); cn++) {

                                                    Cell c = firstRow.getCell(cn);

                                                    if (c != null) {

                                                        if (c.getStringCellValue().equalsIgnoreCase(Status)) {

                                                            status = formatter.formatCellValue(row.getCell(cn));

                                                        }

                                                    }

                                                }

                                                if( status == null || status.isEmpty() || status.trim().contains("not founded") || status.trim().equalsIgnoreCase("not founded")) {

                                                    WebElement searchField = driver.findElementByName("q");

                                                    searchField.sendKeys(companyName);

                                                    searchField.submit();


                                                    //Company Results
                                                    if (driver.findElementsById("companies").size() != 0) {

                                                        WebDriverWait wait  = new WebDriverWait(driver,30);

                                                        wait.until(ExpectedConditions.presenceOfElementLocated (By.xpath("//*[@id='companies']/li/a[2]")));

                                                        driver.findElement(By.xpath("//*[@id='companies']/li/a[2]")).click();

                                                        //company_number
                                                        if (driver.findElementsByClassName("company_number").size() != 0) {

                                                            company_number = driver.findElementByClassName("company_number").getText();

                                                            companyData.put(Company_Number, company_number);

                                                        } else {

                                                            companyData.put(Company_Number, "Not Found");

                                                        }

                                                        //Status
                                                        if (driver.findElementsByClassName("status").size() != 0) {

                                                            status = driver.findElementByClassName("status").getText();

                                                            companyData.put(Status, status);

                                                        } else {

                                                            companyData.put(Status, "Not Found");

                                                        }

                                                        //incorporation_date
                                                        if (driver.findElementsByClassName("incorporation_date").size() != 0) {

                                                            incooperationDate = driver.findElementByClassName("incorporation_date").getText();

                                                            companyData.put(Incorporation_Date, incooperationDate);

                                                        } else {

                                                            companyData.put(Incorporation_Date, "Not Found");

                                                        }


                                                        //company_addresses
                                                        if (driver.findElementsById("company_addresses").size() != 0) {

                                                            mailingAddress = driver.findElementById("company_addresses").findElement(By.className("description")).getText();

                                                            companyData.put(MAILING_ADDRESS, mailingAddress);

                                                        } else {

                                                            companyData.put(MAILING_ADDRESS, "Not Found");

                                                        }

                                                        //telephone_numbers
                                                        if (driver.findElementsByCssSelector("div.assertion.telephone_number").size() != 0) {

                                                            telephone_numbers = driver.findElementByCssSelector("div.assertion.telephone_number").findElement(By.className("description")).getText();

                                                            if (telephone_numbers.contains(":")) {

                                                                telephone_numbers = telephone_numbers.split(":")[1].trim();

                                                            }

                                                            companyData.put(TELEPHONE_NUMBER, telephone_numbers);

                                                        } else {

                                                            companyData.put(TELEPHONE_NUMBER, "Not Found");

                                                        }


                                                        //attribute_list
                                                        if (driver.findElementsByClassName("attribute_list").size() != 0) {

                                                            WebElement comapanyNames = driver.findElementByClassName("attribute_list");

                                                            int index = 1;

                                                            for (WebElement nam : comapanyNames.findElements(By.tagName("li"))) {

                                                                String name = "";

                                                                if(nam.findElements(By.tagName("a")).size() != 0 ) {

                                                                    name = nam.findElement(By.tagName("a")).getText();

                                                                } else {

                                                                    name = nam.getText();

                                                                }

                                                                companyData.put(("NAME " + index ), name);

                                                            }

                                                        }



                                                        //add data to the excel sheet
                                                        for (Map.Entry<String, String> data: companyData.entrySet()) {

                                                            //check if the column exists
                                                            boolean colExits = false;

                                                            for (int cn=0; cn< firstRow.getLastCellNum(); cn++) {

                                                                Cell c = firstRow.getCell(cn);

                                                                if (c != null) {

                                                                    if (c.getStringCellValue().equalsIgnoreCase(data.getKey())) {

                                                                        colExits = true;

                                                                        Cell cellStatus = row.getCell(cn);

                                                                        if (cellStatus == null ) {

                                                                            cellStatus = row.createCell(cn);

                                                                            cellStatus.setCellType(CellType.STRING);

                                                                            cellStatus.setCellValue(data.getValue());

                                                                        }

                                                                    }

                                                                }

                                                            }


                                                            //Create column if it does'nt exist
                                                            if (!colExits) {

                                                                int colIndex = tmpSheet.getRow(0).getLastCellNum() + 1;

                                                                //Create its column
                                                                Cell nameCol = tmpSheet.getRow(0).getCell(colIndex);

                                                                if (nameCol == null ) {

                                                                    nameCol = tmpSheet.getRow(0).createCell(colIndex);

                                                                    nameCol.setCellType(CellType.STRING);

                                                                    nameCol.setCellValue(data.getKey());

                                                                }

                                                                //Set its row value
                                                                Cell naame = row.getCell(colIndex);

                                                                if (naame == null ) {

                                                                    naame = row.createCell(colIndex);

                                                                    naame.setCellType(CellType.STRING);

                                                                    //Bold if not found
                                                                    if (data.getValue().equalsIgnoreCase("not found")) {

                                                                        XSSFCellStyle style = workbook.createCellStyle();

                                                                        XSSFFont font = workbook.createFont();

                                                                        font.setFontName("Arial");

                                                                        font.setFontHeightInPoints((short) 10);

                                                                        font.setBold(true);

                                                                        style.setFont(font);

                                                                        naame.setCellStyle(style);

                                                                    }


                                                                    naame.setCellValue(data.getValue());

                                                                }

                                                            }

                                                        }

                                                        // Write the output to the file
                                                        FileOutputStream fileOut = new FileOutputStream(docFile);

                                                        workbook.write(fileOut);

                                                        fileOut.close();

                                                        System.out.println("Row " + rowIndex + " completed successfully");

                                                    }

                                                } else {

                                                    System.out.println("Row " + rowIndex + " Already Done. Skipping.");

                                                }

                                            }

                                        }


                                        if (rowIndex == tmpSheet.getLastRowNum()) {

                                            System.out.println("Done with Document: " + docFile.getName());

                                            System.out.println(ANSI_BLUE + "Marking Document: " + docFile.getName() + " as Finished");

                                            File renmed = renamedFile(docFile, "Done");
                                            if (renmed != null && renmed.exists()) {
                                                System.out.println("File marked as Finished");

                                            }

                                        }

                                    } catch (Exception e) {

                                        System.out.println("Row " + rowIndex + " has experienced an error... Skipping.");

                                        e.printStackTrace();
//                                    System.err.println(e.getMessage());

                                    }

                                }

                            } catch (Exception e) {
                                e.printStackTrace();

                            }

                        }

                    }

                }

            }

            driver.quit();

        } catch (Exception e) {

            System.err.println(e.getMessage());

            e.printStackTrace();
        }

    }



    public static File renamedFile(File docFile, String status) {

        String docName = docFile.getName();

        if (StringUtils.countMatches(FilenameUtils.getBaseName(docFile.getName()), "InProgress") == 0 &&
                StringUtils.countMatches(FilenameUtils.getBaseName(docFile.getName()), "Done") == 0) {

            if(status.equalsIgnoreCase("Done") && docName.contains("Done")) {

                docName = docName.replace("InProgess", status);

            } else {

                docName = status + "_" + docName;

            }

            //Rename file
            String newFilePath = docFile.getAbsolutePath().replace(docFile.getName(), docName);


            File newFile = new File(newFilePath);

            try {

                FileUtils.moveFile(docFile, newFile);

                return newFile;

            } catch (Exception e) {

                System.out.println("Unable to Mark file! " + e.getMessage());

                System.err.println(e.getMessage());

            }

        }


        return null;
    }

    private static void setUpColumns(XSSFWorkbook myWorkBook, File docFile) {

        //Set up the columns
        XSSFCellStyle style = myWorkBook.createCellStyle();

        XSSFFont font = myWorkBook.createFont();

        font.setFontName("Arial");

        font.setFontHeightInPoints((short) 10);

        font.setBold(true);

        style.setFont(font);


        XSSFSheet mySheet = myWorkBook.getSheetAt(0);

        XSSFRow myRow = mySheet.getRow(0);

        if (myRow == null) {

            myRow = mySheet.createRow(0);

        }

        int colCount = myRow.getLastCellNum();

        for (int colIndex = 0; colIndex < colCount; colIndex++ ) {

            Cell colName= myRow.getCell(colIndex);

            if (colName == null ) {

                colName = myRow.createCell(colCount);


            }

            for (String col : columns) {

                if (colName.getStringCellValue().trim().equalsIgnoreCase(col)) {

                    columns.remove(col);

                    break;

                }

            }

            colName.setCellStyle(style);

            mySheet.setColumnWidth(colIndex, 6000);

        }



        int index = myRow.getLastCellNum();

        for (String col : columns ) {

            Cell newCol = myRow.getCell(index);

            if (newCol == null ){

                newCol = myRow.createCell(index);

            }

            newCol.setCellType(CellType.STRING);

            newCol.setCellValue(col);

            newCol.setCellStyle(style);

            mySheet.setColumnWidth(index, 6000);

            index++;

        }

        try {

            FileOutputStream out = new FileOutputStream(docFile.getPath());

            myWorkBook.write(out);

            out.close();

            System.out.println("Columns Set");
        } catch (Exception e) {

            e.printStackTrace();

        }

    }
}
