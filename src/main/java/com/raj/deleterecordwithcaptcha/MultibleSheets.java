package com.raj.deleterecordwithcaptcha;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.*;
import org.openqa.selenium.OutputType;

public class MultibleSheets {

    private static ITesseract tesseract = new Tesseract();

    static {
        tesseract.setDatapath("src/main/java/tessdata/"); // Set the path to your Tesseract data directory
    }

    private static String readCaptcha(WebElement captchaImage) {
        try {
            // Get the screenshot of the captcha image
            File screenshot = captchaImage.getScreenshotAs(OutputType.FILE);
            return tesseract.doOCR(screenshot);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public static void main(String[] args) throws IOException {
        WebDriverManager.chromedriver().setup();

        // Read data from Excel
        FileInputStream fileInputStream = new FileInputStream(new File("D:\\Test.xlsx"));
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        // Process each sheet concurrently
        ExecutorService executorService = Executors.newCachedThreadPool();
        List<Callable<Void>> tasks = new ArrayList<>();
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            WebDriver driver = new ChromeDriver();
            tasks.add(() -> processSheet(sheet, driver));
        }

        try {
            // Submit all tasks for execution and wait for completion
            executorService.invokeAll(tasks);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            e.printStackTrace();
        } finally {
            executorService.shutdown();
            workbook.close();
            fileInputStream.close();
        }
    }

    private static Void processSheet(Sheet sheet, WebDriver driver) {
        try {
            driver.manage().window().maximize();
            driver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);
            driver.get("https://tn.unitetools.in/");

            driver.findElement(By.id("user")).sendKeys("TN12101048_de@coopsindia.com");
            driver.findElement(By.id("pwd")).sendKeys("Unite@123");
            driver.findElement(By.id("btnvalidatelogin")).click();

            // Click on 'Customer Data' link
            WebElement excelUploadLink = driver.findElement(By.linkText("Customer Data"));
            excelUploadLink.click();

            // Navigate to the delete page
            WebElement deletePage = driver.findElement(By.xpath("//a[contains(@href,'/Utilities/DeleteCustomerData/DeleteCustomerData')]/li[contains(@class,'list-group-item') and contains(@class,'bg-primary')]"));
            deletePage.click();

            // Iterate through each row in the Excel sheet
            DataFormatter dataFormatter = new DataFormatter();
            WebDriverWait wait = new WebDriverWait(driver, 1);
            for (Row row : sheet) {
                Cell cell = row.getCell(0); // Assuming column A contains the values
                if (cell != null && cell.getCellTypeEnum() == CellType.NUMERIC) {
                    String formattedValue = dataFormatter.formatCellValue(cell);
                    System.out.println("Processing value: " + formattedValue);

                    WebElement inputField = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("idamissionno")));
                    inputField.clear(); // Clear the input field
                    inputField.sendKeys(formattedValue); // Enter the value from Excel

                    // Click on search button
                    WebElement searchButton = driver.findElement(By.id("iconMemberSearch"));
                    searchButton.click();

                    try {
                        // Check if the "No Data Found" alert is present
                        WebElement alertBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
                        WebElement alertTitle = alertBox.findElement(By.xpath(".//h2[text()='No Data Found']"));

                        if (alertTitle.isDisplayed()) {
                            WebElement okButton = alertBox.findElement(By.xpath(".//button[@class='confirm']"));
                            okButton.click();
                            continue;// Move to next iteration
                        }
                    } catch (Exception e) {
                        // No alert found, proceed with deletion process
                        System.out.println("No alert found, proceeding with deletion for value: " + e.getMessage());
                    }

                    // Explicitly wait for the product dropdown to be clickable
                    WebElement productDropdown = wait.until(ExpectedConditions.elementToBeClickable(By.id("ProductCode")));
                    Select productSelect = new Select(productDropdown);
                    productSelect.selectByVisibleText("Customer");

                    // Select "Admission No." from the task type dropdown
                    Select taskTypeDropdown = new Select(driver.findElement(By.id("TaskType")));
                    taskTypeDropdown.selectByValue("1");

                    // Click the delete button
                    WebElement deleteButton = driver.findElement(By.id("btnDelete"));
                    deleteButton.click();

                    try {
                        // Check if the "* Marked Fields Are Mandatory" alert is present
                        WebElement alertBox1 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
                        WebElement alertTitle1 = alertBox1.findElement(By.xpath(".//h2[text()='* Marked Fields Are Mandatory']"));

                        if (alertTitle1.isDisplayed()) {
                            WebElement okButton = alertBox1.findElement(By.xpath(".//button[@class='confirm']"));
                            okButton.click();
                            continue;// Move to next iteration
                        }
                    } catch (Exception e) {
                        // No alert found, proceed with deletion process
                        System.out.println("Marked filed---> " + e.getMessage());
                    }

                    // Find and read the captcha image
                    WebElement captchaImage = driver.findElement(By.id("imgcapt"));
                    String captchaText = readCaptcha(captchaImage);

                    System.out.println("Captcha Text: " + captchaText);

                    WebElement captchaInput = driver.findElement(By.id("Captcha"));
                    captchaInput.sendKeys(captchaText);

                    // Click validate button
                    WebElement validateButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(text(), 'Validate')]")));
                    validateButton.click();

                    // Handle SweetAlert confirmation dialog
                    WebElement sweetAlert = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
                    WebElement h2Element = sweetAlert.findElement(By.tagName("h2"));
                    String message = h2Element.getText().trim();
                    System.out.println("messageeeee--------> " + message);

                    while (true) {
                        if (message.equals("Data deleted successfully")) {
                            // Click OK button
                            WebElement okButton = sweetAlert.findElement(By.className("confirm"));
                            okButton.click();
                            break; // Exit the loop if data is successfully deleted
                        } else if (message.equals("Enter valid captcha")) {
                            // Click OK button
                            WebElement okButton = sweetAlert.findElement(By.className("confirm"));
                            okButton.click();

                            // Click Refresh Captcha button
                            WebElement refreshCaptchaButton = driver.findElement(By.xpath("//button[@class='btn btn-primary' and contains(@onclick,'RefreshCaptcha')]"));
                            refreshCaptchaButton.click();

                            // Find and read the new captcha image
                            WebElement captchaImage2 = driver.findElement(By.id("imgcapt"));
                            String captchaText2 = readCaptcha(captchaImage2);

                            System.out.println("Captcha Text 2: " + captchaText2);

                            WebElement captchaInput2 = driver.findElement(By.id("Captcha"));
                            captchaInput2.sendKeys(captchaText2);

                            // Click validate button for the second attempt
                            WebElement validateButton2 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(text(), 'Validate')]")));
                            validateButton2.click();

                            // Wait for the next SweetAlert confirmation dialog
                            sweetAlert = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
                            h2Element = sweetAlert.findElement(By.tagName("h2"));
                            message = h2Element.getText().trim();
                        } else {
                            System.out.println("Unexpected SweetAlert message: " + message);
                            break; // Exit the loop in case of unexpected message
                        }
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit(); // Close the WebDriver session
        }

        return null;
    }
}
