/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.raj.deleterecordwithcaptcha;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

public class ImprovedCode {

    private static String readCaptcha(WebElement captchaImage) {
        try {
            // Get the screenshot of the captcha image
            File screenshot = captchaImage.getScreenshotAs(OutputType.FILE);

            // Use Tesseract to perform OCR on the captcha image
            ITesseract tesseract = new Tesseract();
            tesseract.setDatapath("src\\main\\java\\tessdata\\"); // Set the path to your Tesseract data directory
            return tesseract.doOCR(screenshot);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public static void main(String[] args) throws IOException, InterruptedException {
        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();

        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.get("https://tn.unitetools.in/");

        driver.findElement(By.id("user")).sendKeys("TN12101006_de@coopsindia.com");
        driver.findElement(By.id("pwd")).sendKeys("Unite@123");
        driver.findElement(By.id("btnvalidatelogin")).click();

        // Click on 'Customer Data' link
        driver.findElement(By.linkText("Customer Data")).click();

        // Navigate to the delete page
        driver.findElement(By.xpath("//a[contains(@href,'/Utilities/DeleteCustomerData/DeleteCustomerData')]/li[contains(@class,'list-group-item') and contains(@class,'bg-primary')]")).click();

        // Read data from Excel
        FileInputStream fileInputStream = new FileInputStream(new File("D:\\Test.xlsx"));
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // Iterate through each row in the Excel sheet
        DataFormatter dataFormatter = new DataFormatter();
        WebDriverWait wait = new WebDriverWait(driver, 10);
        JavascriptExecutor js = (JavascriptExecutor) driver;

        for (Row row : sheet) {
            Cell cell = row.getCell(0); // Assuming column A contains the values
            if (cell != null && cell.getCellTypeEnum() == CellType.NUMERIC) {
                String formattedValue = dataFormatter.formatCellValue(cell);
                System.out.println("Processing value: " + formattedValue);

                try {
                    // Perform search
                    performSearch(driver, formattedValue, wait);

                    // Handle search results
                    handleSearchResults(driver, formattedValue, wait, js);

                    // Delete the record with captcha validation
                    handleCaptchaAndDelete(driver, wait, js);

                } catch (Exception e) {
                    System.out.println("Error processing value " + formattedValue + ": " + e.getMessage());
                }
            }
        }

        workbook.close();
        fileInputStream.close();
        driver.quit(); // Close the WebDriver session
    }

    private static void performSearch(WebDriver driver, String formattedValue, WebDriverWait wait) {
        WebElement inputField = driver.findElement(By.id("idamissionno"));
        inputField.clear(); // Clear the input field
        inputField.sendKeys(formattedValue); // Enter the value from Excel

        // Click on search button
        WebElement searchButton = driver.findElement(By.id("iconMemberSearch"));
        searchButton.click();
    }

    private static void handleSearchResults(WebDriver driver, String formattedValue, WebDriverWait wait, JavascriptExecutor js) throws InterruptedException {
        try {
            // Check if the "No Data Found" alert is present
            WebElement alertBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
            WebElement alertTitle = alertBox.findElement(By.xpath(".//h2[text()='No Data Found']"));

            if (alertTitle.isDisplayed()) {
                WebElement okButton = alertBox.findElement(By.xpath(".//button[@class='confirm']"));
                okButton.click();
                return;
            }
        } catch (Exception e) {
            System.out.println("No 'No Data Found' alert, proceeding with deletion: " + e.getMessage());
        }

        // Explicitly wait for the product dropdown to be clickable
        WebElement productDropdown = wait.until(ExpectedConditions.elementToBeClickable(By.id("ProductCode")));
        Select productSelect = new Select(productDropdown);
        productSelect.selectByVisibleText("Customer");

        // Select "Admission No." from the task type dropdown
        Select taskTypeDropdown = new Select(driver.findElement(By.id("TaskType")));
        taskTypeDropdown.selectByValue("1");

        // Wait for any loading overlays to disappear
        wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("div.bg[style*='position: absolute']")));

        // Click the delete button
        WebElement deleteButton = driver.findElement(By.id("btnDelete"));
        wait.until(ExpectedConditions.elementToBeClickable(deleteButton));
        js.executeScript("arguments[0].click();", deleteButton);

        handleMandatoryFieldAlert(driver, formattedValue, wait, js);
    }

    private static void handleMandatoryFieldAlert(WebDriver driver, String formattedValue, WebDriverWait wait, JavascriptExecutor js) throws InterruptedException {
        try {
            WebElement alertBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
            WebElement alertTitle = alertBox.findElement(By.xpath(".//h2[text()='* Marked Fields Are Mandatory']"));

            if (alertTitle.isDisplayed()) {
                WebElement okButton = alertBox.findElement(By.xpath(".//button[@class='confirm']"));
                okButton.click();

                handleMandatoryFields(driver, formattedValue, wait, js);

                WebElement deletePage = driver.findElement(By.xpath("//a[contains(@href,'/Utilities/DeleteCustomerData/DeleteCustomerData')]/li[contains(@class,'list-group-item') and contains(@class,'bg-primary')]"));
                deletePage.click();
            }
        } catch (Exception e) {
            System.out.println("No mandatory field alert: " + e.getMessage());
        }
    }

    private static void handleMandatoryFields(WebDriver driver, String formattedValue, WebDriverWait wait, JavascriptExecutor js) throws InterruptedException {
        WebElement homeButton = driver.findElement(By.xpath("//a[@href='/Home/MenuScreen']"));
        homeButton.click();

        // Locate and click the "Personal Details" tab
        WebElement personalDetailsTab = wait.until(ExpectedConditions.elementToBeClickable(By.id("v-pills-01-tab")));
        personalDetailsTab.click();

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("v-pills-01")));

        WebElement verifyButton = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#v-pills-01 .list-group a[href='/Membership/ModifyCustomerPersonalDeatils/VerifyorEditCustomerPersonalDeatils?formid=1003&moduleid=1']")));
        verifyButton.click();

        WebElement inputField1 = driver.findElement(By.id("idamissionno"));
        inputField1.clear(); // Clear the input field
        inputField1.sendKeys(formattedValue); // Enter the value from Excel

        WebElement statusDropdown = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("Status")));
        Select statusSelect = new Select(statusDropdown);
        statusSelect.selectByVisibleText("Not Matched");

        WebElement reasonDropdown = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ModificationReasonId")));
        Select reasonSelect = new Select(reasonDropdown);
        reasonSelect.selectByVisibleText("Record mismatch");

        WebElement editButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("btnModify")));
        editButton.click();

        wait.until(ExpectedConditions.jsReturnsValue("return document.readyState=='complete'"));

        WebElement surnameInput = driver.findElement(By.id("MemberSurName"));
        surnameInput.clear();
        surnameInput.sendKeys("Surname");

        WebElement nameInput = driver.findElement(By.id("MemberName"));
        nameInput.clear();
        nameInput.sendKeys("Name");

        WebElement stateCode = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("StateCode")));
        if (stateCode.isEnabled()) {
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", stateCode);
            Select stateCodeSelect = new Select(stateCode);
            stateCodeSelect.selectByVisibleText("Tamil Nadu");
        } else {
            System.out.println("State code dropdown is not enabled.");
        }

        try {
            WebElement pinCode = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("Pincode")));
            if (pinCode.isDisplayed() && pinCode.isEnabled()) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", pinCode);
                ((JavascriptExecutor) driver).executeScript("arguments[0].removeAttribute('readonly')", pinCode);
                pinCode.clear();
                pinCode.sendKeys("627602");
            } else {
                System.out.println("Pin code field is not displayed or not enabled.");
            }
        } catch (NoSuchElementException e) {
            System.out.println("Pin code element not found.");
        }

        WebElement submitButton = driver.findElement(By.id("btnSubmit"));
        submitButton.click();

        WebElement alertBox2 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
        WebElement alertTitle2 = alertBox2.findElement(By.xpath(".//h2[text()='Data updated successfully']"));

        if (alertTitle2.isDisplayed()) {
            WebElement okButton2 = alertBox2.findElement(By.xpath(".//button[@class='confirm']"));
            okButton2.click();

            WebElement homeButton2 = driver.findElement(By.xpath("//a[@href='/Home/MenuScreen']"));
            homeButton2.click();

            WebElement customerData1 = driver.findElement(By.linkText("Customer Data"));
            customerData1.click();

            // Navigate to the delete page
            WebElement deletePage2 = driver.findElement(By.xpath("//a[contains(@href,'/Utilities/DeleteCustomerData/DeleteCustomerData')]/li[contains(@class,'list-group-item') and contains(@class,'bg-primary')]"));
            deletePage2.click();

            WebElement inputField2 = driver.findElement(By.id("idamissionno"));
            inputField2.clear(); // Clear the input field
            inputField2.sendKeys(formattedValue); // Enter the value from Excel

            WebElement searchButton1 = driver.findElement(By.id("iconMemberSearch"));
            searchButton1.click();

            try {
                // Check if the "No Data Found" alert is present
                WebElement alertBox3 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
                WebElement alertTitle3 = alertBox3.findElement(By.xpath(".//h2[text()='No Data Found']"));

                if (alertTitle3.isDisplayed()) {
                    WebElement okButton1 = alertBox3.findElement(By.xpath(".//button[@class='confirm']"));
                    okButton1.click();
                    return;
                }
            } catch (Exception e) {
                System.out.println("No 'No Data Found' alert, proceeding with deletion: " + e.getMessage());
            }

            WebElement productDropdown1 = wait.until(ExpectedConditions.elementToBeClickable(By.id("ProductCode")));
            Select productSelect1 = new Select(productDropdown1);
            productSelect1.selectByVisibleText("Customer");

            Select taskTypeDropdown1 = new Select(driver.findElement(By.id("TaskType")));
            taskTypeDropdown1.selectByValue("1");

            wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("div.bg[style*='position: absolute']")));

            WebElement deleteButton1 = driver.findElement(By.id("btnDelete"));
            wait.until(ExpectedConditions.presenceOfElementLocated(By.id("btnDelete")));

            JavascriptExecutor js1 = (JavascriptExecutor) driver;
            js1.executeScript("arguments[0].click();", deleteButton1);
        }
    }

    private static void handleCaptchaAndDelete(WebDriver driver, WebDriverWait wait, JavascriptExecutor js) throws InterruptedException {
        while (true) {
            // Read and enter the captcha
            WebElement captchaImage = driver.findElement(By.id("imgcapt"));
            String captchaText = readCaptcha(captchaImage);
            System.out.println("Captcha Text: " + captchaText);

            WebElement captchaInput = driver.findElement(By.id("Captcha"));
            captchaInput.clear();
            captchaInput.sendKeys(captchaText);

            // Click validate button
            WebElement validateButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(text(), 'Validate')]")));
            validateButton.click();

            // Handle SweetAlert confirmation dialog
            WebElement sweetAlert = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
            WebElement h2Element = sweetAlert.findElement(By.tagName("h2"));
            String message = h2Element.getText().trim();
            System.out.println("message--------> " + message);

            if (message.equals("Data deleted successfully")) {
                WebElement okButton = sweetAlert.findElement(By.className("confirm"));
                okButton.click();
                break; // Exit the loop if data is successfully deleted
            } else if (message.equals("Enter valid captcha")) {
                WebElement okButton = sweetAlert.findElement(By.className("confirm"));
                okButton.click();

                // Click Refresh Captcha button
                WebElement refreshCaptchaButton = driver.findElement(By.xpath("//button[@class='btn btn-primary' and contains(@onclick,'RefreshCaptcha')]"));
                refreshCaptchaButton.click();
                Thread.sleep(2000); // Wait for captcha to refresh
            } else {
                System.out.println("Unexpected SweetAlert message: " + message);
                break; // Exit the loop in case of unexpected message
            }
        }
    }
}
