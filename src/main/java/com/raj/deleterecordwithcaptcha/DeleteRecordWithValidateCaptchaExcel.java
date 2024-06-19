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
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;

public class DeleteRecordWithValidateCaptchaExcel {

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
        driver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);
        driver.get("https://tn.unitetools.in/");

        driver.findElement(By.id("user")).sendKeys("TN12102253_de@coopsindia.com");
        driver.findElement(By.id("pwd")).sendKeys("Unite@123");
        driver.findElement(By.id("btnvalidatelogin")).click();

        // Click on 'Customer Data' link
        WebElement customerData = driver.findElement(By.linkText("Customer Data"));
        customerData.click();

        // Navigate to the delete page
        WebElement deletePage = driver.findElement(By.xpath("//a[contains(@href,'/Utilities/DeleteCustomerData/DeleteCustomerData')]/li[contains(@class,'list-group-item') and contains(@class,'bg-primary')]"));
        deletePage.click();

        // Read data from Excel
        FileInputStream fileInputStream = new FileInputStream(new File("D:\\Test.xlsx"));
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // Iterate through each row in the Excel sheet
        DataFormatter dataFormatter = new DataFormatter();
        for (Row row : sheet) {
            Cell cell = row.getCell(0); // Assuming column A contains the values
            if (cell != null && cell.getCellTypeEnum() == CellType.NUMERIC) {
                String formattedValue = dataFormatter.formatCellValue(cell);
                System.out.println("Processing value: " + formattedValue);

                WebElement inputField = driver.findElement(By.id("idamissionno"));
                inputField.clear(); // Clear the input field
                inputField.sendKeys(formattedValue); // Enter the value from Excel

                WebDriverWait wait = new WebDriverWait(driver, 1);
                

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
                    
                    wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("div.bg[style*='position: absolute']")));
                    
                    // Click the delete button
                    WebElement deleteButton = driver.findElement(By.id("btnDelete"));
                 //   wait.until(ExpectedConditions.elementToBeClickable(deleteButton));
                  //  deleteButton.click();
                   wait.until(ExpectedConditions.presenceOfElementLocated(By.id("btnDelete")));
                   
                   JavascriptExecutor js = (JavascriptExecutor) driver;
                   js.executeScript("arguments[0].click();", deleteButton);
                  
                    try {
                        // Check if the "No Data Found" alert is present
                        WebElement alertBox1 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
                        WebElement alertTitle1 = alertBox1.findElement(By.xpath(".//h2[text()='* Marked Fields Are Mandatory']"));
                         
                        if (alertTitle1.isDisplayed()) {
                            WebElement okButton = alertBox1.findElement(By.xpath(".//button[@class='confirm']"));
                            okButton.click();
                          //  continue;// Move to next iteration
                            // Added by rajmohan
                             
                          //  WebElement homeButton = driver.findElement(By.cssSelector("a.pl-2.pr-2.home-botton.bg-white"));
                           // Click the element
                          //  homeButton.click();
                            
                            
                            WebElement homeButton = driver.findElement(By.xpath("//a[@href='/Home/MenuScreen']"));
                            homeButton.click();
                            
                            // Locate and click the "Personal Details" tab
                            WebElement personalDetailsTab = wait.until(ExpectedConditions.elementToBeClickable(By.id("v-pills-01-tab")));
                            personalDetailsTab.click();
                            
                            wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("v-pills-01")));
                             
                             
                            WebElement verifyButton = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#v-pills-01 .list-group a[href='/Membership/ModifyCustomerPersonalDeatils/VerifyorEditCustomerPersonalDeatils?formid=1003&moduleid=1']")));
                            verifyButton.click();
                            
                            // Locate the input field by its ID and enter the value
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
                            
                            WebElement modifyButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("btnModify")));
                            modifyButton.click();
                            
                            WebElement confirmButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("btnSavedataFRompopup")));
                            confirmButton.click();
                            
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
                                        continue;// Move to next iteration
                                    }
                                } catch (Exception e) {
                                    // No alert found, proceed with deletion process
                                    System.out.println("No alert found, proceeding with deletion for value: " + e.getMessage());
                                }
                                
                                // Explicitly wait for the product dropdown to be clickable
                                 WebElement productDropdown1 = wait.until(ExpectedConditions.elementToBeClickable(By.id("ProductCode")));
                                 Select productSelect1 = new Select(productDropdown1);
                                 productSelect1.selectByVisibleText("Customer");
                                
                                // WebDriverWait wait1 = new WebDriverWait(driver, 1);
                                // Select "Admission No." from the task type dropdown
                                Select taskTypeDropdown1 = new Select(driver.findElement(By.id("TaskType")));
                                taskTypeDropdown1.selectByValue("1");

                                wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("div.bg[style*='position: absolute']")));

                                // Click the delete button
                                WebElement deleteButton1 = driver.findElement(By.id("btnDelete"));
                             //   wait.until(ExpectedConditions.elementToBeClickable(deleteButton));
                              //  deleteButton.click();
                                wait.until(ExpectedConditions.presenceOfElementLocated(By.id("btnDelete")));

                               JavascriptExecutor js1 = (JavascriptExecutor) driver;
                               js1.executeScript("arguments[0].click();", deleteButton1);
                             //  continue;
                             }
                            
                            
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
                    System.out.println("messageeeee--------> "+ message);
                    
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
                        }else{
                            System.out.println("Unexpected SweetAlert message: " + message);
                            break; // Exit the loop in case of unexpected message
                        }
                    }
         
            }
        }

        workbook.close();
        fileInputStream.close();
       // driver.quit(); // Close the WebDriver session
    }
}
