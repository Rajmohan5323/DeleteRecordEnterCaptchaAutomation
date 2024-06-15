/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.raj.deleterecordwithcaptcha;

import io.github.bonigarcia.wdm.WebDriverManager;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.time.Duration;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.StaleElementReferenceException;

/**
 *
 * @author rajmo
 */
public class DeleteRecordWithValidateCaptcha {
    
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

public static void main(String[] args) throws FileNotFoundException, IOException, InterruptedException {
        //System.out.println("Hello World!");
        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();
        
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.get("https://tn.unitetools.in/");
        
        driver.findElement(By.id("user")).sendKeys("TN12101048_de@coopsindia.com");
        driver.findElement(By.id("pwd")).sendKeys("Unite@123");
        driver.findElement(By.id("btnvalidatelogin")).click();
        // Click on 'Excel Upload' link
        WebElement excelUploadLink = driver.findElement(By.linkText("Customer Data"));
        excelUploadLink.click();
        
        WebElement deletePage = driver.findElement(By.xpath("//a[contains(@href,'/Utilities/DeleteCustomerData/DeleteCustomerData')]/li[contains(@class,'list-group-item') and contains(@class,'bg-primary')]"));
        deletePage.click();
         
        WebDriverWait wait = new WebDriverWait(driver, 10);
        
        
         // Loop to fetch and click the first row continuously
        while (true) {
            try {
                
                WebElement searchButton = driver.findElement(By.id("iconMemberSearch"));
                searchButton.click();

                WebElement viewButton = driver.findElement(By.xpath("//button[@class='btn btn-primary' and @value='View']"));
                viewButton.click();
                // Wait for the first row to be clickable and click it
                WebElement firstRow = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//table[@id='Membersearch']/tbody/tr[@class='odd']")));
                firstRow.click();
                System.out.println("Clicked on the first row");

                // Additional actions after clicking the row
                // For example, select options from dropdowns and click buttons
                WebElement productDropdown = wait.until(ExpectedConditions.elementToBeClickable(By.id("ProductCode")));
                Select productSelect = new Select(productDropdown);
                productSelect.selectByVisibleText("Customer");

                Select taskTypeDropdown = new Select(driver.findElement(By.id("TaskType")));
                taskTypeDropdown.selectByValue("1");

                WebElement deleteButton = driver.findElement(By.id("btnDelete"));
                deleteButton.click();

                // Handle captcha
                WebElement captchaImage = driver.findElement(By.id("imgcapt"));
                String captchaText = readCaptcha(captchaImage);
                
                WebElement captchaInput = driver.findElement(By.id("Captcha"));
                captchaInput.sendKeys(captchaText);
                
                
                WebElement validateButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(text(), 'Validate')]")));
                validateButton.click();
                
             //   WebElement sweetAlert = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
              //  WebElement okButton = sweetAlert.findElement(By.className("confirm"));
              //  okButton.click();
              
               WebElement sweetAlert = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sweet-alert")));
               WebElement h2Element = sweetAlert.findElement(By.tagName("h2"));
               String message = h2Element.getText().trim();
               
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
                            System.out.println("message"+message);
                        } else {
                            // Handle unexpected SweetAlert message
                            System.out.println("Unexpected SweetAlert message: " + message);
                            break; // Exit the loop in case of unexpected message
                        }
                    }
               

            } catch (StaleElementReferenceException e) {
                // Handle stale element exception by re-trying to fetch the first row
                System.out.println("StaleElementReferenceException occurred, retrying...");
            }     
                
       
        }
   }
}
