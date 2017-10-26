package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class SeleniumClass {

	/**
	 * @param args
	 * @throws InterruptedException 
	 * @throws IOException 
	 */
	public static void main(String[] args) throws InterruptedException, IOException {
		// TODO Auto-generated method stub
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\1338846\\Desktop\\chromedriver_win32_2.32\\chromedriver.exe");
        WebDriver driver=new ChromeDriver();
        //driver.manage().window().maximize();
        Actions ac=new Actions(driver);
        driver.get("https://www.amazon.in/");
        WebElement we=driver.findElement(By.xpath("//*[@id='nav-xshop']/a[2]"));
        ac.moveToElement(we).click().build().perform();
        
        Thread.sleep(3000);
        driver.findElement(By.cssSelector("span.a-dropdown-prompt")).click();
        driver.findElements(By.className("a-dropdown-link")).get(2).click();
        Thread.sleep(3000);
        WebElement element = driver.findElement(By.name("field-keywords"));
        element.sendKeys("All out");
        element.submit();
       //driver.findElement(By.xpath("//*[@id='s-result-info-bar']/div[1]/div[1]/span/div/div/a/span[1]")).click();
        ((JavascriptExecutor) driver).executeScript("scroll(0,300)");
        driver.findElement(By.xpath("//*[@id='leftNavContainer']/ul[6]/div/li[1]/span/span/div/label/input")).click();
        Thread.sleep(2000);
        TakesScreenshot scrnshot=((TakesScreenshot) driver);
        File srcfile=scrnshot.getScreenshotAs(OutputType.FILE);
        File destfile= new File("C:\\Users\\1338846\\Desktop\\Screenshot\\a.jpg");
        FileUtils.copyFile(srcfile,destfile);
        
		driver.navigate().to("https://www.flipkart.in");
        FileInputStream f=new FileInputStream("C:\\Users\\1338846\\Desktop\\Excelsheet\\exceltype.xls");
        Workbook wb= new HSSFWorkbook(f);
        Sheet s=wb.getSheetAt(0);
        int rowCount = s.getLastRowNum()-s.getFirstRowNum();

        //Create a loop over all the rows of excel file to read it

        for (int i = 0; i < rowCount+1; i++) {

            Row row = s.getRow(i);

            //Create a loop to print cell values in a row

            driver.findElement(By.xpath("//*[@id='container']/div/header/div[1]/div[2]/div/div/div[2]/form/div/div[1]/div/input")).sendKeys((row.get);
			driver.findElement(By.xpath("//*[@id='container']/div/header/div[1]/div[2]/div/div/div[2]/form/div/div[2]/button")).click();
			
			driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
			Thread.sleep(3000);
			driver.findElement(By.xpath("//*[@id='container']/div/header/div[1]/div[2]/div/div/div[2]/form/div/div[1]/div/input")).clear();
            }

           
        }
        
        
       
       
      
        
	}

	

