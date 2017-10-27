package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Trythis {

	/**
	 * @param args
	 * @throws InterruptedException 
	 * @throws IOException 
	 * @throws BiffException 
	 * @throws WriteException 
	 * @throws RowsExceededException 
	 */
	public static void main(String[] args) throws InterruptedException, IOException, BiffException, RowsExceededException, WriteException {
		// TODO Auto-generated method stub
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\1338846\\Desktop\\chromedriver_win32_2.32\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		//driver.manage().window().maximize();
		driver.get("https://www.flipkart.com/");
		Actions ac=new Actions(driver);
		WebElement we= driver.findElement(By.xpath("//*[@id='container']/div/header/div[2]/div/ul/li[7]/a/span"));
         ac.moveToElement(we).build().perform();
         Thread.sleep(2000);
        driver.findElement(By.xpath("//*[@id='container']/div/header/div[2]/div/ul/li[7]/ul/li/ul/li[5]/ul/li[4]/a/span")).click();
        Thread.sleep(2000);
        ((JavascriptExecutor)driver).executeScript("scroll(0,200)");
        driver.findElement(By.xpath("//*[@id='container']/div/div[1]/div/div[2]/div/div[1]/div/div/div[6]/section/div[2]/div[1]/div[2]/div/div/label/div[1]")).click();
        Thread.sleep(2000);
        ((JavascriptExecutor)driver).executeScript("scroll(0,400)");
        driver.findElement(By.xpath("//*[@id='container']/div/div[1]/div/div[2]/div/div[1]/div/div/div[8]/section/div[2]/div/div[2]/div/div/label/div[1]")).click();
        ((JavascriptExecutor)driver).executeScript("scroll(0,625)");
        driver.findElement(By.xpath("//*[@id='container']/div/div[1]/div/div[2]/div/div[2]/div/div[3]/div/div[2]/div[2]/div/a[2]")).click();
        Thread.sleep(2000);
        TakesScreenshot scrnshot=((TakesScreenshot)driver);
        File srcfile=scrnshot.getScreenshotAs(OutputType.FILE);
        File destfile=new File("C:\\Users\\1338846\\Desktop\\Screenshot\\newimg.jpg");
        FileUtils.copyFile(srcfile, destfile);
        Thread.sleep(2000);
        driver.navigate().to("https://www.amazon.in/");
        FileInputStream f=new FileInputStream("C:\\Users\\1338846\\Desktop\\Excelsheet\\exceltype.xls");
        Workbook wb=Workbook.getWorkbook(f);
        Sheet s=wb.getSheet("Sheet1");
        int rows=s.getRows();
        for(int i=0;i<rows;i++){
        	driver.findElement(By.xpath("//*[@id='twotabsearchtextbox']")).sendKeys(s.getCell(0, i).getContents());
        	driver.findElement(By.xpath("//*[@id='nav-search']/form/div[2]/div/input")).click();
        	driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
        	Thread.sleep(3000);
        	driver.findElement(By.xpath("//*[@id='twotabsearchtextbox']")).clear();
        }
        FileOutputStream of=new FileOutputStream("C:\\Users\\1338846\\Desktop\\Excelsheet\\outputtype.xls");
        WritableWorkbook ww= Workbook.createWorkbook(of);
        WritableSheet ws=ww.createSheet("Sheet1", 0);
        Label Content=new Label(0,0,"Apple");
        ws.addCell(Content);
        ww.write();
        ww.close();
        driver.quit();
	}

}
