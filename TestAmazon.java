package Amazon;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class TestAmazon {

	public static void main(String[] args) throws InterruptedException, IOException {

		System.setProperty("webdriver.chrome.driver","C:\\geckodriver\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get("https://www.amazon.in");
		WebElement search,grocery,kitchen,apparel,mob;
		search = driver.findElement(By.id("twotabsearchtextbox"));
		String currentWindow= driver.getWindowHandle(); 
		//Grocery
		search.sendKeys("Almonds");
		search.sendKeys(Keys.ENTER);
		grocery = driver.findElement(By.xpath("//*[@id='search']/div[1]/div[2]/div/span[4]/div[1]/div[1]/div/span/div/div/div/div/span/a/div/img"));
		grocery.click();
		for (String popUpHandle : driver.getWindowHandles()) {  
	        if(popUpHandle.equalsIgnoreCase(currentWindow))
	            continue;
	        driver.switchTo().window(popUpHandle);
	        String sTitle = driver.getTitle();
		}
		
		driver.findElement(By.xpath("//*[@id='submit.add-to-cart']")).click();
	
		//kitchen
		search = driver.findElement(By.id("twotabsearchtextbox"));
		search.sendKeys("chef knife");
		search.sendKeys(Keys.ENTER);
		kitchen = driver.findElement(By.xpath("//*[@id='search']/div[1]/div[2]/div/span[6]/span/div[1]/div[2]/div/span/div/div/span/a/div/img"));
		kitchen.click();
		for (String popUpHandle : driver.getWindowHandles()) {  
	        if(popUpHandle.equalsIgnoreCase(currentWindow))
	            continue;
	        driver.switchTo().window(popUpHandle);
	        String sTitle = driver.getTitle();
		}
		
		driver.findElement(By.xpath("//*[@id='submit.add-to-cart']")).click();
		
		//apparel
		search = driver.findElement(By.id("twotabsearchtextbox"));
		search.sendKeys("nike sneekers for men");
		search.sendKeys(Keys.ENTER);
		apparel = driver.findElement(By.xpath("//*[@id='anonCarousel1']/ol/li[1]/div/div/span/a/div/img"));
		apparel.click();
		for (String popUpHandle : driver.getWindowHandles()) {  
	        if(popUpHandle.equalsIgnoreCase(currentWindow))
	            continue;
	        driver.switchTo().window(popUpHandle);
	        String sTitle = driver.getTitle();
		}
		
		driver.findElement(By.xpath("//*[@id='submit.add-to-cart']")).click();
		
		//mobile
		
		search = driver.findElement(By.id("twotabsearchtextbox"));
		search.sendKeys("one plus 7t");
		search.sendKeys(Keys.ENTER);
		mob = driver.findElement(By.xpath("//*[@id='search']/div[1]/div[2]/div/span[4]/div[1]/div[1]/div/span/div/div/div[2]/div/span/a/div/img"));
		mob.click();
		for (String popUpHandle : driver.getWindowHandles()) {  
	        if(popUpHandle.equalsIgnoreCase(currentWindow))
	            continue;
	        driver.switchTo().window(popUpHandle);
	        String sTitle = driver.getTitle();
		}
		
		driver.findElement(By.xpath("//*[@id='submit.add-to-cart']")).click();
		
		//view cart
//		WebElement cart = driver.findElement(By.xpath("//*[@id='nav-cart']"));
//		cart.click();

		driver.switchTo().window(currentWindow);
		WebElement cart=driver.findElement(By.id("nav-cart"));
		cart.click();
		 java.util.List<WebElement> allitems = driver.findElements(By.className("sc-product-title"));
         int RowCount = allitems.size();
         java.util.List<WebElement> allprice = driver.findElements(By.className("sc-product-price"));
         ArrayList<String> products = new ArrayList<String>();
         ArrayList<String> price =  new ArrayList<String>();
        for(int i=0;i<RowCount;i++)
        {
        	WebElement w1=allitems.get(i);
        	WebElement w2=allprice.get(i);
        	products.add(w1.getText());
        	price.add(w2.getText());
        	System.out.println(w1.getText());
        	System.out.println(w2.getText());
        }
        WebElement total = driver.findElement(By.className("sc-price"));
        System.out.println(total.getText());
		products.add("total");
		price.add(total.getText());
        
        //***************adding to excel sheet*******
		
		 XSSFWorkbook workbook = new XSSFWorkbook();
	        XSSFSheet sheet = workbook.createSheet("products");
	         for(int i=0;i<RowCount+1;i++)
	         {
	        	 Row row = sheet.createRow(i+1);
	        	 for(int j=1;j<=2;j++)
	        	 {
	        		 Cell cell = row.createCell(j);
	        		 if(j==1)
	        		 cell.setCellValue(products.get(i));
	        		 if(j==2)
	        			 cell.setCellValue(price.get(i));
	        	 }
	         }
	         
	        try (FileOutputStream outputStream = new FileOutputStream("D://products.xlsx")) {
	            workbook.write(outputStream);
	        }
	}

}
