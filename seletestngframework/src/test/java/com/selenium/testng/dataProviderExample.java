package com.selenium.testng;

import java.io.File;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook; 

@Test
public class dataProviderExample {
	public static WebDriver driver;
	private String baseUrl;
	
    @BeforeSuite	
    public void setUp() throws Exception {
    	System.setProperty("webdriver.chrome.driver", "C:\\Users\\infosys\\Downloads\\chromedriver_win32\\chromedriver.exe");
    	driver = new ChromeDriver();
    	//driver = new FirefoxDriver();
    	driver.manage().deleteAllCookies();
    	driver.manage().window().maximize();
      baseUrl = "http://www.imdb.com/";
      driver.get(baseUrl + "/");
      driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
    }
    
    @DataProvider(name = "DP1")
    public Object[][] createData1() throws Exception{
        Object[][] retObjArr=getTableArray("test\\Resources\\Data\\data1.xlsx",
                "DataPool", "imdbTestData1");
        return(retObjArr);
    }
    
    public void testDataProviderExample(String movieTitle, 
            String directorName, String moviePlot, String actorName) throws Exception {    
        //enter the movie title 
    	driver.findElement(By.id("q")).sendKeys(movieTitle);
    	
        //they keep switching the go button to keep the bots away
    	if(driver.findElement(By.id("navbar-submit-button")) != null)
    		driver.findElement(By.id("navbar-submit-button")).click();
        else
        	driver.findElement(By.xpath("/html/body/div[1]/div/div[1]/div[2]/form/div/button")).click();
    	driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
    	//click on the movie title in the search result page        
    	driver.findElement(By.xpath("xpath=/descendant::a[text()='"+movieTitle+"']"));
    	driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
    	//verify director name is present in the movie details page
    	String bodyText = driver.findElement(By.tagName("body")).getText();
    	Assert.assertTrue(bodyText.contains(directorName), "Text not found!");
    	//verify movie plot is present in the movie details page
    	Assert.assertTrue(driver.getPageSource().contains(moviePlot));
    	//verify movie actor name is present in the movie details page
    	Assert.assertTrue(driver.getPageSource().contains(actorName));
    }
    
   @AfterSuite
    public void tearDown(){
    	driver.close();
    } 
    
    public String[][] getTableArray(String xlFilePath, String sheetName, String tableName) throws Exception{
        String[][] tabArray=null;
        
            Workbook workbook = Workbook.getWorkbook(new File(xlFilePath));
            Sheet sheet = workbook.getSheet(sheetName); 
            int startRow,startCol, endRow, endCol,ci,cj;
            Cell tableStart=sheet.findCell(tableName);
            startRow=tableStart.getRow();
            startCol=tableStart.getColumn();

            Cell tableEnd= sheet.findCell(tableName, startCol+1,startRow+1, 100, 64000,  false);                

            endRow=tableEnd.getRow();
            endCol=tableEnd.getColumn();
            System.out.println("startRow="+startRow+", endRow="+endRow+", " +
                    "startCol="+startCol+", endCol="+endCol);
            tabArray=new String[endRow-startRow-1][endCol-startCol-1];
            ci=0;

            for (int i=startRow+1;i<endRow;i++,ci++){
                cj=0;
                for (int j=startCol+1;j<endCol;j++,cj++){
                    tabArray[ci][cj]=sheet.getCell(j,i).getContents();
                }
            }
        

        return(tabArray);
    }
    
    
}//end of class