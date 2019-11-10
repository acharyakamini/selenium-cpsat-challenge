package newAutomationCPSAT;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;
import org.testng.asserts.SoftAssert;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Tests {

    @Test
    public void testLink() throws IOException {
        SoftAssert softAssert = new SoftAssert();
        System.setProperty("webdriver.chrome.driver", "./driver/chromedriver");
        WebDriver driver = new ChromeDriver();
        driver.get("http://agiletestingalliance.org/");
        driver.findElement(By.xpath("//a[contains(text(),'Certifications')]")).click();
        List<WebElement> optionCount = driver.findElements(By.xpath("//div[@class='grid_12']//area[@target='_self']"));
        System.out.println(optionCount.size());
        try {
            for (WebElement oneEle : optionCount) {
                System.out.println(oneEle.getAttribute("href"));
                String urlLink = oneEle.getAttribute("href");

                // Getting the Response Code for URL
                URL url = new URL(urlLink);
                HttpURLConnection connection = (HttpURLConnection) url.openConnection();
                connection.setRequestMethod("GET");
                connection.connect();
                int code = connection.getResponseCode();

                // Condition to check whether the URL is valid or Invalid
                if (code == 200) System.out.println("Valid Link:" + urlLink);
                else System.out.println("INVALID Link:" + urlLink);

                //take screenshot
                TakesScreenshot ts = (TakesScreenshot) driver;
                File source = ts.getScreenshotAs(OutputType.FILE);
                FileUtils.copyFile(source, new File("./Screenshot.png"));
                System.out.println("the Screenshot is taken");
            }
        } catch (Exception e) {
            softAssert.fail(e.getMessage());
            softAssert.assertAll();
        }
        //Mouseover on submit button
        Actions action = new Actions(driver);
        WebElement btn = driver.findElement(By.xpath("//div[@class='grid_12']//area[@title='CP-CCT']"));
        action.moveToElement(btn).perform();
        //take screenshot
        TakesScreenshot ts = (TakesScreenshot) driver;
        File source = ts.getScreenshotAs(OutputType.FILE);
        FileUtils.copyFile(source, new File("./Screenshot.png"));
        System.out.println("the Screenshot of CP-CCT is taken");
    }

    @Test
    public void testMinimumCount() {
        System.setProperty("webdriver.chrome.driver", "./driver/chromedriver");
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.nseindia.com/");
        List<WebElement> optionCount = driver.findElements(By.cssSelector(".advanceTab> li > span"));
        List<Integer> values = new ArrayList<Integer>();
        System.out.println(optionCount.size());
        for (WebElement oneEle : optionCount) {
            System.out.println(oneEle.getText());
            values.add(Integer.valueOf(oneEle.getText()));
        }
        Object obj = Collections.min(values);
        System.out.println("Minimum values is " + obj);
    }

    @Test
    public void testEquityPage() throws InterruptedException, IOException {
        System.setProperty("webdriver.chrome.driver", "./driver/chromedriver");
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.nseindia.com/");
        WebElement webElement = driver.findElement(By.xpath("//input[@id='keyword']"));
        webElement.sendKeys("Eicher Motors Limited");
        Thread.sleep(1000);
        webElement.sendKeys(Keys.ENTER);
        //take screenshot
        TakesScreenshot ts = (TakesScreenshot) driver;
        File source = ts.getScreenshotAs(OutputType.FILE);
        FileUtils.copyFile(source, new File("./Screenshot.png"));
        System.out.println("the Screenshot of equity is taken");
        //fetch and print the values
        WebElement webElement1 = driver.findElement(By.xpath("//li[@id='face']"));
        System.out.println(webElement1.getText());
        WebElement webElement2 = driver.findElement(By.xpath("//span[@id='high52']"));
        System.out.println(webElement2.getText());
        WebElement webElement3 = driver.findElement(By.xpath("//span[@id='low52']"));
        System.out.println(webElement3.getText());

    }

    @Test
    public void testQuotePage() throws IOException, InterruptedException {
        System.setProperty("webdriver.chrome.driver", "./driver/chromedriver");
        WebDriver driver = new ChromeDriver();
        FileInputStream fis = new FileInputStream("./File/Company_Name.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        int rowCount = wb.getSheet("sheet1").getLastRowNum();
        for (int i = 0; i <= rowCount; i++) {
            String companyName = wb.getSheet("sheet1").getRow(i).getCell(0).toString();
            System.out.println("Company to be entered in equity search box is :------> " + companyName);
            driver.manage().window().maximize();
            driver.get("https://www.nseindia.com/");
            driver.findElement(By.xpath("//input[@id='keyword']")).sendKeys(companyName);
            Thread.sleep(3000);
            driver.findElement(By.xpath("//input[@id='keyword']")).sendKeys(Keys.ENTER);
            Thread.sleep(3000);
            //fetch and print the values
            WebElement webElement1 = driver.findElement(By.xpath("//span[@id='faceValue']"));
            System.out.println(webElement1.getText());
            WebElement webElement2 = driver.findElement(By.xpath("//span[@id='high52']"));
            System.out.println(webElement2.getText());
            WebElement webElement3 = driver.findElement(By.xpath("//span[@id='low52']"));
            System.out.println(webElement3.getText());
            //takeScreenshot
            TakesScreenshot ts = (TakesScreenshot) driver;
            File source = ts.getScreenshotAs(OutputType.FILE);
            FileUtils.copyFile(source, new File("./Screenshot.png"));
            System.out.println("the Screenshot of quote page is taken");
        }

    }

    @Test
    public void testWriteDataToExcel() throws IOException {
        System.setProperty("webdriver.chrome.driver", "./driver/chromedriver");
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.nseindia.com/products.htm");
        driver.findElement(By.xpath("//a[contains(text(),'Live Market')]")).click();
        driver.findElement(By.xpath("//a[contains(text(),'Top Ten Gainers / Losers')]")).click();
        List<WebElement> irows = driver.findElements(By.xpath("//table[@id='topGainers']/tbody//tr"));
        int iRowsCount = irows.size();
        List<WebElement> icols = driver.findElements(By.xpath("//table[@id='topGainers']/tbody/tr[1]/th"));
        int iColsCount = icols.size();
        System.out.println("Selected web table has " + iRowsCount + " Rows and " + iColsCount + " Columns");
        System.out.println();

        FileOutputStream fos = new FileOutputStream("./File/Bank_Name.xlsx");
        XSSFWorkbook wkb = new XSSFWorkbook();
        XSSFSheet sheet1 = wkb.createSheet("Gainers");

        for (int i = 1; i <= iRowsCount; i++) {
            for (int j = 1; j <= iColsCount; j++) {
                if (i == 1) {
                    WebElement val = driver.findElement(By.xpath("//table[@id='topGainers']/tbody/tr[" + i + "]/th[" + j + "]"));
                    String a = val.getText();
                    System.out.print(a);

                    XSSFRow excelRow = sheet1.createRow(i);
                    XSSFCell excelCell = excelRow.createCell(j);
                    excelCell.setCellValue(a);
                    wkb.write(fos);
                } else {
                    WebElement val = driver.findElement(By.xpath("//table[@id='topGainers']/tbody/tr[" + i + "]/td[" + j + "]"));
                    String a = val.getText();
                    System.out.print(a);

                    XSSFRow excelRow = sheet1.createRow(i);
                    XSSFCell excelCell = excelRow.createCell(j);
                    //excelCell.setCellType(XSSFCell.CELL_TYPE_STRING);
                    excelCell.setCellValue(a);
                    wkb.write(fos);
                }
            }
            System.out.println();
        }
        fos.flush();
        fos.close();
    }

    @Test
    public void testNavigationAndGetAllDropdownValues() throws InterruptedException {
        System.setProperty("webdriver.gecko.driver", "./driver/geckodriver");
        WebDriver driver = new FirefoxDriver();
        driver.get("https://www.shoppersstop.com/");
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.manage().window().maximize();
        // Forward navigation
        for(int i =0;i<5;i++)
        {
            driver.findElement(By.xpath("//div[@class='dy-container-437876 slick-initialized slick-slider slick-dotted']//div[@class='dy-slick-arrow dy-next-arrow slick-arrow']")).click();
            Thread.sleep(1000);
        }
        // back navigation
        for(int j=0;j<5;j++)
        {
            driver.findElement(By.xpath("//div[@class='dy-slick-arrow dy-prev-arrow slick-arrow']")).click();
            Thread.sleep(1000);
        }
        driver.navigate().refresh();
        //get all stores
        driver.findElement(By.xpath("//ul[contains(@class,'text-right')]//a[contains(text(),'All Stores')]")).click();
        Select dropdown = new Select(driver.findElement(By.id("city-name")));
        //Get all options
        List<WebElement> dd = dropdown.getOptions();
        //Get the length
        System.out.println(dd.size());
        // Loop to print one by one
        for (int j = 0; j < dd.size(); j++) {
            System.out.println(dd.get(j).getText());
        }
        //Mouseover on submit button
        Actions action = new Actions(driver);
        WebElement btn = driver.findElement(By.xpath("/html/body/main/nav/div[2]/div/ul/li[4]/a"));
        action.moveToElement(btn).perform();
        List<WebElement> elementCount = driver.findElements(By.xpath("//ul[contains(@class,'lvl1')]/li[4]/div[1]/div[1]/ul[1]/li[6]/div//span"));
        System.out.println(elementCount.size());
        for (WebElement oneEle : elementCount) {
            System.out.println(oneEle.getText());
        }
    }
}