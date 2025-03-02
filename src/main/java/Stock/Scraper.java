package Stock;

import Utilities.ExcelUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import java.time.Duration;


public class Scraper {
    public static void main(String[] args) throws Exception{

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--headless=new");

        WebDriver driver = new ChromeDriver(options);
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

        driver.get("https://www.screener.in/");
        driver.manage().window().maximize();

        String filepath = System.getProperty("user.dir")+"\\ExcelFiles(MarketData)\\StocksData.xlsx";

        int rows = ExcelUtils.getRowCount(filepath,"Sheet1");

        for (int i = 1; i<=rows; i++) {
            String marketCaptxt = null;
            String currentPriceTxt = null;
            String high_lowTxt = null;
            String changeInPercentTxt = null;
            String bookValueTxt = null;
            String dividendYieldTxt = null;
            String roceTxt = null;
            String roeTxt = null;
            String faceValueTxt = null;
            String stock_peTxt = null;
            try {
                //Reading Stocks Name from Excel file
                String stocksSymbols = ExcelUtils.getCellData(filepath, "Sheet1", i, 2);
                System.out.println("Fetching Data for stock: " + stocksSymbols);

                //Searching for stock in the website
                Thread.sleep(1000);
                WebElement searchBox = driver.findElement(By.xpath("//input[@autofocus='true']"));
                searchBox.clear();
                searchBox.sendKeys(stocksSymbols);
                searchBox.sendKeys(Keys.ENTER);


                //Scraping Data from Website
                //MarketCap
                WebElement marketCap = driver.findElement(By.xpath("//*[@id='top-ratios']/li[1]/span[2]"));
                marketCaptxt = marketCap.getText();
                System.out.println("Market Cap: " + marketCaptxt);

                //Current Price of the Stock
                WebElement currentPrice = driver.findElement(By.xpath("//*[@id='top-ratios']/li[2]/span[2]"));
                currentPriceTxt = currentPrice.getText();
                System.out.println("Current Price: " + currentPriceTxt);

                //High/low prices of the stock
                WebElement high_low = driver.findElement(By.xpath("//*[@id='top-ratios']/li[3]/span[2]"));
                high_lowTxt = high_low.getText();
                System.out.println("High/Low: " + high_lowTxt);


                //Change(%) of stock
                WebElement changeInPercent = driver.findElement(By.xpath("//*[@id='top']/div[1]/div/div[2]/div[1]/span[2]"));
                changeInPercentTxt = changeInPercent.getText();
                System.out.println("Change(%): " + changeInPercentTxt);

                //Book Value of stock
                WebElement bookValue = driver.findElement(By.xpath("//*[@id='top-ratios']/li[5]/span[2]"));
                bookValueTxt = bookValue.getText();
                System.out.println("Book Value: " + bookValueTxt);

                //dividendYield of stock
                WebElement dividendYield = driver.findElement(By.xpath("//*[@id='top-ratios']/li[6]/span[2]"));
                dividendYieldTxt = dividendYield.getText();
                System.out.println("Dividend Yield: " + dividendYieldTxt);

                //roce of stock
                WebElement roce = driver.findElement(By.xpath("//*[@id='top-ratios']/li[7]/span[2]"));
                roceTxt = roce.getText();
                System.out.println("ROCE: " + roceTxt);

                //roe of stock
                WebElement roe = driver.findElement(By.xpath("//*[@id='top-ratios']/li[8]/span[2]"));
                roeTxt = roe.getText();
                System.out.println("ROE: " + roeTxt);

                //faceValue of stock
                WebElement faceValue = driver.findElement(By.xpath("//*[@id='top-ratios']/li[9]/span[2]"));
                faceValueTxt = faceValue.getText();
                System.out.println("Face Value: " + faceValueTxt);

                //stock_pe of stock
                WebElement stock_pe = driver.findElement(By.xpath("//*[@id='top-ratios']/li[4]/span[2]"));
                stock_peTxt = stock_pe.getText();
                System.out.println("Stock P/E: " + stock_peTxt);

                System.out.println("------------------------------------------------------------------------------------------------------------");

                Thread.sleep(2000);

                //Back To HomePage
                driver.get("https://www.screener.in/");

            } catch (Exception e) {
                System.out.println("Error in row " + i + ": " + e.getMessage());
            }

            //Writing Data Into Excel
            DataSaver.writeStockData(filepath,i,marketCaptxt,currentPriceTxt,high_lowTxt,changeInPercentTxt,bookValueTxt,dividendYieldTxt,roceTxt,roeTxt,faceValueTxt,stock_peTxt);
        }

        driver.quit();
    }
}
