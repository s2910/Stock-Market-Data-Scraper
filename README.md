# Stock Market Data Scraper

## Overview

This project is a **Stock Market Data Scraper** that automatically extracts financial data from [Screener.in](https://www.screener.in/) using **Selenium WebDriver** and saves it into an Excel file. The script updates the data in the existing file and also creates a backup with a timestamp for record-keeping.

## Features

- **Automated stock data extraction** from Screener.in
- **Excel file handling**: Read & update stock data
- **Headless browser mode** for fast execution
- **Handles missing data** gracefully

## Technologies Used

- **Java** (JDK 11+)
- **Selenium WebDriver**
- **Apache POI** (for Excel file operations)
- **Maven** (for dependency management)

## Prerequisites

Make sure you have the following installed:

- **Java JDK 11+**
- **Maven**
- **Google Chrome**
- **ChromeDriver** (Ensure it's compatible with your Chrome version)

## Installation & Setup

### 1. Clone the Repository

```sh
git clone https://github.com/your-username/StockMarketScraper.git
cd StockMarketScraper
```

### 2. Install Dependencies

```sh
mvn clean install
```

### 3. Configure ChromeDriver Path (If Needed)

Ensure that the **ChromeDriver** is correctly set up in the system `PATH`. Alternatively, update the path in `Scraper.java`:

```java
System.setProperty("webdriver.chrome.driver", "path/to/chromedriver");
```

### 4. Run the Scraper

```sh
mvn exec:java -Dexec.mainClass="Stock.Scraper"
```

## File Structure

```
📦 StockMarketScraper
├── 📂 src
│   ├── 📂 main
│   │   ├── 📂 java
│   │   │   ├── 📂 Stock
│   │   │   │   ├── Scraper.java  # Main scraper script
│   │   │   │   ├── DataSaver.java  # Writes data to Excel
│   │   │   ├── 📂 Utilities
│   │   │   │   ├── ExcelUtils.java  # Handles Excel read/write operations
├── 📂 ExcelFiles(MarketData)
│   ├── StocksData.xlsx  # Main data file  
├── pom.xml  # Maven dependencies
└── README.md  # Project documentation
```

## How It Works

1. Reads stock names from `StocksData.xlsx` (Column C)
2. Searches each stock on Screener.in
3. Extracts data points such as:
   - Market Cap
   - Current Price
   - High/Low
   - Change (%)
   - Book Value, Dividend Yield, ROCE, ROE, Face Value, Stock P/E
4. Updates `StocksData.xlsx` with new values

## Contributing

Feel free to open an **issue** or **pull request** if you’d like to improve this project.

## License

This project is licensed under the **MIT License**.

