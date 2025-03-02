package Stock;

import Utilities.ExcelUtils;

public class DataSaver {
    public static void writeStockData(String filepath,int i,String marketCaptxt,String currentPriceTxt,
                                      String high_lowTxt,String changeInPercentTxt,String bookValueTxt,
                                      String dividendYieldTxt, String roceTxt,String roeTxt,String faceValueTxt,String stock_peTxt) throws Exception{
        try {
            // Writing Data to Excel File
            ExcelUtils.setCellData(filepath, "Sheet1", i, 3, marketCaptxt);
            ExcelUtils.setCellData(filepath, "Sheet1", i, 4, currentPriceTxt);
            ExcelUtils.setCellData(filepath, "Sheet1", i, 5, high_lowTxt);
            ExcelUtils.setCellData(filepath, "Sheet1", i, 6, changeInPercentTxt);
            ExcelUtils.setCellData(filepath, "Sheet1", i, 7, bookValueTxt);
            ExcelUtils.setCellData(filepath, "Sheet1", i, 8, dividendYieldTxt);
            ExcelUtils.setCellData(filepath, "Sheet1", i, 9, roceTxt);
            ExcelUtils.setCellData(filepath, "Sheet1", i, 10, roeTxt);
            ExcelUtils.setCellData(filepath, "Sheet1", i, 11, faceValueTxt);
            ExcelUtils.setCellData(filepath, "Sheet1", i, 12, stock_peTxt);
        }catch (Exception e){
            System.out.println("Error while writing data to "+i+ e.getMessage());
        }

    }
}
