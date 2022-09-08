package com.seetech.automation.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Hashtable;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLib {
    public static String path = System.getProperty("user.dir") + File.separator + "TestData" + File.separator + "TestData.xlsx";
    public  FileInputStream fis = null;
    public  FileOutputStream fileOut =null;
    private XSSFWorkbook workbook = null;
    private XSSFSheet sheet = null;
    private XSSFRow row   =null;
    private XSSFCell cell = null;
    //String path="";

    
    public ExcelLib() {
      try {
          fis = new FileInputStream(path);
          workbook = new XSSFWorkbook(fis);
          sheet = workbook.getSheetAt(0);
          fis.close();
      } catch (Exception e) {
          e.printStackTrace();
      }
    }
    
    /**
     * @param path
     */
    public ExcelLib(String filePath) {

        try {
        	  path = filePath;
            fis = new FileInputStream(path);
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(0);
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Function to get data from xls sheet in 2 dimensional array
     *
     * @param testCase - testCase name
     * @param xls      - Xls_Reader Object
     * @return 2 dimensional array
     * @throws Throwable 
     */
    public Hashtable<String, String> getData(String testCase, String sheetName) throws Throwable {
      int testCaseStartRowNum = 0;
      Hashtable<String, String> table = null;
      // iterate through all rows from the sheet Test Data
      for (int rNum = 1; rNum <= getRowCount(sheetName); rNum++) {
        // to identify testCase starting row number
        if (testCase.equals(getCellData(sheetName, 0, rNum))) {
          testCaseStartRowNum = rNum;
          break;
        }
      }

      int colStartRowNum = testCaseStartRowNum + 1;
      int cols = 0;
      // Get the total number of columns for which test data is present
      while (!getCellData(sheetName, cols, colStartRowNum).equals("")) {
        cols++;
      }

      // rows
      int rowStartRowNum = testCaseStartRowNum + 2;
      int rows = 0;
      // Get the total number of rows for which test data is present
      while (!getCellData(sheetName, 0, (rowStartRowNum + rows)).equals("")) {
        rows++;
      }

      Object[][] data = new Object[rows][1];
      for (int rNum = rowStartRowNum; rNum < (rows + rowStartRowNum); rNum++) {
        table = new Hashtable<String, String>();
        for (int cNum = 0; cNum < cols; cNum++) {
          table.put(getCellData(sheetName, cNum, colStartRowNum), getCellData(sheetName, cNum, rNum));
        }
        data[rNum - rowStartRowNum][0] = table;
      }
      return table;
    }
    
    // returns the row count in a sheet
    public int getRowCount(String sheetName) throws Throwable{
      int index = workbook.getSheetIndex(sheetName);
      if (index == -1)
        return 0;
      else {
        sheet = workbook.getSheetAt(index);
        int number = sheet.getLastRowNum() + 1;
        return number;
      }
    }
    
    public String getCellData(String sheetName, String colName, int rowNum) {
      try {
        if (rowNum <= 0) {
          return "";
        }
        int index = workbook.getSheetIndex(sheetName);
        int col_Num = -1;
        if (index == -1)
          return "";

        sheet = workbook.getSheetAt(index);
        row = sheet.getRow(0);
        for (int i = 0; i < row.getLastCellNum(); i++) {
          if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
            col_Num = i;
        }
        if (col_Num == -1)
          return "";

        sheet = workbook.getSheetAt(index);
        row = sheet.getRow(rowNum - 1);
        
        if(row == null){
          return "";
        }
        
        cell = row.getCell(col_Num);

        if(cell == null){
          return "";
          }
        
        if(cell.getCellType() == Cell.CELL_TYPE_STRING){
          return cell.getStringCellValue();
        }else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
            String cellText = String.valueOf(cell.getNumericCellValue());
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
              double d = cell.getNumericCellValue();
              Calendar cal = Calendar.getInstance();
              cal.setTime(HSSFDateUtil.getJavaDate(d));
              cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
              cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" + cal.get(Calendar.MONTH) + 1 + "/"+ cellText;
            }
            return cellText;
      } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK){
          return "";
        }else {
          return String.valueOf(cell.getBooleanCellValue());
        }
      } catch (Exception e) {
        e.printStackTrace();
        return "row " + rowNum + " or column " + colName + " does not exist in xls";
      }
    }
    
    public String getCellData(String sheetName, int colNum, int rowNum) {
      try {
        if (rowNum <= 0)
          return "";

        int index = workbook.getSheetIndex(sheetName);

        if (index == -1)
          return "";

        sheet = workbook.getSheetAt(index);
        row = sheet.getRow(rowNum - 1);
        if (row == null)
          return "";
        cell = row.getCell(colNum);
        if (cell == null)
          return "";

        if (cell.getCellType() == Cell.CELL_TYPE_STRING)
          return cell.getStringCellValue();
        else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

          String cellText = String.valueOf(cell.getNumericCellValue());
          if (HSSFDateUtil.isCellDateFormatted(cell)) {
            // format in form of M/D/YY
            double d = cell.getNumericCellValue();
            Calendar cal = Calendar.getInstance();
            cal.setTime(HSSFDateUtil.getJavaDate(d));
            cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
            cellText = cal.get(Calendar.MONTH) + 1 + "/" +cal.get(Calendar.DAY_OF_MONTH) + "/" +cellText;
          }

          return cellText;
        } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
          return "";
        else
          return String.valueOf(cell.getBooleanCellValue());
      } catch (Exception e) {

        e.printStackTrace();
        return "row " + rowNum + " or column " + colNum + " does not exist  in xls";
      }
    }
    
    /**
     * @param arg
     * @throws IOException
     */
    public static void main(String arg[]) throws IOException{
  	ExcelLib obj = new ExcelLib();
  	//obj.getTestCaseData("");
  	try {
  	  Hashtable<String, String> testData = obj.getData("CreateAccount", "RegressionSuite");
  	  System.out.println(testData.get("Phone").toString());
  	  System.out.println(testData.get("Revenue").toString());
  	  System.out.println(testData.get("Billing_Zip").toString());
  	} catch (Throwable e) {
      e.printStackTrace();
    }
        
	} 
    
    
}
