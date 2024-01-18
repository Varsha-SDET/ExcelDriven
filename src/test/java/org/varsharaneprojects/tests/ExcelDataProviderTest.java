package org.varsharaneprojects.tests;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
/*
 *@className : ExcelDataProviderTest
 *@Objective : to run the test case 3 times with different test data driven from excel sheet (main test)
 */
public class ExcelDataProviderTest {
    DataFormatter dataFormatter = new DataFormatter(); //for storing in multidimention array

    @Test(dataProvider = "driveTest")
    public void testCaseData(String greeting,String communication, String id){
  //  will run testcase 1 time by taking only 1 set of data from [][]array
        System.out.println(greeting + " "+ communication + " " + id);
    }

//    @DataProvider(name = "driveTest")
//    public Object[][] getData() throws IOException {
//    Object[][] data ={{"hello","text","1"},{"bye","message","143"},{"solo","call","456"}};
//    return data;
    //same logic by excel sheet (every row should be send as 1 array)


    @DataProvider(name = "driveTest")
    public Object[][] getData() throws IOException {
        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\varsh\\OneDrive\\Desktop\\Documents\\excelDriven.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = wb.getSheetAt(0);
        int rowCount =sheet.getPhysicalNumberOfRows();
        XSSFRow row =sheet.getRow(0);
        int colCount = row.getLastCellNum();
        //creating multidimentional array (memory allocation)
        Object data[][] = new Object[rowCount-1][colCount];
        for (int i=0;i<rowCount-1;i++){
           // System.out.println("Outer loop started:------------");
            //-1 becoz 1st row is neglacted
            row = sheet.getRow(i+1);

            for (int j=0;j<colCount;j++){
                XSSFCell cell = row.getCell(j);
                //format cell value into string and store in [][]
                data[i][j]=dataFormatter.formatCellValue(cell);
               // System.out.println("Outer loop ended:------------");

            }
        }

        return data;
    }
}
