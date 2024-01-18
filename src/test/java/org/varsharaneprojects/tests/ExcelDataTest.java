package org.varsharaneprojects.tests;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

/*
 *@className : ExcelDataTest
 *@Objective: The objective of class is to drive data from excel sheet (get all the purchase row values)
 */
public class ExcelDataTest {
    @Test
    public void printData() throws IOException {
        DataDriven dataDriven = new DataDriven();
        ArrayList<String> dataList = dataDriven.getDataFromExcel("Add Profile");
        System.out.println(dataList.get(0));
        System.out.println(dataList.get(1));
        System.out.println(dataList.get(2));
        System.out.println(dataList.get(3));

    }
    @Test
    public void printExcelData() throws IOException {
        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\varsh\\OneDrive\\Desktop\\Documents\\excelDriven.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = wb.getSheetAt(0);
        int rowCount =sheet.getPhysicalNumberOfRows();
        XSSFRow row =sheet.getRow(0);
        int colCount = row.getLastCellNum(); //column = last row cell count
        //creating multidimentional array (memory allocation)
        Object data[][] = new Object[rowCount-1][colCount];
        for (int i=0;i<rowCount-1;i++){
           // System.out.println("outer loop started.......");
            //-1 becoz 1st row is neglacted(header row)
            row = sheet.getRow(i+1);
            //System.out.println(row);

            for (int j=0;j<colCount;j++){

                System.out.println(row.getCell(j));
            }
          //  System.out.println("outer loop ended.......");

        }
    }

}
