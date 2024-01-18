package org.varsharaneprojects.tests;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

/*
 *@className : DataDriven
 *@Objective: The objective of class is to return data from excel sheet in form of arraylist(get all the purchase row values)
 */
public class DataDriven {

    public ArrayList<String> getDataFromExcel (String testCaseName ) throws IOException {
        ArrayList<String> arrayList = new ArrayList<String>();
        //fileInputStream argument
        FileInputStream fileInputStream= new FileInputStream("C://Users//varsh//OneDrive//Desktop//Documents//testDemo.xlsx");
        //get access of excel
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        //read sheet no.
        int sheetCount =workbook.getNumberOfSheets();
        for (int i=0;i<sheetCount;i++){
            //sheet  name data access
            if(workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                //identify testcases column by scanning the entire 1st row
               Iterator<Row> rows =  sheet.iterator(); //sheet is collection of rows
               Row firstRow= rows.next();
               Iterator<Cell> cell =firstRow.cellIterator(); //row is collection of cells
                int k=0;//interates after every while loop to check cell no.in row
                int columnIndex = 0;
                //iterate in cell
                while(cell.hasNext()){
                   Cell value =cell.next();
                   //desired column
                   if(value.getStringCellValue().equalsIgnoreCase("TestCases")){
                       columnIndex =k; //to get column index
                   }
                    k++; //row ++
                }
              //  System.out.println(columnIndex);

                //iterate for rows
                while (rows.hasNext()){
                    Row r = rows.next();
                    if(r.getCell(columnIndex).getStringCellValue().equalsIgnoreCase(testCaseName)){ //get access of purchase row
                        //get all the values of cells
                        Iterator<Cell> cv = r.cellIterator();
                        while (cv.hasNext()){
                            //on iterating data stored in arraylist
                            arrayList.add(cv.next().getStringCellValue());
                        }

                    }
                }
            }
        }
        return arrayList;
    }

}
