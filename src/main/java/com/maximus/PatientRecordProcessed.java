package com.maximus;

import com.maximus.utilities.Driver;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class PatientRecordProcessed {

    private static final File sourceFile= new File("D:\\xlsx\\DataGen\\PatientRecord_01.xlsx"); // original
    private static final File processedFile= new File("D:\\xlsx\\DataGen\\ProcessedPatientRecord_01.xlsx"); // processed

    public static void main(String[] args) throws IOException {

        // For reading, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisReading = new FileInputStream(sourceFile);
        XSSFWorkbook workbookR = new XSSFWorkbook(fisReading);
        XSSFSheet sheetR = workbookR.getSheetAt(0);

        // For writing, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisWriting = new FileInputStream(processedFile);
        Workbook workbookW = new XSSFWorkbook (fisWriting);
        Sheet sheetW = workbookW.getSheetAt(0);

        System.out.println("============== Table Begin ===================");
        // row numbers of the sheet:
        int rowNums= sheetR.getLastRowNum(); // for real job
        //int rowNums=2; //0,1, ... // for testing only
        System.out.println("rowNums = " + rowNums);

        // column numbers of the sheet:
        int colNums=sheetR.getRow(0).getLastCellNum(); // 0,1,2, ...,9 // sheet colNums=10 :::
        System.out.println("colNums = " + colNums);

        // change two columns of the sourFile


        fisWriting.close();
        FileOutputStream fos =new FileOutputStream(sourceFile);
        workbookW.write(fos);
        fos.close();
        System.out.println("Done: values are written in "+processedFile);

        fisReading.close();
        Driver.closeDriver();

    }

}
