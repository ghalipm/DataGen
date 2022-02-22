package com.maximus;

import com.github.javafaker.Faker;
import com.maximus.utilities.Driver;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Locale;

public class PatientRecord {

    private static final File dataFile= new File("D:\\xlsx\\DataGen\\PatientRecord_01.xlsx"); // for writing


    public static void main(String[] args) throws IOException {

        // For writing, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisWriting = new FileInputStream(dataFile);
        Workbook workbookW = new XSSFWorkbook(fisWriting);
        Sheet sheetW = workbookW.getSheetAt(0);

        System.out.println("============== Table Begin ===================");

        String[] columnNames={"Id", "firstName", "lastName", "DoB", "profession", "Address","DrfName", "DrlName", "DayofVisit", "Disease", "Medication"};
        int colNums=columnNames.length;
        //System.out.println("colNums = " + colNums);
        int rowNums=20; //0,1, ... // for testing only
        //System.out.println("rowNums = " + rowNums);


        // for writing the column names from "columnNames" on the zero row
        Row rowW = sheetW.createRow(0);
        for (int i = 0; i < colNums; i++) {
            Cell cellW = rowW.createCell(i);
            cellW.setCellValue(columnNames[i]);
            System.out.println("cellW["+i+"] = " + cellW);
        }

        System.out.println( "============Row zero -- Column names=============");

        // create random values for the cells of "i" row:
        for (int i = 1; i <=rowNums; i++) {

            // new column content
            Faker faker=new Faker(new Locale("en-CA"));
            /**
             * The languages supported are as follows:
             *
             * bg、ca、ca-CAT、da-DK、de、de-AT、de-CH、en、en-AU、en-au-ocker、
             * en-BORK、en-CA、en-GB、en-IND、en-MS、en-NEP、en-NG、en-NZ、en-PAK、
             * en-SG、en-UG、en-US、en-ZA、es、es-MX、fa、fi-FI、fr、he、hu、in-ID、
             * it、ja、ko、nb-NO、nl、pt、pt-BR、ru、sk、sv、sv-SE、tr、uk、vi、zh-CN、zh-TW
             * Reference: https://developpaper.com/how-to-generate-test-data-gracefully/
             */
            columnNames[0]=""+i; // Id
            columnNames[1]=faker.name().firstName(); //"firstName"
            columnNames[2]=faker.name().lastName();//"lastName"
            columnNames[3]=faker.date().birthday().toString();// "DoB"
            columnNames[4]=faker.company().profession(); //"profession"
            columnNames[5]=faker.address().fullAddress();//"Address"
            columnNames[6]=faker.name().firstName();// "DrFirstName"
            columnNames[7]=faker.name().lastName();// "DrLastName"
            columnNames[8]=faker.date().birthday(18,78).toString();// "DayofVisit"
            columnNames[9]=faker.medical().diseaseName(); //"Disease"
            columnNames[10]=faker.medical().medicineName();//"Medication"
            rowW = sheetW.createRow(i);

            // for writing, starting column zero to "colNums":
            for (int j = 0; j <colNums ; j++) {
                //create a cell to write the content
                Cell cellW = rowW.createCell(j);
                cellW.setCellValue(columnNames[j]);
                System.out.println("cellW["+j+"] = " + cellW);

            }
            System.out.println("============== End of Row " + i + " ===================");
        }

        FileOutputStream fos =new FileOutputStream(dataFile);
        workbookW.write(fos);
        fos.close();
        System.out.println("Done: values are written in "+dataFile);

        fisWriting.close();
        Driver.closeDriver();


    }

}

