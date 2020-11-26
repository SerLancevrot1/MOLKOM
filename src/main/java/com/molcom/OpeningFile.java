package com.molcom;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class OpeningFile {

    public XSSFSheet readXlsx(String file) throws IOException {

        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("Richiesta di deposito");
        return myExcelSheet;
    }

    public  void readSheetHead(XSSFSheet myExcelSheet){
// proverka na 0
        XSSFRow rowNumOfOrder = myExcelSheet.getRow(12);
        String numOfOrder = rowNumOfOrder.getCell(4).getRawValue();

        XSSFRow rowCar = myExcelSheet.getRow(10);
        String nameOfCar = rowCar.getCell(4).getStringCellValue();

        System.out.println("1) " + numOfOrder);
        System.out.println("2) " + nameOfCar);
    }

    public  void readSheetBody(XSSFSheet myExcelSheet){
// proverka na 0

        int countBattary = 0;
        int countPallet = 0;
        System.out.println("3) ");
        for(int i = 22; i<31; i++ ){
            XSSFRow row = myExcelSheet.getRow(i);
            String cellA = row.getCell(0).getStringCellValue();
            String cellB = row.getCell(1).getRawValue();
            String cellC = row.getCell(2).getRawValue();
            String cellG = row.getCell(6).getRawValue();

            Integer c = Integer.valueOf(cellC);
            Integer g = Integer.valueOf(cellG);
            countBattary += c;
            countPallet += g;

            System.out.println( cellA + "   |   " + cellB + "   |   " + cellC + "   |   " + cellG);


        }

        if(countBattary != 238){
            System.out.println("4) Колличество батарей не совпадает с итоговым(238): " + countBattary);
        }
        else {
            System.out.println("4) Колличество батарей  совпадает с итоговым(238): " + countBattary);
        }

        if(countPallet != 24){
            System.out.println("5) Колличество паллет не совпадает с итоговым(24): " + countPallet);
        }
        else {
            System.out.println("5) Колличество паллет  совпадает с итоговым(24): " + countPallet);
        }


    }

}
