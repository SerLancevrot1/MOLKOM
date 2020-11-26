package com.molcom;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class OpeningFile {

    public XSSFSheet readXlsx(String file) {

        XSSFWorkbook myExcelBook = null;
        try {
            myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            XSSFSheet myExcelSheet = myExcelBook.getSheet("Richiesta di deposito");
            return myExcelSheet;
        }
        catch (Exception e){
            System.out.println("неверное название листа: " +  e.getMessage());
            return null;
        }


    }

    public  void readSheetHead(XSSFSheet myExcelSheet){

        // 2 первых значения
        XSSFRow rowNumOfOrder = myExcelSheet.getRow(12);
        String numOfOrder = rowNumOfOrder.getCell(4).getRawValue();

        XSSFRow rowCar = myExcelSheet.getRow(10);
        String nameOfCar = rowCar.getCell(4).getStringCellValue();

        System.out.println("1) " + numOfOrder);
        System.out.println("2) " + nameOfCar);
    }

    public  void readSheetBody(XSSFSheet myExcelSheet){
        System.out.println("3) ");

        //привязался к таблице для начала цикла
        int startBody = FindRow.findWithName(myExcelSheet, "codice merce /код товара");

        //счетчик total
        int countBattary = 0;
        int countPallet = 0;
        int rowWithTotal = 0;

        for(int i = startBody + 1; i<1000; i++ ){
            XSSFRow row = myExcelSheet.getRow(i);

            String cellA = null;
            double cellB = 0;
            double cellC = 0;
            double cellG = 0;

            //читаем колонки
                cellA = row.getCell(0).getStringCellValue();
                cellB = row.getCell(1).getNumericCellValue();
                cellC = row.getCell(2).getNumericCellValue();

            try {
                //проверка на текст паллет
                 cellG = row.getCell(6).getNumericCellValue();
            }
            catch (Exception e){
                cellG = 0;
            }

            //выход из цикла при отсутствии записей

            if (cellA == ""){
                //допустил что тотал идет через 2 пробела как в образце
                rowWithTotal = row.getRowNum() + 2;
                break;
            }

            countBattary += cellC;
            countPallet += cellG;

            //System.out.println( cellA + "   |   " + cellB + "   |   " + cellC + "   |   " + cellG);
            System.out.format("%33s%20s%10s%10s",cellA,(int)cellB,(int)cellC,(int)cellG );
            System.out.println();


        }



        //берем знасения из тотала
        XSSFRow row = myExcelSheet.getRow(rowWithTotal );
        double batteryTotal = row.getCell(2).getNumericCellValue();
        double palletTotal = row.getCell(6).getNumericCellValue();

        //сравниваем
        if(countBattary != batteryTotal){
            System.out.println("4) Колличество батарей не совпадает с итоговым("+(int)batteryTotal+"): " + countBattary);
        }
        else {
            System.out.println("4) Колличество батарей  совпадает с итоговым("+(int)batteryTotal+"): " + countBattary);
        }

        if(countPallet != palletTotal){
            System.out.println("5) Колличество паллет не совпадает с итоговым("+(int)palletTotal+"): " + countPallet);
        }
        else {
            System.out.println("5) Колличество паллет  совпадает с итоговым("+(int)palletTotal+"): " + countPallet);
        }


    }

}
