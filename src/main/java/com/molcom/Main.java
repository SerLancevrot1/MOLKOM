package com.molcom;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        OpeningFile openingFile = new OpeningFile();
        String file = "C:\\Users\\max\\Desktop\\Заказ 2200000483.xlsx";

        XSSFSheet myExcelSheet = openingFile.readXlsx(file);
        openingFile.readSheetHead(myExcelSheet);
        openingFile.readSheetBody(myExcelSheet);




    }
}
