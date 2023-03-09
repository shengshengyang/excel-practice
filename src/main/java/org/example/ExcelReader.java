package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {

    public static void main(String[] args) {

        try {
            // Load the Excel file
            File file = new File("workbook.xlsx");
            FileInputStream fis = new FileInputStream(file);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate over each row in the sheet starting from the second row
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                if (row != null) {
                    Cell cell1 = row.getCell(0);
                    Cell cell2 = row.getCell(1);

                    if (cell1 != null && cell2 != null) {
                        String name = cell1.getStringCellValue();
                        int age = (int) cell2.getNumericCellValue();

                        System.out.println("Name: " + name + ", Age: " + age);
                    }
                }
            }

            // Close the input stream and workbook
            fis.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
