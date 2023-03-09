package org.example;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelExporter {
    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("My Sheet");

        // Create some sample data
        Row row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("Name");
        row1.createCell(1).setCellValue("Age");

        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("John");
        row2.createCell(1).setCellValue(30);

        Row row3 = sheet.createRow(2);
        row3.createCell(0).setCellValue("Jane");
        row3.createCell(1).setCellValue(25);

        if (sheet.getPhysicalNumberOfRows() > 0) {
            Row row = sheet.getRow(sheet.getFirstRowNum());
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                int columnIndex = cell.getColumnIndex();
                sheet.autoSizeColumn(columnIndex);
                int currentColumnWidth = sheet.getColumnWidth(columnIndex);
                try {
                    sheet.setColumnWidth(columnIndex, (currentColumnWidth + 500));
                } catch (IllegalArgumentException iae) {
                    sheet.setColumnWidth(columnIndex, currentColumnWidth);
                }
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream("workbook.xlsx");
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
