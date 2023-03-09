package org.example;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("表格一");
        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cellA1 = row.createCell(j);
                cellA1.setCellValue("生成" + i + "行" + j + "列");
            }
        }
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
                } catch(IllegalArgumentException iae) {
                    sheet.setColumnWidth(columnIndex, currentColumnWidth);
                }
            }
        }

        try{
            FileOutputStream outputStream = new FileOutputStream("測試.xlsx");
            wb.write(outputStream);
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
