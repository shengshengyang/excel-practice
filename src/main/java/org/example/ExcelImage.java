package org.example;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 讀取png後寫入第一格
 */
public class ExcelImage {
    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        // Create a row for the logo
        XSSFRow logoRow = sheet.createRow(0);

        try {
            // Load the logo image file
            byte[] logoBytes = IOUtils.toByteArray(new FileInputStream("psvm.png"));

            // Add the logo to the sheet
            int logoIndex = workbook.addPicture(logoBytes, Workbook.PICTURE_TYPE_PNG);
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 1, 1);
            drawing.createPicture(anchor, logoIndex);

            // Write the workbook to a file
            FileOutputStream fos = new FileOutputStream("workbook.xlsx");
            workbook.write(fos);

            // Close the output stream and workbook
            fos.close();
            workbook.close();
        }catch (IOException e){
            e.printStackTrace();
        }
    }
}
