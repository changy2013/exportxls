package com.ctrip.test.ch;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.concurrent.atomic.AtomicStampedReference;

/**
 * @ Author : yangchang@ctrip.com
 * @ Desc ：
 * @ Date : Created in 2018/1/23 11:11
 * @ Modified By ：
 */
public class excelTransfer {

    public static void main(String[] args) throws IOException {
        File file = new File("d:/seo.xls");
        InputStream inputStream = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

        HSSFSheet sheet1 = workbook.getSheetAt(2);
        int rows = sheet1.getPhysicalNumberOfRows();

        for (int i=1; i < rows; i++) {
            HSSFRow row = sheet1.getRow(i);
            String url = row.getCell(0).getStringCellValue();
            String name = row.getCell(1).getStringCellValue();
            StringBuilder sb = new StringBuilder("<a href=\"").append(url).append("\">");
            sb.append(name).append("</a>");
            if (i != rows-1) {
                sb.append(" | ");
            }
            System.out.print(sb.toString());
        }
    }
}
