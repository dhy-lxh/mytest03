package com.itheima.health.text;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiTest {

    //@Test
    public void poiTest() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\qq404\\Desktop\\基金投入.xlsx");
        XSSFSheet sheetAt = workbook.getSheetAt(0);
        for (Row row : sheetAt) {
            for (Cell cell : row) {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                String value = cell.getStringCellValue();
                System.out.println(value);
            }
        }
        workbook.close();
    }

    //@Test
    public void poiTest_lastRow() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook("c:/Users/qq404/Desktop/模版计划11.xlsx");
        XSSFSheet xssfSheet = workbook.getSheetAt(0);
        int lastRowNum = xssfSheet.getLastRowNum();
        for (int i = 0; i < lastRowNum; i++) {
            XSSFRow row = xssfSheet.getRow(i);
            short lastCellNum = row.getLastCellNum();
            for (short j = 0; j < lastCellNum; j++){
                row.getCell(j).setCellType(Cell.CELL_TYPE_STRING);
                String value = row.getCell(j).getStringCellValue();
                System.out.println(value);
            }
        }
        workbook.close();
    }

    //@Test
    public void writeTest() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("传智播客");
        XSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("编号");
        row.createCell(1).setCellValue("姓名");
        row.createCell(2).setCellValue("年龄");

        XSSFRow row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("1");
        row1.createCell(1).setCellValue("小明");
        row1.createCell(2).setCellValue("18");

        XSSFRow row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("2");
        row2.createCell(1).setCellValue("小红");
        row2.createCell(2).setCellValue("22");

        FileOutputStream fileOutputStream = new FileOutputStream("c:/Users/qq404/Desktop/传智健康.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();
        workbook.close();
    }
}
