/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.raj.deleterecordwithcaptcha;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.*;

import static org.apache.poi.ss.usermodel.CellType.BOOLEAN;
import static org.apache.poi.ss.usermodel.CellType.FORMULA;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.*;
import org.apache.poi.ss.usermodel.DataFormatter;
/**
 *
 * @author rajmo
 */
public class Excel {
    
    public static List<String> getRow(int index, Sheet sheet) {
        List<String> list = new ArrayList<>();

        for (Cell cell : sheet.getRow(index)) {
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    list.add(cell.getStringCellValue());
                    break;
                case NUMERIC:
                    list.add(String.valueOf(cell.getNumericCellValue()));
                    break;
                case BOOLEAN:
                    list.add(String.valueOf(cell.getBooleanCellValue()));
                    break;
                case FORMULA:
                    list.add(cell.getCellFormula().toString());
                    break;
            }
        }

        return list;
    }
    
    public static void main(String[] args) throws FileNotFoundException, IOException {
        FileInputStream fileInputStream = new FileInputStream(new File("D:\\Test.xlsx"));
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);
       // List<String> headerList = getRow(0, sheet);
        
        int columnIndex = 0; // Column A
        int lastRowNum = sheet.getLastRowNum();
         DataFormatter dataFormatter = new DataFormatter();
        for (int i = 0; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellTypeEnum() == CellType.NUMERIC) {
                    String formattedValue = dataFormatter.formatCellValue(cell);
                    System.out.println(formattedValue);
                }
            }
            workbook.close();
            fileInputStream.close();
    }
}
