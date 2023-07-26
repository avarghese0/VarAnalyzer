package com.analyzer.varanalyzer.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class ExcelColumnReader {
    public static void main(String[] args) {
        String filePath = "src/main/resources/VARCalculator.xlsx";

        try (FileInputStream fileInputStream = new FileInputStream(new File(filePath))) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet dataSheet = workbook.getSheet("data"); // Assuming we are reading the first sheet (index 0)
            Sheet indexSheet = workbook.getSheet("index");

            Map<String, List<String>> data = new HashMap<>();

            //data = loadData(dataSheet, data);
            data = loadData(indexSheet, data);

//            data.
            workbook.close();

            // Print column data
            for (Map.Entry<String, List<String>> entry : data.entrySet()) {
                System.out.println("Column Header: " + entry.getKey());
                System.out.println("Column Data: " + entry.getValue());
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static Map<String, List<String>> loadData(Sheet sheet, Map<String, List<String>> columnData) {
        Row headerRow = sheet.getRow(0);

        // Get column headers from the first row
        List<String> columnHeaders = new ArrayList<>();
        for (Cell cell : headerRow) {
            columnHeaders.add(cell.getStringCellValue());
        }

        // Read column data and store it in a map
        for (String header : columnHeaders) {
            List<String> dataSheet = new ArrayList<>();
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    Cell cell = row.getCell(columnHeaders.indexOf(header));
                    if(cell != null){
                        CellType cellType = cell.getCellType();

                        if (cellType == CellType.NUMERIC) {
                            if (DateUtil.isCellDateFormatted(cell))
                                dataSheet.add(cell.getDateCellValue().toString());
                            else
                                dataSheet.add(String.valueOf(cell.getNumericCellValue()));
                        } else if (cellType == CellType.STRING) {
                            dataSheet.add(cell.getStringCellValue());
                        }
                    }
                }
            }
            columnData.put(header, dataSheet);
        }
        return columnData;
    }


}

