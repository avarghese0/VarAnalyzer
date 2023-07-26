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


public class ExcelRowReader {
    public static void main(String[] args) {
        String filePath = "src/main/resources/VARCalculator.xlsx";

        try (FileInputStream fileInputStream = new FileInputStream(new File(filePath))) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet dataSheet = workbook.getSheet("data"); // Assuming we are reading the first sheet (index 0)
            Sheet indexSheet = workbook.getSheet("index");

            List<Map<String, String>> data = new ArrayList<>();
            List<Map<String, String>> index = new ArrayList<>();

            data = loadData(dataSheet, data);
            index = loadData(indexSheet, index);

            workbook.close();

            // Print the loaded data
            for (Map<String, String> rowData : data) {
                for (Map.Entry<String, String> entry : rowData.entrySet()) {
                    System.out.print(entry.getKey() + ": " + entry.getValue() + "\t");
                }
                System.out.println(); // Move to the next row
            }


            for (Map<String, String> rowData : index) {
                for (Map.Entry<String, String> entry : rowData.entrySet()) {
                    System.out.print(entry.getKey() + ": " + entry.getValue() + "\t");
                }
                System.out.println(); // Move to the next row
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static List<Map<String, String>> loadData(Sheet dataSheet, List<Map<String, String>> data) {
        Row headerRow = dataSheet.getRow(0);

        for (int rowIndex = 1; rowIndex <= dataSheet.getLastRowNum(); rowIndex++) {
            Row currentRow = dataSheet.getRow(rowIndex);
            Map<String, String> rowData = new HashMap<>();

            for (int columnIndex = 0; columnIndex < headerRow.getLastCellNum(); columnIndex++) {
                Cell headerCell = headerRow.getCell(columnIndex);
                Cell currentCell = currentRow.getCell(columnIndex);

                String columnHeader = headerCell.getStringCellValue();
                String cellValue = "";

                if (currentCell != null) {
                    CellType cellType = currentCell.getCellType();

                    if (cellType == CellType.NUMERIC && DateUtil.isCellDateFormatted(currentCell)) {
                        cellValue = currentCell.getDateCellValue().toString();
                    } else if (cellType == CellType.NUMERIC) {
                        cellValue = String.valueOf(currentCell.getNumericCellValue());
                    } else if (cellType == CellType.STRING) {
                        cellValue = currentCell.getStringCellValue();
                    }
                }

                rowData.put(columnHeader, cellValue);
            }

            data.add(rowData);
        }
        return data;
    }


}

