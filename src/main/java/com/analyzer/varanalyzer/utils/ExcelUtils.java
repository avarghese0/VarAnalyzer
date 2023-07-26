package com.analyzer.varanalyzer.utils;

import com.analyzer.varanalyzer.dto.ExcelDataDto;
import com.analyzer.varanalyzer.dto.Stock;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author avarghese
 * Holds all excel manipulation methods.
 */
@Component
public class ExcelUtils {

    @Autowired
    Stock stock;
    @Autowired
    ExcelDataDto excelDataDto;

    public ExcelDataDto excelReader(String filePath) {
        Map<String, List<String>> data = new HashMap<>();
        Map<String, String> commonIndices = new HashMap<>();
        ExcelDataDto excelDataDto = new ExcelDataDto();

        try (FileInputStream fileInputStream = new FileInputStream(new File(filePath))) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet dataSheet = workbook.getSheet("data"); // Assuming we are reading the first sheet (index 0)
            Sheet indexSheet = workbook.getSheet("index");


            data = loadDataSheet(dataSheet, data);
            commonIndices = excelRowDataReader(filePath, "index");
            excelDataDto.setStockData(data);
            excelDataDto.setConfigParams(commonIndices);

            workbook.close();

            // Print column data
            for (Map.Entry<String, List<String>> entry : data.entrySet()) {
                System.out.println("Column Header: " + entry.getKey());
                System.out.println("Column Data: " + entry.getValue());
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return excelDataDto;
    }

    private Map<String, List<String>> loadDataSheet(Sheet sheet, Map<String, List<String>> columnData) {
        Row headerRow = sheet.getRow(0);

        // Get column headers from the first row
        List<String> columnHeaders = new ArrayList<>();
        for (Cell cell : headerRow) {
            columnHeaders.add(cell.getStringCellValue());
        }

        // Read column data and store it in a map
        for (String header : columnHeaders) {
            if (!header.equalsIgnoreCase("date")) {
                List<String> dataSheet = new ArrayList<>();
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        Cell cell = row.getCell(columnHeaders.indexOf(header));
                        if (cell != null) {
                            CellType cellType = cell.getCellType();

                            if (cellType == CellType.NUMERIC) {
                                if (DateUtil.isCellDateFormatted(cell))
                                    dataSheet.add(cell.getDateCellValue().toString());
                                else
                                    dataSheet.add(String.valueOf(cell.getNumericCellValue()));
                            } else if (cellType == CellType.STRING) {
                                if (header.equals("Investment_Split")) {
                                    for (String split : cell.getStringCellValue().split(":")) {
                                        dataSheet.add(split);
                                    }
                                } else
                                    dataSheet.add(cell.getStringCellValue());
                            }
                        }
                    }
                }
                columnData.put(header, dataSheet);
            }
        }
        return columnData;
    }

    public Map<String, String> excelRowDataReader(String filepath, String sheetName) {
        Map<String, String> keyValueData = new HashMap<>();

        try (FileInputStream fileInputStream = new FileInputStream(new File(filepath))) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheet(sheetName); // Assuming we are reading the first sheet (index 0)

            // Create a map to store key-value data

            for (Row row : sheet) {
                // Get the cell values from the first two columns
                Cell keyCell = row.getCell(0);
                Cell valueCell = row.getCell(1);
                keyCell.setCellType(CellType.STRING);
                valueCell.setCellType(CellType.STRING);

                CellType cellType = valueCell.getCellType();
                if (keyCell != null && valueCell != null) {
                    String key = keyCell.getStringCellValue();
                    String value = valueCell.getStringCellValue();

                    // Add the key-value pair to the map
                    keyValueData.put(key, value);
                }
            }

            // Print the key-value data
            for (Map.Entry<String, String> entry : keyValueData.entrySet()) {
                System.out.println("Key: " + entry.getKey() + ", Value: " + entry.getValue());
            }

            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return keyValueData;
    }
    private String getStringCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else if (cell.getCellType() == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        }
        return "";
    }

    public void stockExcelWriter(String filePath, List<Stock> stocks, List<Double> portfolios, Double alpha, double valueAtRisk) {
        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fileInputStream)) {

            int sheetIndex = workbook.getSheetIndex("results");
            if (sheetIndex != -1) {
                workbook.removeSheetAt(sheetIndex);
            }
            Sheet resultSheet = workbook.createSheet("results");

            int columnCounter = 0;
            for(Stock stock : stocks) {
                // Create the header row
                setListData(resultSheet, stock.getStockName() + " ReturnInPercent", stock.getDailyReturnInPercent(), columnCounter++);
                setListData(resultSheet, stock.getStockName() + " ReturnInDollars", stock.getDailyReturnInDollars(), columnCounter++);

            }

            setListData(resultSheet, "Portfolio Total", portfolios, columnCounter++);

            setStringData(resultSheet, "Historical VAR Index", String.valueOf((int) Math.ceil((alpha * portfolios.size()) / 100)), columnCounter++);
            setStringData(resultSheet, "VAR in Dollar Amount", String.valueOf(valueAtRisk), columnCounter++);
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }

            System.out.println("Data written to Excel file successfully!");

            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void setListData(Sheet resultSheet, String columnHeader, List<Double> columnData, int columnCounter) {

        int rowIndex = 0;

        Row headerRow = resultSheet.getRow(rowIndex);
        if (headerRow == null) {
            headerRow = resultSheet.createRow(rowIndex);
        }
        headerRow.createCell(columnCounter).setCellValue(columnHeader);

        rowIndex++;

        rowIndex = 1;
        for (Double value : columnData) {
            Row dataRow = resultSheet.getRow(rowIndex);
            if (dataRow == null) {
                dataRow = resultSheet.createRow(rowIndex);
            }
            dataRow.createCell(columnCounter).setCellValue(value);
            rowIndex++;
        }
    }

    private void setStringData(Sheet resultSheet, String columnHeader, String columnData, int columnCounter) {

        int rowIndex = 0;

        Row headerRow = resultSheet.getRow(rowIndex);
        if (headerRow == null) {
            headerRow = resultSheet.createRow(rowIndex);
        }
        headerRow.createCell(columnCounter).setCellValue(columnHeader);

        rowIndex++;

        rowIndex = 1;

        Row dataRow = resultSheet.getRow(rowIndex);
        if (dataRow == null) {
            dataRow = resultSheet.createRow(rowIndex);
        }
        dataRow.createCell(columnCounter).setCellValue(columnData);
        rowIndex++;

    }
}