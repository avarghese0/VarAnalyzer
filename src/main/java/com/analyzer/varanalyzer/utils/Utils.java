package com.analyzer.varanalyzer.utils;

import com.analyzer.varanalyzer.dto.ExcelDataDto;
import com.analyzer.varanalyzer.dto.Stock;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

@Component
public class Utils {

    @Autowired
    Stock stock;

    @Autowired
    ExcelDataDto excelDataDto;

    @Autowired
    private ApplicationContext applicationContext;

    public static Double calculateVaR(List<Double> portfolio, double alpha) {
        Collections.sort(portfolio);
        int index = (int) Math.ceil((alpha * portfolio.size()) / 100);
        Double dollarValueAtRisk = portfolio.get(index);
        return dollarValueAtRisk;
    }

    public static List<Double> dailyReturnInPercent(List<Double> values) {

        List<Double> dailyReturnInpercent = new ArrayList<>();
        double currentPercent;

        for (int i = 1; i < values.size(); i++) {
            currentPercent = (values.get(i) - values.get(i - 1)) / values.get(i - 1);
            dailyReturnInpercent.add(currentPercent);
        }
        return dailyReturnInpercent;
    }


    public static List<Double> calcuateDailyReturnValueInDollar(List<Double> dailyReturnValue, double investAmount) {
        //calculate the daily return in dollar amounts
        return dailyReturnValue.stream().map(percentValue -> percentValue * investAmount).collect(Collectors.toList());
//        return null;
    }

    public static List<Double> calculatePortfolioValues(List<Stock> stocks) {

        int size = stocks.get(0).getDailyReturnInDollars().size();

        List<Double> result = new ArrayList<>();

        for (int i = 0; i < size; i++) {
            double sum = 0;
            for (Stock stock : stocks) {
                sum += stock.getDailyReturnInDollars().get(i);
            }
            result.add(sum);
        }


        return result;
    }

    public static ExcelDataDto excelReader(String filePath) {
        //String filePath = "src/main/resources/VARCalculator.xlsx";
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

    private static Map<String, List<String>> loadDataSheet(Sheet sheet, Map<String, List<String>> columnData) {
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

    public static List<Double> strToDouble(List<String> listOfStrings) {
        // Convert the list of strings to a list of doubles
        List<Double> listOfDoubles = new ArrayList<>();
        for (String str : listOfStrings) {
            try {
                double value = Double.parseDouble(str);
                listOfDoubles.add(value);
            } catch (NumberFormatException e) {
                // Handle parsing errors if necessary
                System.err.println("Error parsing string as double: " + str);
            }
        }
        return listOfDoubles;
    }

    public static List<String> loadCommonIndices() {
        List<String> commonIndices = new ArrayList<>();
        commonIndices.add("Date");
        commonIndices.add("Investment");
        commonIndices.add("Investment_Split");
        commonIndices.add("Confidence");
        commonIndices.add("Alpha");
        commonIndices.add("Porfolio");
        commonIndices.add("Historical VAR Index");
        commonIndices.add("Historical VAR Index Rounded");
        commonIndices.add("VAR in Dollar Amount");
        return commonIndices;
    }

    public static Double getStockInvestment(Map<String, String> configParameters, Double investment, String stockName ) {
        Double stockInvestment = 0.0;
        stockInvestment = (Double.valueOf(configParameters.get("Investment_" + stockName)) / 100) * investment;

        return stockInvestment;
    }

    public static Map<String, String> excelRowDataReader(String filepath, String sheetName) {
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
    private static String getStringCellValue(Cell cell) {
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

    public ApplicationContext getApplicationContext() {
        return applicationContext;
    }
}