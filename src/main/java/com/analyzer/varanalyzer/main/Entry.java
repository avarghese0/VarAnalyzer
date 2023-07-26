package com.analyzer.varanalyzer.main;

import com.analyzer.varanalyzer.dto.ExcelDataDto;
import com.analyzer.varanalyzer.dto.Stock;
import com.analyzer.varanalyzer.utils.ExcelUtils;
import com.analyzer.varanalyzer.utils.Utils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;

@Component
public class Entry {

    @Autowired
    Stock stock;
    @Autowired
    ExcelUtils excelUtils;

    public void excecute() {
        /**
         * Reading the data file
         * input data file contains two tabs - one for data and one for indexes
         * data file should be of xlsx format
         */
        String filePath = "src/main/resources/VARCalculator.xlsx";
        ExcelDataDto excelDataDto = excelUtils.excelReader(filePath);

        /**
         * Total Investment Amount
         */
        double totalInvestmentAmount = Double.valueOf(excelDataDto.getConfigParams().get("Investment"));

        /**
         * Load confidence level and calculate alpha
         * confidence value is loaded from the index sheet in the data file
         */
        double confidenceLevel = Double.valueOf(excelDataDto.getConfigParams().get("Confidence"));
        double alpha = 100 - confidenceLevel;

        List<Stock> listOfStocks = new ArrayList<>();

        /**
         * Build stock dto objects and calculate daily return investment in percent and in dollars
         * Stock builder uses builder design pattern to build stock objects.
         * The stock dto used to build stock objects are available in the dto package.
         *
         *  Parameters required to build stock object - stock name, stock value, investment amount
         *  stock level Investment amount is being calculated from the total investment amount and
         *  stock split values read from the index sheet of the data file.
         *
         *  Then the stock builder.build() method will creat stock objects which are needed for var calculation.
         *  Then these stocks are added to the stock list of stocks.
         */
        for (String stockName : excelDataDto.getStockData().keySet()) {

            List<Double> stockValue = Utils.strToDouble((excelDataDto.getStockData().get(stockName)));

            Double stockInvestment = Utils.getStockInvestment(excelDataDto.getConfigParams(), totalInvestmentAmount, stockName);

            Stock stock = new Stock.StockBuilder()
                    .stockName(stockName)
                    .stockValues(stockValue)
                    .investAmount(stockInvestment)
                    .build();

            listOfStocks.add(stock);
        }

        /**
         * Add stock values across stock indices based on date to generate portfolio.
         */
        List<Double> portfolios = Utils.calculatePortfolioValues(listOfStocks);

        /**
         * Calculate the Value at risk using alpha, index and portfolios generated above.
         */
        double valueAtRisk = Utils.calculateVaR(portfolios, alpha);

        System.out.println("VaR value " + valueAtRisk);

        /**
         * Write results in the results sheet the same data file.
         */
        excelUtils.stockExcelWriter(filePath, listOfStocks, portfolios, alpha, valueAtRisk);
    }
}