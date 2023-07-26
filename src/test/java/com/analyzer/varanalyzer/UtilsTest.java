package com.analyzer.varanalyzer;

import com.analyzer.varanalyzer.dto.Stock;
import com.analyzer.varanalyzer.utils.Utils;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.*;

import static org.junit.jupiter.api.Assertions.assertEquals;

@SpringBootTest
/**
 * @author avarghese
 * Holds the unit test cases for the var calculation logic.
 */
public class UtilsTest {

    @Test
    public void testCalculateVaR_SortedPortfolio() {
        // Given
        List<Double> portfolio = Arrays.asList(1000.0, 1500.0, 2000.0, 2500.0, 3000.0);
        Double alpha = 5.0; // 5% confidence level

        // When

        Double calculatedVaR = Utils.calculateVaR(portfolio, alpha);

        // Then
        assertEquals(1500.0, calculatedVaR, 0.001);
    }

    @Test
    public void testCalculateVaR_UnsortedPortfolio() {
        // Given
        List<Double> portfolio = Arrays.asList(2000.0, 2500.0, 1000.0, 3000.0, 1500.0);
        Double alpha = 5.0; // 5% confidence level

        // When
        Double calculatedVaR = Utils.calculateVaR(portfolio, alpha);

        // Then
        assertEquals(1500.0, calculatedVaR, 0.001);
    }

    @Test
    public void testDailyReturnInPercent_PositiveValues() {
        // Given
        List<Double> values = Arrays.asList(100.0, 110.0, 120.0, 105.0, 108.0);
        List<Double> expectedDailyReturn = Arrays.asList(0.1, 0.0909, -0.125, 0.0286);

        // When
        List<Double> calculatedDailyReturn = Utils.dailyReturnInPercent(values);

        // Then
        assertEquals(expectedDailyReturn, calculatedDailyReturn);
    }

    @Test
    public void testDailyReturnInPercent_NegativeValues() {
        // Given
        List<Double> values = Arrays.asList(20.0, 15.0, 10.0, 8.0, 5.0);
        List<Double> expectedDailyReturn = Arrays.asList(-0.25, -0.3333, -0.2, -0.375);

        // When
        List<Double> calculatedDailyReturn = Utils.dailyReturnInPercent(values);

        // Then
        assertEquals(expectedDailyReturn, calculatedDailyReturn);
    }

    @Test
    public void testDailyReturnInPercent_EmptyList() {
        // Given
        List<Double> values = Arrays.asList();
        List<Double> expectedDailyReturn = Arrays.asList();

        // When
        List<Double> calculatedDailyReturn = Utils.dailyReturnInPercent(values);

        // Then
        assertEquals(expectedDailyReturn, calculatedDailyReturn);
    }

    @Test
    public void testDailyReturnInPercent_SingleValue() {
        // Given
        List<Double> values = Arrays.asList(100.0);
        List<Double> expectedDailyReturn = Arrays.asList();

        // When
        List<Double> calculatedDailyReturn = Utils.dailyReturnInPercent(values);

        // Then
        assertEquals(expectedDailyReturn, calculatedDailyReturn);
    }

    @Test
    public void testCalcuateDailyReturnValueInDollar_PositiveValues() {
        // Given
        List<Double> dailyReturnValue = Arrays.asList(0.1, 0.09, 0.08, 0.07, 0.06);
        double investAmount = 1000.0;

        // When
        List<Double> calculatedDailyReturnInDollar = Utils.calcuateDailyReturnValueInDollar(dailyReturnValue, investAmount);

        // Then
        List<Double> expectedDailyReturnInDollar = Arrays.asList(100.0, 90.0, 80.0, 70.0, 60.0);
        assertEquals(expectedDailyReturnInDollar, calculatedDailyReturnInDollar);
    }

    @Test
    public void testCalcuateDailyReturnValueInDollar_NegativeValues() {
        // Given
        List<Double> dailyReturnValue = Arrays.asList(-0.25, -0.3333, -0.2, -0.375, -0.375);
        double investAmount = 500.0;

        // When
        List<Double> calculatedDailyReturnInDollar = Utils.calcuateDailyReturnValueInDollar(dailyReturnValue, investAmount);

        // Then
        List<Double> expectedDailyReturnInDollar = Arrays.asList(-125.0, -166.65, -100.0, -187.5, -187.5);
        assertEquals(expectedDailyReturnInDollar, calculatedDailyReturnInDollar);
    }

    @Test
    public void testCalcuateDailyReturnValueInDollar_EmptyList() {
        // Given
        List<Double> dailyReturnValue = Arrays.asList();
        double investAmount = 1000.0;

        // When
        List<Double> calculatedDailyReturnInDollar = Utils.calcuateDailyReturnValueInDollar(dailyReturnValue, investAmount);

        // Then
        List<Double> expectedDailyReturnInDollar = Arrays.asList();
        assertEquals(expectedDailyReturnInDollar, calculatedDailyReturnInDollar);
    }

    @Test
    public void testCalcuateDailyReturnValueInDollar_ZeroInvestAmount() {
        // Given
        List<Double> dailyReturnValue = Arrays.asList(0.1, 0.09, 0.08, 0.07, 0.06);
        double investAmount = 0.0;

        // When
        List<Double> calculatedDailyReturnInDollar = Utils.calcuateDailyReturnValueInDollar(dailyReturnValue, investAmount);

        // Then
        List<Double> expectedDailyReturnInDollar = Arrays.asList(0.0, 0.0, 0.0, 0.0, 0.0);
        assertEquals(expectedDailyReturnInDollar, calculatedDailyReturnInDollar);
    }

    //    @Test
//    public void testCalculatePortfolioValues() {
//        // Given
//        Stock stock1 = new Stock(Arrays.asList(100.0, 110.0, 120.0, 105.0, 108.0));
//        Stock stock2 = new Stock(Arrays.asList(50.0, 60.0, 55.0, 65.0, 70.0));
//        Stock stock3 = new Stock(Arrays.asList(75.0, 70.0, 80.0, 85.0, 90.0));
//        List<Stock> stocks = Arrays.asList(stock1, stock2, stock3);
//
//        // When
//        List<Double> calculatedPortfolioValues = Utils.calculatePortfolioValues(stocks);
//
//        // Then
//        List<Double> expectedPortfolioValues = Arrays.asList(225.0, 240.0, 255.0, 255.0, 268.0);
//        assertEquals(expectedPortfolioValues, calculatedPortfolioValues);
//    }
//
    @Test
    public void testCalculatePortfolioValues_EmptyStockList() {
        // Given
        List<Stock> stocks = new ArrayList<>();

        // When
        List<Double> calculatedPortfolioValues = Utils.calculatePortfolioValues(stocks);

        // Then
        List<Double> expectedPortfolioValues = new ArrayList<>();
        assertEquals(expectedPortfolioValues, calculatedPortfolioValues);
    }

//    @Test
//    public void testCalculatePortfolioValues_DifferentSizesOfStocks() {
//        // Given
//        Stock stock1 = new Stock(Arrays.asList(100.0, 110.0, 120.0));
//        Stock stock2 = new Stock(Arrays.asList(50.0, 60.0, 55.0, 65.0, 70.0));
//        Stock stock3 = new Stock(Arrays.asList(75.0, 70.0, 80.0, 85.0));
//        List<Stock> stocks = Arrays.asList(stock1, stock2, stock3);
//
//        // When
//        List<Double> calculatedPortfolioValues = Utils.calculatePortfolioValues(stocks);
//
//        // Then
//        List<Double> expectedPortfolioValues = Arrays.asList(225.0, 240.0, 255.0);
//        assertEquals(expectedPortfolioValues, calculatedPortfolioValues);
//    }

    @Test
    public void testStrToDouble_ValidStrings() {
        // Given
        List<String> listOfStrings = Arrays.asList("10.5", "20.25", "30.75", "40.0");

        // When
        List<Double> convertedDoubles = Utils.strToDouble(listOfStrings);

        // Then
        List<Double> expectedDoubles = Arrays.asList(10.5, 20.25, 30.75, 40.0);
        assertEquals(expectedDoubles, convertedDoubles);
    }

    @Test
    public void testStrToDouble_InvalidStrings() {
        // Given
        List<String> listOfStrings = Arrays.asList("10.5", "abc", "30.75", "xyz");

        // When
        List<Double> convertedDoubles = Utils.strToDouble(listOfStrings);

        // Then
        List<Double> expectedDoubles = Arrays.asList(10.5, 30.75);
        assertEquals(expectedDoubles, convertedDoubles);
    }

    @Test
    public void testStrToDouble_EmptyList() {
        // Given
        List<String> listOfStrings = new ArrayList<>();

        // When
        List<Double> convertedDoubles = Utils.strToDouble(listOfStrings);

        // Then
        List<Double> expectedDoubles = new ArrayList<>();
        assertEquals(expectedDoubles, convertedDoubles);
    }

    @Test
    public void testStrToDouble_NullStrings() {
        // Given
        List<String> listOfStrings = Arrays.asList("10.5", null, "30.75", null);

        // When
        List<Double> convertedDoubles = Utils.strToDouble(listOfStrings);

        // Then
        List<Double> expectedDoubles = Arrays.asList(10.5, 30.75);
        assertEquals(expectedDoubles, convertedDoubles);
    }

    @Test
    public void testGetStockInvestment_ValidStockName() {
        // Given
        Map<String, String> configParameters = new HashMap<>();
        configParameters.put("Investment_Stock1", "30.0");
        configParameters.put("Investment_Stock2", "20.0");
        Double investment = 1000.0;
        String stockName = "Stock1";

        // When
        Double stockInvestment = Utils.getStockInvestment(configParameters, investment, stockName);

        // Then
        Double expectedInvestment = 0.3 * investment;
        assertEquals(expectedInvestment, stockInvestment, 0.001);
    }

    @Test
    public void testGetStockInvestment_InvalidStockName() {
        // Given
        Map<String, String> configParameters = new HashMap<>();
        configParameters.put("Investment_Stock1", "30.0");
        configParameters.put("Investment_Stock2", "20.0");
        Double investment = 1000.0;
        String stockName = "Stock3"; // Stock3 does not exist in the configParameters

        // When
        Double stockInvestment = Utils.getStockInvestment(configParameters, investment, stockName);

        // Then
        Double expectedInvestment = 0.0; // Invalid stockName should result in 0 investment
        assertEquals(expectedInvestment, stockInvestment, 0.001);
    }

    @Test
    public void testGetStockInvestment_ZeroInvestment() {
        // Given
        Map<String, String> configParameters = new HashMap<>();
        configParameters.put("Investment_Stock1", "30.0");
        configParameters.put("Investment_Stock2", "20.0");
        Double investment = 0.0;
        String stockName = "Stock1";

        // When
        Double stockInvestment = Utils.getStockInvestment(configParameters, investment, stockName);

        // Then
        Double expectedInvestment = 0.0; // Zero investment should result in 0 stockInvestment
        assertEquals(expectedInvestment, stockInvestment, 0.001);
    }

    @Test
    public void testGetStockInvestment_MissingConfigParameter() {
        // Given
        Map<String, String> configParameters = new HashMap<>();
        configParameters.put("Investment_Stock1", "30.0");
        Double investment = 1000.0;
        String stockName = "Stock2"; // Stock2 exists in configParameters, but not Stock1

        // When
        Double stockInvestment = Utils.getStockInvestment(configParameters, investment, stockName);

        // Then
        Double expectedInvestment = 0.0; // Missing configParameter should result in 0 stockInvestment
        assertEquals(expectedInvestment, stockInvestment, 0.001);
    }
}
