package com.analyzer.varanalyzer.dto;

import com.analyzer.varanalyzer.utils.Utils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
public class Stock {

    private String stockName;
    private List<Double> stockValues;
    private List<Double> dailyReturnInDollars;
    private Double investAmount;
    private List<Double> dailyReturnInPercent;

    // Private constructor to create a Stock object using the builder
    private Stock(StockBuilder builder) {
        this.stockName = builder.stockName;
        this.stockValues = builder.stockValues;
        this.investAmount = builder.investAmount;
        this.dailyReturnInDollars = builder.dailyReturnInDollars;
        this.dailyReturnInPercent = builder.dailyReturnInPercent;
    }

    // Getters for the attributes

    public String getStockName() {
        return stockName;
    }

    public List<Double> getStockValues() {
        return stockValues;
    }

    public List<Double> getDailyReturnInDollars() {
        return dailyReturnInDollars;
    }

    public Double getInvestAmount() {
        return investAmount;
    }

    public List<Double> getDailyReturnInPercent() {
        return dailyReturnInPercent;
    }


    @Component
    public static class StockBuilder {
        private String stockName;
        private List<Double> stockValues;
        private List<Double> dailyReturnInDollars;
        private Double investAmount;
        private List<Double> dailyReturnInPercent;

        public StockBuilder stockName(String stockName) {
            this.stockName = stockName;
            return this;
        }

        public StockBuilder stockValues(List<Double> stockValues) {
            this.stockValues = stockValues;
            return this;
        }

        public StockBuilder dailyReturnInDollars(List<Double> dailyReturnInDollars) {
            this.dailyReturnInDollars = dailyReturnInDollars;
            return this;
        }

        public StockBuilder investAmount(Double investAmount) {
            this.investAmount = investAmount;
            return this;
        }

        public StockBuilder dailyReturnInPercent(List<Double> dailyReturnInPercent) {
            this.dailyReturnInPercent = dailyReturnInPercent;
            return this;
        }

        public void calucateDailyReturnInDollars() {
            this.dailyReturnInDollars = Utils.calcuateDailyReturnValueInDollar(this.dailyReturnInPercent, this.investAmount);
        }
        private void calucateDailyReturnPercent() {
            this.dailyReturnInPercent = Utils.dailyReturnInPercent(this.stockValues);
        }

        public Stock build() {
            if(this.investAmount != null){
                this.calucateDailyReturnPercent();
                this.calucateDailyReturnInDollars();
                return new Stock(this);
            }
            else {
                System.out.println("Investment values must exist");
                System.exit(0);
            }
            return null;
        }
    }
}
