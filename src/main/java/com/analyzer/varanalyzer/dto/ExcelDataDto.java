package com.analyzer.varanalyzer.dto;

import org.springframework.stereotype.Component;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
@Component
public class ExcelDataDto {
    Map<String, List<String>> stockData;

    Map<String, String> configParams;

    public ExcelDataDto() {

    }

    public ExcelDataDto(Map<String, List<String>> stockData, Map<String, String> configParams) {
        this.stockData = stockData;
        this.configParams = configParams;
    }

    public Map<String, List<String>> getStockData() {
        return stockData;
    }

    public void setStockData(Map<String, List<String>> stockData) {
        this.stockData = stockData;
    }

    public Map<String, String> getConfigParams() {
        return configParams;
    }

    public void setConfigParams(Map<String, String> configParams) {
        this.configParams = configParams;
    }
    public ExcelDataDto getMockData(){

        Map<String, List<String>> stockData = new HashMap<>();

        stockData.put("VFINX", Arrays.asList("195.2", "194.64" ,"195.45" , "195.01" , "194.53", "193.54", "193.83", "192.05", "194.25"));

        stockData.put("KELLOGG", Arrays.asList("195.2", "194.64" ,"195.45" , "195.01" , "194.53", "193.54", "193.83", "192.05", "194.25"));


        Map<String, List<String>> configParams = new HashMap<>();

        //configParams.put("Investment", "1000000");

        return null;
    }
}
