package com.analyzer.varanalyzer;

import com.analyzer.varanalyzer.main.Entry;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;


/**
 * @author avarghese
 * Main class to calculate value at risk based on historical data.
 */
@SpringBootApplication
public class VarAnalyzerApplication implements CommandLineRunner {

    @Autowired
    Entry entry;

    public static void main(String[] args) {

        SpringApplication.run(VarAnalyzerApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
        entry.excecute();
    }
}

