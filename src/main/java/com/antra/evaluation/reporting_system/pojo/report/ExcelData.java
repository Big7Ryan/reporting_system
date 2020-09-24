package com.antra.evaluation.reporting_system.pojo.report;

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

public class ExcelData {
    private String title;
    private LocalDateTime generatedTime;
    private List<ExcelDataSheet> sheets;

    public ExcelData() {}
    
    /**
     * 
     * @param title
     * @param generatedTime
     * @param sheet
     */
    public ExcelData(String title, LocalDateTime generatedTime, ExcelDataSheet sheet) {
    	setTitle(title);
    	setGeneratedTime(generatedTime);
    	sheets = new ArrayList<>();
    	getSheets().add(sheet);
    }
    
    
    /**
     * 
     * @param title
     * @param generatedTime
     * @param sheets
     */
    public ExcelData(String title, LocalDateTime generatedTime, List<ExcelDataSheet> sheets) {
    	setTitle(title);
    	setGeneratedTime(generatedTime);
    	setSheets(sheets);
    }
    
    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public LocalDateTime getGeneratedTime() {
        return generatedTime;
    }

    public void setGeneratedTime(LocalDateTime generatedTime) {
        this.generatedTime = generatedTime;
    }

    public List<ExcelDataSheet> getSheets() {
        return sheets;
    }

    public void setSheets(List<ExcelDataSheet> sheets) {
        this.sheets = sheets;
    }
}
