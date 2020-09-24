package com.antra.evaluation.reporting_system.pojo.report;


import java.util.List;

public class ExcelDataSheet {
    private String title;
    private List<ExcelDataHeader> headers;
    private List<List<Object>> dataRows;

    public ExcelDataSheet() {}
    
    public ExcelDataSheet(String title, List<ExcelDataHeader> headers, List<List<Object>> dataRows) {
    	setTitle(title);
    	setHeaders(headers);
    	setDataRows(dataRows);
    }
    
    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public List<ExcelDataHeader> getHeaders() {
        return headers;
    }

    public void setHeaders(List<ExcelDataHeader> headers) {
        this.headers = headers;
    }

    public List<List<Object>> getDataRows() {
        return dataRows;
    }

    public void setDataRows(List<List<Object>> dataRows) {
        this.dataRows = dataRows;
    }
}
