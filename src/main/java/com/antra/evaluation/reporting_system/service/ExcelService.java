package com.antra.evaluation.reporting_system.service;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import com.antra.evaluation.reporting_system.pojo.report.ExcelData;
import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;

public interface ExcelService {
    InputStream getExcelBodyById(String id);
    File generateExcelReport(ExcelData data) throws IOException;
    List<ExcelFile> getExcelFiles();
    void saveExcelFiles(ExcelFile file);
	String createId();
	void deleteExcelBodyById(String id); 
}
