package com.antra.evaluation.reporting_system.pojo.api;

import java.util.ArrayList;
import java.util.List;

import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;

public class ExcelResponse {
	
	private ExcelFile excelFile;

	private List<ExcelFile> excelFiles = new ArrayList<>();
	
	public ExcelResponse() {}
	
	public ExcelResponse(ExcelFile file) {
		setExcelFile(file);
	}
	
	
	public List<ExcelFile> getExcelFiles() {
		return excelFiles;
	}

	public void setExcelFiles(List<ExcelFile> excelFiles) {
		this.excelFiles = excelFiles;
	}

	public ExcelFile getExcelFile() {
		return excelFile;
	}

	public void setExcelFile(ExcelFile excelFile) {
		this.excelFile = excelFile;
	}
}
