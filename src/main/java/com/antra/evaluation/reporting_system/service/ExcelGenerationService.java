package com.antra.evaluation.reporting_system.service;

import com.antra.evaluation.reporting_system.pojo.api.ExcelRequest;
import com.antra.evaluation.reporting_system.pojo.api.MultiSheetExcelRequest;
import com.antra.evaluation.reporting_system.pojo.report.ExcelData;
import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;
import com.antra.evaluation.reporting_system.repo.ExcelRepository;

import java.io.File;
import java.io.IOException;

public interface ExcelGenerationService {
    File generateExcelReport(ExcelData data) throws IOException;

	ExcelFile createExcel(ExcelRequest request, ExcelRepository excelRepository) throws IOException;

	ExcelFile createMultiSheetExcel(MultiSheetExcelRequest request, ExcelRepository excelRepository) throws IOException;
}
