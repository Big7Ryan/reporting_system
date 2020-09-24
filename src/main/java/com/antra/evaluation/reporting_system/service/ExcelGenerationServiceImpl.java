package com.antra.evaluation.reporting_system.service;

import com.antra.evaluation.reporting_system.pojo.api.ExcelRequest;
import com.antra.evaluation.reporting_system.pojo.api.MultiSheetExcelRequest;
import com.antra.evaluation.reporting_system.pojo.report.ExcelData;
import com.antra.evaluation.reporting_system.pojo.report.ExcelDataHeader;
import com.antra.evaluation.reporting_system.pojo.report.ExcelDataSheet;
import com.antra.evaluation.reporting_system.pojo.report.ExcelDataType;
import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;
import com.antra.evaluation.reporting_system.repo.ExcelRepository;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Data Stucture
 * data - title, generatedTime
 * - sheets
 *      -sheet1 - title (required)
 *              - headers
 *                   - name
 *                   - width
 *                   - type
 *              - dataRows
 *                   - List of objects/values
 */
@Service
public class ExcelGenerationServiceImpl implements ExcelGenerationService {

    private void validateDate(ExcelData data) {
        if (data.getSheets().size() < 1) {
            throw new RuntimeException("Excel Data Error: no sheet is defined");
        }
        for (ExcelDataSheet sheet : data.getSheets()) {
            if (StringUtils.isEmpty(sheet.getTitle())) {
                throw new RuntimeException("Excel Data Error: sheet name is missing");
            }
            if(sheet.getHeaders() != null) {
                int columns = sheet.getHeaders().size();
                for (List<Object> dataRow : sheet.getDataRows()) {
                    if (dataRow.size() != columns) {
                        throw new RuntimeException("Excel Data Error: sheet data has difference length than header number");
                    }
                }
            }
        }
    }

    @Override
    public File generateExcelReport(ExcelData data) throws IOException {
        validateDate(data);
        XSSFWorkbook workbook = new XSSFWorkbook();

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);


        for (ExcelDataSheet sheetData : data.getSheets()) {
            Sheet sheet = workbook.createSheet(sheetData.getTitle());

            Row header = sheet.createRow(0);
            List<ExcelDataHeader> headersData = sheetData.getHeaders();
            for (int i = 0; i < headersData.size(); i++) {
                ExcelDataHeader headerData = headersData.get(i);
                Cell headerCell = header.createCell(i);
                headerCell.setCellValue(headerData.getName());
                if(headerData.getWidth() > 0) sheet.setColumnWidth(i, headerData.getWidth());
                headerCell.setCellValue(headerData.getName());
                headerCell.setCellStyle(headerStyle);
            }
            var rowData = sheetData.getDataRows();
            for (int i = 0; i < rowData.size(); i++) {
                Row row = sheet.createRow(1 + i);
                var eachRow = rowData.get(i);
                for (int j = 0; j < eachRow.size(); j++) {
                    Cell cell = row.createCell(j);
//                    switch (headersData.get(j).getType()) {
//                        case STRING:cell.setCellValue(String.valueOf(eachRow.get(j))); cell.setCellType(CellType.STRING);break;
//                        case NUMBER: cell.setCellValue(eachRow.get(j));cell.setCellType(CellType.NUMERIC);break;
//                        case DATE:cell.setCellValue((Date)eachRow.get(j));break;
//                        default:cell.setCellValue(String.valueOf(eachRow.get(j)));break;
//                    }
                    cell.setCellValue(String.valueOf(eachRow.get(j)));
                    cell.setCellStyle(style);
                }
            }
            for (int i = 0; i < headersData.size(); i++) {
                sheet.autoSizeColumn(i);
            }
        }


        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + data.getTitle() + ".xlsx";

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new File(fileLocation);
    }

    @Override
    public ExcelFile createExcel(ExcelRequest request, ExcelRepository excelRepository) throws IOException {
        List<ExcelDataHeader> headers = createHeaders(request.getHeaders());
        List<List<Object>> dataRows = request.getData();
        ExcelDataSheet excelSheet = new ExcelDataSheet("Sheet1", headers, dataRows);
        ExcelData excelData = new ExcelData(excelRepository.createId(), LocalDateTime.now(), excelSheet);
        File file = generateExcelReport(excelData);
        ExcelFile excelFile = new ExcelFile(excelData.getTitle(), LocalDateTime.now(), file.length(), "/excel/" + excelData.getTitle() +"/content");
        excelRepository.saveFile(excelFile);
        return excelFile;
    }
    
    @Override
    public ExcelFile createMultiSheetExcel(MultiSheetExcelRequest request, ExcelRepository excelRepository) throws IOException {
        List<ExcelDataHeader> headers = createHeaders(request.getHeaders());
        List<List<Object>> dataRows = request.getData();
        
        int headerIndex = 0;
        for(ExcelDataHeader header : headers) {
        	if(header.getName().equals(request.getSplitBy())) break;
        	headerIndex++;
        }
        Map<String, List<List<Object>>> spliterMap = new HashMap<>();
        for(List<Object> dataRow : dataRows) {
        	String spliter = dataRow.get(headerIndex).toString();
        	if(!spliterMap.containsKey(spliter)) spliterMap.put(spliter, new ArrayList<>());
        	spliterMap.get(spliter).add(dataRow);
        }
        
        List<ExcelDataSheet> excelSheets = new ArrayList<>();
        for(Map.Entry<String, List<List<Object>>> entry : spliterMap.entrySet()) {
        	ExcelDataSheet excelSheet = new ExcelDataSheet(entry.getKey(), headers, entry.getValue());
        	excelSheets.add(excelSheet);
        }
        ExcelData excelData = new ExcelData(excelRepository.createId(), LocalDateTime.now(), excelSheets);
        
        File file = generateExcelReport(excelData);
    	
    	
        ExcelFile excelFile = new ExcelFile(excelData.getTitle(), LocalDateTime.now(), file.length(), "/excel/" + excelData.getTitle() +"/content");
        excelRepository.saveFile(excelFile);
        return excelFile;
    }
    
    private List<ExcelDataHeader> createHeaders(List<String> contents) {
    	List<ExcelDataHeader> headers = new ArrayList<>();
    	for(String content : contents) {
    		ExcelDataHeader header = new ExcelDataHeader(content, ExcelDataType.STRING, 64);
    		headers.add(header);
    	}
    	return headers;
    }
}
