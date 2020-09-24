package com.antra.evaluation.reporting_system.service;

import com.antra.evaluation.reporting_system.repo.ExcelRepository;
import com.antra.evaluation.reporting_system.endpoint.ExcelGenerationController;
import com.antra.evaluation.reporting_system.pojo.api.ExcelRequest;
import com.antra.evaluation.reporting_system.pojo.api.MultiSheetExcelRequest;
import com.antra.evaluation.reporting_system.pojo.report.ExcelData;
import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;

import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Optional;

@Service
public class ExcelServiceImpl implements ExcelService {

    @Autowired
    ExcelRepository excelRepository;

    @Autowired
    ExcelGenerationService excelGenerationService;
    
    @Override
    public InputStream getExcelBodyById(String id) {

        Optional<ExcelFile> fileInfo = excelRepository.getFileById(id);
       // if (fileInfo.isPresent()) {
//            File file = new File("temp.xlsx");
//            try {
//                return new FileInputStream(file);
//            } catch (FileNotFoundException e) {
//                e.printStackTrace();
//            }
      //  }
        if(fileInfo.isPresent() && !fileInfo.isEmpty()) {
        	File file = new File(fileInfo.get().getFileId() + ".xlsx");
        	try {
        		return new FileInputStream(file);
        	}
        	catch (FileNotFoundException e) {
        		LoggerFactory.getLogger(ExcelGenerationController.class).info("Cannot find the file with the given id. File may not exist.");
        		e.printStackTrace();
        	}
        }
    	LoggerFactory.getLogger(ExcelGenerationController.class).info("Cannot find the file with the given id. File may not exist.");
    	return null;
        
    }
    
    @Override
    public void deleteExcelBodyById(String id){
    	Optional<ExcelFile> fileInfo = excelRepository.getFileById(id);
    	if(fileInfo.isPresent()) {
    		File file = new File(fileInfo.get().getFileId() + ".xlsx");
    		file.delete();
    	}
    	else {
    		LoggerFactory.getLogger(ExcelGenerationController.class).info("Cannot delete the file with the given id. File may not exist.");
    	}
    	excelRepository.deleteFileById(id);
    }
    
    
    @Override
    public File generateExcelReport(ExcelData data) throws IOException {
    	return excelGenerationService.generateExcelReport(data);
    }
    
    @Override
    public void saveExcelFiles(ExcelFile file) {
    	excelRepository.saveFile(file);
    }
    
    @Override
    public List<ExcelFile> getExcelFiles() {
    	return excelRepository.getFiles();
    }
    
    @Override
    public String createId() {
    	return excelRepository.createId();
    }
    
    @Override
    public ExcelFile createExcel(ExcelRequest request) throws IOException {
    	return excelGenerationService.createExcel(request, excelRepository);
    }
    
    @Override
    public ExcelFile createMultiSheetExcel(MultiSheetExcelRequest request) throws IOException {
    	return excelGenerationService.createMultiSheetExcel(request, excelRepository);
    }

}
