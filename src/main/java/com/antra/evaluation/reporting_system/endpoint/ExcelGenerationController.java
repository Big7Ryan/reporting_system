package com.antra.evaluation.reporting_system.endpoint;

import com.antra.evaluation.reporting_system.pojo.api.ExcelRequest;
import com.antra.evaluation.reporting_system.pojo.api.ExcelResponse;
import com.antra.evaluation.reporting_system.pojo.api.MultiSheetExcelRequest;
import com.antra.evaluation.reporting_system.pojo.report.*;
import com.antra.evaluation.reporting_system.service.ExcelService;

import io.swagger.annotations.ApiOperation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.util.FileCopyUtils;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RestController
public class ExcelGenerationController {

    private static final Logger log = LoggerFactory.getLogger(ExcelGenerationController.class);

    ExcelService excelService;

    @Autowired
    public ExcelGenerationController(ExcelService excelService) {
        this.excelService = excelService;
    }

    
    
    @PostMapping("/excel")
    @ApiOperation("Generate Excel")
    public ResponseEntity<ExcelResponse> createExcel(@RequestBody @Validated ExcelRequest request) throws IOException {
    	log.info("Start Creating an excel.");
        ExcelResponse response = new ExcelResponse();
        ExcelFile excelFile = excelService.createExcel(request);
        response.setExcelFile(excelFile);
        log.info("Excel created successfully");
        return new ResponseEntity<>(response, HttpStatus.CREATED);
    }

    @PostMapping("/excel/auto")
    @ApiOperation("Generate Multi-Sheet Excel Using Split field")
    public ResponseEntity<ExcelResponse> createMultiSheetExcel(@RequestBody @Validated MultiSheetExcelRequest request) throws IOException {
        
    	log.info("Start creating a multi-sheet excel.");
    	ExcelResponse response = new ExcelResponse();
    	ExcelFile excelFile = excelService.createMultiSheetExcel(request);
        response.setExcelFile(excelFile);
        log.info("Multi-sheet Excel created successfully");
        return new ResponseEntity<>(response, HttpStatus.CREATED);
    }
    


    @GetMapping("/excel")
    @ApiOperation("List all existing files")
    public ResponseEntity<List<ExcelResponse>> listExcels() {
    	log.info("Start listing all Excels' infomation");
    	var response = new ArrayList<ExcelResponse>();
    	List<ExcelFile> files = excelService.getExcelFiles();
    	for(ExcelFile file : files) response.add(new ExcelResponse(file));
    	log.info("List all Excels' infomation successfully. There are " + files.size() + " excel(s) in total.");
        return new ResponseEntity<>(response, HttpStatus.OK);
    }
 
    @GetMapping("/excel/{id}/content")
    @ApiOperation("Download a existing file with a given id")
    public void downloadExcel(@PathVariable String id, HttpServletResponse response) throws IOException {
    	log.info("Try to Download " + id + ".xls");
        InputStream fis = excelService.getExcelBodyById(id);
        response.setHeader("Content-Type","application/vnd.ms-excel");
        response.setHeader("Content-Disposition","attachment; filename=" + id + ".xls");
        FileCopyUtils.copy(fis, response.getOutputStream());
        log.info("Copied " + id + ".xls succussfully");
    }

    @GetMapping("/excel/all")
    @ApiOperation("Down all existring files as a zip")
    public void downloadZip(HttpServletResponse response) throws IOException {
    	log.info("Try to Download all excel file as a zip file");
        response.setHeader("Content-Type","application/vnd.ms-excel");
        response.setHeader("Content-Disposition","attachment; filename=all.zip");
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + "all.zip";
        ZipOutputStream out = new ZipOutputStream(new FileOutputStream(fileLocation));
        List<ExcelFile> files = excelService.getExcelFiles();
        System.out.println("before for");
    	for(ExcelFile file : files) {
    		System.out.println("try to find: " + file.getFileId());
    		InputStream fis = excelService.getExcelBodyById(file.getFileId());
    		ZipEntry e = new ZipEntry(file.getFileId() + ".xls");
    		out.putNextEntry(e);
    		byte[] data = fis.readAllBytes();
    		out.write(data, 0, data.length);
    		out.closeEntry();
    	}
        out.close();
        File zipFile = new File(fileLocation);
        FileCopyUtils.copy(new FileInputStream(zipFile), response.getOutputStream());
        log.info("Copied all.zip succussfully");
    }
    
    @DeleteMapping("/excel/{id}")
    @ApiOperation("Delete a excel file with a given id")
    public ResponseEntity<ExcelResponse> deleteExcel(@PathVariable String id) {
    	log.info("Try to delete " + id + ".xls");
		excelService.deleteExcelBodyById(id);
    	
    	
        var response = new ExcelResponse();
        return new ResponseEntity<>(response, HttpStatus.OK);
    }

    
    
 // Exception handling
    
    @ExceptionHandler()
    @ResponseStatus(HttpStatus.NOT_FOUND)
    public void illegalArgumentExceptionHandler(IllegalArgumentException ex) {
    	log.error(ex.toString());
    }
    
    @ExceptionHandler
    @ResponseStatus(HttpStatus.NOT_FOUND)
    public void ioExceptionHandler(IOException ex) {
    	log.error(ex.toString());
    }
    
    @ExceptionHandler
    @ResponseStatus(HttpStatus.NOT_FOUND)
    public void nullPointerExceptionHandler(NullPointerException ex) {
    	log.error(ex.toString());
    }
    
    @ExceptionHandler
    @ResponseStatus(HttpStatus.NOT_FOUND)
    public void runtimeExceptionHandler(RuntimeException ex) {
    	log.error(ex.toString());
    }
    
    @ExceptionHandler
    @ResponseStatus(HttpStatus.NOT_FOUND)
    public void generalExceptionHandler(Exception ex) {
    	log.error(ex.toString());
    }
}
// Log

// Validation


