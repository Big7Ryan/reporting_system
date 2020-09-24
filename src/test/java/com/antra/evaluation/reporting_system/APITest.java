package com.antra.evaluation.reporting_system;

import com.antra.evaluation.reporting_system.endpoint.ExcelGenerationController;
import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;
import com.antra.evaluation.reporting_system.service.ExcelService;
import io.restassured.http.ContentType;
import io.restassured.module.mockmvc.RestAssuredMockMvc;
import org.hamcrest.Matchers;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.MockitoAnnotations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static io.restassured.module.mockmvc.RestAssuredMockMvc.given;
import static org.mockito.ArgumentMatchers.anyString;

public class APITest {
    @Mock
    ExcelService excelService;

    @BeforeEach
    public void configMock() {
        MockitoAnnotations.initMocks(this);
        RestAssuredMockMvc.standaloneSetup(new ExcelGenerationController(excelService));
    }

    @Test
    public void testFileDownload() throws FileNotFoundException {
    	Mockito.when(excelService.getExcelBodyById(anyString())).thenReturn(new FileInputStream("temp.xlsx"));
        given().accept("application/json").get("/excel/123/content").peek().
                then().assertThat()
                .statusCode(200);
    }

    
    @Test
    public void testListFiles() throws FileNotFoundException {
    	List<ExcelFile> excelFiles = new ArrayList<>();
    	ExcelFile file1 = new ExcelFile("1"); 
    	ExcelFile file2 = new ExcelFile("2");
    	ExcelFile file3 = new ExcelFile("3");
    	excelFiles.add(file1);
    	excelFiles.add(file2);
    	excelFiles.add(file3);
    	Mockito.when(excelService.getExcelFiles()).thenReturn(excelFiles);
        given().accept("application/json").get("/excel").peek().
                then().assertThat()
                .statusCode(200)
                .body("excelFile.fileId", Matchers.hasItems("1", "2", "3"));
    }

    @Test
    @Disabled
    public void testExcelGeneration() throws FileNotFoundException {
        // Mockito.when(excelService.getExcelBodyById(anyString())).thenReturn(new FileInputStream("temp.xlsx"));
        given().accept("application/json").contentType(ContentType.JSON).body("{\"headers\":[\"Name\",\"Age\"], \"data\":[[\"Teresa\",\"5\"],[\"Daniel\",\"1\"]]}").post("/excel").peek().
                then().assertThat()
                .statusCode(200)
                .body("fileId", Matchers.notNullValue());
    }

    @Test
    public void testCreateExcel() throws IOException {
    	ExcelFile excelFile = new ExcelFile("1");
    	Mockito.when(excelService.createMultiSheetExcel(Mockito.any())).thenReturn(excelFile);
    	given().accept("application/json").contentType(ContentType.JSON).body("{\"headers\":[\"Name\",\"Age\"], \"data\":[[\"Teresa\",\"5\"],[\"Daniel\",\"1\"]]}").post("/excel/auto").peek().
    		then().assertThat()
    		.statusCode(201)
    		.body("excelFile.fileId", Matchers.equalTo("1"));
    }
    
    @Test
    public void testCreateMultiSheetExcel() throws IOException {
    	ExcelFile excelFile = new ExcelFile("1");
    	Mockito.when(excelService.createMultiSheetExcel(Mockito.any())).thenReturn(excelFile);
    	given().accept("application/json").contentType(ContentType.JSON).body("{\"headers\":[\"Name\",\"Age\"], \"data\":[[\"Teresa\",\"5\"],[\"Daniel\",\"1\"]], \"splitBy\":\"Age\"}}").post("/excel/auto").peek().
    		then().assertThat()
    		.statusCode(201)
    		.body("excelFile.fileId", Matchers.equalTo("1"));
    }
    
    @Test
    public void testDeleteExcel() throws IOException {
    	Mockito.doNothing().when(excelService).deleteExcelBodyById(anyString());
    	given().accept("application/json").delete("/excel/1").peek().
    		then().assertThat()
    		.statusCode(200);
    }
}
