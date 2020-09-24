package com.antra.evaluation.reporting_system.repo;

import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;
import org.springframework.stereotype.Repository;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.LongAdder;

@Repository
public class ExcelRepositoryImpl implements ExcelRepository {

    Map<String, ExcelFile> excelDataMap = new ConcurrentHashMap<>();

    LongAdder count = new LongAdder();
    
    @Override
    public Optional<ExcelFile> getFileById(String id) {
        return Optional.ofNullable(excelDataMap.get(id));
    }

    @Override
    public ExcelFile saveFile(ExcelFile file) {
    	excelDataMap.put(file.getFileId(), file);
        return null;
    }

    @Override
    public ExcelFile deleteFileById(String id) {
    	excelDataMap.remove(id);
        return null;
    }

    @Override
    public List<ExcelFile> getFiles() {
        List<ExcelFile> files = new ArrayList<>();
        for(Map.Entry<String, ExcelFile> entry : excelDataMap.entrySet()) files.add(entry.getValue());
        return files;
    }
    
    @Override
    public String createId() {
    	count.add(1);
    	return count.toString();
    }

}

