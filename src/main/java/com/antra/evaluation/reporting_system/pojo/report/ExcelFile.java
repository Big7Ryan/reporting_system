package com.antra.evaluation.reporting_system.pojo.report;

import java.io.File;
import java.time.LocalDateTime;

public class ExcelFile{
	
	private String fileId;
	private LocalDateTime generatedTime;
	private long size;
	private String downloadLink;
	
	public ExcelFile() {}
	
	public ExcelFile(String fileId) {
		setFileId(fileId);
		setGeneratedTime(LocalDateTime.now());
		setSize(128);
		setDownloadLink("");
	}
	
	public ExcelFile(String fileId, LocalDateTime generatedTime, long size, String downloadLink) {
		super();
		this.fileId = fileId;
		this.generatedTime = generatedTime;
		this.size = size;
		this.downloadLink = downloadLink;
	}

	public String getFileId() {
		return fileId;
	}
	public void setFileId(String fileId) {
		this.fileId = fileId;
	}
	public LocalDateTime getGeneratedTime() {
		return generatedTime;
	}
	public void setGeneratedTime(LocalDateTime generatedTime) {
		this.generatedTime = generatedTime;
	}
	public long getSize() {
		return size;
	}
	public void setSize(int size) {
		this.size = size;
	}
	public String getDownloadLink() {
		return downloadLink;
	}
	public void setDownloadLink(String downloadLink) {
		this.downloadLink = downloadLink;
	}
}
