package com.antra.evaluation.reporting_system.pojo.api;

import java.util.List;

public class ExcelRequest {
    private List<String> headers;
    private List<List<Object>> data;
    private String description;
    private String submitter;

    public String getDescription() {
        return description;
    }

    public List<List<Object>> getData() {
		return data;
	}

	public void setData(List<List<Object>> data) {
		this.data = data;
	}

	public String getSubmitter() {
		return submitter;
	}

	public void setSubmitter(String submitter) {
		this.submitter = submitter;
	}

	public void setDescription(String description) {
        this.description = description;
    }

    public List<String> getHeaders() {
        return headers;
    }

    public void setHeaders(List<String> headers) {
        this.headers = headers;
    }
}
