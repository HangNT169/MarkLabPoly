package com.poly.hangnt169.constant;


import com.poly.hangnt169.util.PropertiesReader;

public enum StatusCode {

    SUCCESS(0, "Success"),
    ERROR_UNKNOWN(1, "Error Unknown"),
    FILE_EMPTY(401, PropertiesReader.getProperty(PropertyKeys.FILE_EMPTY));
    private Integer status;

    private String message;

    StatusCode(Integer status, String message) {
        this.status = status;
        this.message = message;
    }

    public Integer getStatus() {
        return status;
    }

    public void setStatus(Integer status) {
        this.status = status;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

}
