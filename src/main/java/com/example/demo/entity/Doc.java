package com.example.demo.entity;

import java.io.Serializable;

public class Doc implements Serializable {
    private static final long serialVersionUID = 4601357741959494322L;
    private int id;
    private String fileName;
    private String subject;
    private String submitTime;

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getSubject() {
        return subject;
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public String getSubmitTime() {
        return submitTime;
    }

    public void setSubmitTime(String submitTime) {
        this.submitTime = submitTime;
    }

    @Override
    public String toString() {
        return "Doc{" +
                "id='" + id + '\'' +
                ", fileName='" + fileName + '\'' +
                ", subject='" + subject + '\'' +
                ", submitTime='" + submitTime + '\'' +
                '}';
    }
}
