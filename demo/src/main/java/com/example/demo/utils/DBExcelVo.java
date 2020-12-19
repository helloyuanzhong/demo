package com.example.demo.utils;

import java.io.Serializable;
import java.util.*;

public class DBExcelVo implements Serializable {
  private static final long serialVersionUID = 9157693679135715144L;
  private String fileName;
  private String tableName;
  private Integer maxColumnNum = -1;
  private List<ExcelVo> data = new ArrayList();
  private Set<Integer> skip = new HashSet();
  private Map<Integer, Map<String, String>> commentMap = new HashMap();

  public DBExcelVo() {}

  public String getFileName() {
    return this.fileName;
  }

  public void setFileName(String fileName) {
    this.fileName = fileName;
  }

  public String getTableName() {
    return this.tableName;
  }

  public void setTableName(String tableName) {
    this.tableName = tableName;
  }

  public Integer getMaxColumnNum() {
    return this.maxColumnNum;
  }

  public void setMaxColumnNum(Integer maxColumnNum) {
    this.maxColumnNum = maxColumnNum;
  }

  public List<ExcelVo> getData() {
    return this.data;
  }

  public void setData(List<ExcelVo> data) {
    this.data = data;
  }

  public Set<Integer> getSkip() {
    return this.skip;
  }

  public void setSkip(Set<Integer> skip) {
    this.skip = skip;
  }

  public Map<Integer, Map<String, String>> getCommentMap() {
    return this.commentMap;
  }

  public void setCommentMap(Map<Integer, Map<String, String>> commentMap) {
    this.commentMap = commentMap;
  }
}
