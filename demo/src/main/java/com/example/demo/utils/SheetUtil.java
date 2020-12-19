package com.example.demo.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.util.*;
import java.util.regex.Pattern;

public class SheetUtil {
  public static final String SUFFIX_XLS = "xls";
  public static final String SUFFIX_XLSX = "xlsx";
  private static final Logger logger = LoggerFactory.getLogger(SheetUtil.class);

  public SheetUtil() {}

  public static void main(String[] args) throws Exception {
    sheetAllToList("c:\\2.xlsx", (Map) null);
  }

  public static Map<Integer, DBExcelVo> sheetAllToList(
      String fileName, Map<Integer, Map<String, Object>> startMap) throws Exception {
    Map<Integer, DBExcelVo> dbBExcelVoMap = new HashMap();
    Workbook wb = WorkbookFactory.create(new FileInputStream(new File(fileName)));

    for (int k = 0; k < wb.getNumberOfSheets(); ++k) {
      Integer startRow = 1;
      if (startMap != null
          && startMap.get(k) != null
          && ((Map) startMap.get(k)).get("startRow") != null) {
        startRow = (Integer) ((Map) startMap.get(k)).get("startRow");
      }

      Integer commentRow = startRow.equals(0) ? 0 : startRow - 1;
      if (startMap != null
          && startMap.get(k) != null
          && ((Map) startMap.get(k)).get("commentRow") != null) {
        commentRow = (Integer) ((Map) startMap.get(k)).get("commentRow");
      }

      Sheet sheet = wb.getSheetAt(k);
      DBExcelVo dbExcelVo =
          sheetToList((Integer) startRow, commentRow, (String) null, (String) null, (Sheet) sheet);
      dbBExcelVoMap.put(k, dbExcelVo);
    }

    return dbBExcelVoMap;
  }

  private static Map<ExcelMergedRegion, Map<String, String>> getSheetMergedRegionComment(
      Sheet sheet) {
    Map<ExcelMergedRegion, Map<String, String>> sheetMergedRegionComment = new HashMap();
    List<CellRangeAddress> cellRangeAddresses = sheet.getMergedRegions();
    Iterator var3 = cellRangeAddresses.iterator();

    while (var3.hasNext()) {
      CellRangeAddress cellRangeAddress = (CellRangeAddress) var3.next();
      Cell cell =
          sheet
              .getRow(cellRangeAddress.getFirstRow())
              .getCell(cellRangeAddress.getFirstColumn(), MissingCellPolicy.CREATE_NULL_AS_BLANK);
      Map<String, String> map = getCellComment(cell);
      if (map != null) {
        sheetMergedRegionComment.put(new ExcelMergedRegion(cellRangeAddress), map);
      }
    }

    return sheetMergedRegionComment;
  }

  public static DBExcelVo sheetToList(
      String fileName, Integer sheetNo, Integer startRow, String name, String tableName)
      throws Exception {
    return sheetToList(fileName, sheetNo, startRow, 0, name, tableName);
  }

  public static DBExcelVo sheetToList(
      String fileName,
      Integer sheetNo,
      Integer startRow,
      Integer commentRow,
      String name,
      String tableName)
      throws Exception {
    Workbook wb = WorkbookFactory.create(new FileInputStream(new File(fileName)));
    Sheet sheet = wb.getSheetAt(sheetNo);
    DBExcelVo dbExcelVo = sheetToList(startRow, commentRow, name, tableName, sheet);
    return dbExcelVo;
  }

  private static DBExcelVo sheetToList(
      Integer startRow, Integer commentRow, String name, String tableName, Sheet sheet) {
    DBExcelVo dbExcelVo = new DBExcelVo();
    if (name != null) {
      dbExcelVo.setFileName(name);
    }

    if (tableName != null) {
      dbExcelVo.setTableName(tableName);
    }

    int rowMax = sheet.getLastRowNum();

    for (int i = startRow; i < rowMax + 1; ++i) {
      int sizeTemp = sheet.getRow(i).getLastCellNum() + 0;
      if (sizeTemp > dbExcelVo.getMaxColumnNum()) {
        dbExcelVo.setMaxColumnNum(sizeTemp);
      }
    }

    Map<ExcelMergedRegion, Map<String, String>> sheetMergedRegionCommentMap =
        getSheetMergedRegionComment(sheet);
    Set<ExcelMergedRegion> excelMergedRegions = sheetMergedRegionCommentMap.keySet();

    for (int j = 0; j < commentRow + 1; ++j) {
      Row rowJ = sheet.getRow(j);
      if (rowJ != null) {
        if (rowJ.getLastCellNum() + 0 > dbExcelVo.getMaxColumnNum()) {
          dbExcelVo.setMaxColumnNum(rowJ.getLastCellNum() + 0);
        }

        for (int i = 0; i < dbExcelVo.getMaxColumnNum(); ++i) {
          Cell cell = rowJ.getCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);
          if (cell != null) {
            if (cell.getCellComment() == null) {
              Iterator var18 = excelMergedRegions.iterator();

              while (var18.hasNext()) {
                ExcelMergedRegion excelMergedRegion = (ExcelMergedRegion) var18.next();
                if (excelMergedRegion.isMergedRegion(j, i)) {
                  Map<String, String> map =
                      (Map) sheetMergedRegionCommentMap.get(excelMergedRegion);
                  if (!dbExcelVo.getCommentMap().containsKey(i)) {
                    dbExcelVo.getCommentMap().put(i, new HashMap());
                  }

                  ((Map) dbExcelVo.getCommentMap().get(i)).putAll(map);
                  break;
                }
              }
            } else {
              Map<String, String> map = getCellComment(cell);
              if (map != null) {
                if (dbExcelVo.getCommentMap().get(i) != null) {
                  ((Map) dbExcelVo.getCommentMap().get(i)).putAll(map);
                } else {
                  dbExcelVo.getCommentMap().put(i, map);
                }
              }
            }
          }
        }
      }
    }

    return dbExcelVo;
  }

  public static Map<String, String> getCellComment(Cell cell) {
    Map<String, String> map = null;
    if (cell != null && cell.getCellComment() != null) {
      map = new HashMap();
      String comment = cell.getCellComment().getString().getString();
      String[] attributes = comment.split("\n");

      for (int m = 0; m < attributes.length; ++m) {
        if (!StringUtils.isEmpty(attributes[m])) {
          String[] nameValue = attributes[m].split("=");
          map.put(nameValue[0], nameValue[1]);
        }
      }
    }

    return map;
  }

  public static DBExcelVo sheetToList(String fileName, Integer sheetNo, Integer startRow)
      throws Exception {
    return sheetToList(
        (String) fileName, sheetNo, (Integer) startRow, (String) null, (String) null);
  }

  public static void listToFile(String fileName, List<ExcelVo> excelVoList) throws Exception {
    File fileOut = new File(fileName);
    if (!fileOut.getParentFile().isDirectory()) {
      fileOut.getParentFile().mkdirs();
    }

    if (fileOut.isFile()) {
      fileOut.delete();
    }

    FileWriter fw = new FileWriter(fileOut);

    for (int i = 0; i < excelVoList.size(); ++i) {
      fw.write(((ExcelVo) excelVoList.get(i)).toString());
      fw.write(System.getProperty("line.separator"));
    }

    fw.flush();
    fw.close();
  }

  public static void listToFile(String fileName, List<ExcelVo> excelVoList, Boolean highExtension)
      throws Exception {
    Workbook wbOut = null;
    if (highExtension) {
      wbOut = new XSSFWorkbook();
    } else {
      wbOut = new HSSFWorkbook();
    }

    Sheet sheetOut = ((Workbook) wbOut).createSheet();

    for (int i = 0; i < excelVoList.size(); ++i) {
//      System.out.println(excelVoList.size() - i);
      fillStringDataToRow(i, sheetOut, (ExcelVo) excelVoList.get(i), (Workbook) wbOut);
    }

    File fileOut = new File(fileName);
    FileOutputStream fileOutputStream =
        new FileOutputStream(fileOut + (highExtension ? ".xlsx" : ".xls"));
    ((Workbook) wbOut).write(fileOutputStream);
    fileOutputStream.close();
  }

  public static ExcelVo fillStringDataToExcelVo(Row row) {
    Integer total = row.getLastCellNum() + 0;
    ExcelVo evo = new ExcelVo();
    if (total.equals(-1)) {
      total = 0;
    }

    evo.expand(total);

    for (int i = 0; i < total; ++i) {
      Cell cell = row.getCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);
      String value = null;
      value = parse(cell);
      evo.setValue(i, value == null ? null : value.trim());
    }

    System.gc();
    return evo;
  }

  public static Row fillStringDataToRow(Integer i, Sheet sheet, ExcelVo evo, Workbook wbOut) {
    CellStyle cellStyle = wbOut.createCellStyle();
    Row row = getOrCreateRow(i, sheet, evo.getSize(), cellStyle);

    for (int j = 0; j < evo.getSize(); ++j) {
      row.getCell(j).setCellValue(evo.getValue(j));
    }

    System.gc();
    return row;
  }

  public static void createMergedRegion(
      int startRow, int endRow, int startColumn, int endColumn, Sheet sheet) {
    for (int j = startColumn; j < endColumn; ++j) {
      sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, j, j));
    }
  }

  public static Row getOrCreateRow(int i, Sheet sheet, int max, CellStyle cfaf) {
    Row row = sheet.getRow(i);
    if (row == null) {
      row = sheet.createRow(i);

      for (int j = 0; j < max; ++j) {
        Cell c = row.createCell(j);
        c.setCellStyle(cfaf);
        c.setCellValue("");
      }
    }

    return row;
  }

  public static boolean ifRowNullOrEmpty(Row row) {
    return row == null || row.getLastCellNum() == 0;
  }

  public static boolean ifSheetNullOrEmpty(Sheet sheet) {
    return sheet == null || sheet.getLastRowNum() == 0;
  }

  private static String parse(Cell cell) {
    String value = null;
    if (cell != null) {
      double valueTemp;
      CellStyle style;
      DecimalFormat format;
      String temp;
      switch (cell.getCellType()) {
        case 0:
          if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
            value = DateUtil.DTF_RSJ19.print(cell.getDateCellValue().getTime());
          } else {
            valueTemp = cell.getNumericCellValue();
            style = cell.getCellStyle();
            format = new DecimalFormat();
            temp = style.getDataFormatString();
            if (temp.equals("General")) {
              format.applyPattern("#");
            }

            format.setGroupingUsed(false);
            value = format.format(valueTemp);
          }
          break;
        case 1:
          value = cell.getStringCellValue();
          break;
        case 2:
          CellValue cellValue =
              cell.getRow()
                  .getSheet()
                  .getWorkbook()
                  .getCreationHelper()
                  .createFormulaEvaluator()
                  .evaluate(cell);
          value = cell.getCellFormula();
          switch (cellValue.getCellType()) {
            case 0:
              if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                value = DateUtil.DTF_RSJ19.print(cell.getDateCellValue().getTime());
              } else {
                valueTemp = cell.getNumericCellValue();
                style = cell.getCellStyle();
                format = new DecimalFormat();
                temp = style.getDataFormatString();
                if (temp.equals("General")) {
                  format.applyPattern("#");
                }

                format.setGroupingUsed(false);
                value = format.format(valueTemp);
              }
              break;
            case 1:
              value = cell.getStringCellValue();
            case 2:
            case 3:
            default:
              break;
            case 4:
              if (cell.getBooleanCellValue()) {
                value = "1";
              } else {
                value = "0";
              }
              break;
            case 5:
              value = cell.getErrorCellValue() + "";
          }
        case 3:
        default:
          break;
        case 4:
          if (cell.getBooleanCellValue()) {
            value = "1";
          } else {
            value = "0";
          }
          break;
        case 5:
          value = cell.getErrorCellValue() + "";
      }
    }

    return value;
  }

  private static boolean validExcelFormat(String fileName, String type) {
    return valid(fileName) && getFileNameSuffix(fileName).equalsIgnoreCase(type);
  }

  private static String getFileNameSuffix(String fileName) {
    return fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
  }

  private static boolean valid(String fileName) {
    if (StringUtils.isEmpty(fileName)) {
      logger.error("文件路径有误");
      return false;
    } else {
      return true;
    }
  }

  public static Workbook getBookByPath(String fileName) {
    Object book = null;

    try {
      FileInputStream is = new FileInputStream(new File(fileName));
      if (validExcelFormat(fileName, "xls")) {
        book = new HSSFWorkbook(is);
      } else if (validExcelFormat(fileName, "xlsx")) {
        book = new XSSFWorkbook(is);
      }
    } catch (IOException var3) {
      logger.error("没有此文件或该文件有问题");
      var3.printStackTrace();
    }

    return (Workbook) book;
  }

  public static Workbook getBookByPath(InputStream is) {
    Workbook book = null;

    try {
      book = WorkbookFactory.create(is);
    } catch (Exception var3) {
      logger.error("没有此文件或该文件有问题");
      var3.printStackTrace();
    }

    return book;
  }

  public static void writeToExcel(Workbook book, String filePath) {
    OutputStream out = null;
    File file = new File(filePath);
    File parentfile = file.getParentFile();
    if (!parentfile.exists() || !parentfile.isDirectory()) {
      parentfile.mkdirs();
    }

    if (file.exists() && !file.isFile()) {
      file.delete();
    }

    label89:
    {
      try {
        out = new FileOutputStream(filePath);
        book.write(out);
        break label89;
      } catch (Exception var15) {
        var15.printStackTrace();
      } finally {
        if (out != null) {
          try {
            out.flush();
            out.close();
          } catch (IOException var14) {
            var14.printStackTrace();
          }
        }
      }

      return;
    }

    logger.info("数据导出成功");
  }

  public static String getCellValue(Sheet sheet, int rowNum, int columnNum) {
    Cell cell = org.apache.poi.ss.util.SheetUtil.getCellWithMerges(sheet, rowNum, columnNum);
    return parse(cell);
  }

  public static String getCellValue(Sheet sheet, int rowNum, String columnName) {
    int columnNum = getHeadLineIndex(sheet, columnName);
    return columnNum == -1 ? "" : getCellValue(sheet, rowNum, columnNum);
  }

  public static String getCellValue(Row row, int columnNum) {
    Cell cell = row.getCell(columnNum);
    return parse(cell);
  }

  public static String getCellValue(Row row, String columnName) {
    int columnNum = getHeadLineIndex(row.getSheet(), columnName);
    if (columnNum == -1) {
      return "";
    } else {
      Cell cell = row.getCell(columnNum);
      return parse(cell);
    }
  }

  public static String getHeadLine(Sheet sheet, int columnNum) {
    return getCellValue(sheet, 0, columnNum);
  }

  public static int getHeadLineIndex(Sheet sheet, String columnName) {
    columnName = columnName.toUpperCase().replace("_", "").trim();
    Row firstRow = sheet.getRow(0);
    int columnNum = -1;

    for (int i = 0; i <= firstRow.getLastCellNum(); ++i) {
      Cell cell = firstRow.getCell(i);
      if (cell != null) {
        String value = parse(cell).toUpperCase().replace("_", "").trim();
        if (value.equals(columnName)) {
          columnNum = i;
          break;
        }
      }
    }

    return columnNum;
  }

  public static <T> T rowToObject(Row row, Class<T> classType, List<String> excludeList) {
    Object object = null;

    try {
      object = classType.newInstance();
      Field[] fields = classType.getDeclaredFields();
      Field[] var5 = fields;
      int var6 = fields.length;

      for (int var7 = 0; var7 < var6; ++var7) {
        Field field = var5[var7];
        String fieldName = field.getName();
        if (excludeList == null || excludeList.indexOf(fieldName) == -1) {
          field.setAccessible(true);
          field.set(object, stringToOtherType(getCellValue(row, fieldName), field.getType()));
        }
      }
    } catch (Exception var10) {
      var10.printStackTrace();
    }

    return (T) object;
  }

  public static <T> List<T> sheetToObjects(
      Sheet sheet, Class<T> classType, List<String> excludeList) {
    ArrayList resultList = null;

    try {
      resultList = new ArrayList();

      for (int i = 0; i <= sheet.getLastRowNum(); ++i) {
        Row row = sheet.getRow(i);
        T object = rowToObject(row, classType, excludeList);
        resultList.add(object);
      }
    } catch (Exception var7) {
      var7.printStackTrace();
    }

    return resultList;
  }

  private static <T> T stringToOtherType(String str, Class<T> classType) {
    if (str != null && !str.equals("")) {
      T result = null;
      String dataType = classType.getSimpleName();
      byte var5 = -1;
      switch (dataType.hashCode()) {
        case -1808118735:
          if (dataType.equals("String")) {
            var5 = 0;
          }
          break;
        case 2122702:
          if (dataType.equals("Date")) {
            var5 = 4;
          }
          break;
        case 2374300:
          if (dataType.equals("Long")) {
            var5 = 1;
          }
          break;
        case 79860828:
          if (dataType.equals("Short")) {
            var5 = 3;
          }
          break;
        case 1438607953:
          if (dataType.equals("BigDecimal")) {
            var5 = 2;
          }
      }

      switch (var5) {
        case 0:
          result = (T) str;
          break;
        case 1:
          result = (T) Long.valueOf(str);
          break;
        case 2:
          result = (T) new BigDecimal(str);
          break;
        case 3:
          result = (T) Short.valueOf(str);
          break;
        case 4:
          String[] dateTime = str.split("\\s+");
          if (dateTime.length == 1) {
            result = (T) DateUtil.toDate(str);
          } else if (dateTime.length == 2) {
            try {
              result = (T) DateUtil.toDateFull(str);
            } catch (ParseException var8) {
              result = null;
              var8.printStackTrace();
            }
          }
          break;
        default:
          result = (T) str;
      }

      return result;
    } else {
      return null;
    }
  }

  private static <T> String otherTypeToString(T object) {
    String result = "";
    if (object == null) {
      return result;
    } else {
      String dataType = object.getClass().getSimpleName();
      byte var4 = -1;
      switch (dataType.hashCode()) {
        case -1808118735:
          if (dataType.equals("String")) {
            var4 = 0;
          }
          break;
        case 2122702:
          if (dataType.equals("Date")) {
            var4 = 1;
          }
      }

      switch (var4) {
        case 0:
          result = (String) object;
          break;
        case 1:
          result = DateUtil.formatChineseFull((Date) object);
          break;
        default:
          result = String.valueOf(object);
      }

      return result;
    }
  }

  public static <T> Sheet objectsToSheet(
      List<T> objects, Sheet sheet, Class<T> classType, List<String> excludeList) {
    for (int i = 0; i < objects.size(); ++i) {
      Row row = getOrCreateRow(sheet, i + 1);
      objectToRow(objects.get(i), row, classType, excludeList);
    }

    return sheet;
  }

  public static <T> Row objectToRow(
      T object, Row row, Class<T> classType, List<String> excludeList) {
    Field[] fields = classType.getDeclaredFields();

    for (int i = 0; i < fields.length; ++i) {
      Field field = fields[i];
      String fieldName = field.getName();
      if (excludeList == null || excludeList.indexOf(fieldName) == -1) {
        Cell cell = null;
        if (i != 0 && (i <= 0 || row.getCell(i - 1) == null)) {
          cell = row.createCell(i - 1);
        } else {
          cell = getOrCreateCell(row, i);
        }

        field.setAccessible(true);

        try {
          cell.setCellValue(otherTypeToString(field.get(object)));
        } catch (Exception var10) {
          var10.printStackTrace();
        }
      }
    }

    return row;
  }

  public static Row getOrCreateRow(Sheet sheet, int i) {
    Row row = sheet.getRow(i);
    if (row == null) {
      row = sheet.createRow(i);
    }

    return row;
  }

  public static Cell getOrCreateCell(Row row, int i) {
    Cell cell = row.getCell(i);
    if (cell == null) {
      cell = row.createCell(i);
    }

    return cell;
  }

  private static String formatFieldHeadLine(String str) {
    if (str != null && !str.equals("")) {
      String pattern = "[A-Z]";
      StringBuffer buffer = new StringBuffer();

      for (int i = 0; i < str.length(); ++i) {
        buffer.append(str.charAt(i));
        if (i < str.length() - 1 && Pattern.matches(pattern, String.valueOf(str.charAt(i + 1)))) {
          buffer.append('_');
        }
      }

      return buffer.toString().toUpperCase();
    } else {
      return "";
    }
  }

  public static <T> Row setHeadLine(Sheet sheet, Class<T> classType, List<String> excludeList) {
    Row firstRow = getOrCreateRow(sheet, 0);
    Field[] fields = classType.getDeclaredFields();

    for (int i = 0; i < fields.length; ++i) {
      Field field = fields[i];
      String fieldName = field.getName();
      if (excludeList == null || excludeList.indexOf(fieldName) == -1) {
        Cell cell = null;
        if (i != 0 && (i <= 0 || firstRow.getCell(i - 1) == null)) {
          cell = firstRow.createCell(i - 1);
        } else {
          cell = getOrCreateCell(firstRow, i);
        }

        field.setAccessible(true);
        cell.setCellValue(formatFieldHeadLine(fieldName));
      }
    }

    return firstRow;
  }

  public static Row setHeadLine(Sheet sheet, List<String> headLines) {
    Row firstRow = getOrCreateRow(sheet, 0);

    for (int i = 0; i < headLines.size(); ++i) {
      setRowCellValue(firstRow, i, headLines.get(i));
    }

    return firstRow;
  }

  public static Workbook createWorkBook(String type) {
    Workbook wb = null;
    if (type.equalsIgnoreCase("xlsx")) {
      wb = new XSSFWorkbook();
    } else if (type.equalsIgnoreCase("xls")) {
      wb = new HSSFWorkbook();
    }

    return (Workbook) wb;
  }

  public static <T> void exportDataToExcel(
      Sheet sheet, List<T> objects, Class<T> classType, List<String> excludeList) {
    setHeadLine(sheet, classType, excludeList);
    objectsToSheet(objects, sheet, classType, excludeList);
  }

  public static <T> void setRowCellValue(Row row, int i, T object) {
    Cell cell = getOrCreateCell(row, i);
    cell.setCellValue(otherTypeToString(object));
  }

  public static <T> void setSheetCellValue(Sheet sheet, int rowNum, int columnNum, T object) {
    Row row = getOrCreateRow(sheet, rowNum);
    setRowCellValue(row, columnNum, object);
  }
}
