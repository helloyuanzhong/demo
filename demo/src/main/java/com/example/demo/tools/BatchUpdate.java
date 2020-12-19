package com.example.demo.tools;

import com.example.demo.utils.SheetUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

/**
 * @Author: zhong.yuan
 * @Date: 2020-12-19 12:46
 */
public class BatchUpdate {

  private static String str = "-";
  private static String str3= "+";

  public static void main(String[] args) {

    String source = "/Users/yuanzhong/Desktop/source";
    String dest = "/Users/yuanzhong/Desktop/dest";

    File excel = new File("/Users/yuanzhong/Desktop/excel.xlsx");
    try{

      InputStream inputStream = new FileInputStream(excel);
      Workbook book = SheetUtil.getBookByPath(inputStream);
      Sheet dataSheet = book.getSheetAt(0);
      HashMap<Object, Object> nameOrderMap = new HashMap<>();
      for (int i = 2 ; i <= dataSheet.getLastRowNum(); i++) {
        Row row = dataSheet.getRow(i);
        if (row == null) {
          continue;
        }
        String order = getCellValueByColumn(row, 0); // 总序号
        String name = getCellValueByColumn(row, 1); // 姓名
        if (StringUtils.isNotEmpty(name)){
          nameOrderMap.put(name,order);
        }
//         System.out.println("顺序：" + i + "姓名： "+name +" 总序号："+ order);
        }
//      System.out.println("map size:"+nameOrderMap.values().size());
      // 读取目录下pdf文件

      File sourceFile = new File(source);
      if (sourceFile.isDirectory()){

        File[] pdfFiles = sourceFile.listFiles();
        System.out.println("需要匹配的数量:" +  pdfFiles.length);
        for (int i = 0 ;i< pdfFiles.length;i++){
          File old = pdfFiles[i];
          String originalName = old.getName();
          String name = "";
          String newName  = "";

          if (originalName.contains(str)){
            String[] split = originalName.split(str);
            name = split[0].trim();

            if (nameOrderMap.containsKey(name)){
              // 复制到新文件
               newName = nameOrderMap.get(name)+ str3 + originalName;
              copyFile(source +  File.separator + originalName,dest + File.separator + newName);
              System.out.println("顺序: " + i +"姓名:" + name +" 替换成功");
              continue;
            }

          }else{
            name = originalName.substring(0,3);
            if (nameOrderMap.containsKey(name)){
              // 复制到新文件
               newName = nameOrderMap.get(name)+ str3 + originalName;
              copyFile(source + File.separator + originalName,dest + File.separator + newName);
              System.out.println("顺序: " + i +"姓名:" + name +" 替换成功");
              continue;
            }else {
              name = originalName.substring(0,2);
              if (nameOrderMap.containsKey(name)){
                newName = nameOrderMap.get(name)+ str3 + originalName;
                copyFile(source + File.separator + originalName,dest + File.separator + newName);
                System.out.println("顺序: " + i +"姓名:" + name +" 替换成功");
                continue;
              }

            }
          }
            System.out.println("没有匹配的文件名称:" + originalName);

        }
        System.out.println("替换成功。。。");
      }

    }catch (Exception e){

      e.printStackTrace();

    }

  }

  public static void copyFile(String fromFile,String toFile){
    /*
     * 创建输入输出流对象，值为null
     */
    FileInputStream from = null;
    FileOutputStream to = null;
    /*
     * 复制文件操作
     */
    try {
      /*
       * 输入输出流对象引用指向源文件和目标文件
       */
      from = new FileInputStream(fromFile);
      to = new FileOutputStream(toFile);

      byte[] buffbyte = new byte[10*1024];//定义复制缓冲区
      /*
       * 复制操作
       */
      int len = -1;
      while((len = from.read(buffbyte)) != -1) {
        to.write(buffbyte, 0, len);
      }
    } catch (Exception e) {
      e.printStackTrace();
    } finally {//关闭流
      if(from != null) {
        try {
          from.close();
        }catch (Exception e) {
          System.out.println(e.getMessage());
        }
      }
      if(to != null) {
        try {
          to.close();
        }catch(Exception e) {
          System.out.println(e.getMessage());
        }
      }
    }




  }



  public static String getCellValueByColumn(Row row, int columnNum) {
    Cell cell = row.getCell(columnNum);
    return getCellValue(cell);
  }


  //获取单元格内容(支持日期及小数)
  private static String getCellValue(Cell cell) {
    if(cell == null){
      return null;
    }

    String cellValue = "";
    //like12 modified bug,20171124,不能保留小数
    //DecimalFormat df = new DecimalFormat("#");//不保留小数
    DecimalFormat df = new DecimalFormat("#.##");//最多保留x位小数
    switch (cell.getCellType()) {
      case HSSFCell.CELL_TYPE_STRING:
        cellValue = cell.getRichStringCellValue().getString().trim();
        break;
      case HSSFCell.CELL_TYPE_NUMERIC:
        // like12 add,20180622,支持日期格式
        if (HSSFDateUtil.isCellDateFormatted(cell)) {
          Date d = cell.getDateCellValue();
          DateFormat df2 = new SimpleDateFormat("yyyy-MM-dd");// HH:mm:ss
          cellValue = df2.format(d);
        }
        // 数字
        else {
          cellValue = df.format(cell.getNumericCellValue()).toString();
        }
        break;
      case HSSFCell.CELL_TYPE_BOOLEAN:
        cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
        break;
      case HSSFCell.CELL_TYPE_FORMULA:
        cellValue = cell.getCellFormula();
        break;
      default:
        cellValue = "";
    }
    return cellValue;
  }
}
