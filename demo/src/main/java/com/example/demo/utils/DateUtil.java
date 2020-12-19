package com.example.demo.utils;

import org.apache.commons.lang3.StringUtils;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.DecimalFormat;
import java.text.ParseException;
import java.util.Calendar;
import java.util.Date;

public class DateUtil {
  private static final Logger logger = LoggerFactory.getLogger(DateUtil.class);
  public static final DateTimeFormatter DTF_RJ10 = DateTimeFormat.forPattern("yyyy-MM-dd");
  public static final DateTimeFormatter DTF_RJ8 = DateTimeFormat.forPattern("yyyyMMdd");
  public static final DateTimeFormatter DTF_SJ8 = DateTimeFormat.forPattern("HH:mm:ss");
  public static final DateTimeFormatter DTF_SJ6 = DateTimeFormat.forPattern("HHmmss");
  public static final DateTimeFormatter DTF_RSJ15 = DateTimeFormat.forPattern("yyyyMMdd_HHmmss");
  public static final DateTimeFormatter DTF_RSJ22 =
      DateTimeFormat.forPattern("yyyy-MM-dd HH:mm:ss.SS");
  public static final DateTimeFormatter DTF_RSJ19 =
      DateTimeFormat.forPattern("yyyy-MM-dd HH:mm:ss");
  private static final DateTimeFormatter YYYYMMDDHH24MISS =
      DateTimeFormat.forPattern("yyyyMMddHHmmss");
  private static final String LONG_DATE_REGEX = "\\d{4}\\d{2}\\d{2}\\d{2}\\d{2}\\d{2}";
  private static final String SHORT_DATE_REGEX = "\\d{4}\\d{2}\\d{2}";

  public DateUtil() {}

  public static DateTimeFormatter createDateTimeFormatter(String format) {
    return DateTimeFormat.forPattern(format);
  }

  public static Date getYesterday() {
    Calendar c = Calendar.getInstance();
    c.add(5, -1);
    return c.getTime();
  }

  public static Date getTomorrow() {
    Calendar c = Calendar.getInstance();
    c.add(5, 1);
    return c.getTime();
  }

  public static Date getTomorrow(Date date) {
    return changeDate(date, 1);
  }

  public static Date changeDate(Date date, Integer days) {
    Calendar c = Calendar.getInstance();
    c.setTime(date);
    c.add(5, days);
    return c.getTime();
  }

  public static String changeDateDTF_RJ10(String dateStr, int days) {
    Date date = toDate(dateStr);
    Calendar c = Calendar.getInstance();
    c.setTime(date);
    c.add(5, days);
    Date resultDate = c.getTime();
    return formatChinese(resultDate);
  }

  public static Date changeTime(Date date, String time) {
    Calendar c = Calendar.getInstance();
    c.setTime(date);
    String[] times = time.split(":");
    c.add(10, Integer.parseInt(times[0]));
    c.add(12, Integer.parseInt(times[1]));
    c.add(13, Integer.parseInt(times[2]));
    return c.getTime();
  }

  public static Long getTimeDifference(String dateString) {
    Date date = DTF_RSJ19.parseLocalDateTime(dateString).toDate();
    Date dateNow = new Date();
    return date.getTime() - dateNow.getTime();
  }

  public static Long getTimeDifference(String date, String type) {
    String[] time = date.split(":");
    Long millis = 0L;
    if ("time".equals(type)) {
      millis =
          (((long) Integer.parseInt(time[0]) * 60L + (long) Integer.parseInt(time[1])) * 60L
                  + (long) Integer.parseInt(time[2]))
              * 1000L;
    } else if ("date".equals(type)) {
      Calendar c1 = Calendar.getInstance();
      c1.set(
          0, 0, 0, Integer.parseInt(time[0]), Integer.parseInt(time[1]), Integer.parseInt(time[2]));
      Calendar c2 = Calendar.getInstance();
      c2.set(0, 0, 0, c2.get(10), c2.get(12), c2.get(13));
      if (c2.after(c1)) {
        c1.set(
            0,
            0,
            1,
            Integer.parseInt(time[0]),
            Integer.parseInt(time[1]),
            Integer.parseInt(time[2]));
      }

      millis = c1.getTimeInMillis() - c2.getTimeInMillis();
    } else {
      millis = 86400000L;
      logger.warn("类型异常");
    }

    return millis;
  }

  public static Term getTerm(Term term, String changeType, Integer num, Date date) {
    if (num <= 0) {
      return null;
    } else if (!term.contain(date)) {
      return term;
    } else {
      Date startDate = term.getStartDate();
      Date endDate = changeDateByType(startDate, changeType, num);
      if (endDate.after(term.getEndDate())) {
        endDate = term.getEndDate();
      }

      Term resuleTerm;
      for (resuleTerm = new Term(startDate, endDate);
          !resuleTerm.contain(date);
          resuleTerm = new Term(startDate, endDate)) {
        startDate = endDate;
        endDate = changeDateByType(endDate, changeType, num);
        if (endDate.after(term.getEndDate())) {
          endDate = term.getEndDate();
        }
      }

      return resuleTerm;
    }
  }

  public static Date changeDateByType(Date startDate, String changeType, Integer num) {
    Calendar cEndDate = Calendar.getInstance();
    cEndDate.setTime(startDate);
    if ("Y".equals(changeType)) {
      cEndDate.add(1, num);
    } else if ("M".equals(changeType)) {
      cEndDate.add(2, num);
    } else if ("D".equals(changeType)) {
      cEndDate.add(6, num);
    }

    return cEndDate.getTime();
  }

  public static Date toDate(String d, Boolean full) {
    if (StringUtils.isEmpty(d)) {
      return null;
    } else if (full) {
      return d.length() == 14
          ? YYYYMMDDHH24MISS.parseLocalDateTime(d).toDate()
          : YYYYMMDDHH24MISS.parseLocalDateTime(d + "000000").toDate();
    } else {
      return DTF_RJ8.parseLocalDateTime(d).toDate();
    }
  }

  public static String format(Date d, Boolean full) {
    if (d == null) {
      return null;
    } else {
      return full ? YYYYMMDDHH24MISS.print(d.getTime()) : DTF_RJ8.print(d.getTime());
    }
  }

  public static Date toDate(String d) {
    return StringUtils.isEmpty(d) ? null : DTF_RJ10.parseLocalDateTime(d).toDate();
  }

  public static String formatChinese(Date d) {
    return d == null ? null : DTF_RJ10.print(d.getTime());
  }

  public static Date toDateFull(String d) throws ParseException {
    return StringUtils.isEmpty(d) ? null : DTF_RSJ19.parseLocalDateTime(d).toDate();
  }

  public static String formatChineseFull(Date d) {
    return d == null ? null : DTF_RSJ19.print(d.getTime());
  }

  public static boolean matchDate(String date, boolean b) {
    if (b && date.matches("\\d{4}\\d{2}\\d{2}\\d{2}\\d{2}\\d{2}")) {
      return true;
    } else {
      return !b && date.matches("\\d{4}\\d{2}\\d{2}");
    }
  }

  public static boolean vaildToday(String date, boolean isLong) {
    Date d = toDate(date, isLong);
    Calendar cal = Calendar.getInstance();
    cal.setTime(new Date());
    cal.add(11, 1);
    Date e = cal.getTime();
    cal.add(11, -2);
    Date s = cal.getTime();
    return !d.after(e) && !d.before(s);
  }

  public static String getBirthdayFromPersonCardNo(String cardNo) {
    if (StringUtils.isEmpty(cardNo)) {
      return null;
    } else {
      return cardNo.length() != 18 ? null : cardNo.substring(6, 14);
    }
  }

  public static String getYear(Date date) {
    Calendar calendar = Calendar.getInstance();
    calendar.setTime(date);
    return String.valueOf(calendar.get(1));
  }

  public static String getMonth(Date date) {
    Calendar calendar = Calendar.getInstance();
    calendar.setTime(date);
    DecimalFormat df = new DecimalFormat();
    df.applyPattern("00;00");
    return df.format((long) (calendar.get(2) + 1));
  }

  public static Date cleanHMS(Date d) {
    Calendar cal = Calendar.getInstance();
    cal.setTime(d);
    cal.set(11, 0);
    cal.set(12, 0);
    cal.set(13, 0);
    cal.set(14, 0);
    return cal.getTime();
  }

  public static String getDay(Date date) {
    Calendar calendar = Calendar.getInstance();
    calendar.setTime(date);
    DecimalFormat df = new DecimalFormat();
    df.applyPattern("00;00");
    return df.format((long) calendar.get(5));
  }

  public static boolean isSameDay(Date startDate, Date endDate) {
    if (startDate != null && endDate != null) {
      return cleanHMS(startDate).getTime() == cleanHMS(endDate).getTime();
    } else {
      return false;
    }
  }
}
