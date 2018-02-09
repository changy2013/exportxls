package org.feiyu.export.common;

import com.google.common.base.Strings;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

/**
 * 日期常用方法类
 * @author kim
 */
public class DateUtils {

    /**
     * 防止构造
     */
    private DateUtils() {
    }

    private static final Logger logger = LoggerFactory.getLogger(DateUtils.class);

    /**
     * 获取指定的日期的Date格式
     * @param year 年
     * @param month 月
     * @param day 日
     * @return Date 指定年月日的日期格式
     */
    public static Date getDate(int year, int month, int day) {

        Calendar cal = Calendar.getInstance();
        cal.set(year, month - 1, day);
        return cal.getTime();
    }

    /**
     * 将日期格式化输出为yyyy-MM-dd HH:mm:ss的String格式
     * @param date
     * @return String
     */
    public static String format(Date date) {
        return format(date, "yyyy-MM-dd HH:mm:ss");
    }

    /**
     * 判断时间格式 格式必须为“YYYY-MM-dd”
     * 2004-2-30 是无效的
     * 2003-2-29 是无效的
     * @param str
     * @return
     */
    public static boolean isValidDate(String str) {
        return isValidDate(str, "yyyy-MM-dd");
    }

    /**
     * 判断时间格式
     * @param str
     * @return
     */
    public static boolean isValidDate(String str, String format) {
        if(Strings.isNullOrEmpty(str)) return false;
        DateFormat formatter = new SimpleDateFormat(format);
        try {
            Date date = formatter.parse(str);
            return str.equals(formatter.format(date));
        } catch (Exception e) {
            return false;
        }
    }


    /**
     * 将日期格式化输出为String格式
     * @param date 日期
     * @param format 日期格式
     * @return String
     */
    public static String format(Date date, String format) {
        if (date == null) return "";
        DateFormat df;
        if (Strings.isNullOrEmpty(format)) {
            df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        } else {
            df = new SimpleDateFormat(format);
        }
        try {
            return df.format(date);
        } catch (Exception e) {
            logger.error(e.getMessage());
        }
        return "";
    }

    public static String convertDateString(long timestamp) {
        Date date = new Date(timestamp);
        return format(date,"yyyy-MM-dd");
    }

    public static String convertDateStringCn(long timestamp) {
        Date date = new Date(timestamp);
        return format(date,"yyyy年MM月dd日");
    }

    /**
     * 得到和当前时间指定日期差时间
     * @param amount 相差的天数
     * @return 日期
     */
    public static Date addDays(int amount) {
        return addDays(new Date(), amount);
    }

    /**
     * 得到指定相差天数的日期
     * @param date 日期
     * @param amount 天数
     * @return 返回相加后的日期
     */
    public static Date addDays(Date date, int amount) {
        Calendar c = Calendar.getInstance();
        c.setTime(date);
        c.add(Calendar.DAY_OF_YEAR, amount);
        return c.getTime();
    }

}
