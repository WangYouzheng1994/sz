package com.wyz.common.commons.util.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @Description: Excel类型枚举
 * @Author: WangYouzheng
 * @Date: 2023/1/29 11:36
 * @Version: V1.0
 */
public enum ExcelTypeEnum {
    /**
     * 97~2003
     */
    XLS("xls", HSSFWorkbook.class),

    /**
     * 2007+
     */
    XLSX("xlsx", XSSFWorkbook.class),

    /**
     * 2007+ 低内存写入版本，避免OOM，吃磁盘I/O 速度不一定快。
     */
    XLSX_LOW_MEM("xlsx",SXSSFWorkbook.class);

    private String extra;
    private Class<? extends Workbook> workBookType;

    ExcelTypeEnum(String extra, Class<? extends Workbook> workBookType) {
        this.extra = extra;
        this.workBookType = workBookType;
    }


}