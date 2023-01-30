package com.wyz.common.commons.util.excel;

import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.EnumSet;

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
    XLS("xls", HSSFWorkbook.class, 65534L, 255),

    /**
     * 2007+
     */
    XLSX("xlsx", XSSFWorkbook.class, 1048573L, 255),

    /**
     * 2007+ 低内存写入版本，避免OOM，吃磁盘I/O 速度不一定快。
     */
    XLSX_LOW_MEM("xlsx",SXSSFWorkbook.class,1048573L, 255);

    /**
     * 后缀
     */
    private final String extra;

    /**
     * 全路径类名
     */
    private final Class<? extends Workbook> workBookType;

    /**
     * 单sheet 行数 上限
     */
    private final Long rowSize;

    /**
     * 工作簿 sheet 个数上限
     */
    private final Integer sheetSize;

    static {

    }

    public static Long getRowSize(Class<? extends Workbook> workBookType) {
        ArrayList<ExcelTypeEnum> excelTypeEnums = Lists.newArrayList(values());
        for (ExcelTypeEnum enumType: excelTypeEnums) {
            if (enumType.getWorkBookType() == workBookType) {
                return enumType.getRowSize();
            }
        }
        return 0l;
    }

    Long getRowSize() {
        return rowSize;
    }

    Class<? extends Workbook> getWorkBookType() {
        return workBookType;
    }

    ExcelTypeEnum(String extra, Class<? extends Workbook> workBookType, Long rowSize, Integer sheetSize) {
        this.extra = extra;
        this.workBookType = workBookType;
        this.rowSize = rowSize;
        this.sheetSize = sheetSize;
    }
}