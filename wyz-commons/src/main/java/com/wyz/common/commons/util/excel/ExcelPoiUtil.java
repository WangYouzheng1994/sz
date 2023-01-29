package com.wyz.common.commons.util.excel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

/**
 * HSSF 97~2003 .xls 一个sheet最大行数65536，最大列数256。
 * XSSF 2007~ .xlsx 一个sheet最大行数1048576，最大列数16384。
 * SXSSF 缓存写入版本 .xlsx
 *
 * @Description: POI原生系Excel操作
 * @Author: WangYouzheng
 * @Date: 2023/1/29 10:55
 * @Version: V1.0
 */
public class ExcelPoiUtil {
    /**
     * 获取默认的表头格式类型
     *
     * @return
     */
    public static CellStyle getDefaultTitleRowStyle() {
        return null;
    }

    /**
     * 获取默认的值格式类型
     *
     * @return
     */
    public static CellStyle getDefaultValueRowStyle() {
        return null;
    }


    /**
     * 构造对应版本的WorkBook
     * @param titles 标题
     * @param infoList 内容
     * @param sheetName sheet名字
     * @param enumType excel文件类型
     * @return
     */
    public static Workbook createWorkBook(String[] titles, List<Map<String, String>> infoList, String sheetName, ExcelTypeEnum enumType) {
        Workbook workbook = null;
        Sheet sheet = null;


        switch (enumType) {
            case XLS: workbook = new HSSFWorkbook(); break;
            case XLSX: workbook = new XSSFWorkbook(); break;
            case XLSX_LOW_MEM: workbook = new SXSSFWorkbook(); break;
            default: workbook = new HSSFWorkbook();
        }

        // 创建sheet
        sheet = workbook.createSheet(sheetName);

        // 创建表头
        Row row = sheet.createRow(0);

        if (infoList == null || infoList.isEmpty()) {
            for (int celLength = 0; celLength < titles.length; celLength++) {
                // 创建相应的单元格
                Cell cell = row.createCell(celLength);
                cell.setCellValue(titles[celLength]);
            }
            return workbook;
        }

        return workbook;
    }

    /**
     * 创建导出的工作簿
     * 2003 版本 .xls
     *
     * @param titles
     * @param sheetName
     * @return
     */
    public static HSSFWorkbook createHSSFWorkbookInfo(String[] titles, List<Map<String, String>> infoList, String sheetName) {
        // 生成一个表格
        //        XSSFWorkbook xBook = new XSSFWorkbook();// 2007版本用此workbook解析
        HSSFWorkbook workbook = new HSSFWorkbook();
        if (infoList == null || infoList.isEmpty()) {
            // 创建sheet
            HSSFSheet sheet = workbook.createSheet(sheetName);
            // 创建表头
            Row row = sheet.createRow(0);
            for (int celLength = 0; celLength < titles.length; celLength++) {
                // 创建相应的单元格
                Cell cell = row.createCell(celLength);
                cell.setCellValue(titles[celLength]);
            }

            return workbook;
        }

        // 数据总件数
        int listSize = infoList.size();
        // sheet总长度
        int sheetSize = listSize;
        if (listSize > 65534 || listSize < 1) {
            sheetSize = 65534;
        }
        //样式
        CellStyle style = workbook.createCellStyle();
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("@"));

        //因为2003的Excel一个工作表最多可以有65536条记录，除去列头剩下65535条
        //所以如果记录太多，需要放到多个工作表中，其实就是个分页的过程
        // 计算一共有多少个工作表
        int sheetNum = (int) Math.ceil((double) listSize / (double) sheetSize);
        for (int i = 0; i < sheetNum; i++) {
            // 行号
            int rowNumber = 0;
            // 创建sheet
            HSSFSheet sheet = workbook.createSheet(sheetName + (i + 1));

            //获取开始索引和结束索引
            int firstIndex = i * sheetSize;
            int lastIndex = (i + 1) * sheetSize > listSize ? listSize : (i + 1) * sheetSize;

            // 构建临时数据
            List<Map<String, String>> tempList = infoList.subList(firstIndex, lastIndex);

            // 创建表头
            Row row = sheet.createRow(rowNumber);
            for (int celLength = 0; celLength < titles.length; celLength++) {
                // 创建相应的单元格
                Cell cell = row.createCell(celLength);
                cell.setCellValue(titles[celLength]);
                sheet.setDefaultColumnStyle(celLength, style);
            }

            if (!tempList.isEmpty()) {
                for (Map<String, String> info : tempList) {
                    rowNumber++;
                    row = sheet.createRow(rowNumber);
                    for (int celLength = 0; celLength < titles.length; celLength++) {
                        // 创建相应的单元格
                        Cell cell = row.createCell(celLength);
                        cell.setCellValue(info.get(titles[celLength]));

                        sheet.setDefaultColumnStyle(celLength, style);
                    }
                }
            }
        }

        return workbook;
    }




    /**
     * 创建Excel的Sheet，只包含标题
     *
     * @param titles
     * @param sheetName
     * @return
     */
    public static HSSFSheet createHSSFWorkbookTitle(HSSFWorkbook workbook, String[] titles, String sheetName) {

        // 生成一个表格
        HSSFSheet sheet = workbook.createSheet(sheetName);
        Row row = sheet.createRow(0);
        for (int celLength = 0; celLength < titles.length; celLength++) {
            // 创建相应的单元格
            Cell cell = row.createCell(celLength);
            cell.setCellValue(titles[celLength]);
        }

        return sheet;
    }


}
