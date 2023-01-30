package com.wyz.common.commons.util.excel;

import cn.hutool.core.lang.Assert;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
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
@Slf4j
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
     * 获取表头行
     *
     * @param titles 表头数组
     * @param sheet sheet页面
     * @return
     */
    public static Row getTitleRow(String[] titles, Sheet sheet) {
        // 创建表头
        Row row = sheet.createRow(0);
        CellStyle titleStyle = getDefaultTitleRowStyle();

        for (int celLength = 0; celLength < titles.length; celLength++) {
            // 创建相应的单元格
            Cell cell = row.createCell(celLength);
            cell.setCellValue(titles[celLength]);
            cell.setCellStyle(titleStyle);
        }
        return row;
    }


    /**
     * 设置数据
     * 自动识别 低版本与高版本并进行自动切分
     *
     * @param workbook
     * @param firstSheet
     * @param infoList
     * @param autosplit
     * @return
     */
    public static Boolean setDataList(Workbook workbook, Sheet firstSheet, String sheetName, List<Map<String, Object>> infoList, boolean autosplit) {
        Assert.notNull(workbook);
        Assert.notNull(firstSheet);

        Long rowSize = ExcelTypeEnum.getRowSize(workbook.getClass());

        if (infoList != null && !infoList.isEmpty()) {
            // TODO: 低版本需要分割~
            // 如果是允许继续分割 那就继续 否则停止迭代 并返回。
            double sheetNum = 1;
            if (autosplit) {
                sheetNum = Math.ceil(infoList.size() / rowSize);
            }
/*
            firstSheet();
            workbook.createSheet(sheetName + (i + 1));*/
        }

        return true;
    }

    /**
     * 构造对应版本的WorkBook
     * @param titles 标题
     * @param infoList 内容
     * @param sheetName sheet名字
     * @param enumType excel文件类型
     * @return
     */
    public static Workbook createWorkBook(String[] titles, List<Map<String, Object>> infoList, String sheetName, ExcelTypeEnum enumType) {
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

        // 如果为空就直接设置表头好了
        if (infoList != null && !infoList.isEmpty()) {
            // 设置表头
            getTitleRow(titles, sheet);
        } else {
            setDataList(workbook, sheet, sheetName, infoList, true);
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

    /**
     * 通过Http流去输出附件
     *
     * @param fileName
     * @return
     */
    public static void writeExcelFile(HttpServletRequest request, HttpServletResponse response, Workbook workbook, String fileName) {
        // 根据列名填充相应的数据
        try (ServletOutputStream out = response.getOutputStream();) {
            response.reset();
            response.setContentType("application/msexcel;charset=utf-8");
            //response.setHeader("content-disposition", "attachment;filename=" + new String((fileName).getBytes("UTF-8"), "ISO8859-1"));
            String userAgent = request.getHeader("user-agent");
            if (userAgent != null && userAgent.indexOf("Edge") >= 0) {
                fileName = URLEncoder.encode(fileName, "UTF8");
            } else if (userAgent.indexOf("Firefox") >= 0 || userAgent.indexOf("Chrome") >= 0
                    || userAgent.indexOf("Safari") >= 0) {
                fileName = new String((fileName).getBytes("UTF-8"), "ISO8859-1");
            } else {
                fileName = URLEncoder.encode(fileName, "UTF8"); //其他浏览器
            }
            response.setHeader("content-disposition", "attachment;filename=" + fileName);

            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 多线程文件下载~
     *
     * @param threadNum 线程数
     * @param request httpservletrequest
     * @param response httpservletResponse
     * @param workbook excel文件对象
     * @param fileName 文件名称
     */
    public static void multiThreadWrite(int threadNum, HttpServletRequest request, HttpServletResponse response, Workbook workbook, String fileName) {

    }

    /**
     * 断点续传文件
     *
     * @param response
     * @param request
     * @param location
     * @return
     */
    public static boolean downFile(HttpServletResponse response, HttpServletRequest request, String location) {
        try {
            File file = new File(location);
            if (file.exists()) {
                long p = 0L;
                long toLength = 0L;
                long contentLength = 0L;
                int rangeSwitch = 0; // 0,从头开始的全文下载；1,从某字节开始的下载（bytes=27000-）；2,从某字节开始到某字节结束的下载（bytes=27000-39000）
                long fileLength;
                String rangBytes = "";
                fileLength = file.length();

                // get file content
                try (InputStream ins = new FileInputStream(file);
                     BufferedInputStream bis = new BufferedInputStream(ins);
                     OutputStream out = response.getOutputStream();) {

                    // tell the client to allow accept-ranges
                    response.reset();
                    response.setHeader("Accept-Ranges", "bytes");

                    // client requests a file block download start byte
                    String range = request.getHeader("Range");
                    if (range != null && range.trim().length() > 0 && !"null".equals(range)) {
                        response.setStatus(javax.servlet.http.HttpServletResponse.SC_PARTIAL_CONTENT);
                        rangBytes = range.replaceAll("bytes=", "");
                        if (rangBytes.endsWith("-")) {  // bytes=270000-
                            rangeSwitch = 1;
                            p = Long.parseLong(rangBytes.substring(0, rangBytes.indexOf("-")));
                            contentLength = fileLength - p;  // 客户端请求的是270000之后的字节（包括bytes下标索引为270000的字节）
                        } else { // bytes=270000-320000
                            rangeSwitch = 2;
                            String temp1 = rangBytes.substring(0, rangBytes.indexOf("-"));
                            String temp2 = rangBytes.substring(rangBytes.indexOf("-") + 1, rangBytes.length());
                            p = Long.parseLong(temp1);
                            toLength = Long.parseLong(temp2);
                            contentLength = toLength - p + 1; // 客户端请求的是 270000-320000 之间的字节
                        }
                    } else {
                        contentLength = fileLength;
                    }

                    // 如果设设置了Content-Length，则客户端会自动进行多线程下载。如果不希望支持多线程，则不要设置这个参数。
                    // Content-Length: [文件的总大小] - [客户端请求的下载的文件块的开始字节]
                    response.setHeader("Content-Length", new Long(contentLength).toString());

                    // 断点开始
                    // 响应的格式是:
                    // Content-Range: bytes [文件块的开始字节]-[文件的总大小 - 1]/[文件的总大小]
                    if (rangeSwitch == 1) {
                        String contentRange = new StringBuffer("bytes ").append(new Long(p).toString()).append("-")
                                .append(new Long(fileLength - 1).toString()).append("/")
                                .append(new Long(fileLength).toString()).toString();
                        response.setHeader("Content-Range", contentRange);
                        bis.skip(p);
                    } else if (rangeSwitch == 2) {
                        String contentRange = range.replace("=", " ") + "/" + new Long(fileLength).toString();
                        response.setHeader("Content-Range", contentRange);
                        bis.skip(p);
                    } else {
                        String contentRange = new StringBuffer("bytes ").append("0-")
                                .append(fileLength - 1).append("/")
                                .append(fileLength).toString();
                        response.setHeader("Content-Range", contentRange);
                    }

                    String fileName = file.getName();
                    response.reset();
                    //response.setHeader("content-disposition", "attachment;filename=" + new String((fileName).getBytes("UTF-8"), "ISO8859-1"));
                    String userAgent = request.getHeader("user-agent");
                    if (userAgent != null && userAgent.indexOf("Edge") >= 0) {
                        fileName = URLEncoder.encode(fileName, "UTF8");
                    } else if (userAgent.indexOf("Firefox") >= 0 || userAgent.indexOf("Chrome") >= 0
                            || userAgent.indexOf("Safari") >= 0) {
                        fileName = new String((fileName).getBytes("UTF-8"), "ISO8859-1");
                    } else {
                        fileName = URLEncoder.encode(fileName, "UTF8"); //其他浏览器
                    }
                    response.setContentType("application/octet-stream");
                    response.addHeader("Content-Disposition", "attachment;filename=" + fileName);

                    int n = 0;
                    long readLength = 0;
                    int bsize = 1024;
                    byte[] bytes = new byte[bsize];
                    if (rangeSwitch == 2) {
                        // 针对 bytes=27000-39000 的请求，从27000开始写数据
                        while (readLength <= contentLength - bsize) {
                            n = bis.read(bytes);
                            readLength += n;
                            out.write(bytes, 0, n);
                        }
                        if (readLength <= contentLength) {
                            n = bis.read(bytes, 0, (int) (contentLength - readLength));
                            out.write(bytes, 0, n);
                        }
                    } else {
                        while ((n = bis.read(bytes)) != -1) {
                            out.write(bytes, 0, n);
                        }
                    }
                }
            } else {
                throw new Exception("file not found");
            }
        } catch (IOException ie) {
            // 忽略 ClientAbortException 之类的异常
            log.error(ie.getMessage(), ie);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            return false;
        }
        return true;
    }
}
