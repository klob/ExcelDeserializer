package com.diandi.klob.sdk.excel;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * *******************************************************************************
 * *********    Author : klob(kloblic@gmail.com) .
 * *********    Date : 2015-10-24  .
 * *********    Time : 14:32 .
 * *********    Version : 1.0
 * *********    Copyright © 2015, klob, All Rights Reserved
 * *******************************************************************************
 */
public class ExcelDeserializer {

    /**
     * 总行数
     */

    private int totalRows = 0;

    /**
     * 总列数
     */

    private int totalCells = 0;


    /**
     * 首行
     */
    private int mHeadRowIndex = 0;

    /**
     * 错误信息
     */

    private String errorInfo;

    /**
     * Excel头
     */
    private ExcelHeader mExcelHeader;

    public ExcelDeserializer() {

    }

    public ExcelDeserializer(int headRow) {
        mHeadRowIndex = headRow;
    }

    /**
     * 依据内容判断是否为excel2003及以下
     */
    public static boolean isExcel2003(String filePath) {
        try {
            BufferedInputStream bis = new BufferedInputStream(new FileInputStream(filePath));
            if (POIFSFileSystem.hasPOIFSHeader(bis)) {
                System.out.println("Excel版本为excel2003及以下");
                return true;
            }
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
        return false;
    }

    /**
     * 依据内容判断是否为excel2007及以上
     */
    public static boolean isExcel2007(String filePath) {
        try {
            BufferedInputStream bis = new BufferedInputStream(new FileInputStream(filePath));
            if (POIXMLDocument.hasOOXMLHeader(bis)) {
                System.out.println("Excel版本为excel2007及以上");
                return true;
            }
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
        return false;
    }

    /**
     * @return 得到excel总行数
     */

    public int getTotalRows() {

        return totalRows;

    }

    /**
     * @return 得到excel总列数
     */

    public int getTotalCells() {

        return totalCells;

    }

    /**
     * @return ：得到错误信息
     */

    public String getErrorInfo() {

        return errorInfo;

    }

    /**
     * 验证excel文件
     */

    public boolean validateExcel(String filePath) {

        /** 检查文件名是否为空或者是否是Excel格式的文件 */

        if (filePath == null || !(isExcel2003(filePath) || isExcel2007(filePath))) {

            errorInfo = "文件名不是excel格式";

            return false;

        }

        /** 检查文件是否存在 */

        File file = new File(filePath);

        if (file == null || !file.exists()) {

            errorInfo = "文件不存在";

            return false;

        }

        return true;

    }

    /**
     * 读取文件并发反序列化为对象
     */

    public <T> List read(File file, Class<T> clz) {
        List<List<String>> data = readFile(file.getAbsolutePath());
        List<T> list = new ArrayList<>();
        for (List<String> row : data) {
            JSONObject jsonObject = new JSONObject();
            for (int i = 0; i < row.size(); i++) {
                jsonObject.put(mExcelHeader.getField(i), row.get(i) + "");
            }
            try {
                T model = parseData(jsonObject.toJSONString(), clz);
                list.add(model);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        L.D(list.toString());
        return list;

    }

    /**
     * 解析数据
     */

    public <T> T parseData(String json, Class<T> clz) {
        return JSON.parseObject(json, clz);
    }

    public List<List<String>> readFile(String filePath) {

        List<List<String>> dataLst = new ArrayList<List<String>>();

        InputStream is = null;

        try {

            /** 验证文件是否合法 */

            if (!validateExcel(filePath)) {

                System.out.println(errorInfo);

                return null;

            }

            /** 判断文件的类型，是2003还是2007 */

            boolean isExcel2003 = false;

            if (isExcel2003(filePath)) {
                isExcel2003 = true;
            }

            // L.d("是否是2003版本"+isExcel2003);
            /** 调用本类提供的根据流读取的方法 */

            File file = new File(filePath);

            is = new FileInputStream(file);

            dataLst = read(is, isExcel2003);

            is.close();

        } catch (Exception ex) {

            ex.printStackTrace();

        } finally {

            if (is != null) {

                try {

                    is.close();

                } catch (IOException e) {

                    is = null;

                    e.printStackTrace();

                }

            }

        }

        /** 返回最后读取的结果 */

        return dataLst;

    }

    /**
     * 根据流读取Excel文件
     */

    public List<List<String>> read(InputStream inputStream, boolean isExcel2003) {

        List<List<String>> dataLst = null;

        try {

            /** 根据版本选择创建Workbook的方式 */

            Workbook wb = null;

            if (isExcel2003) {
                wb = new HSSFWorkbook(inputStream);
            } else {
                wb = new XSSFWorkbook(inputStream);
            }
            dataLst = readFirstSheet(wb);

        } catch (IOException e) {

            e.printStackTrace();

        }

        return dataLst;

    }

    /**
     * 读取一张表格数据
     */

    public List<List<String>> readFirstSheet(Workbook wb) {

        /** 得到第一个sheet */

        Sheet sheet = wb.getSheetAt(0);
        return readSheet(sheet);

    }

    /**
     * 读取所有表格数据
     */


    public List<List<String>> readAllSheet(Workbook wb) {
        List<List<String>> mExcelData = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            List<List<String>> sheetData = readSheet(wb.getSheetAt(i));
            mExcelData.addAll(sheetData);
        }

        return mExcelData;
    }

    /**
     * 读取表格
     */

    public List<List<String>> readSheet(Sheet sheet) {
        List<List<String>> dataLst = new ArrayList<List<String>>();

        /** 得到Excel的行数 */

        this.totalRows = sheet.getPhysicalNumberOfRows();

        /** 得到Excel的列数 */

        if (this.totalRows >= 1 && sheet.getRow(0) != null) {

            this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();

        }

        /** 循环Excel的行 */

        readHeader(sheet.getRow(mHeadRowIndex));

        for (int r = mHeadRowIndex + 1; r < this.totalRows; r++) {

            Row row = sheet.getRow(r);

            if (row == null) {

                continue;

            }

            /** 读取一行数据 */

            dataLst.add(readRow(row));

        }

        return dataLst;

    }

    /**
     * 读取一行数据
     */

    public List<String> readRow(Row row) {
        List<String> rowData = new ArrayList<String>();


        for (int c = 0; c < this.getTotalCells(); c++) {

            Cell cell = row.getCell(c);

            String cellValue = "";

            if (null != cell) {
                // 以下是判断数据的类型
                switch (cell.getCellType()) {

                    case HSSFCell.CELL_TYPE_NUMERIC:
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                            Date date = cell.getDateCellValue();
                            if (date != null) {
                                cellValue = new SimpleDateFormat("hh:mm")
                                        .format(date);
                            } else {
                                cellValue = "";
                            }
                        } else {
                            cellValue = new DecimalFormat("0").format(cell
                                    .getNumericCellValue());
                        }
                        break;
                   /* case HSSFCell.CELL_TYPE_NUMERIC: // 数字
                        cellValue = cell.getNumericCellValue() + "";
                        break;
*/
                    case HSSFCell.CELL_TYPE_STRING: // 字符串
                        cellValue = cell.getStringCellValue();
                        break;
                    case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                        cellValue = cell.getBooleanCellValue() + "";
                        break;
                    case HSSFCell.CELL_TYPE_FORMULA:
                        // 导入时如果为公式生成的数据则无值
                        if (!cell.getStringCellValue().equals("")) {
                            cellValue = cell.getStringCellValue();
                        } else {
                            cellValue = cell.getNumericCellValue() + "";
                        }
                        break;

                    case HSSFCell.CELL_TYPE_BLANK: // 空值
                        cellValue = "";
                        break;

                    case HSSFCell.CELL_TYPE_ERROR: // 故障
                        cellValue = "非法字符";
                        break;

                    default:
                        cellValue = "未知类型";
                        break;
                }
            }

            rowData.add(cellValue);

        }
        return rowData;
    }

    /**
     * 得到表头
     */

    public void readHeader(Row headRow) {
        if (headRow == null) {
            L.Warn("headRow is null");
            return;
        }

        List<String> rowLst = new ArrayList<String>();

        /** 循环Excel的列 */

        for (int c = 0; c < this.getTotalCells(); c++) {

            Cell cell = headRow.getCell(c);

            String cellValue = "";

            if (null != cell) {
                // 以下是判断数据的类型
                switch (cell.getCellType()) {
                    case HSSFCell.CELL_TYPE_NUMERIC: // 数字
                        cellValue = cell.getNumericCellValue() + "";
                        break;

                    case HSSFCell.CELL_TYPE_STRING: // 字符串
                        cellValue = cell.getStringCellValue();
                        break;

                    case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                        cellValue = cell.getBooleanCellValue() + "";
                        break;

                    case HSSFCell.CELL_TYPE_FORMULA: // 公式
                        cellValue = cell.getCellFormula() + "";
                        break;

                    case HSSFCell.CELL_TYPE_BLANK: // 空值
                        cellValue = "";
                        break;

                    case HSSFCell.CELL_TYPE_ERROR: // 故障
                        cellValue = "非法字符";
                        break;

                    default:
                        cellValue = "未知类型";
                        break;
                }
            }

            rowLst.add(cellValue);

        }

        mExcelHeader = new ExcelHeader(rowLst);

    }
}
