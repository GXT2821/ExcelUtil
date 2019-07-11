package com.telek.gt.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExeclUtil {

    /**
     * @param path 文件地址
     * @return
     * @Description 解析Execl文件中的数据，此方法只解析了execl文件中的第一个sheet页，并且没有考虑第一行表头
     */
    public static List<List<String>> readExecl(String path) {
        List<List<String>> result = new ArrayList<>();
        try {
            //获取execl文件的个工作簿
            Workbook wb = getWorkbook(path);
            if (wb != null) {
                //获取execl文件的的第一个sheet页
                Sheet sheet = wb.getSheetAt(0);
                for (Row row : sheet) {
                    ArrayList<String> list = new ArrayList<String>();
                    for (Cell cell : row) {
                        list.add(cell.toString());
                    }
                    result.add(list);
                }
            } else {
                System.out.println("找不到指定文件");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * @param path
     * @return
     * @Description 解析Execl文件中的数据，此方法只解析了execl文件中的第一个sheet页，并且将第一列表头作为每行数据的key返回
     */
    public static List<Map<String, Object>> readExeclByHead(String path) {
        List<Map<String, Object>> result = new ArrayList<>();
        try {
            //获取execl文件的个工作簿
            Workbook wb = getWorkbook(path);
            if (wb != null) {
                //获取execl文件的的第一个sheet页
                Sheet sheet = wb.getSheetAt(0);
                //获取第一行
                Row row = sheet.getRow(0);
                int firstRowNum = sheet.getFirstRowNum();
                int lastRowNum = sheet.getLastRowNum();
                for (int i = 1; i <= lastRowNum; i++) {
                    Map<String, Object> map = new HashMap<>();
                    Row row1 = sheet.getRow(i);
                    if (row1 != null) {
                        int firstCellNum = row1.getFirstCellNum();
                        int lastCellNum = row.getLastCellNum();
                        for (int j = firstCellNum; j <= lastCellNum; j++) {
                            Cell cellKey = row.getCell(j);
                            Cell cellValue = row1.getCell(j);
                            if (cellKey != null && cellValue != null) {
                                map.put(cellKey.toString(), cellValue.toString());
                            }
                        }
                        result.add(map);
                    }
                }
            } else {
                System.out.println("找不到指定文件");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    private static Workbook getWorkbook(String path) {
        InputStream is = null;
        try {
            File excel = new File(path);
            if (excel.isFile() && excel.exists()) {
                String fileType = path.substring(path.lastIndexOf(".") + 1);
                //读取excel文件流
                is = new FileInputStream(path);
                //获取工作薄
                Workbook wb = null;
                if (fileType.equals("xls")) {
                    wb = new HSSFWorkbook(is);
                } else if (fileType.equals("xlsx")) {
                    wb = new XSSFWorkbook(is);
                } else {
                    return null;
                }
                return wb;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e){
            e.printStackTrace();
        }finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;
    }
}
