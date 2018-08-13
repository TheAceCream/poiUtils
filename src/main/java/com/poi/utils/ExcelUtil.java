package com.poi.utils;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * Created with IntelliJ IDEA.
 * Description:
 * User: Ace
 * Date: 2018/8/13 10:26
 * Time: 14:15
 *
 * excle导出/导入
 */
public class ExcelUtil {

    private static final int DEFAULT_WIDTH = 5*256*5;

    /**
     * 导入excel
     * @param filePath  文件物理地址
     * @param keys      字段名称数组 如  ["id", "name", ... ]
     * @return
     * @throws Exception
     */
    public static List<Map<String, Object>> defaultImport(String filePath, String[] keys) throws Exception {
        List<Map<String, Object>> list = new ArrayList<>();
        Map<String, Object> map;
        if (keys==null){
            throw new Exception("keys不能为空!");
        }
        if (!filePath.endsWith(".xls") && !filePath.endsWith(".xlsx")) {
            throw new Exception("文件格式有误!");
        }
        //读
        FileInputStream fis = null;
        Workbook workbook = null;
        try {
            fis = new FileInputStream(filePath);
            if(filePath.endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else if(filePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            }

            // 获取第一个工作表信息
            Sheet sheet = workbook.getSheetAt(0);
            //获得数据的总行数
            int totalRowNum = sheet.getLastRowNum();

            // 获得表头
            Row rowHead = sheet.getRow(0);
            // 获得表头总列数
            int cols = rowHead.getPhysicalNumberOfCells();

            if(keys.length != cols) {
                throw new Exception("传入的key数组长度与表头长度不一致!");
            }
            Row row = null;
            Cell cell = null;
            Object value = null;
            // 遍历所有行
            for (int i = 1; i <= totalRowNum; i++) {
                // 清空数据，避免遍历时读取上一次遍历数据
                row = null;
                cell = null;
                value = null;
                map = new HashMap<String, Object>();

                row = sheet.getRow(i);
                if(null == row) continue;	// 若该行第一列为空，则默认认为该行就是空行

                // 遍历该行所有列
                for (short j = 0; j < cols; j++) {
                    cell = row.getCell(j);
                    if(null == cell) continue;	// 为空时，下一列

                    // 根据poi返回的类型，做相应的get处理
                    if(Cell.CELL_TYPE_STRING == cell.getCellType()) {
                        value = cell.getStringCellValue();
                    } else if(Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                        value = cell.getNumericCellValue();

                        // 由于日期类型格式也被认为是数值型，此处判断是否是日期的格式，若时，则读取为日期类型
                        if(cell.getCellStyle().getDataFormat() > 0)  {
                            value = cell.getDateCellValue();
                        }
                    } else if(Cell.CELL_TYPE_BOOLEAN == cell.getCellType()) {
                        value = cell.getBooleanCellValue();
                    } else if(Cell.CELL_TYPE_BLANK == cell.getCellType()) {
                        value = cell.getDateCellValue();
                    } else {
                        throw new Exception("At row: %s, col: %s, can not discriminate type!");
                    }

                    map.put(keys[j], value);
                }
                list.add(map);
            }
        }catch (Exception e){
            throw new Exception("导入表格出错!", e);
        }finally {
            //关闭流
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return list;
    }


    /**
     * 导出最普通的 excel表格
     * @param fileNamePath  导出文件名称
     * @param sheetName     sheet名
     * @param list          数据
     * @param titles        第一行表头标题 数组
     * @param fieldNames    字段名称 数组
     *
     */
    public static <T> File defaultExport(String fileNamePath, String sheetName,List<T> list, String[] titles, String[] fieldNames) throws Exception{
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet;
        if (sheetName==null){
            sheetName = "DefaultSheet";
        }
        sheet = wb.createSheet(sheetName);
        //表头
        HSSFRow topRow = sheet.createRow(0);
        for (int i=0; i<titles.length; i++){
            fillCellWithValue(topRow.createCell(i),titles[i]);
        }
        String methodNameWithGet = "";
        String methodNameWithIs = "";
        Method method = null;
        T t = null;
        Object ret = null;
        // 遍历生成数据行，通过反射获取字段的get方法
        for (int i = 0; i < list.size(); i++) {
            t = list.get(i);
            HSSFRow row = sheet.createRow(i+1);
            Class<? extends Object> clazz = t.getClass();
            for(int j = 0; j < fieldNames.length; j++){
                methodNameWithGet = "get" + capitalize(fieldNames[j]);
                methodNameWithIs = "is" + capitalize(fieldNames[j]);
                try {
                    //先用get前缀试，如果获取不到试一下is前缀
                    method = clazz.getDeclaredMethod(methodNameWithGet);
                    if (method==null){
                        method =  clazz.getDeclaredMethod(methodNameWithIs);
                    }

                } catch (NoSuchMethodException e) {	//	不存在该方法，查看父类是否存在。此处只支持一级父类，若想支持更多，建议使用while循环
                    if(null != clazz.getSuperclass()) {
                        method = clazz.getSuperclass().getDeclaredMethod(methodNameWithGet);
                        if (method==null){
                            method =  clazz.getDeclaredMethod(methodNameWithIs);
                        }
                    }
                }
                if(null == method) {
                    throw new Exception(clazz.getName() + " don't have menthod --> " + methodNameWithGet + " or " + methodNameWithIs);
                }
                ret = method.invoke(t);
                fillCellWithValue(row.createCell(j), ret + "");
            }
        }
        //自定义表格宽度
        for (int i=0; i<fieldNames.length; i++){
            sheet.autoSizeColumn(i);
        }
        File file = null;

        OutputStream os = null;
        file = new File(fileNamePath);
        try {
            os = new FileOutputStream(file);
            wb.write(os);
            System.out.println("成功导出Excel");
        } catch (Exception e) {
            throw new Exception("write excel file error!", e);
        } finally {
            if(null != os) {
                os.flush();
                os.close();
            }
        }
        return file;
    }

    /**
     * 填用value值填充cell小格子
     * @param cell  格子
     * @param value 待填数据
     */
    private static void fillCellWithValue(HSSFCell cell, String value) {
        cell.setCellValue(new HSSFRichTextString(value));
    }

    /**
     * 拼接
      * @param str
     * @return
     */
    private static String capitalize(final String str) {
        int strLen;
        if (str == null || (strLen = str.length()) == 0) {
            return str;
        }

        final char firstChar = str.charAt(0);
        final char newChar = Character.toTitleCase(firstChar);
        if (firstChar == newChar) {
            // already capitalized
            return str;
        }

        char[] newChars = new char[strLen];
        newChars[0] = newChar;
        str.getChars(1,strLen, newChars, 1);
        return String.valueOf(newChars);
    }


    /**
     * 向已有的Excel表格中 插入一行数据，合并单元格
     * @param filePath      文件路径
     * @param theRow        行数
     * @param sheetName     sheet名字
     * @param note          文本
     * @param rowHeight     行高
     * @throws Exception
     */
    public static void insertRows(String filePath,int theRow,String sheetName,String note,int cellTotal,Short rowHeight)throws Exception{
        //得到POI对象
        HSSFWorkbook workbook = getWorkBook(filePath);
        //得到sheet
        HSSFSheet sheet = workbook.getSheet(sheetName);
        HSSFRow row = createRow(sheet,theRow,rowHeight);
        createCell(row,note);
        //合并单元格
        mergeCells(sheet,theRow,theRow,0,cellTotal);
        //保存文件
        updateFile(filePath,workbook);
    }



    /**
     * 找到需要插入的行数，并新建一个POI的row对象
     * @param sheet        sheet对象
     * @param rowIndex      行数
     * @return
     */
    private static HSSFRow createRow(HSSFSheet sheet, Integer rowIndex,Short height) {
        HSSFRow row = null;
        if (sheet.getRow(rowIndex) != null) {
            int lastRowNo = sheet.getLastRowNum();
            sheet.shiftRows(rowIndex, lastRowNo, 1);
        }
        row = sheet.createRow(rowIndex);
        if (height!=null){
            row.setHeight(height);
        }
        return row;
    }


    /**
     * 补一行
     * @param row   行数
     * @param note  文本
     * @return
     */
    private static HSSFCell createCell(HSSFRow row,String note) {
        HSSFCell cell = row.createCell(0);
        fillCellWithValue(cell,note);
        return cell;
    }

    /**
     * 合并单元格
     * @param sheet     所在的sheet
     * @param startRow  开始行
     * @param endRow    结束行
     * @param startCell 开始列
     * @param endCell   结束列
     */
    public static void mergeCells(Sheet sheet,int startRow,int endRow,int startCell,int endCell){
        sheet.addMergedRegion(new CellRangeAddress(startRow,endRow,startCell,endCell));
    }

    /**
     * 通过文件路径，获取workBook POI对象
     * @param filePath 文件路径
     * @return
     * @throws Exception
     */
    public static HSSFWorkbook getWorkBook(String filePath)throws Exception{
        HSSFWorkbook workbook;
        FileInputStream fis = null;
        File file = new File(filePath);
        try {
            if (!filePath.endsWith(".xls") && !filePath.endsWith(".xlsx")) {
                throw new Exception("文件格式有误!");
            }
            if (file!=null){
                fis = new FileInputStream(file);
                workbook = new HSSFWorkbook(fis);
            }else {
                throw new Exception("文件不存在！");
            }
        }catch (Exception e){
            throw new Exception("插入数据失败!", e);
        }finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return workbook;
    }

    /**
     * 通过文件路径和sheet名字 获取Sheet对象
     * @param filePath      文件路径
     * @param sheetName     sheet名称
     * @return
     * @throws Exception
     */
    public static Sheet getSheet(String filePath,String sheetName)throws Exception{
        HSSFWorkbook workbook = getWorkBook(filePath);
        Sheet sheet = workbook.getSheet(sheetName);
        return sheet;
    }


    /**
     * 保存workbook到文件，保存文件
     * @param filePath  保存文件路径
     * @param workbook  poi对象
     * @throws Exception
     */
    public static void updateFile(String filePath,Workbook workbook)throws Exception{
        //保存
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(filePath);
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("成功插入数据！");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
