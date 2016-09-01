package com.jd.excel.util;

import com.jd.excel.annotation.Excel;
import com.jd.excel.factory.ExtendWorkbookFactory;
import com.jd.excel.model.ExcelVo;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * EXCEL文件创建工具类
 * Created by caozhifei on 2016/7/20.
 */
public class ExcelUtil {
    private final static String DATE_FORMAT = "m/d/yy h:mm";

    /**
     * 创建excel文件压缩文件
     *
     * @param lists
     * @return
     * @throws Exception
     */
    public static void createExcelZip(List<List> lists, String zipName) throws Exception {
        FileOutputStream fileOut = new FileOutputStream(zipName);
        try {
            byte[] bytes = createExcelZip(lists);
            fileOut.write(bytes);
        } finally {
            fileOut.flush();
            fileOut.close();
        }
    }

    /**
     * 创建excel文件压缩流，并返回字节数组
     *
     * @param lists
     * @return
     * @throws Exception
     */
    public static byte[] createExcelZip(List<List> lists) throws Exception {
        if (lists == null || lists.isEmpty()) {
            return null;
        }
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ZipOutputStream out = new ZipOutputStream(bos);
        int i = 0;
        for (List<?> list : lists) {
            i++;
            byte[] bytes = createExcel(list, true);
            out.putNextEntry(new ZipEntry(i + ".xlsx"));
            out.write(bytes);
            out.flush();
        }
        out.close();
        return bos.toByteArray();
    }


    /**
     * 根据元数据生成excel文件
     *
     * @param fileName
     * @param list
     * @throws Exception
     */
    public static void createExcelFile(String fileName, List<?> list) throws Exception {
        FileOutputStream fileOut = new FileOutputStream(fileName);
        try {
            byte[] bytes = createExcel(list, true);
            fileOut.write(bytes);
        } finally {
            fileOut.flush();
            fileOut.close();
        }
    }

    /**
     * 创建excel文件，并返回字节数组
     *
     * @param list
     * @return
     * @throws Exception
     */
    public static byte[] createExcel(List<?> list) throws Exception {
        return createExcel(list, true);
    }

    /**
     * 生成excel文件
     *
     * @param fileName 生成的文件名
     * @param list     需要生成excel文件的元数据
     * @param isXSSF
     * @throws Exception
     */
    public static void createExcelFile(String fileName, List<?> list, boolean isXSSF) throws Exception {
        FileOutputStream fileOut = new FileOutputStream(fileName);
        try {
            byte[] bytes = createExcel(list, isXSSF);
            fileOut.write(bytes);
        } finally {
            fileOut.flush();
            fileOut.close();
        }
    }

    /**
     * 基于类属性主键创建excel文件字节数组
     *
     * @param list
     * @throws Exception
     */
    public static byte[] createExcel(List<?> list, boolean isXSSF) throws Exception {
        if (list == null || list.isEmpty()) {
            throw new Exception("list is null");
        }
        Class type = list.get(0).getClass();
        Field[] fields = type.getDeclaredFields();
        LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();
        for (Field field : fields) {
            if (field.isAnnotationPresent(Excel.class)) {
                Excel excel = field.getAnnotation(Excel.class);
                map.put(excel.columnName(), field.getName());
            }
        }
        if (map.isEmpty()) {
            throw new Exception("can not find Excel Annotation");
        }
        ExcelVo excelVo = new ExcelVo(map, list);
        return createExcel(excelVo, isXSSF);
    }

    public static void createExcelFile(String fileName, ExcelVo excelVo) throws Exception {
        createExcelFile(fileName, excelVo, true);
    }

    /**
     * 创建excel文件
     *
     * @param fileName
     * @param excelVo
     * @param isXSSF
     * @throws Exception
     */
    public static void createExcelFile(String fileName, ExcelVo excelVo, boolean isXSSF) throws Exception {
        FileOutputStream fileOut = new FileOutputStream(fileName);
        try {
            byte[] bytes = createExcel(excelVo, isXSSF);
            fileOut.write(bytes);
        } finally {
            fileOut.flush();
            fileOut.close();
        }
    }

    /**
     * @param excelVo
     * @return
     * @throws Exception
     */
    public static byte[] createExcel(ExcelVo excelVo) throws Exception {
        return createExcel(excelVo, true);
    }

    /**
     * 创建excel文件字节数组
     *
     * @param excelVo
     * @param isXSSF
     * @return
     * @throws Exception
     */
    public static byte[] createExcel(ExcelVo excelVo, boolean isXSSF) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            Workbook wb = null;
            if (isXSSF) {
                wb = ExtendWorkbookFactory.createXSSF();
            } else {
                wb = ExtendWorkbookFactory.createHSSF();
            }
            CreationHelper createHelper = wb.getCreationHelper();
            Sheet sheet = wb.createSheet("sheet1");
            //填充标题行
            fillTitle(sheet, createHelper, excelVo);

            CellStyle cellStyle = wb.createCellStyle();
            cellStyle.setDataFormat(
                    createHelper.createDataFormat().getFormat(DATE_FORMAT));
            //填充内容行
            fillContent(sheet, createHelper, excelVo, cellStyle);
            wb.write(out);
        } finally {
            out.flush();
            out.close();
        }
        return out.toByteArray();
    }

    public static List<Map> parseExcel(String path) throws Exception {
        if (path == null || path == "") {
            throw new Exception("path is null;" + path);
        }
        if (path.endsWith(".xls") || path.endsWith(".xlsx")) {
            FileInputStream is = new FileInputStream(new File(path));
            boolean isXSSF = path.endsWith(".xlsx") ? true : false;
            List<Map> list = parseExcel(is, isXSSF);
            return list;
        } else {
            throw new Exception("file is invalid;" + path);
        }
    }

    public static List<Map> parseExcel(InputStream is, boolean isXSSF) throws IOException {
        List<Map> list = new ArrayList<Map>();
        Workbook book;
        if (isXSSF) {
            book = new XSSFWorkbook(is);
        } else {
            book = new HSSFWorkbook(is);
        }
        Sheet sheet = book.getSheetAt(0);
        for (int i = 1; i < sheet.getLastRowNum() + 1; i++) {
            Row row = sheet.getRow(i);
            if (null == row) {
                break;
            }
            Map map = new HashMap();
            for (int j = 0; j < (row.getLastCellNum() + 1); j++) {
                Cell cell = sheet.getRow(i).getCell(j);
                if (cell == null) {
                    break;
                }
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    map.put(sheet.getRow(sheet.getFirstRowNum()).getCell(j).getStringCellValue(), cell.getNumericCellValue());
                } else {
                    map.put(sheet.getRow(sheet.getFirstRowNum()).getCell(j).getStringCellValue(), cell.getStringCellValue());
                }
            }
            list.add(map);
        }
        return list;
    }

    /**
     * 根据传入的map映射关系，制作标题行
     *
     * @param sheet
     * @param createHelper
     */
    private static void fillTitle(Sheet sheet, CreationHelper createHelper, ExcelVo excelVo) {
        LinkedHashMap<String, String> titleFiledMap = excelVo.getTitleFiledMap();
        Row row = sheet.createRow(0);
        int column = 0;
        for (String key : titleFiledMap.keySet()) {
            row.createCell(column).setCellValue(createHelper.createRichTextString(key));
            column++;
        }

    }

    /**
     * 填充内容行
     *
     * @param sheet
     * @param createHelper
     * @param excelVo
     */
    private static void fillContent(Sheet sheet, CreationHelper createHelper, ExcelVo excelVo, CellStyle cellStyle) throws Exception {
        List list = excelVo.getList();
        LinkedHashMap<String, String> titleFiledMap = excelVo.getTitleFiledMap();
        int totalRow = list.size();
        for (int rowNum = 0; rowNum < totalRow; rowNum++) {
            Row row = sheet.createRow(rowNum + 1);//从第二行开始
            int column = 0;
            for (String key : titleFiledMap.keySet()) {
                String filed = titleFiledMap.get(key);
                Class filedType = PropertyUtils.getPropertyType(list.get(rowNum), filed);
                Object filedValue = PropertyUtils.getProperty(list.get(rowNum), filed);
                if (can2String(filedType)) {
                    if (String.class.equals(filedType)) {
                        row.createCell(column).setCellValue(createHelper.createRichTextString((String) filedValue));
                    } else {
                        row.createCell(column).setCellValue(String.valueOf(filedValue));
                    }
                } else if (Date.class.equals(filedType)) {
                    cellStyle.setDataFormat(
                            createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
                    Cell cell = row.createCell(column);
                    cell.setCellValue((Date) filedValue);
                    cell.setCellStyle(cellStyle);
                } else if (Double.class.equals(filedType) || double.class.equals(filedType) || Float.class.equals(filedType) || float.class.equals(filedType)) {
                    row.createCell(column).setCellValue((Double) filedValue);
                } else if (Boolean.class.equals(filedType) || boolean.class.equals(filedType)) {
                    row.createCell(column).setCellValue((Boolean) filedValue);
                }
                column++;
            }
        }
    }

    /**
     * 是否可以按照字符串形式写入单元格
     *
     * @param filedType
     * @return
     */
    private static boolean can2String(Class filedType) {
        if (String.class.equals(filedType) || Integer.class.equals(filedType) || Long.class.equals(filedType) || Byte.class.equals(filedType) || Short.class.equals(filedType) || int.class.equals(filedType) || long.class.equals(filedType) || byte.class.equals(filedType) || short.class.equals(filedType)) {
            return true;
        }
        return false;
    }

    public static void main(String[] args) throws Exception {
        String path = "d:/util.xlsx";
        List list = parseExcel(path);
        System.out.println(list);
    }
}
