package com.jd.excel.factory;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**扩展工作簿工厂类
 * Created by caozhifei on 2016/7/20.
 */
public class ExtendWorkbookFactory extends WorkbookFactory {
    /**
     * 创建HSSF类型的excek工作簿，office2003
     * @return
     */
    public static Workbook createHSSF(){
        return  new HSSFWorkbook();
    }

    /**
     * 创建HSSF类型的excek工作簿，office2007
     * @return
     */
    public static Workbook createXSSF(){
        return  new XSSFWorkbook();
    }
}
