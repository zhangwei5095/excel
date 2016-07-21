package com.jd.excel.model;

import java.util.LinkedHashMap;
import java.util.List;

/**
 * Created by caozhifei on 2016/7/20.
 */
public class ExcelVo<T> {
    /**
     * 创建的excel标题和传入的实体类中字段的对应关系，创建的excel标题顺序以map中的顺序为准
     */
    private LinkedHashMap<String,String> titleFiledMap;
    /**
     * 需要生成excel文件的集合数据
     */
    private List<T> list;

    public LinkedHashMap<String, String> getTitleFiledMap() {
        return titleFiledMap;
    }

    public void setTitleFiledMap(LinkedHashMap<String, String> titleFiledMap) {
        this.titleFiledMap = titleFiledMap;
    }

    public List<T> getList() {
        return list;
    }

    public void setList(List<T> list) {
        this.list = list;
    }

    public ExcelVo(LinkedHashMap<String, String> titleFiledMap, List<T> list) {
        this.titleFiledMap = titleFiledMap;
        this.list = list;
    }
}
