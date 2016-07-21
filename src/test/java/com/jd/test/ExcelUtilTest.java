package com.jd.test;

import com.jd.excel.util.ExcelUtil;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by caozhifei on 2016/7/20.
 */
public class ExcelUtilTest {
    public static void main(String[] args) throws Exception {
        List<OrderPo> list = new ArrayList<OrderPo>();
        for(int i=0;i<5000;i++){
            OrderPo po = new OrderPo();
            po.setId(i);
            po.setName("jerry-"+i);
            po.setUid(1000+i);
            po.setCreated(new Date());
            list.add(po);
        }
//        LinkedHashMap<String,String> map = new LinkedHashMap<String, String>();
//        map.put("序号","id");
//        map.put("姓名","name");
//        map.put("用户ID","uid");
//        map.put("创建时间","created");
//        ExcelVo<OrderPo> excelVo = new ExcelVo<OrderPo>(map,list);
        ExcelUtil.createExcelFile("d:/util.xlsx", list);
    }
}
