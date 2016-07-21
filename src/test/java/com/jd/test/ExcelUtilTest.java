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
        for(int i=0;i<500;i++){
            OrderPo po = new OrderPo();
            po.setId(i);
            po.setName("jerry-"+i);
            po.setUid(1000+i);
            po.setCreated(new Date());
            list.add(po);
        }

        List<OrderPo> list1 = new ArrayList<OrderPo>();
        for(int i=0;i<500;i++){
            OrderPo po = new OrderPo();
            po.setId(i);
            po.setName("tomcat-"+i);
            po.setUid(1000+i);
            po.setCreated(new Date());
            list1.add(po);
        }

        List<OrderPo> list2 = new ArrayList<OrderPo>();
        for(int i=0;i<500;i++){
            OrderPo po = new OrderPo();
            po.setId(i);
            po.setName("lucky-"+i);
            po.setUid(1000+i);
            po.setCreated(new Date());
            list2.add(po);
        }
//        LinkedHashMap<String,String> map = new LinkedHashMap<String, String>();
//        map.put("序号","id");
//        map.put("姓名","name");
//        map.put("用户ID","uid");
//        map.put("创建时间","created");
//        ExcelVo<OrderPo> excelVo = new ExcelVo<OrderPo>(map,list);
        //ExcelUtil.createExcelFile("d:/util.xlsx", list);
        List<List> lists = new ArrayList<List>();
        lists.add(list);
        lists.add(list1);
        lists.add(list2);
        ExcelUtil.createExcelZip(lists,"d:/list.zip");
        System.out.println("create ok");
    }
}
