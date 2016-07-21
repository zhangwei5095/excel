package com.jd.test;

import com.jd.excel.annotation.Excel;

import java.util.Date;

/**
 * Created by caozhifei on 2016/7/20.
 */
public class OrderPo {
    @Excel(columnName = "主键ID")
    private int id;
    @Excel(columnName = "商品名称")
    private String name;
    @Excel(columnName = "用户的ID")
    private Integer uid;
    @Excel(columnName = "下单时间")
    private Date created;
    private String image;


    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getUid() {
        return uid;
    }

    public void setUid(Integer uid) {
        this.uid = uid;
    }

    public Date getCreated() {
        return created;
    }

    public void setCreated(Date created) {
        this.created = created;
    }

    public String getImage() {
        return image;
    }

    public void setImage(String image) {
        this.image = image;
    }

    public OrderPo(int id, String name, Integer uid, Date created) {
        this.id = id;
        this.name = name;
        this.uid = uid;
        this.created = created;
    }

    public OrderPo() {
    }
}
