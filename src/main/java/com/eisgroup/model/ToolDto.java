package com.eisgroup.model;

import com.eisgroup.annotation.ClassTypeDesc;
import com.eisgroup.annotation.FieldDesc;

/**
 * @Description: 解析excel是用来接收每一行的信息
 * @Date: 2019/10/31 14:46
 * @author: xfei
 */
@ClassTypeDesc(value = "总表")
public class ToolDto {

    @FieldDesc(value = "序号")
    private String index;

    @FieldDesc(value = "产品型号")
    private String typeNumber;

    @FieldDesc(value = "产品描述(中文)")
    private String desc;

    @FieldDesc(value = "系列")
    private String type;

    @FieldDesc(value = "颜色")
    private String color;

    @FieldDesc(value = "19年市场价")
    private String price;

    @FieldDesc(value = "团购价")
    private String partyPrice;

    @FieldDesc(value = "数量")
    private String number;

    @FieldDesc(value = "总价")
    private String countPrice;

    public String getIndex() {
        return index;
    }

    public void setIndex(String index) {
        this.index = index;
    }

    public String getTypeNumber() {
        return typeNumber;
    }

    public void setTypeNumber(String typeNumber) {
        this.typeNumber = typeNumber;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public String getPrice() {
        return price;
    }

    public void setPrice(String price) {
        this.price = price;
    }

    public String getPartyPrice() {
        return partyPrice;
    }

    public void setPartyPrice(String partyPrice) {
        this.partyPrice = partyPrice;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public String getCountPrice() {
        return countPrice;
    }

    public void setCountPrice(String countPrice) {
        this.countPrice = countPrice;
    }
}
