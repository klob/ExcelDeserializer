package com.diandi.klob.sdk.excel;

import java.util.ArrayList;
import java.util.List;

/**
 * *******************************************************************************
 * *********    Author : klob(kloblic@gmail.com) .
 * *********    Date : 2015-10-22  .
 * *********    Time : 17:04 .
 * *********    Version : 1.0
 * *********    Copyright © 2015, klob, All Rights Reserved
 * *******************************************************************************
 */
public class ExcelHeader {
    public List<String> header=new ArrayList<>();

    public ExcelHeader(List<String> header) {
        this.header = header;
    }

    public void add(String filed)
    {
        header.add(filed);
    }

    public String getField(int index)
    {
        return header.get(index);
    }
    @Override
    public String toString() {
        return "ExcelHeader{" +
                "header=" + header +
                '}';
    }
}
