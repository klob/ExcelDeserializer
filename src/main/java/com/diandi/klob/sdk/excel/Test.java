package com.diandi.klob.sdk.excel;

import java.io.File;
import java.util.List;

/**
 * *******************************************************************************
 * *********    Author : klob(kloblic@gmail.com) .
 * *********    Date : 2015-10-22  .
 * *********    Time : 17:54 .
 * *********    Version : 1.0
 * *********    Copyright © 2015, klob, All Rights Reserved
 * *******************************************************************************
 */
public class Test {
    public static void main(String[] args) {
        File file = new File(System.getProperty("user.dir") + "\\excel" + '\\' + "transcript.xls");
        ExcelDeserializer controller = new ExcelDeserializer();
        List<TestModel> models=controller.read(file, TestModel.class);
    }
}
