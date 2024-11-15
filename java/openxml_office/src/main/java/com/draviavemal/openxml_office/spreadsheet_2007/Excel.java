package com.draviavemal.openxml_office.spreadsheet_2007;

import openxml_office.spreadsheet_2007.*;
import com.google.flatbuffers.FlatBufferBuilder;

public class Excel {

    static {
        System.loadLibrary("src/main/resources/lib/openxmloffice_ffi");
    }

    private native byte create_excel(String fileName, byte[] buffer, int bufferSize, long[] outExcel);

    public Excel() {
        FlatBufferBuilder builder = new FlatBufferBuilder(1024);
        byte[] inventory = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
        ExcelPropertiesModel.startExcelPropertiesModel(builder);
        ExcelPropertiesModel.addIsInMemory(builder, false);
        int excelPropertiesModel = ExcelPropertiesModel.endExcelPropertiesModel(builder);
    }
}