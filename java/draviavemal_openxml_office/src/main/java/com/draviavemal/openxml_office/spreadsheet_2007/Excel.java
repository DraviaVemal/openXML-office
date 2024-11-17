package com.draviavemal.openxml_office.spreadsheet_2007;

import openxml_office_ffi.spreadsheet_2007.*;
import com.google.flatbuffers.FlatBufferBuilder;

public class Excel {

    static {
        System.loadLibrary("src/main/resources/lib/openxmloffice_ffi");
    }

    public native byte excelCreate(String fileName, byte[] buffer, int bufferSize, long[] outExcel,
            long[] outError);

    private native byte create_excel(String fileName, byte[] buffer, int bufferSize, long[] outExcel);

    public Excel() {
        FlatBufferBuilder builder = new FlatBufferBuilder(1024);
        ExcelPropertiesModel.startExcelPropertiesModel(builder);
        ExcelPropertiesModel.addIsInMemory(builder, false);
        int excelPropertiesModel = ExcelPropertiesModel.endExcelPropertiesModel(builder);
    }
}