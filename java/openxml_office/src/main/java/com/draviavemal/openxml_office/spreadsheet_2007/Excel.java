package com.draviavemal.openxml_office.spreadsheet_2007;

import com.google.flatbuffers.FlatBufferBuilder;

public class Excel {

    static {
        System.loadLibrary("src/lib/openxmloffice_ffi");
    }

    private native byte create_excel(String fileName, byte[] buffer, int bufferSize, long[] outExcel);

    public Excel() {
        FlatBufferBuilder builder = new FlatBufferBuilder(1024);

    }

}