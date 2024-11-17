from openxml_office_ffi.spreadsheet_2007 import ExcelPropertiesModel;
from flatbuffers import Builder;

class Excel:
    def __init__(self):
        builder = Builder(1024)
        ExcelPropertiesModel.ExcelPropertiesModelStart(builder)
        ExcelPropertiesModel.AddIsInMemory(builder,True)
        ExcelPropertiesModel.ExcelPropertiesModelEnd(builder)   
        pass
    def SaveAs(self):
        pass