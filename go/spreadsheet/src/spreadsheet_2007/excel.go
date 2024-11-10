package spreadsheet_2007

/*
#cgo LDFLAGS: -L./lib -lopenxmloffice_ffi
#include <stdint.h>
#include <string.h>
*/
import "C"
import "unsafe"

type Excel struct {
	excelPtr *C.int32_t
}

type ExcelProperties struct {
}

func NewExcel(excel_properties Excel) *Excel {
	cFileName := C.CString(fileName)
	defer C.free(unsafe.Pointer(cFileName))

	cBuffer := C.CBytes(buffer)
	defer C.free(unsafe.Pointer(cBuffer))

	var outExcel *C.Excel

	result := C.create_excel(cFileName, (*C.uchar)(cBuffer), C.size_t(len(buffer)), (**C.Excel)(&outExcel))
	return int8(result)
}
