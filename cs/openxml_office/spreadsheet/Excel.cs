using System;
using System.Runtime.InteropServices;
using Google.FlatBuffers;
using draviavemal.openxml_office.global_2007;
using openxml_office.spreadsheet_2007;

namespace draviavemal.openxml_office.spreadsheet_2007
{
    /// <summary>
    /// This class serves as a versatile tool for working with Excel spreadsheets, built upon the
    /// foundation of the OpenXML SDK. This class offers a wide range of functionalities for
    /// handling Excel-related objects and operation It is designed to simplify tasks related to
    /// Excel file manipulation, including the creation of new Excel files, reading and updating
    /// existing files, and processing Excel data from stream
    /// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
    /// </summary>
    public class Excel : PrivacyProperties
    {
        private readonly IntPtr ffiExcel;

        /// <summary>
        /// 
        /// </summary>
        [DllImport("lib/openxmloffice_ffi", CallingConvention = CallingConvention.Cdecl)]
        public static extern sbyte create_excel(
            [MarshalAs(UnmanagedType.LPStr)] string optional_string,
            IntPtr buffer,
            int length,
            out IntPtr excel
        );

        /// <summary>
        /// 
        /// </summary>
        [DllImport("lib/openxmloffice_ffi", CallingConvention = CallingConvention.Cdecl)]
        public static extern sbyte save_as(
            IntPtr excelPtr,
            [MarshalAs(UnmanagedType.LPStr)] string file_name
        );

        /// <summary>
        /// Create New file in the system
        /// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
        /// </summary>
        public Excel(ExcelProperties excelProperties)
        {
            FlatBufferBuilder builder = new(1024);
            ExcelPropertiesModel.AddIsInMemory(builder, excelProperties.IsInMemory);
            Offset<ExcelPropertiesModel> excelPropertiesModel = ExcelPropertiesModel.EndExcelPropertiesModel(builder);
            builder.Finish(excelPropertiesModel.Value);
            byte[] buffer = builder.SizedByteArray();
            int bufferSize = buffer.Length;
            unsafe
            {
                fixed (byte* bufferPtr = buffer)
                {
                    IntPtr bufferIntPtr = (IntPtr)bufferPtr;
                    StatusCode.ProcessStatusCode(create_excel(null, bufferIntPtr, bufferSize, out ffiExcel));
                }
            }
        }

        /// <summary>
		/// Works with in memory object can be saved to file at later point.
		/// Source file will be cloned and released. hence can be replace by saveAs method if you want to update the same file.
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
        public Excel(string fileName, ExcelProperties excelProperties)
        {
            FlatBufferBuilder builder = new(1024);
            builder.StartTable(1);
            ExcelPropertiesModel.AddIsInMemory(builder, excelProperties.IsInMemory);
            Offset<ExcelPropertiesModel> excelPropertiesModel = ExcelPropertiesModel.EndExcelPropertiesModel(builder);
            builder.Finish(excelPropertiesModel.Value);
            byte[] buffer = builder.SizedByteArray();
            int bufferSize = buffer.Length;
            unsafe
            {
                fixed (byte* bufferPtr = buffer)
                {
                    IntPtr bufferIntPtr = (IntPtr)bufferPtr;
                    StatusCode.ProcessStatusCode(create_excel(fileName, bufferIntPtr, bufferSize, out ffiExcel));
                }
            }
        }

        /// <summary>
        /// Even on edit file OpenXML-Office Will clone the source and work on top of it to protect the integrity of source file.
        /// You can save the document at the end of lifecycle targeting the edit file to update or new file.
        /// This is supported for both file path and data stream
        /// </summary>
        public void SaveAs(string filePath)
        {
            StatusCode.ProcessStatusCode(save_as(ffiExcel, filePath));
        }

    }
}
