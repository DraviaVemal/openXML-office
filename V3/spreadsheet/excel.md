---
layout:
  title:
    visible: true
  description:
    visible: false
  tableOfContents:
    visible: true
  outline:
    visible: true
  pagination:
    visible: true
---

# Excel

The `Excel` class, an integral part of the `OpenXMLOffice.Spreadsheet` library, facilitates seamless interaction with Excel workbooks. Designed to simplify the creation and manipulation of Excel (.xlsx) files, this class provides a user-friendly interface for developers to efficiently handle data, worksheets, and formatting.

#### Key Features

1. **Effortless Initialization:** Initializing a new Excel workbook is simplified with the `Excel` class. Developers can swiftly create new workbooks or open existing ones, setting the stage for easy data management.
2. **Worksheet Manipulation:** The class offers intuitive methods for adding, deleting, and manipulating worksheets within a workbook. Developers can efficiently organize data and structure it across multiple sheets.
3. **Cell-Level Operations:** Granular control over individual cells is provided, allowing developers to set values, apply formatting, and perform various operations on specific cells within a worksheet.
4. **Data Import and Export:** The `Excel` class supports seamless data import from external sources and export to various formats. This enables efficient integration with external data sets and applications.

### Basic Code Samples

{% tabs %}
{% tab title="C#" %}
```csharp
Excel excel = new();
Worksheet worksheet = excel.AddSheet();
excel.SaveAs(string.Format("../../test-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
```
{% endtab %}
{% endtabs %}
