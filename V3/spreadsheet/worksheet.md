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

# Worksheet

Note: New features related to this feature is been worked on release v2.8.x. Document is not upto date

Adding, Modifying a sheet from spreadsheet is handled by this class object

### Methods

<table><thead><tr><th width="185">Method</th><th width="263">Parameter/Return</th><th>Function</th></tr></thead><tbody><tr><td>GetSheetId</td><td>/string</td><td>Return current sheet id</td></tr><tr><td>GetSheetName</td><td>/string</td><td>Return current sheet name </td></tr><tr><td>SetColumn</td><td>cilumn,ColumnProperty</td><td>Set column property</td></tr><tr><td>SetRow</td><td>cellid,cellData,RowProperty</td><td>Set row property and data</td></tr><tr><td>AddPicture</td><td>filePath,PictureSetting/Picture</td><td>Add Picture to current slide</td></tr><tr><td>AddChart</td><td>DataRange,chartSetting/Chart</td><td>Add Chart to current slide</td></tr><tr><td>GetMergeCellList</td><td>/List&#x3C;MergeCellRange></td><td>Get existing merge range from current sheet</td></tr><tr><td>SetMergeCell</td><td>MergeCellRange/bool</td><td>Set new merge range if not affecting existing</td></tr><tr><td>RemoveMergeCell</td><td>MergeCellRange/bool</td><td>Remove any existing range within the caller range</td></tr></tbody></table>

### Sheet Code Samples

To add, remove and get sheet from excel

{% tabs %}
{% tab title="C#" %}
```csharp
// Adding new sheet to excel
Worksheet worksheet = excel.AddSheet();
Worksheet worksheet = excel.AddSheet("Data Sheet 2");
// Get an existing sheet from Excel
Worksheet worksheet = excel.GetWorksheet("Data Sheet 3");
// Remove existing sheet from Excel
Worksheet worksheet = excel.RemoveSheet("Sheet 1");
// Rename existing sheet
Worksheet worksheet = excel.RenameSheet("Data Sheet 2", "Sheet 1");
```
{% endtab %}
{% endtabs %}

### Sheet Column Settings Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
Worksheet worksheet = excel.AddSheet();
// Set Column property
worksheet.SetColumn("A1", new ColumnProperties()
	{
		width = 30
	});
```
{% endtab %}
{% endtabs %}

### `ColumnProperties` Options

<table><thead><tr><th width="120">Property</th><th width="107">Type</th><th>Details</th></tr></thead><tbody><tr><td>bestFit</td><td>bool</td><td>Auto bit column width based on content.</td></tr><tr><td>hidden</td><td>bool</td><td>Hide the column</td></tr><tr><td>width</td><td>double?</td><td>Set manual column width.</td></tr></tbody></table>

### Sheet Row Data and Settings Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
Worksheet worksheet = excel.AddSheet();
// Set Row data and setting starting from A1 Cell and move right
worksheet.SetRow("A1", 
	new DataCell[6]{
		new DataCell(){
			cellValue = "test1",
			dataType = CellDataType.STRING
		},
		 new DataCell(){
			cellValue = "test2",
			dataType = CellDataType.STRING
		},
		 new DataCell(){
			cellValue = "test3",
			dataType = CellDataType.STRING
		},
		 new DataCell(){
			cellValue = "test4",
			dataType = CellDataType.STRING,
			styleSetting = new(){
				fontSize = 20
			}
		},
		 new DataCell(){
			cellValue = "2.51",
			dataType = CellDataType.NUMBER,
			styleSetting = new(){
				numberFormat = "00.000",
			}
		},new(){
			cellValue = "5.51",
			dataType = CellDataType.NUMBER,
			styleSetting = new(){
				numberFormat = "₹ #,##0.00;₹ -#,##0.00",
			}
		}
	}, new RowProperties()
	{
		height = 20
	});
```
{% endtab %}
{% endtabs %}

### `DataCell` Options.

<table><thead><tr><th width="191">Property</th><th width="179">Type</th><th>Details</th></tr></thead><tbody><tr><td>cellValue</td><td>string?</td><td>Can be any value or null. Will be parsed based on <code>dataType</code></td></tr><tr><td>dataType</td><td>CellDataType</td><td>Refer to the data type present in <code>cellValue</code> property</td></tr><tr><td>styleSetting</td><td><a href="style.md#cellstylesetting-options">CellStyleSetting</a>?</td><td><strong>AVOID USING THIS.</strong> Used to set specific cell style. For optimised performance refer <a href="style.md">Style Component</a></td></tr><tr><td>styleId</td><td>uint?</td><td>Insert the style Id returened from <a href="style.md">Style Componenet</a></td></tr><tr><td>hyperlinkProperties</td><td><a href="shared.md#hyperlinkproperties-setting">HyperlinkProperties</a></td><td>Set hyperlink property for the current cell</td></tr></tbody></table>

### `RowProperties` Options

<table><thead><tr><th width="116">Property</th><th>Type</th><th>Details</th></tr></thead><tbody><tr><td>height</td><td>double?</td><td>Set row height property</td></tr><tr><td>hidden</td><td>bool</td><td>Hide the row</td></tr></tbody></table>
