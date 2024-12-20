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

# Style

Style is instance object per excel created to maintain for performance of the document creation process. This istance provides a common handle to set a style combination once refer it across the spreadsheet.

It is highly recomended to use this stratgy and reduce the number of cycle the `openXMLOffice.Excel` cell row insert life take to find the style id for the repeated setting you are makking using the cell style property. Use the styleId thats given as response from this object to optimise the document creation process. &#x20;

### Sample Code

{% tabs %}
{% tab title="C#" %}
```csharp
// To Get the style Id
uint styleId = excel.GetStyleId(new CellStyleSetting()
	{
		isBold = true,
		borderLeft = new()
		{
			style = BorderSetting.StyleValues.THICK
		},
		backgroundColor = "112233"
	});
// Use the Style Id
worksheet.SetRow("A1", new DataCell[6]{
	new(){
		cellValue = "test1",
		dataType = CellDataType.STRING,
		styleId = styleId
	}, new RowProperties());

```
{% endtab %}
{% endtabs %}

### `CellStyleSetting` Options

<table><thead><tr><th width="202">Property</th><th width="216">Type</th><th>Details</th></tr></thead><tbody><tr><td>backgroundColor</td><td>string?</td><td>Cell background color</td></tr><tr><td>borderBottom</td><td><a href="style.md#bordersetting-options">BorderSetting</a></td><td>Bottom Border Setting</td></tr><tr><td>borderLeft</td><td><a href="style.md#bordersetting-options">BorderSetting</a></td><td>Left Border Setting</td></tr><tr><td>borderRight</td><td><a href="style.md#bordersetting-options">BorderSetting</a></td><td>Right Border Setting</td></tr><tr><td>borderTop</td><td><a href="style.md#bordersetting-options">BorderSetting</a></td><td>Top Border Setting</td></tr><tr><td>fontFamily</td><td>string</td><td>Font family of the cell content</td></tr><tr><td>fontSize</td><td>uint</td><td>Font size of the cell content</td></tr><tr><td>foregroundColor</td><td>string</td><td>Cell foreground color</td></tr><tr><td>isBold</td><td>bool</td><td>Set cell content bold</td></tr><tr><td>isDoubleUnderline</td><td>bool</td><td>Set cell content underline</td></tr><tr><td>isItalic</td><td>bool</td><td>Set cell content italic</td></tr><tr><td>isUnderline</td><td>bool</td><td>Set cell content underline</td></tr><tr><td>isWrapText</td><td>bool</td><td>Set cell content auto wrap</td></tr><tr><td>numberFormat</td><td>string</td><td>Number formating for the cell content</td></tr><tr><td>textColor</td><td>string</td><td>Cell Text Color</td></tr><tr><td>horizontalAlignment</td><td>horizontalAlignment</td><td>Cell content horizontal alignment</td></tr><tr><td>verticalAlignment</td><td>VerticalAlignmentValues</td><td>Cell content vertical alignment</td></tr></tbody></table>

### `BorderSetting` Options

<table><thead><tr><th width="109">Property</th><th width="122">Type</th><th>Details</th></tr></thead><tbody><tr><td>color</td><td>string</td><td>Boder Color</td></tr><tr><td>style</td><td>StyleValues</td><td>Border Line Style</td></tr></tbody></table>
