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

# Table

The `Table` class, a dynamic feature within the `OpenXMLOffice.Presentation` library, provides developers with a robust toolset for effortlessly incorporating tables into PowerPoint presentations. This class offers extensive support for diverse table configurations, allowing users to create, modify, and enhance tables in a slide with ease. Developers can leverage the Table class to add new tables.

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
Slide slide = powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
slide.AddTable(CreateTableRowPayload(10), new TableSetting()
{
	name = "New Table",
	widthType = TableSetting.WidthOptionValues.PERCENTAGE,
	tableColumnWidth = new() { 80, 20 },
	x = (uint)G.ConverterUtils.PixelsToEmu(10),
	y = (uint)G.ConverterUtils.PixelsToEmu(10)
});
```
{% endtab %}
{% endtabs %}

### `TableSetting` Options

<table><thead><tr><th width="188">Property</th><th width="176">Type</th><th>Details</th></tr></thead><tbody><tr><td>name</td><td>string</td><td>Table Internal name for reference. Default: Table 1</td></tr><tr><td>height</td><td>uint</td><td>Table overall height in EMU.<br>Note: If the specified height is insufficient, the PowerPoint (PPT) application will automatically use the minimum required height.<br>Default : 741680</td></tr><tr><td>width</td><td>uint</td><td>Table overall width in EMU<br>Default : 8128000</td></tr><tr><td>x</td><td>uint</td><td>This parameter precisely determines the X position of the entire chart. Default: 0</td></tr><tr><td>y</td><td>uint</td><td>This parameter precisely determines the Y position of the entire chart. Default : 0</td></tr><tr><td>widthType</td><td><a href="table.md#widthoptionvalues-width-type-additional-details">WidthOptionValues</a></td><td>This parameter affects how value in <code>tableColumnWidth</code> is been used. Refer <a href="table.md#width-type-additional-details">additional details</a>.<br>Default : Auto</td></tr><tr><td>tableColumnWidth</td><td>List&#x3C;float></td><td>The float value is accessed or ignored based on <code>widthType</code> setting.</td></tr></tbody></table>

### `TableRow` Options

<table><thead><tr><th width="169">Property</th><th width="146">Type</th><th>Details</th></tr></thead><tbody><tr><td>height</td><td>int</td><td>Configures row height, measured in EMU (English Metric Unit).<br>Default : 370840</td></tr><tr><td>rowBackground</td><td>string?</td><td>Configure row background color using hex code (without #).</td></tr><tr><td>textColor</td><td>string</td><td>Configure row text color using hex code (without #).</td></tr><tr><td>tableCells</td><td>List&#x3C;<a href="table.md#tablecell-options">TableCell</a>></td><td>Contains the list of cell that needs to be inserted into the current row.</td></tr></tbody></table>

### `TableCell` Options

<table><thead><tr><th width="203">Property</th><th width="242">Type</th><th>Details</th></tr></thead><tbody><tr><td>value</td><td>string?</td><td>Value of specific table cell</td></tr><tr><td>rowSpan</td><td>uint</td><td>Number of cell Rows of same column to be merged</td></tr><tr><td>columnSpan</td><td>uint</td><td>Number of cell Columns to be merged in same row</td></tr><tr><td>horizontalAlignment</td><td>HorizontalAlignmentValues?</td><td>Horizontal alignement property</td></tr><tr><td>verticalAlignment</td><td>VerticalAlignmentValues?</td><td>Vertical alignement property</td></tr><tr><td>borderSettings</td><td><a href="table.md#tablebordersettings-options">TableBorderSettings</a></td><td>Table Border setting options</td></tr><tr><td>cellBackground</td><td>string?</td><td>Cell Specific background color</td></tr><tr><td>textColor</td><td>string</td><td>Cell Specific text color</td></tr><tr><td>fontFamily</td><td>string</td><td>Cell Specific font family</td></tr><tr><td>fontSize</td><td>int</td><td>Cell Specific font size</td></tr><tr><td>isBold</td><td>bool</td><td>Set cell content bold</td></tr><tr><td>isItalic</td><td>bool</td><td>Set cell content Italic</td></tr><tr><td>isUnderline</td><td>bool</td><td>Set cell content underline status</td></tr><tr><td></td><td></td><td></td></tr></tbody></table>

### `TableBorderSettings` Options

<table><thead><tr><th width="265">Property</th><th width="180">Type</th><th>Details</th></tr></thead><tbody><tr><td>leftBorder</td><td><a href="table.md#tablebordersetting-options">TableBorderSetting</a></td><td>Left Border Related Setting</td></tr><tr><td>topBorder</td><td><a href="table.md#tablebordersetting-options">TableBorderSetting</a></td><td>Top Border Related Setting</td></tr><tr><td>rightBorder</td><td><a href="table.md#tablebordersetting-options">TableBorderSetting</a></td><td>Right Border Related Setting</td></tr><tr><td>bottomBorder</td><td><a href="table.md#tablebordersetting-options">TableBorderSetting</a></td><td>Bottom Border Related Setting</td></tr><tr><td>topLeftToBottomRightBorder</td><td><a href="table.md#tablebordersetting-options">TableBorderSetting</a></td><td>TODO</td></tr><tr><td>bottomLeftToTopRightBorder</td><td><a href="table.md#tablebordersetting-options">TableBorderSetting</a></td><td>TODO</td></tr></tbody></table>

### `TableBorderSetting` Options

<table><thead><tr><th>Property</th><th width="259">Type</th><th></th></tr></thead><tbody><tr><td>showBorder</td><td>bool</td><td>Set border status</td></tr><tr><td>borderColor</td><td>string</td><td>Set border color</td></tr><tr><td>width</td><td>float</td><td>Set border width</td></tr><tr><td>borderStyle</td><td>BorderStyleValues</td><td>Border line style</td></tr><tr><td>dashStyle</td><td>DrawingPresetLineDashValues</td><td>Border dash style</td></tr></tbody></table>

#### `WidthOptionValues` \*Width Type additional Details

<table><thead><tr><th width="156">Option</th><th>Behaviour</th></tr></thead><tbody><tr><td>AUTO</td><td>Ignore User Width value and space the colum equally.</td></tr><tr><td>EMU</td><td>(English Metric Units) Direct PPT standard Sizing 1 Inch * 914400 EMU's</td></tr><tr><td>PIXEL</td><td>Based on Target DPI the pixel is converted to EMU and used when running</td></tr><tr><td>PERCENTAGE</td><td>0-100 Width percentage split for each column</td></tr><tr><td>RATIO</td><td>0-10 Width ratio of each column</td></tr></tbody></table>
