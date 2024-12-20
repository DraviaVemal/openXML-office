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

# Picture

Details about adding and manipulating picture to a worksheet

### Basic Code Sample

```csharp
public void AddPicture()
	{
		Worksheet worksheet = excel.AddSheet("Data4");
		worksheet.AddPicture("./TestFiles/tom_and_jerry.jpg", new()
		{
			imageType = ImageType.JPEG,
			from = new()
			{
				column = 6,
				row = 6
			},
			to = new()
			{
				column = 8,
				row = 8
			}
		});
		Assert.IsTrue(true);
	}
```

### `ExcelPictureSetting` Options

<table><thead><tr><th width="197">Property</th><th width="203">Type</th><th>Details</th></tr></thead><tbody><tr><td>hyperlinkProperties</td><td><a href="shared.md#hyperlinkproperties-setting">HyperlinkProperties</a></td><td>Attach hyperlink to the image</td></tr><tr><td>imageType</td><td>ImageType</td><td>Image extension type</td></tr><tr><td>anchorEditType</td><td>AnchorEditType</td><td>Mode of picture starting point</td></tr><tr><td>from</td><td><a href="picture.md#anchorposition-options">AnchorPosition</a></td><td>Top Left coordinate</td></tr><tr><td>to</td><td><a href="picture.md#anchorposition-options">AnchorPosition</a></td><td>Bottom right X coordinate</td></tr></tbody></table>

### `AnchorPosition` Options

<table><thead><tr><th width="158">Property</th><th width="67">Type</th><th>Details</th></tr></thead><tbody><tr><td>column</td><td>uint</td><td></td></tr><tr><td>columnOffset</td><td>uint</td><td></td></tr><tr><td>row</td><td>uint</td><td></td></tr><tr><td>rowOffset</td><td>uint</td><td></td></tr></tbody></table>
