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

# Textbox

Textbox Control to add and update Text Box

{% tabs %}
{% tab title="C#" %}
```csharp
Slide slide = powerPoint.GetSlideByIndex(0);
shapes3[0].ReplaceTextBox(slide, new TextBox(new G.TextBoxSetting()
	{
		textBlocks = new List<G.TextBlock>(){
			new(){
				text = "Move Slide To ",
				fontFamily = "Bernard MT Condensed"
			},
			new(){
				text = "Prev",
				fontSize = 25,
				isBold = true,
				textColor = "AAAAAA",
				hyperlinkProperties = new(){
					hyperlinkPropertyType = G.HyperlinkPropertyType.PREVIOUS_SLIDE,
				}
			}
		}.ToArray()
	}));
```
{% endtab %}
{% endtabs %}

### `TextBoxSetting` Options

<table><thead><tr><th width="200">Property</th><th width="245">Type</th><th>Details</th></tr></thead><tbody><tr><td>x</td><td>uint</td><td>Textbox Top Left X</td></tr><tr><td>y</td><td>uint</td><td>Textbox Top Left y</td></tr><tr><td>height</td><td>uint</td><td>Texbox Total Height</td></tr><tr><td>width</td><td>uint</td><td>Texbox Total Width</td></tr><tr><td>horizontalAlignment</td><td>HorizontalAlignmentValues?</td><td></td></tr><tr><td>textBlocks</td><td><a href="textbox.md#textblock-options">TextBlock</a>[]</td><td>Text box content as parts to have different style setting</td></tr><tr><td>shapeBackground</td><td>string?</td><td>Entire share background color</td></tr></tbody></table>

### `TextBlock` Options

|                     |                                                                             |                                          |
| ------------------- | --------------------------------------------------------------------------- | ---------------------------------------- |
| fontFamily          | string                                                                      | This section font family                 |
| fontSize            | int                                                                         | This section font size                   |
| isBold              | bool                                                                        | This section font family                 |
| isItalic            | bool                                                                        | This section font italic                 |
| isUnderline         | bool                                                                        | This section font underline              |
| text                | string                                                                      | This section text value                  |
| textBackground      | string                                                                      | This section text hightlight color       |
| textColor           | string                                                                      | This section text color                  |
| hyperlinkProperties | [HyperlinkProperties](../spreadsheet/shared.md#hyperlinkproperties-setting) | Hyperlink properties for each text block |
