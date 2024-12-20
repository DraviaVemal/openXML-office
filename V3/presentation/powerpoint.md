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

# PowerPoint

The `Powerpoint` class, a core component of the `OpenXMLOffice.Presentation` library, empowers developers to create, open, and manipulate PowerPoint (.pptx) files with ease. Whether generating new presentations or working with existing ones, this class provides a simple yet powerful interface for efficient content manipulation. Once modifications are complete, users can effortlessly save the updated presentation.

### Usage, Options and Examples

Create or open a pptx file from path

{% tabs %}
{% tab title="C#" %}
```csharp
public static CreateNew(){
    PowerPoint powerPoint = new();
    powerPoint.SaveAs("test.pptx");
}

public static OpenExisting(){
    PowerPoint powerPoint = new("test.pptx",true, null);
    powerPoint.SaveAs("/NewPath/file.pptx");
}
```
{% endtab %}
{% endtabs %}

Create or open a pptx object using a stream

{% tabs %}
{% tab title="C#" %}
```csharp
public static CreateUsingStream(Stream stream){
    PowerPoint powerPoint = new();
    powerPoint.SaveAs(stream);
}
```
{% endtab %}
{% endtabs %}

Sample using most of the exposed functions

{% tabs %}
{% tab title="C#" %}
```csharp
public static CreateNew(){
    PowerPoint powerPoint = new();
    // Add Blank Slide To the Blank Presentation
    // Return Slide Object that can be used to do slide level operation
    Slide slide = powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
    powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
    Slide slide1 = powerPoint.GetSlideByIndex(1);
    // Move the Slide Order
    powerPoint.MoveSlideByIndex(1,0);
    // Remove Slide and its content from Presentation
    powerPoint.RemoveSlideByIndex(0);
    // Save the Opened Presentation
    powerPoint.SaveAs("test.pptx");
}
```
{% endtab %}
{% endtabs %}

### `PresentationProperties` Options

| Property     | Type                                                               | Details                                  |
| ------------ | ------------------------------------------------------------------ | ---------------------------------------- |
| settings     | [PresentationSettings](powerpoint.md#presentationsettings-options) | Provides Presentation setting options    |
| slideMasters | Dictionary\<string, PresentationSlideMaster>?                      | Multislide master support is in pipeline |
| theme        | ThemePallet                                                        | Color template for overall presentation  |

### `PresentationSettings` Options

<table><thead><tr><th width="318">Property</th><th width="85">Type</th><th>Details</th></tr></thead><tbody><tr><td>isMultiSlideMasterPartPresentation</td><td>bool</td><td>Get or Set Multslide master option</td></tr><tr><td>isMultiThemePresentation</td><td>bool</td><td>Get or Set Mult theme option</td></tr></tbody></table>

### TODO

* Multi Slide Master Support
* Each Slide Master Theme Support
