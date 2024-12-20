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

# Slide

### Slide Class Overview

The `Slide` class, a vital component of the `OpenXMLOffice.Presentation` library, serves as a versatile tool for manipulating individual slides within a PowerPoint presentation. Whether extracted from an existing presentation or added new, this class provides a convenient handle for developers to modify the content of a single slide.

#### Key Features

1. **Content Manipulation:** The `Slide` class facilitates the addition of various elements to a slide, including charts, tables, and text. Developers can seamlessly enhance the visual richness of the presentation by incorporating diverse content types.
2. **Integration with PowerPoint Templates:** This class seamlessly integrates with existing PowerPoint templates, allowing developers to maintain a consistent look and feel throughout the presentation.

#### Example Usage

{% tabs %}
{% tab title="C#" %}
<pre class="language-csharp"><code class="lang-csharp">using OpenXMLOffice.Presentation;

public static SlideManipulation(Slide slide, DataCell[][] DataCells, AreaChartSetting AreaChartSetting){
    // Add New Chart To the Slide
    // Follow chart document for more info
    Chart chart = slide.AddChart(DataCells, AreaChartSetting);
    // To Add picture to the slide
    Picture picture = slide.AddPicture("sample.jpg", PictureSetting PictureSetting);
<strong>    // To Add Table
</strong><strong>    Table table = slide.AddTable(DataCells, TableSetting TableSetting);
</strong><strong>    // To Find a shape based on text from PPTX
</strong><strong>    Shape shape = slide.FindShapeByText("shape_id");
</strong><strong>}
</strong></code></pre>
{% endtab %}
{% endtabs %}

### TODO

Slide Setting Object for slide options

* Background
* Header/Footer
* Size
* Border, etc...
