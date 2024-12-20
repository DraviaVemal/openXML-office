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

# Shape

Note: New features related to this feature is been worked on release v3.x.x. Document is not upto date.

### Shape Class Overview

The `Shape` class, an essential component of the `OpenXMLOffice.Presentation` library, plays a pivotal role in template-based operations within PowerPoint presentations. This class serves as a fundamental object that enables developers to locate and manipulate shapes within a slide. When a shape is retrieved from a slide, the `Shape` object provides a powerful mechanism to precisely position and customize its properties based on the layout defined in the template.

#### Key Features

1. **Template-Based Operation:** The `Shape` class facilitates template-driven operations by allowing developers to locate and work with shapes in accordance with the predefined layout. This ensures consistency and adherence to the specified design.
2. **Positioning and Customization:** Developers can leverage the `Shape` object to precisely position and customize properties of shapes. This includes attributes such as size, color, and text content, providing granular control over the visual elements.

#### Example Usage

{% tabs %}
{% tab title="C#" %}
```csharp
using OpenXMLOffice.Presentation;

public static ShapeManipulation(Shape shape,Chart chart,Table table, TextBox textbox,Picture picture){
    // This will replace the shape object with passed picture object
    shape.ReplacePicture(picture);
    // This will replace the shape object with passed table object
    shape.ReplaceTable(table);
    // This will replace the shape object with passed chart object
    shape.ReplaceChart(chart);
    // This will replace the shape object with passed textbox object
    shape.ReplaceTextBox(textbox);
    // Just to Remove the Shape
    share.RemoveShape();
}
```
{% endtab %}
{% endtabs %}
