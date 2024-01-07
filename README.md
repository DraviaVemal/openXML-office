# Status

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/5b420a599805426ab8a990a1a741247a)](https://app.codacy.com/gh/DraviaVemal/OpenXMLOffice/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade) [![Downloads](https://img.shields.io/nuget/dt/OpenXMLOffice.Presentation.svg)](https://www.nuget.org/packages/OpenXMLOffice.Presentation)

# OpenXMLOffice

OpenXMLOffice is an advanced .NET library that leverages the power of OpenXML SDK 3.0 to streamline the creation and manipulation of Office documents, with a primary focus on Excel, Word, and PowerPoint files. Our mission is to enhance the document creation experience for developers by providing intuitive namespaces, classes, and utilities. The library is designed to offer maximum efficiency and ease of use, ensuring a seamless workflow. Please note that a minimum Microsoft Office support version of 2013 is required for optimal compatibility.

## Scope Details

- **Easy Creation of Office Documents**: Create and manipulate Excel, Word, and PowerPoint files with ease.
- **OpenXML SDK 3.0**: Built on the robust foundation of the OpenXML SDK.
- **Modular Architecture**: Dedicated modules for each Office application for better manageability.

## Getting Started [Link](https://draviavemal.gitbook.io/openxmloffice/getting-started)

The library is available on NuGet. You can install it using the following command:

```shell
Install-Package OpenXMLOffice
```

```shell
dotnet add package OpenXMLOffice.Presentation
```

## Package Details

The official release NuGet packages for OpenXMLOffice on NuGet.org:

| Package | Dev Status | Download | Prerelease |
|---------|---|----------|------------|
| OpenXMLOffice.Presentation | Active | [![NuGet](https://img.shields.io/nuget/v/OpenXMLOffice.Presentation.svg)](https://www.nuget.org/packages/OpenXMLOffice.Presentation) | [![NuGet](https://img.shields.io/nuget/vpre/OpenXMLOffice.Presentation.svg)](https://www.nuget.org/packages/OpenXMLOffice.Presentation) |
| OpenXMLOffice.Excel | Active | [![NuGet](https://img.shields.io/nuget/v/OpenXMLOffice.Excel.svg)](https://www.nuget.org/packages/OpenXMLOffice.Excel) | [![NuGet](https://img.shields.io/nuget/vpre/OpenXMLOffice.Excel.svg)](https://www.nuget.org/packages/OpenXMLOffice.Excel)  |
| OpenXMLOffice.Document | Not Active | [![NuGet](https://img.shields.io/nuget/v/OpenXMLOffice.Document.svg)](https://www.nuget.org/packages/OpenXMLOffice.Document) | [![NuGet](https://img.shields.io/nuget/vpre/OpenXMLOffice.Document.svg)](https://www.nuget.org/packages/OpenXMLOffice.Document) |


## Documentation [Link](https://draviavemal.gitbook.io/openxmloffice/)

All project documentation is completed and regularly updated in Gitbooks. The maintained branch for documentation is the "Documents" branch within the project repository. We welcome any contributions or updates through pull requests. Your assistance is highly appreciated.

## Active Features In Different DLL

### OpenXMLOffice.Presentation

| Control  | Description |
|----------|-------------|
| Slide    | Allows addition of blank slides, removal of slides based on index, and rearrangement of slide positions. |
| Shape    | Enables the location and replacement of shapes by text. Facilitates size and position updates within a slide. |
| Table    | Adds and replaces existing charts. Future plans include finding and updating existing charts. Currently supports creating and updating tables, inheriting all shape functionality. |
| TextBox  | Updates or replaces existing shape text content. Adds new text boxes based on slide control with inherited options from shapes. |
| Picture  | Adds images in PNG, JPG, BMP, TIFF formats to slides. Supports replacement in existing shapes or creation of new controls within a slide. |
| Chart    | Allows users to add charts based on slide control, directly insert them, or replace existing shape controls. Supports major and sub-chart types such as column, line, pie, bar, area, and scatter. Can update the data excel from control |

For charts, the following types are supported:

- **Column Chart:**
  - Cluster
  - Stacked
  - 100% Stacked

- **Line Chart:**
  - Cluster
  - Stacked
  - 100% Stacked
  - Cluster Marker
  - Stacked Marker
  - 100% Stacked Marker

- **Pie Chart:**
  - Pie
  - Doughnut

- **Bar Chart:**
  - Cluster
  - Stacked
  - 100% Stacked

- **Area Chart:**
  - Cluster
  - Stacked
  - 100% Stacked

- **X Y (Scatter) Chart:**
  - Scatter
  - Scatter Smooth Line Marker
  - Scatter Smooth Line
  - Scatter Line Marker
  - Scatter Line
  - Bubble

### OpenXMLOffice.Excel

| Control      | Description |
|--------------|-------------|
| Spreadsheet  | Enables manipulation of worksheets, including retrieval, addition, removal, and renaming of sheets. |
| Worksheet    | Facilitates manipulation of cell data, row properties (e.g., height), column properties (e.g., width), and cell data formatting. |

**Future Plans:**
- **Styling:**
  - Future releases will introduce styling options for enhanced visual representation.

- **Shared String Data Loading:**
  - Planned for memory optimization, shared string data loading will be implemented in upcoming releases.

**Important Note:**
- It is advised to avoid using the library for the creation of large, repeated data files at this point.


## Version History

- **v0.1.0**: Cover Spreadsheet data loading and saving.
- **v0.2.0**: Power Point Exisitng File Shape Based Manipulation of Tables, Text, Charts (Primary)
- Follow project roadmap to get upto date info [Link](https://github.com/users/DraviaVemal/projects/2)

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/DraviaVemal/OpenXMLOffice/blob/main/LICENSE) file for details.

## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement". Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (git checkout -b feature/AmazingFeature)
3. Commit your Changes (git commit -m 'Add some AmazingFeature')
4. Push to the Branch (git push origin feature/AmazingFeature)
5. Open a Pull Request

Please ensure you follow our PR and issue templates for quicker resolution.

## Support

Your feedback and support are important. Feel free to reach out to us with any questions or suggestions.
