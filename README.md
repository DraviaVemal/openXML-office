# Effortless Creation of Excel, PowerPoint, and Word Documents

This project aims to provide a streamlined, efficient way to create and manipulate Excel, PowerPoint, and Word documents using the OpenXML format. By leveraging the power of modern programming languages, this library simplifies the process of document creation, enabling developers to focus on functionality rather than the intricacies of the OpenXML standard.

## V2/Stable Version

Please refer [stable branch](https://github.com/DraviaVemal/OpenXML-Office/tree/stable) for V2/Stable Package related details. This is active development Branch

## Stable Status
| Detail            | Status                                                                                                                                                                                                                                       | Detail                | Status                                                                                                                                                                                                                                            |
| ----------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Stable Release    | [![GitHub Release](https://img.shields.io/github/v/release/DraviaVemal/OpenXML-Office?sort=semver&label=Stable%20Release)](https://openxml-office.draviavemal.com/)                                                                          | Stable Build Status   | [![Package Build and Publish to NuGet](https://github.com/DraviaVemal/openxmloffice/actions/workflows/nuget-publish-stable.yml/badge.svg?branch=stable)](https://github.com/DraviaVemal/openxmloffice/actions/workflows/nuget-publish-stable.yml) |
| Alpha Release     | [![NuGet](https://img.shields.io/nuget/vpre/openxmloffice.Presentation.svg)](https://www.nuget.org/packages/openxmloffice.Presentation)                                                                                                      | Alpha Build Status    | [![Package Build and Publish to NuGet](https://github.com/DraviaVemal/openxmloffice/actions/workflows/nuget-publish-alpha.yml/badge.svg?branch=alpha)](https://github.com/DraviaVemal/openxmloffice/actions/workflows/nuget-publish-alpha.yml)    |
| Code Quality      | [![Codacy Badge](https://app.codacy.com/project/badge/Grade/5b420a599805426ab8a990a1a741247a)](https://app.codacy.com/gh/DraviaVemal/OpenXML-Office/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)       | PPT Files Generated   | [![Generated](https://draviavemal.com/openxml-office/powerpoint-count.svg)](https://openxml-office.draviavemal.com/)                                                                                                                              |
| Code Coverage     | [![Codacy Badge](https://app.codacy.com/project/badge/Coverage/5b420a599805426ab8a990a1a741247a)](https://app.codacy.com/gh/DraviaVemal/OpenXML-Office/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_coverage) | Excel Files Generated | [![Generated](https://draviavemal.com/openxml-office/excel-count.svg)](https://openxml-office.draviavemal.com/)                                                                                                                                   |
| Package Downloads | [![Downloads](https://img.shields.io/nuget/dt/openxmloffice.Presentation.svg)](https://www.nuget.org/packages/openxmloffice.Presentation)                                                                                                    | Word Files Generated  | [![Generated](https://draviavemal.com/openxml-office/word-count.svg)](https://openxml-office.draviavemal.com/)                                                                                                                                    |

## Important Note about V3

After thorough analysis, I have concluded that the full OpenXML format relations and connections are adequately addressed. Therefore, to eliminate duplicate work, V3 is no longer under development. Please follow V4 updates for final results.

# Version 4 Goals and Objectives

| Supported Languages | Min.Support Version | Readme link | Packages   | package link                                             | Description                                           |
| ------------------- | ------------------- | ----------- | ---------- | -------------------------------------------------------- | ----------------------------------------------------- |
| Rust                | 1.32                | TODO        | Rust       | [Crates](https://crates.io/)                             | Rust crate directly connecting to core lib            |
| C#                  | .net6.0             | TODO        | C#         | [Nuget](https://www.nuget.org/)                          | C# wrapper package wrote around FFI layer of rust     |
| Python              |                     | TODO        | Python     | [PyPi](https://pypi.org/)                                | Python Wrapper package wrote around FFI layer of rust |
| ----------------    | ----------          | ---------   | --------   | ---------                                                | -------                                               |
| PHASE 2             | PHASE 2             | PHASE 2     | PHASE 2    | PHASE 2                                                  | PHASE 2                                               |
| ---------------     | -----------         | ---------   | --------   | ----------                                               | -------                                               |
| Java                | 1.8                 | TODO        | Java       | [Maven Central](https://mvnrepository.com/)              | Java wrapper package wrote around FFI layer of rust   |
| Go                  | 1.22                | TODO        | Go         | [Github](https://github.com/DraviaVemal/OpenXML-Office/) | Go wrapper package wrote around FFI layer of rust     |
| TypeScript          |                     | TODO        | TypeScript | [npm](https://www.npmjs.com/)                            | NAPI-RS is used to expose the core lib as node addon  |
|                     |                     |             | Rust-API   | [Docker Hub](https://hub.docker.com/)                    | API container running rust crate for HTTP support     |

## V4 Status Development

| Language    | Package      | Progress |
| ----------- | ------------ | -------- |
| Flatbuffers | common       | 🟩⬜⬜⬜⬜    |
| Rust        | fbs          | 🟩⬜⬜⬜⬜    |
| Rust        | xml          | 🟩🟩⬜⬜⬜    |
| Rust        | global       | 🟩⬜⬜⬜⬜    |
| Rust        | spreadsheet  | 🟩🟩⬜⬜⬜    |
| Rust        | presentation | 🟩⬜⬜⬜⬜    |
| Rust        | document     | 🟩⬜⬜⬜⬜    |
| Rust        | FFI          | 🟩⬜⬜⬜⬜    |
| C#          | fbs          | 🟩⬜⬜⬜⬜    |
| C#          | spreadsheet  | 🟩⬜⬜⬜⬜    |
| C#          | presentation | 🟩⬜⬜⬜⬜    |
| C#          | document     | 🟩⬜⬜⬜⬜    |
| Python      | fbs          | 🟩⬜⬜⬜⬜    |
| Python      | spreadsheet  | ⬜⬜⬜⬜⬜    |
| Python      | presentation | ⬜⬜⬜⬜⬜    |
| Python      | document     | ⬜⬜⬜⬜⬜    |
| ----------- | -----------  | -------- |
| PHASE 2     | PHASE 2      | PHASE 2  |
| ----------- | -----------  | -------- |
| Java        | fbs          | ⬜⬜⬜⬜⬜    |
| Java        | spreadsheet  | ⬜⬜⬜⬜⬜    |
| Java        | presentation | ⬜⬜⬜⬜⬜    |
| Java        | document     | ⬜⬜⬜⬜⬜    |
| Go          | fbs          | ⬜⬜⬜⬜⬜    |
| Go          | spreadsheet  | ⬜⬜⬜⬜⬜    |
| Go          | presentation | ⬜⬜⬜⬜⬜    |
| Go          | document     | ⬜⬜⬜⬜⬜    |
| TypeScript  | fbs          | ⬜⬜⬜⬜⬜    |
| TypeScript  | spreadsheet  | ⬜⬜⬜⬜⬜    |
| TypeScript  | presentation | ⬜⬜⬜⬜⬜    |
| TypeScript  | document     | ⬜⬜⬜⬜⬜    |

## Repo & Design Block Diagram
![Design](design.svg)


## Technical Details

This release marks a significant evolution for the OpenXML-office project. The upcoming version, V4, will be a complete rewrite of this package. It aims to maintain previous release structures as much as possible, with a strong focus on minimizing migration efforts for adopters.

## Inspiration

This project has been in the works for nearly a year, driven by a desire to explore the OpenXML standards for office documents. Initially developed as a C# project using the OpenXML-SDK as a baseline, I have now gathered sufficient knowledge to transition into a cross-platform, multi-language supported package. This effort is a way to give back to the community that has been instrumental in my professional and personal growth.

### Architecture

The core system is written in Rust, ensuring optimal performance and memory usage. This system is then exposed as a "C" extern FFI, facilitating interaction with other languages. Wrappers for each supported language have been created, and the package is published in the respective package managers. For TypeScript, `napi-rs` is utilized to create a Node.js addon, preserving performance advantages. The data transmission is handled using FlatBuffer, following a central schema that ensures consistent patterns across all supported languages and facilitates code organization and documentation maintenance.

## Support Scope

This package supports `.xlsx`, `.pptx`, and `.docx` formats starting from Office 2007. Features are organized into respective modules, namespaces, and directories to provide clarity on the minimum supported version for each feature. The package is designed to be compatible with all applications that open standard OpenXML documents, including online solutions.

## Project Timeline

I will be halting work on the V3 version to prevent duplicate efforts. I have gathered all foundational information necessary for designing the system from the ground up, leading to this decision. 

I anticipate a timeline of 6-8 months for migrating all existing features from the repository to the new codebase. The same functionality will be available across all supported languages and operating systems. While timelines may vary based on my availability, I am committed to maintaining consistent progress, so there should not be significant surprises.

Until then, V2 will remain the stable version for use, and any issues related to Excel and PowerPoint will be prioritized until V4 is ready for release.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are greatly appreciated. If you have suggestions that could improve this project, please fork the repo and create a pull request. Alternatively, you can open an issue tagged "enhancement." Don’t forget to star the project!

### How to Contribute

1. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
2. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
3. Push to the Branch (`git push origin feature/AmazingFeature`)
4. Open a Pull Request 

Please ensure you follow the PR and issue templates for quicker resolution.

## Support

Your feedback and support are important. Feel free to reach out with any questions or suggestions.
