# Effortless Creation of Excel, PowerPoint, and Word Documents

This project aims to provide a streamlined, efficient way to create and manipulate Excel, PowerPoint, and Word documents using the OpenXML format. By leveraging the power of modern programming languages, this library simplifies the process of document creation, enabling developers to focus on functionality rather than the intricacies of the OpenXML standard.

## Important Note about V3

After thorough analysis, I have concluded that the full OpenXML format relations and connections are adequately addressed. Therefore, to eliminate duplicate work, V3 is no longer under development. Please follow V4 updates for final results.

# Version 4 Goals and Objectives

| Supported Languages | Document link | Packages   | package link                                             | Description                                          |
| ------------------- | ------------- | ---------- | -------------------------------------------------------- | ---------------------------------------------------- |
| Rust                | TODO          | Rust       | [Crates](https://crates.io/)                             | Rust crate directly connecting to core lib           |
| C#                  | TODO          | C#         | [Nuget](https://www.nuget.org/)                          | C# wrapper package wrote around FFI layer of rust    |
| Java                | TODO          | Java       | [Maven Central](https://mvnrepository.com/)              | Java wrapper package wrote around FFI layer of rust  |
| Go                  | TODO          | Go         | [Github](https://github.com/DraviaVemal/OpenXML-Office/) | Go wrapper package wrote around FFI layer of rust    |
| TypeScript          | TODO          | TypeScript | [npm](https://www.npmjs.com/)                            | NAPI-RS is used to expose the core lib as node addon |
|                     |               | Rust-API   | [Docker Hub](https://hub.docker.com/)                    | API container running rust crate for HTTP support    |



## Status

: TODO

## Technical Details

This release marks a significant evolution for the OpenXML-office project. The upcoming version, V4, will be a complete rewrite of this package. It aims to maintain previous release structures as much as possible, with a strong focus on minimizing migration efforts for adopters.

## Inspiration

This project has been in the works for nearly a year, driven by a desire to explore the OpenXML standards for office documents. Initially developed as a C# project using the OpenXML-SDK as a baseline, I have now gathered sufficient knowledge to transition into a cross-platform, multi-language supported package. This effort is a way to give back to the community that has been instrumental in my professional and personal growth.

### Architecture

The core system is written in Rust, ensuring optimal performance and memory usage. This system is then exposed as a "C" extern FFI, facilitating interaction with other languages. Wrappers for each supported language have been created, and the package is published in the respective package managers. For TypeScript, `napi-rs` is utilized to create a Node.js addon, preserving performance advantages.

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
