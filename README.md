# Effortless Creation of Excel, PowerPoint, and Word Documents

This is rust experimental branch rethinking the design idea.

The intension I created this package is to have a good understand of openXML world and to have a cross programming language supported project
that showcase my skill set.

In that motivation, and the knowledge i got from working around openXML-SDK for almost a year now my next action is to drop the dependency of openXML-SDK.

The idea behind this branch is to build the core module in low level language rust and create a wrapper world around it so that the same can be shipped and used in multiple language at almost no additional performance or memory cost.

The intended target language that I want to cover as result.
- Rust - Core
- C/C++ - C API
- C# - FFI - P/Invoke
- Go Lang - FFI - cgo
- Java - FFI - JNI
- Swift - FFI - Swift's C
- Ruby - FFI
- TS - NAPI (NodeJS)
- JS - NAPI (NodeJS)