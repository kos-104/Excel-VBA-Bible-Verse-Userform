# Excel VBA Bible Userform

**Bringing scripture to Excel — a dynamic, MVVM-driven VBA userform for looking up Bible verses on demand.**

![VBA](https://img.shields.io/badge/Language-VBA-blue.svg)
![Excel](https://img.shields.io/badge/Platform-Excel-217346.svg)
![Architecture](https://img.shields.io/badge/Architecture-MVVM-orange.svg)
![Web Scraping](https://img.shields.io/badge/Feature-Web%20Scraping-brightgreen.svg)

## Description

The **Excel VBA Bible Userform** is an internet-connected tool that uses web-scraping to display Bible verses directly within Microsoft Excel.

It employs the **Model-View-ViewModel (MVVM)** architecture to enhance maintainability and scalability, patterned after the *RubberduckSwagShop MVVM-Lite* project. This structure separates the user interface from business logic, enabling flexible data binding and command execution. **Interfaces** are used throughout to exhibit **polymorphism**, allowing different class modules to share a common contract while implementing their own specific behavior — a key OOP technique that reinforces the flexibility of the MVVM design.

**Key features include:**
- A fully resizable form, controls, and downloaded verse text — ensuring a responsive experience across screen sizes
- Intuitive navigation via tooltips and spin buttons
- Mouse-scrollable listboxes for added convenience
- Support for several popular Bible versions
- Simultaneous download of multiple verses from the same chapter
- Hover-responsive command buttons for clearer target identification
- Thorough event handling, with custom commands for adding, clearing, and confirming line items — all backed by robust error handling

## Requirements
- Excel with macros enabled (developed and tested on Excel 365)
- An active internet connection, for retrieving verse text
- The following VBA References must be enabled (**Tools → References** in the VBA Editor):
  - Microsoft Scripting Runtime
  - Microsoft HTML Object Library
  - Microsoft Internet Controls
  - Microsoft XML, v6.0

## How to Use
1. Launch the userform from the button provided on the Excel sheet
2. Navigate the form using the labeled controls and hover tooltips — the interface is intuitive by design
3. Select a Bible version, then specify the book, chapter, and verse(s) to retrieve

## Acknowledgments

Special thanks to:

- **David Hager**, for his extensive groundwork on the dynamic named ranges/formulas and pivot tables underpinning this form's framework.
  📖 [Lookup a Bible Verse Using Excel w/o VBA](https://dhexcel1.wordpress.com/2017/07/03/lookup-a-bible-verse-using-excel-wo-vba-by-david-hager/)

- **Mathieu Guindon and the Rubberduck team**, for creating the *SwagShop* MVVM-Lite program used as the architectural foundation for this project.
  🦆 [Lightweight MVVM in VBA](https://rubberduckvba.blog/2023/04/11/lightweight-mvvm-in-vba/)

---

*Built to bring scripture into the spreadsheet.*
