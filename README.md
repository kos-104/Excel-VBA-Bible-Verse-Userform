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

## Under the Hood

This project's MVVM structure exists to solve a common problem in VBA userforms: business logic, UI controls, and event handling tend to blur together until every procedure knows too much about everything else. MVVM keeps these concerns separate, so the form's controls don't need to know how a verse gets scraped from the web, and the scraping logic doesn't need to know which textbox displays the result.

### The Three Layers

**The ViewModel** (`BibleViewModel`) holds the application's state and logic — the selected Bible version, book, chapter, verse selections, and the operations that act on them. It knows nothing about the userform, its controls, or their positions. This isolation is what makes the logic testable and reusable independent of the UI.

**The View** is the userform itself — a collection of textboxes, listboxes, spin buttons, and command buttons. The controls are intentionally "dumb": they display values and raise events, but they don't contain business logic or know how to fetch a verse.

**The glue in between** is where data binding and commands do their work, allowing the ViewModel and View to communicate without ever referencing each other directly.

### Data Binding: Keeping the ViewModel and Form in Sync

Each bindable control (textboxes, listboxes, spin buttons, and command buttons) has a corresponding binding class — `TextBoxValueBinding`, `ListBoxValueBinding`, `SpinBttnValueBinding`, and `CommandBttnValueBinding` — created through a central `PropertyBindings` factory. Each binding class implements `IHandlePropertyChanged`, allowing it to listen for changes on the ViewModel.

Using `TextBoxValueBinding` as an example, the round-trip works like this:

1. The ViewModel's property changes (e.g., new verse text is retrieved)
2. `PropertyChangeNotification` raises an event through the `INotifyPropertyChanged` interface
3. The binding's `IHandlePropertyChanged_OnPropertyChanged` handler catches this and updates the TextBox

The reverse path handles user input:

1. The user edits the TextBox
2. The control's `Change` event fires
3. The binding reads the new value and pushes it back to the ViewModel property using `CallByName`

This two-way flow means the ViewModel and the View stay synchronized without either one holding a direct reference to the other's internals.

### Commands: Decoupling Button Clicks from Logic

Rather than wiring a button's `Click` event directly to a ViewModel method, each user action is wrapped in a command class implementing a shared `ICommand` interface. Using `AddLineItemCommand` as an example: when the button is clicked, `CommandBinding` calls the command's `CanExecute` method first, confirming the ViewModel is in a valid state to receive the action. Only then does it call `Execute`, which forwards the request to the corresponding ViewModel method.

This indirection means the button doesn't need to know anything about the ViewModel's methods, and the ViewModel doesn't need to know anything about which button was clicked. Similar commands (`ClearLineItemsCommand`, `ClearOptionsCommand`) follow the same pattern for their respective actions.

### Web Scraping: Retrieving Verse Text

Once a book, chapter, and verse selection are confirmed, a dedicated module handles the HTTP request using `Microsoft XML v6.0`, retrieving the raw HTML from the source page. A second pair of modules then parses that HTML using the `Microsoft HTML Object Library`, locating the specific elements containing the verse text and passing the cleaned result back to the ViewModel — which, in turn, pushes it to the View through the data binding layer already described.

### Supporting Utilities

A few additional components round out the user experience:

- **`CFormResizer`** — handles proportional resizing of the form and all its controls, adapted from Stephen Bullen and Rob Bovey's *Professional Excel Development*, with supporting Win32 API calls (via a dedicated `APIs` module) used to position the form precisely at the active cell on launch
- **`Scroll`** — enables mousewheel scrolling within listboxes, based on a solution originally developed by Jaafar Tribak

Both are widely used, community-vetted utilities adapted here to fit the project's needs.

## Acknowledgments

Special thanks to:

- **David Hager**, for his extensive groundwork on the dynamic named ranges/formulas and pivot tables underpinning this form's framework.
  📖 [Lookup a Bible Verse Using Excel w/o VBA](https://dhexcel1.wordpress.com/2017/07/03/lookup-a-bible-verse-using-excel-wo-vba-by-david-hager/)

- **Mathieu Guindon and the Rubberduck team**, for creating the *SwagShop* MVVM-Lite program used as the architectural foundation for this project.
  🦆 [Lightweight MVVM in VBA](https://rubberduckvba.blog/2023/04/11/lightweight-mvvm-in-vba/)

- **Stephen Bullen and Rob Bovey**, for the resizable userform framework adapted for this project's `CFormResizer` class, from their book *Professional Excel Development*.

- **Jaafar Tribak**, for the mousewheel-scrolling solution for VBA userform ComboBoxes, adapted for this project's listboxes.

---

*Built to bring scripture into the spreadsheet.*
