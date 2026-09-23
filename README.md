# Excel VBA Bible Userform

**Bringing scripture to Excel — a dynamic, MVVM-driven VBA userform for looking up Bible verses on demand.**

![VBA](https://img.shields.io/badge/Language-VBA-blue.svg)
![Excel](https://img.shields.io/badge/Platform-Excel-217346.svg)
![Architecture](https://img.shields.io/badge/Architecture-MVVM-orange.svg)
![Web Scraping](https://img.shields.io/badge/Feature-Web%20Scraping-brightgreen.svg)

## Description

The **Excel VBA Bible Userform** is an internet-connected tool that uses web-scraping to display Bible verses directly within Microsoft Excel.

It employs the **Model-View-ViewModel (MVVM)** architecture to enhance maintainability and scalability, patterned after the *RubberduckSwagShop MVVM-Lite* project. This structure separates the user interface from business logic, enabling flexible data binding and command execution. Interfaces are used throughout to exhibit **polymorphism** — different class modules sharing a common contract, each implementing its own specific behavior.

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

MVVM keeps business logic, UI controls, and event handling from blurring together — a common problem in VBA userforms once every procedure starts knowing too much about everything else. The form's controls don't need to know how a verse gets scraped from the web, and the scraping logic doesn't need to know which textbox displays the result.

This project implements a **hybrid** MVVM approach rather than a textbook-pure one, as a deliberate tradeoff in favor of simplicity for a single-form, single-consumer tool:

- List population, selection syncing, and change notification follow strict MVVM separation
- Verse-capture and clearing commands take a more direct, pragmatic route instead, described below

### The Three Layers

**The ViewModel** (`BibleViewModel`) holds the application's bindable state — selected version, book, chapter, and verse lists — exposed as properties with change notification. For most of the form's list-driven behavior, this keeps the ViewModel properly decoupled from the userform, its controls, and their positions.

That isolation isn't absolute across the whole class, though. Command operations — capturing a verse, clearing the form, adjusting the spin position — take a more direct route:

- These methods receive the userform as a parameter
- They forward it into standalone procedures
- Those procedures read control values and write results straight back into the form

This keeps the code straightforward, at the cost of the ViewModel being fully agnostic of the View for that portion of its behavior.

**The View** is the userform itself — textboxes, listboxes, spin buttons, and command buttons. Most controls are intentionally "dumb": they display values and raise events without containing business logic. The exception is the verse-capture and clearing flow, where the command layer hands the form directly to the logic that manipulates it, bypassing the binding layer for that operation.

**The glue in between** is where data binding and commands do their work for everything else, letting the ViewModel and View communicate without referencing each other directly. List population and selection syncing for versions, books, and chapters follow this path faithfully, even where the command operations don't.

### Interfaces and Polymorphism

VBA doesn't support inheritance the way languages like C# or Java do, but it does support interface implementation through `Implements` — and this project leans on that heavily:

- Each binding class (`TextBoxValueBinding`, `ListBoxValueBinding`, `SpinBttnValueBinding`, `CommandBttnValueBinding`) implements the same `IHandlePropertyChanged` interface
- Each command class implements the same `ICommand` interface, with its own `CanExecute` and `Execute` methods

This matters because the rest of the system — `PropertyChangeNotification`, `CommandBinding`, `PropertyBindings` — never needs to know which concrete class it's holding. A collection of bindings can be looped through and notified identically, regardless of whether each one wraps a TextBox, a ListBox, or a SpinButton, because they all satisfy the same contract. Swap in a new control type tomorrow, and as long as its binding class implements `IHandlePropertyChanged`, nothing else in the system needs to change.

This is **polymorphism**: different class modules, sharing a common interface, each implementing that interface's methods according to their own control's needs — letting a single `PropertyChangeNotification` object call `OnPropertyChanged` on a whole collection of handlers uniformly, with no `If TypeOf... Then` chain required to sort out what kind of control it's dealing with.

### Data Binding: Keeping the ViewModel and Form in Sync

Each bindable control has a corresponding binding class — `TextBoxValueBinding`, `ListBoxValueBinding`, `SpinBttnValueBinding`, `CommandBttnValueBinding` — created through a central `PropertyBindings` factory. Each implements `IHandlePropertyChanged`, allowing it to listen for changes on the ViewModel.

Using `TextBoxValueBinding` as an example, the round-trip works like this:

1. The ViewModel's property changes (e.g., new verse text is retrieved)
2. `PropertyChangeNotification` raises an event through the `INotifyPropertyChanged` interface
3. The binding's `IHandlePropertyChanged_OnPropertyChanged` handler catches this and updates the TextBox

The reverse path handles user input:

1. The user edits the TextBox
2. The control's `Change` event fires
3. The binding reads the new value and pushes it back to the ViewModel property using `CallByName`

This two-way flow keeps the ViewModel and View synchronized without either holding a direct reference to the other's internals — for the properties and controls that follow this path.

### Commands: Decoupling Button Clicks from Logic

Rather than wiring a button's `Click` event directly to a ViewModel method, each user action is wrapped in a command class implementing a shared `ICommand` interface. Using the verse-capture command as an example:

- The button click triggers `CommandBinding`, which calls the command's `CanExecute` method first, confirming the ViewModel is in a valid state to receive the action
- Only then does it call `Execute`, forwarding the request to the corresponding `BibleViewModel` method
- `BibleViewModel` passes the userform reference into a standalone procedure
- That procedure reads the current listbox selections, performs the lookup, and writes the result straight back into the form

Similar commands (clearing line items, clearing options) follow the same pattern for their respective actions.

### Web Scraping: Retrieving Verse Text

Once a book, chapter, and verse selection are confirmed, retrieving and displaying the verse text happens in two steps:

1. A dedicated module handles the HTTP request using `Microsoft XML v6.0`, retrieving the raw HTML from the source page
2. A second pair of modules parses that HTML using the `Microsoft HTML Object Library`, locating the specific elements containing the verse text and passing the cleaned result back for display on the form

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
