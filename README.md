[![Codacy Badge](https://api.codacy.com/project/badge/Grade/ee4775d9ccc540c6bf7750d8490b2ed8)](https://www.codacy.com/app/uribench/SheetProperties)
[![BCH compliance](https://bettercodehub.com/edge/badge/uribench/SheetProperties?branch=master)](https://bettercodehub.com/)

# Spreadsheet Cells Properties Actions

TOC:

- [Objectives](#objectives)
- [Installation](#installation)
- [Usage](#usage)
- [Gist of this Macro](#gist-of-this-macro)
- [Examples and Tests](#examples-and-tests)

## Objectives

The [`Spreadsheet Workbench`][1] of FreeCAD allows including spreadsheets in your model. The support externalizing parametric values of your model. The values in the spreadsheet can be manipulated and control geometric properties of your model.

Value cells in the spreadsheet can be assigned with properties (e.g., `units`, `alias`). 

The problems I have encountered with the way FreeCAD is natively handling the properties are:

1. Properties of a cell are not directly visible when editing a spreadsheet. You need to access them via the `Properties...` dialog.
2. Properties tend to be lost, and a lot of efforts are needed to restore all the lost assigned properties.

Therefore, I have created the `SheetProperties` macro just to handle the above two problems. When using this macro, one can create spreadsheets with visible properties that can be automatically copied and assigned to cell values in the native spreadsheets of FreeCAD.

This macro is not changing the internal structure of the native spreadsheets of FreeCAD, nor their use.

## Installation

As the `SheetProperties` macro is slightly more complex than usual, I decided to split it into multiple files, each handling a single concern. The entry point is still `SheetProperties.FCMacro`.

The `SheetProperties` macro has to be installed manually as it was not officially released for use with the `FreeCAD Addon Manager`. 

In order to install this macro manually all you need is to copy all the files under the `src` folder of this repository into the FreeCAD Macro folder. In the default installation of FreeCAD under Windows the Macro folder is located at: `C:\Users\<*username*>\AppData\Roaming\FreeCAD\Macro`

For more details on installing FreeCAD macros see [this wiki page][2].

## Usage

### Practice First with FreeCAD Spreadsheets without the `SheetProperties` macro

Before using the `SheetProperties` macro, you need to understand how spreadsheets are used in FreeCAD. Read [this wiki][1] for general knowledge on how to use spreadsheets in FreeCAD.

I suggest that you start practicing with the general use of spreadsheets in FreeCAD before using the `SheetProperties` macro.

If you have done so, you have noticed that a value cell in a native spreadsheet of FreeCAD may be assigned with properties, such as `units` and `alias`. You do that by right-clicking on a cell and then selecting the `Properties...` dialog. Once an `alias` has been associated with a value cell, the value can be used on geometries using the assigned `alias`.

### Preparing a Spreadsheet for use with

Unlike the FreeCAD native spreadsheets in which the properties are hidden, with this macro the properties are visible in the spreadsheet.

This Macro helps performing actions (e.g., set, clear) on cells properties of FreeCAD spreadsheet. This is done by having properties data (i.e., `units`, `alias`) in the same row as the target value cell and under their respective columns headers (i.e., Alias, Units).

As can be seen in the following simple spreadsheet, the properties are visible data in respective columns of the same spreadsheet and same row as the target value cell.

![SimpleSpreadsheet.jpg][3]

The data source columns for the cells properties, as well as the target column, need to have appropriate headers (i.e., Alias, Units, Value).

These columns with their headers (i.e., Alias, Units, Value) can be included in the spreadsheet in any order and in any column position, as long as each header with its respective data are given in the same column.

The headers can be placed at any row, as long as all the headers are on the same row. The target column header (i.e., Value) is mandatory, the headers for the data source columns for the cells properties (i.e., Alias, Units) can be configured each as mandatory or optional.

Empty rows can be placed anywhere, including inside the range occupied by the cells of these columns with their headers (i.e., Alias, Units, Value).

Missing or invalid source data for the cells properties are ignored.

Additional data (e.g., comments, description, tabular data) can be placed above or below the range occupied by the cells of these columns with their headers (i.e., Alias, Units, Value).

Additional columns can be placed to the left, right, or between these columns with their headers. This additional data may freely include any name of the headers, as long as it is part of a longer text.

The location of the headers, and the starting and ending rows of the range of these columns are discovered automatically after the target spreadsheet is selected. The user may however, limit this range if it is desired so.

The Macro supports multiple spreadsheets included in a single FreeCAD document. It allows selecting the target spreadsheet using a ComboBox (aka, pop-up menu), or from the tree view. Both methods of selecting a target spreadsheet can be used interchangeably. The Macro syncs between the tree view selection and the ComboBox, bi-directionally.

### Executing the `SheetProperties` macro

Once installed, and an appropriate spreadsheet is ready, you can start using the `SheetProperties` macro. 

Follow the next steps to execute and use the macro:

1. Execute the `SheetProperties.FCMacro` macro from the `Macros | Macros...` dialog using the `User macros` tab.
2. A dialog named `Sheet Properties Actions` will show. 
3. All the existing spreadsheets included in the FreeCAD model will be identified by the macro.
4. The currently active spreadsheet will be selected as the target spreadsheet, but you can switch to any other one using the drop-down menu.
5. The target spreadsheet will be analyzed and the results will be shown in the `Status` panel.
6. The actions `Set` and `Clear` will set or clear the properties of all the cells in the `Value` column based on the content in the respective cells in the  `Alias` and `Units` columns.

From now on you can use the spreadsheet as any native spreadsheets of FreeCAD.

If for some reason FreeCAD will lose the properties, you can always use this macro to restore them easily, just by activating the `Set` action.

Checkout the examples included in the file: `test/TestAll-SheetProperties.FCStd`. Start by experimenting with the 6 spreadsheets under the `Good Data` folder. As you load the file `test/TestAll-SheetProperties.FCStd`, the cells in the `Value` column are without properties. If you execute the `SheetProperties` macro and trigger the `Set` action, you will see that the cells in the `Value` column will then be assigned with the respective properties.

## Gist of this Macro

The `ActiveDocumentSheets` class holds the context of this Macro. It maintains 
information about the active document, such as a list of all the available 
spreadsheets, useful document level constants (e.g., header names for common 
columns), and useful maps (e.g., from spreadsheet label to spreadsheet object 
reference). One instance of `ActiveDocumentSheets` is needed for this Macro and 
it should not have any spreadsheet specific information.

The `RequestParameters` class holds spreadsheet specific information. This 
includes for instance, the location of the common columns, the ranges of usable 
data source rows. One instance is created for every spreadsheet found in the 
active document. The `RequestParameters` has to be ready with all of its 
information prior to performing any action on the associated spreadsheet.

The `SheetPropertiesActions` class provides the possible actions on a spreadsheet 
(e.g., setting and clearing cell properties). It requires a concrete 
RequestParameters instance prior to performing any of its actions.

The `SheetPropertiesActionsForm` is one way of consuming the above. When the 
`SheetPropertiesActionsForm` is instantiated, all the `RequestParameters` instances 
of all the sheets found for the active document, have been set to reflect the 
current state of the sheets. The `SheetPropertiesActionsForm` allows selecting 
interactively one spreadsheet from the list of known spreadsheets of the active 
document, identify the appropriate RequestParameters and pass it to the respective 
`SheetPropertiesActions`. However, RequestParameters and `SheetPropertiesActions` can 
be set and consumed without using a GUI.

## Examples and Tests

### Main Dialog

![MainDialogScreenShot.jpg][4]

The above main dialog of the **SheetProperties** Macro refers to the following 
FreeCAD spreadsheet example:

![Multiple empty rows blocks in the middle.jpg][5]

In the above example, the spreadsheet includes several comment rows at the beginning, 
followed by an empty row and then the headers line, and finally the main data rows with 
multiple empty rows left intentionally here and there.

The **SheetProperties** Macro automatically identified the row location of the headers, 
as well as the ranges of the source data for setting the properties of the relevant cells 
under the Value header column. These discovered ranges are displayed in the above main dialog.

### Test Cases

For test cases (both valid and invalid examples) see: `test/TestAll-SheetProperties.FCStd`

---

[1]: https://wiki.freecadweb.org/Manual:Using_spreadsheets
[2]: https://wiki.freecadweb.org/How_to_install_macros
[3]: assets/SimpleSpreadsheet.jpg
[4]: assets/MainDialogScreenShot.jpg
[5]: assets/MultipleEmptyRowsBlocksInTheMiddle.jpg