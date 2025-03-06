# Notes on the Office VBA editor

We enjoy our lovely VSCode but it seems working with the Microsoft editor is unavoidable, we shall save notes on it in this markdown file. These notes shall help us if we need a refresher on how to accomplish specific tasks.

### VBA objects
Examples of objects and their relation to eachother:

```
Application (Excel)
│
├── Workbooks (Collection)
│   ├── Workbook (Single)
│   │   ├── Sheets (Collection)
│   │   │   ├── Worksheet (Single)
│   │   │   │   ├── Range (Collection of Cells)
│   │   │   │   ├── ListObjects (Tables)
│   │   │   │   ├── PivotTables
│   │   │   │   ├── Shapes (Images, Textboxes, etc.)
│   │   │   │   ├── Hyperlinks
│   │   │   │   ├── Comments
│   │   │   │   ├── OLEObjects (Embedded Files)
│   │   │   │   ├── QueryTables (External Data)
│   │   │   │   ├── ChartObjects (Embedded Charts)
│   │   │   ├── Chart (Standalone Chart Sheet)
│   │   │   ├── PivotTable (Standalone)
│   │   ├── Names (Named Ranges)
│   │   ├── Connections (External Data Connections)
│   │   ├── VBProject (VBA Modules)
│   │   │   ├── VBComponents (Modules, Forms, Classes)
│   │   ├── CommandBars (Legacy Menus)
│   ├── ActiveWorkbook (Currently Opened Workbook)
│
├── UserForms (Collection)
│   ├── UserForm (Single)
│   │   ├── Controls (Buttons, Textboxes, etc.)
│
├── FileDialog (File Selection)
│   ├── SelectedItems (Chosen Files)
│
└── CommandBars (Legacy UI)
```


Built in objects that can be manipulated :
| Object             | Description |
|--------------------|-------------|
| `Err`             | Handles runtime errors by providing details such as `Number`, `Description`, and methods like `Raise` and `Clear`. |
| `Application`     | Represents the current instance of the host application (Excel, Word, PowerPoint, etc.), providing access to properties and methods. |
| `Workbook`        | Represents an open Excel workbook (`ThisWorkbook` refers to the active file running VBA). |
| `Worksheet`       | Represents an individual sheet in a workbook. |
| `Range`           | Represents a range of cells in Excel, allowing for reading, writing, and formatting operations. |
| `Cells`           | Represents all cells in a worksheet and allows referencing individual cells dynamically. |
| `Selection`       | Represents the currently selected object in Excel (cells, chart, shape, etc.). |
| `ActiveSheet`     | Refers to the currently active worksheet. |
| `ActiveCell`      | Refers to the currently active cell. |
| `FileSystemObject (FSO)` | Provides file and folder manipulation methods, like reading, writing, and creating files. |
| `Dictionary`      | Allows key-value pair storage similar to a hash table (requires `Scripting.Dictionary` reference). |
| `Collection`      | Represents a group of related objects that can be iterated using `For Each`. |
| `Clipboard`       | Provides access to the Windows clipboard (requires API calls). |
| `Debug`           | Used to print output to the **Immediate Window** using `Debug.Print`. |
| `InputBox`        | Displays a dialog box that prompts the user for input. |
| `MsgBox`          | Displays a message box for user interaction. |
| `Shell`           | Allows running external commands or applications from VBA. |
| `CreateObject`    | Creates an instance of an external application (e.g., `CreateObject("Scripting.FileSystemObject")`). |
| `Timer`           | Returns the number of seconds elapsed since midnight. Useful for performance measurement. |









### Operators 

# VBA Operators

| Operator | Category         | Description                                    | Example (`A = 10, B = 3, X = True, Y = False`) | Result        |
|----------|------------------|------------------------------------------------|--------------------------------|--------------|
| `+`      | Arithmetic       | Addition                                       | `A + B`                        | `13`         |
| `-`      | Arithmetic       | Subtraction                                    | `A - B`                        | `7`          |
| `*`      | Arithmetic       | Multiplication                                 | `A * B`                        | `30`         |
| `/`      | Arithmetic       | Division (Returns Float)                       | `A / B`                        | `3.3333`     |
| `\`      | Arithmetic       | Integer Division (Drops Decimals)              | `A \ B`                        | `3`          |
| `Mod`    | Arithmetic       | Modulus (Remainder)                            | `A Mod B`                      | `1`          |
| `^`      | Arithmetic       | Exponentiation                                 | `A ^ B`                        | `1000`       |
| `=`      | Comparison       | Equal to                                       | `A = B`                        | `False`      |
| `<>`     | Comparison       | Not equal to                                   | `A <> B`                       | `True`       |
| `>`      | Comparison       | Greater than                                   | `A > B`                        | `True`       |
| `<`      | Comparison       | Less than                                      | `A < B`                        | `False`      |
| `>=`     | Comparison       | Greater than or equal                          | `A >= B`                       | `True`       |
| `<=`     | Comparison       | Less than or equal                             | `A <= B`                       | `False`      |
| `And`    | Logical          | True if **both** are True                      | `X And Y`                      | `False`      |
| `Or`     | Logical          | True if **at least one** is True               | `X Or Y`                       | `True`       |
| `Not`    | Logical          | Inverts Boolean value                          | `Not X`                        | `False`      |
| `Xor`    | Logical          | True if **only one** is True                   | `X Xor Y`                      | `True`       |
| `&`      | Concatenation    | Joins two strings                              | `"Hel" & "lo"`                 | `"Hello"`    |
| `=`      | Assignment       | Assigns a value                                | `A = 5`                        | `A = 5`      |






### Variables

Instantiation :
`Dim variable1 As Type`
Set a value : 
`variable1 = ...`

Types:

| **Type**      | **Size**        | **Description** |
|--------------|---------------|----------------|
| `Boolean`    | 2 bytes       | `True` or `False` |
| `Byte`       | 1 byte        | Integer from `0` to `255` |
| `Integer`    | 2 bytes       | Whole number from `-32,768` to `32,767` |
| `Long`       | 4 bytes       | Whole number from `-2,147,483,648` to `2,147,483,647` |
| `Single`     | 4 bytes       | Floating-point number (precision up to 7 digits) |
| `Double`     | 8 bytes       | Floating-point number (precision up to 15 digits) |
| `Currency`   | 8 bytes       | Fixed-point number with 4 decimal places (used for currency values) |
| `Decimal`    | 14 bytes      | Up to 28 decimal places (only available via `Variant`) |
| `String`     | 1 byte per char | Holds text (max 2 billion characters for variable-length strings) |
| `Date`       | 8 bytes       | Stores date/time values (range: `100` AD to `9999` AD) |
| `Object`     | 4 bytes       | Reference to an object (e.g., `Workbook`, `Worksheet`, `Range`) |
| `Variant`    | Varies        | Can store any data type (inefficient for large datasets) |
| `User-Defined Type` | Varies | Custom structure (defined with `Type...End Type`) |
| `Array`      | Varies        | Collection of values stored in a single variable |




### Arrays

```VB
Sub StaticArrayExample()
    Dim arr(3) As String
    arr(0) = "Apple"
    arr(1) = "Banana"
    arr(2) = "Cherry"
    MsgBox arr(1)   '//Shows "Banana"
End Sub
```






### Loops
There is no continue keyword, or pass keyword, in VBA loops. Repeating an iteration requires an If verification and an iterator manipulation inside the loop. The `Next` keyword will increment the iterator and do a goto to its `For To` loop, they must be affecting the same variable. Nested loops are required to manipulate multiple variables at the same time.


For loop :
```VB
Sub ForLoopExample()
    'Iterator is instantiated before the loop
    Dim i As Integer

    For i = 1 To 5 'Condition
        Debug.Print "Iteration " & i
    Next i  'Increments
End Sub
```

While loop :
```VB
Sub DoWhileExample()
    Dim x As Integer
    x = 1
    Do While x <= 5  'Do Until works in a similar way, but repeats when condition is false
        Debug.Print "x = " & x
        x = x + 1
    Loop
End Sub
```


For each:
```VB
Sub ForEachLoopExample()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        Debug.Print ws.Name
    Next ws
End Sub
```

List of iterable elements :

| Iterable Element      | Description |
|-----------------------|-------------|
| `Worksheet`          | Iterates through all worksheets in a workbook (`ThisWorkbook.Sheets`). |
| `Workbook`           | Iterates through all open workbooks (`Application.Workbooks`). |
| `Range`             | Iterates through each cell in a range (`Range("A1:A10")`). |
| `ChartObject`       | Iterates through all charts in a worksheet (`ActiveSheet.ChartObjects`). |
| `Shape`            | Iterates through all shapes (e.g., textboxes, images) in a worksheet (`ActiveSheet.Shapes`). |
| `PivotTable`        | Iterates through all pivot tables in a worksheet (`ActiveSheet.PivotTables`). |
| `PivotField`        | Iterates through fields in a PivotTable (`PivotTable.PivotFields`). |
| `ListObject`        | Iterates through tables in a worksheet (`ActiveSheet.ListObjects`). |
| `Name`             | Iterates through named ranges (`ThisWorkbook.Names`). |
| `Chart`           | Iterates through all charts in a workbook (`ThisWorkbook.Charts`). |
| `WorkbookConnection` | Iterates through all data connections in a workbook (`ThisWorkbook.Connections`). |
| `Comment`         | Iterates through all cell comments in a worksheet (`ActiveSheet.Comments`). |
| `Hyperlink`       | Iterates through all hyperlinks in a worksheet (`ActiveSheet.Hyperlinks`). |
| `QueryTable`      | Iterates through query tables (external data connections) in a worksheet (`ActiveSheet.QueryTables`). |
| `CommandBar`      | Iterates through all command bars (used in older versions of Excel, pre-Ribbon) (`Application.CommandBars`). |
| `VBAComponent`    | Iterates through all VBA modules in a workbook (`ThisWorkbook.VBProject.VBComponents`). |
| `FileDialogSelectedItems` | Iterates through selected files from a file dialog (`Application.FileDialog(msoFileDialogOpen).SelectedItems`). |
| `OLEObject`      | Iterates through all embedded OLE objects (e.g., embedded PDFs) in a worksheet (`ActiveSheet.OLEObjects`). |
| `UserForm`       | Iterates through all open UserForms (`VBA.UserForms`). |
| `Control`        | Iterates through all controls in a UserForm (`UserForm1.Controls`). |





# If Elif Else
Straightforward. Do not forget the `End if` to close the statement.

If-Elif-Else:
```VB
Sub IfElseIfExample()
    Dim score As Integer
    score = 85
    If score >= 90 Then
        MsgBox "Grade: A"
    ElseIf score >= 75 Then
        MsgBox "Grade: B"
    Else
        MsgBox "Grade: C"
    End If
End Sub
```

Switch-Case:
```VB
Select Case expression
    Case value1
        ' Code to execute if expression = value1
    Case value2
        ' Code to execute if expression = value2
    Case value3, value4
        ' Code to execute if expression = value3 OR value4
    Case Else
        ' Code to execute if none of the cases match
End Select
```





### Subroutines

Subroutines are void functions. They will not return a value, which makes them impossible to call directly in a spreadsheet.

```VB
Sub GreetUser()
    'This is a comment
    MsgBox "Hello, welcome to VBA!"
End Sub
```





### Functions
Functions can be called inside of your spreadsheet to expand the capabilities of Excel. The return value is given by setting the name of the Function to a value. There is no return keyword.

```VB
Function AddNumbers(a As Integer, b As Integer) As Integer
    "this is a comment 
    block"
    AddNumbers = a + b  ' Return the sum
End Function
```





### Errors
We can use the Err built in object to handle errors.


Raising an error manually:
```VB
Sub CustomError()
    On Error GoTo ErrorHandler

    Err.Raise 9999, "CustomModule", "Something went wrong!"

Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Custom Error"
    Err.Clear
End Sub
```


Handling an error explicitly:
```VB
Sub ExampleErrorHandling()
    On Error Resume Next ' Enable error handling

    Dim x As Integer
    x = 10 / 0 ' This will cause a division by zero error

    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error Occurred"
        Err.Clear ' Reset the error object
    End If
End Sub
```


Err Attributes and Methods :
| Attribute/Method    | Type     | Description |
|---------------------|----------|-------------|
| `Err.Number`       | Integer  | Returns the error number (0 if no error). |
| `Err.Description`  | String   | Provides a textual description of the error. |
| `Err.Source`       | String   | Returns the name of the object or application that generated the error. |
| `Err.HelpFile`     | String   | Returns the path to the associated help file (if available). |
| `Err.HelpContext`  | Integer  | Returns the context ID for a specific error in the help file. |
| `Err.LastDllError` | Integer  | Returns the last system error from a DLL call (used in API error handling). |
| `Err.Clear`        | Method   | Resets (`Err.Number` to `0`) and clears error details. |
| `Err.Raise(Number, [Source], [Description], [HelpFile], [HelpContext])` | Method | Generates a runtime error with a specified number and optional details. |






### File manipulation


Read File :
```VB
Sub ReadFile()
    Dim fileNum As Integer
    Dim lineText As String

    fileNum = FreeFile  'VBA object that gives us a number / ID to associate to a file
    Open "C:\path\file.txt" For Input As #fileNum '`For Input` mode Opens the file in read only mode and associates the content to the ID (# is for file numbers)

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText  ' `Line Input` represents a full line of text without \n and stores it in lineText
        Debug.Print lineText
    Loop

    Close #fileNum
End Sub
```

Other `Open` modes:

| Mode         | Purpose                                      | Read/Write | Behavior |
|-------------|----------------------------------------------|------------|----------|
| **For Input**  | Opens a file for reading only             | Read-only  | Cannot modify the file. |
| **For Output** | Opens a file for writing (creates a new file or overwrites) | Write-only | Deletes existing content if the file exists. |
| **For Append** | Opens a file for writing at the end       | Write-only | Adds new data without overwriting existing content. |
| **For Binary** | Opens a file in binary mode               | Read/Write | Reads and writes **byte-level data**. |
| **For Random** | Opens a file for structured data access   | Read/Write | Uses **fixed-length records** for organized storage. |


Other `Input` modes:
| Method        | Use Case                         | Reads What?          | Handles Line Breaks? |
|--------------|--------------------------------|---------------------|---------------------|
| `Line Input #` | Reading full lines of text    | Full line (String)  |  Yes (but removes newline) |
| `Input #`      | Reading structured data       | Values (Comma-separated) |  No |
| `Get #`        | Reading binary data           | Byte/Fix-Length Record | No |
| `Input( )`     | Reading fixed characters      | Fixed number of chars |  No |
| `Seek + Input$` | Position-based reading       | Chars from a position |  No |


Read File + Array:
```VB
Sub ReadFileToArray()
    Dim fileNum As Integer
    Dim lineText As String
    Dim lines() As String 'dynamic array because no size specified
    Dim i As Integer
    
    fileNum = FreeFile
    Open "C:\path\file.txt" For Input As #fileNum
    
    i = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText 'full line
        ReDim Preserve lines(i) ' ReDim resizes the array, Preserve keeps existing values, lines(i) sets the new size 
        lines(i) = lineText
        i = i + 1
    Loop
    
    Close #fileNum
End Sub
```


Read file + Collection:
```VB
Sub ReadFileToCollection()
    Dim fileNum As Integer
    Dim lineText As String
    Dim lines As New Collection 'Collection is similar to a python dictionary, it allows indexing with custom keys, but allows duplicate keys, O(n) search
    
    fileNum = FreeFile
    Open "C:\path\file.txt" For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lines.Add lineText  ' Add each line to the collection without a key
    Loop
    
    Close #fileNum
End Sub
```



Array + Write to sheet:
```VB
Sub ParseLinesToSheet()
    'Excel sheets
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Target Sheet is the sheet1 of our Worksheet
    'Data
    Dim lines As String() = Array("Name,Age,Country", "Alice,30,USA", "Bob,25,UK", "Charlie,35,Canada") 'or from ReadFile
    Dim i As Integer, j As Integer
    Dim values As Variant

    ' Loop through each line and split it
    For i = LBound(lines) To UBound(lines) ' LBound = lower bound of the array, UBound = upper bound, useful for a For To loop
        values = Split(lines(i), ",")

        ' Loop through each value and place it in the Excel sheet
        For j = LBound(values) To UBound(values)
            ws.Cells(i + 1, j + 1).Value = values(j) 'x,y coordinates 
        Next j
    Next i
End Sub
```






### Auto Filter
The following line can be read as "On the worksheet, find the region of squares adjacent to A1 and filter the column with id colIndex cased on the string criteria item, using the mode xlFilterValues".
`ws.range("A1").CurrentRegion.AutoFilter Field:=colIndex, Criteria1:=item, Operator:=xlFilterValues`

Filters in Excel (via AutoFilter) are cumulative as long as they affect different columns. However, if you're applying multiple filters to the same column, the filters will override each other, and only the last filter will be active for that column.


Filtering modes:

| **Operator**           | **Meaning**                                                            | **Use Case**                                                               |
|------------------------|------------------------------------------------------------------------|----------------------------------------------------------------------------|
| `xlAnd`                | Combines multiple criteria where both conditions must be true.        | Filters for values that meet **all** conditions (logical AND).            |
| `xlOr`                 | Combines multiple criteria where either condition can be true.       | Filters for values that meet **at least one** of the conditions (logical OR). |
| `xlTop10Items`         | Filters for the top `n` items based on a specific column.             | Display the top `n` highest values in the column (e.g., top 10 sales).    |
| `xlBottom10Items`      | Filters for the bottom `n` items based on a specific column.          | Display the bottom `n` lowest values in the column (e.g., bottom 10 sales). |
| `xlTop10Percent`       | Filters for the top `n%` of items based on a specific column.         | Display the top `n%` highest values (e.g., top 10% of income earners).    |
| `xlBottom10Percent`    | Filters for the bottom `n%` of items based on a specific column.      | Display the bottom `n%` lowest values (e.g., bottom 10% of grades).       |
| `xlFilterValues`       | Filters based on a list of predefined values.                          | Display rows that match any of a set of specific values (e.g., product categories). |
| `xlFilterDynamic`      | Filters based on dynamic criteria, such as dates or times.            | Filters based on time ranges or dynamic data comparisons.                 |
| `xlCustom`             | Allows defining custom filters with complex conditions.               | Useful for advanced conditions like **contains**, **does not contain**, etc. |
