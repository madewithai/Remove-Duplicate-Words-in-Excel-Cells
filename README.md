# Remove Duplicate Words in Excel Cells

This VBA script removes duplicate words within cell contents in a specified range of columns in an Excel worksheet. In this example, the script processes columns D to AD.

## How It Works

The script iterates through each cell in the specified range of columns (D to AD), splits the cell content into individual words, and uses a `Collection` to identify and keep only unique words. The cell content is then updated with the unique words.

## Usage

1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module by selecting `Insert > Module`.
4. Copy and paste the VBA code into the module.
5. Adjust the sheet name if necessary (`Sheet1`).
6. Run the script by pressing `F5` or by selecting `Run > Run Sub/UserForm`.

## VBA Code

```vba
Sub RemoveDuplicateWords()
    Dim cell As Range
    Dim ws As Worksheet
    Dim i As Integer
    Dim words() As String
    Dim uniqueWords As String
    Dim wordDict As Collection
    Dim word As Variant
    Dim col As Integer

    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change this to your sheet name

    For col = 4 To 30 ' From column D (4) to column AD (30)
        For Each cell In ws.Columns(col).SpecialCells(xlCellTypeConstants, xlTextValues)
            words = Split(cell.Value, " ")
            Set wordDict = New Collection
            uniqueWords = ""

            On Error Resume Next
            For i = LBound(words) To UBound(words)
                wordDict.Add words(i), CStr(words(i))
                If Err.Number = 0 Then
                    uniqueWords = uniqueWords & words(i) & " "
                End If
                Err.Clear
            Next i
            On Error GoTo 0

            cell.Value = Trim(uniqueWords)
            Set wordDict = Nothing
        Next cell
    Next col
End Sub
```

## Contributing

If you have any suggestions or improvements, feel free to create a pull request or open an issue.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
