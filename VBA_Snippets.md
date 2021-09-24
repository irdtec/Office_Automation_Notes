# VBA scripts and tips #

## Get last used ROW and COLUMN ##

https://excelchamps.com/vba/find-last-row-column-cell/#Last_Row_Column_and_Cell_using_the_Find_Method
  ```vba
  lastRow = sheet.Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    lastColumn = sheet.Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Column
```

## Speed up the macro / Make the macro to run faster ##

Turn screen updating off to speed up your macro code. You won't be able to see what the macro is doing, but it will run faster.
Remember to set the ScreenUpdating property back to True when your macro ends.
```vb
Application.ScreenUpdating = False 
'Some heavy code here
Application.ScreenUpdating = True
```
From <https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.ScreenUpdating>


## Change text on status bar of Excel application ##
```vb
Application.StatusBar  ="Some text"
```

## Get a table as a range and get column values ##

```vb
Dim myRange as Range
Set myRange = worksheet.Range("tblZariGroup[Formato De Negocio]")

'Once the range as been assigned, we use Offset(row,col) to get column  values
myRange.Offset(0,1).Text  
```

## Apply filters to a table ##

```vb
'Set a worksheet
Set ws = ThisWorkbook.Worksheets("Solicitudes")

'Remove all filters from a table.
ws.AutoFilter.ShowAllData
ws.ListObjects("tblSolicitudes").AutoFilter.ShowAllData

'Set filters on different columns.
'Criteria1, Criteria2, etc, are the filters search criteria
ws.ListObjects("tblSolicitudes").Range.AutoFilter Field:=4, _
	Criteria1:="Validar Estructura", Operator:=xlOr, Criteria2:="Zari Descargada"

ws.ListObjects("tblSolicitudes").Range.AutoFilter Field:=13, _
	Criteria1:=sTDoc, Operator:=xlAnd

ws.ListObjects("tblSolicitudes").Range.AutoFilter Field:=14, _
	Criteria1:=sFNegocio, Operator:=xlAnd

ws.ListObjects("tblSolicitudes").Range.AutoFilter Field:=37, _
Criteria1:="=*" & sZari & "*", Operator:=xlAnd
```

## Get the rows that only visible after a filter has been applied ##

We use the following method: .SpecialCells(xlCellTypeVisible) 
```vb
Dim rRequests as range
Set rRequests = worksheet.Range("tblSolicitudes[Id Solicitud]").SpecialCells(xlCellTypeVisible)
```

## Remove filters/table filters from a specific sheet ##

If worksheet.AutoFilterMode Then worksheet.AutoFilterMode = False
```vb
Add a table (list) to an excel sheet programmatically/code
Dim cel As Range, rng As Range
Dim lstRow As Long

'Get the first and last cell positions that hold the table data
'Search for firstCell
Set cel = ws.Cells.Find(What:="Record Type - L", After:=ws.Range("A1"), LookIn:= _
            xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False)
'Get last used row on sheet
        lstRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row

'Set Range with found cell and last row
        Set rng = ws.Range(cel.Address & ":W" & lstRow)

'Create Table(list) object
ws.ListObjects.Add(xlSrcRange, rng, , xlYes, TableStyleName:="TableStyleLight8").Name = _
            "TABLE_NAME"
```

## Add a row to a table (list) ##

```vb
Sub AddDataRow(tableName As String, values() As Variant)
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim col As Integer
    Dim lastRow As Range
Set sheet = ActiveWorkbook.Worksheets("Sheet1")
    Set table = sheet.ListObjects.Item(tableName)
'First check if the last row is empty; if not, add a row
    If table.ListRows.Count > 0 Then
        Set lastRow = table.ListRows(table.ListRows.Count).Range
        For col = 1 To lastRow.Columns.Count
            If Trim(CStr(lastRow.Cells(1, col).Value)) <> "" Then
                table.ListRows.Add
                Exit For
            End If
        Next col
    Else
        table.ListRows.Add
    End If
'Iterate through the last row and populate it with the entries from values()
    Set lastRow = table.ListRows(table.ListRows.Count).Range
    For col = 1 To lastRow.Columns.Count
        If col <= UBound(values) + 1 Then lastRow.Cells(1, col) = values(col - 1)
    Next col
End Sub
```
**Example of use**
```vb
Dim x(2)
x(0) = 1
x(1) = "apple"
x(2) = 2
AddDataRow "Table1", x
```
From <https://stackoverflow.com/questions/8295276/function-or-sub-to-add-new-row-and-data-to-table#14591924>


## Format a Column ##

```vb
'Date to Col H: 
    sheet.Columns(8).NumberFormat = "yyyy/mm/dd"

'Custom Number format
Sheets("Sheet1").Columns(3).NumberFormat = "#,##0"
```
**NumberFormat coding guide:**
https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-us&rs=en-us&ad=us


## SQL Query, NOT IN function not working ##

There could be some problems with the data, on some occasions, because data is NULL, the query could fail with no errors to be shown.

**Fix:** Include  **WHERE  [FIELD NAME] IS NOT NULL**

Source: https://stackoverflow.com/questions/5231712/sql-not-in-not-working

## VBA SQL Get Field names from a record set ##
```vb
Sub FieldNames()
        Dim Rst As Recordset
        Dim s As Field
  
     Set Rst = YourDatabase.OpenRecordset("YourTableName")
  
         For Each s In Rst.Fields
         MsgBox (s.name)
         Next
     Rst.Close
 End Sub
```
From <https://bytes.com/topic/visual-basic/answers/649054-how-do-i-get-column-names-recordset> 


## VBA Open a file dialog ##

```vb
Dim fDialog As FileDialog
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
 
'Show the dialog. -1 means success!
If fDialog.Show = -1 Then
   Debug.Print fDialog.SelectedItems(1) 'The full path to the file selected by the user
End If
```
From <https://analystcave.com/vba-application-filedialog-select-file/> 



## Connect to Access DB ##

How to Connect Excel to Access Database using VBA (exceltip.com)
```vb
Execute an  MS ACCESS Stored procedure
Function Sproc()
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim cnnStr As String
    Dim Rs As New ADODB.Recordset
    Dim StrSproc As String
cnnStr = "Provider=SQLOLEDB;Data Source=DBSource;" & "Initial Catalog=CurrentDb;" & _
             "Integrated Security=SSPI;"
    With cnn
        .CommandTimeout = 900
        .ConnectionString = cnnStr
        .Open
    End With
    With cmd
        .ActiveConnection = cnn
        .CommandType = adCmdStoredProc
        .CommandText = "[StoredProcedureName]"
        .Parameters.Append .CreateParameter("@parameter1", adInteger, adParamInput, , 0)
        .Parameters.Append .CreateParameter("@parameter2", adInteger, adParamInput, , 0)
        .Parameters.Append .CreateParameter("@parameter2", adInteger, adParamInput, , 0)
    End With
    With Rs
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open cmd
    End With
    Application.StatusBar = "Running stored procedure..."
    Set rst = cmd.Execute
End Function
```
From <https://stackoverflow.com/questions/31986552/how-do-i-run-a-stored-procedure-with-parameters-from-excel-vba-string#31991476> 



## Fix issues with field types, Conversion errors, Numbers stored as Text ##

```vb
'Solve issues with field types
   Set sheet = ThisWorkbook.worksheets("Sheet1")
    lastRow = GetLastRow(sheet.Name)
    lastCol = GetLastColumn(sheet.Name)
    With sheet.Range(sheet.Cells(2, 1), sheet.Cells(lastRow, lastCol))
        .NumberFormat = "General"
        .Value = .Value
    End With

''Helper functions
Public Function GetLastRow(sheetName As String) As Long
    Dim tempSheet As Worksheet
    Dim result As Long
    
    Set tempSheet = ThisWorkbook.Worksheets(sheetName)
    If tempSheet.UsedRange.Rows.Count = 1 Then
        result = 1
    Else
    result = tempSheet.Cells.Find(What:="*", _
                            After:=Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    End If
    GetLastRow = result
End Function

'Get the last column from the given sheet on the workbook
Public Function GetLastColumn(sheetName As String) As Long
    Dim tempSheet As Worksheet
    Dim result As Long
    
    Set tempSheet = ThisWorkbook.Worksheets(sheetName)
    If tempSheet.UsedRange.Columns.Count = 1 Then
        result = 1
    Else
    result = tempSheet.Cells.Find(What:="*", _
                           After:=Range("A1"), _
                           LookAt:=xlPart, _
                           LookIn:=xlFormulas, _
                           SearchOrder:=xlByRows, _
                           SearchDirection:=xlPrevious, _
                           MatchCase:=False).Column
    End If
    GetLastColumn = result  
End Function
```

