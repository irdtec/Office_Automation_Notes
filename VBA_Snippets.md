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

## Copy & Paste Over Existing Row / Column ##

This will copy row 1 and paste it into the existing row 5:
```vb
Range("1:1").Copy Range("5:5")
This will copy column C and paste it into column E:
Range("C:C").Copy Range("E:E")
```
From <https://www.automateexcel.com/vba/copy-column-row>

## Copy & Insert Row / Column ##

Instead you can insert the copied row or column and shift the existing rows or columns to make room.
This will copy row 1 and insert it into row 5, shifting the existing rows down:
```vb
Range("1:1").Copy
Range("5:5").Insert
```
This will copy column C and insert it into column E, shifting the existing columns to the right:
```vb
Range("C:C").Copy
Range("E:E").Insert
```

## Copy Entire Row ##

Below we will show you several ways to copy row 1 and paste into row 5.
```vb
Range("1:1").Copy Range("5:5")
Range("A1").EntireRow.Copy Range("A5")
Rows(1).Copy Rows(5)
```

## Cut and Paste Rows ##

Simply use Cut instead of Copy to cut and paste rows:
```vb
Rows(1).Cut Rows(5)
```

## Copy Multiple Rows ##

Here are examples of copying multiple rows at once:
```vb
Range("5:7").Copy Range("10:13")
Range("A5:A7").EntireRow.Copy Range("A10:A13")
Rows(5:7).Copy Rows(10:13)
```
From <https://www.automateexcel.com/vba/copy-column-row> 


## Copy Entire Column ##

You can copy entire columns similarily to copying entire rows:
```vb
Range("C:C").Copy Range("E:E")
Range("C1").EntireColumn.Copy Range("C1").EntireColumn
Columns(3).Copy Range(5)
```
## Cut and Paste Columns ##

Simply use Cut instead of Copy to cut and paste columns:
```vb
Range("C:C").Cut Range("E:E")
```

## Copy Multiple Columns ##

Here are examples of copying multiple columns at once:
```vb
Range("C:E").Copy Range("G:I")
Range("C1:E1").EntireColumn.Copy Range("G1:I1")
Columns(3:5).Copy Columns(7:9)
```

## Copy Rows or Columns to Another Sheet ##

To copy to another sheet, simply use the Sheet Object:
```vb
Sheets("sheet1").Range("C:E").Copy Sheets("sheet2").Range("G:I")
```
From <https://www.automateexcel.com/vba/copy-column-row> 

## Cut Rows or Columns to Another Sheet ##

You can use the exact same technique to cut and paste rows or columns to another sheet.
```vb
Sheets("sheet1").Range("C:E").Cut Sheets("sheet2").Range("G:I")
```
From <https://www.automateexcel.com/vba/copy-column-row>


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


## Finds local path for a OneDrive file URL, using environment variables of OneDrive ##
```vb
Private Function LocalFullName$(ByVal fullPath$)
    'Finds local path for a OneDrive file URL, using environment variables of OneDrive
    'Reference https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive
    'Authors: Philip Swannell 2019-01-14, MatChrupczalski 2019-05-19, Horoman 2020-03-29, P.G.Schild 2020-04-02

    Dim ii&
    Dim iPos&
    Dim oneDrivePath$
    Dim endFilePath$

    If Left(fullPath, 8) = "https://" Then 'Possibly a OneDrive URL
        If InStr(1, fullPath, "my.sharepoint.com") <> 0 Then 'Commercial OneDrive
            'For commercial OneDrive, path looks like "https://companyName-my.sharepoint.com/personal/userName_domain_com/Documents" & file.FullName)
            'Find "/Documents" in string and replace everything before the end with OneDrive local path
            iPos = InStr(1, fullPath, "/Documents") + Len("/Documents") 'find "/Documents" position in file URL
            endFilePath = Mid(fullPath, iPos) 'Get the ending file path without pointer in OneDrive. Include leading "/"
        Else 'Personal OneDrive
            'For personal OneDrive, path looks like "https://d.docs.live.net/d7bbaa#######1/" & file.FullName
            'We can get local file path by replacing "https.." up to the 4th slash, with the OneDrive local path obtained from registry
            iPos = 8 'Last slash in https://
            For ii = 1 To 2
                iPos = InStr(iPos + 1, fullPath, "/") 'find 4th slash
            Next ii
            endFilePath = Mid(fullPath, iPos) 'Get the ending file path without OneDrive root. Include leading "/"
        End If
        endFilePath = Replace(endFilePath, "/", Application.PathSeparator) 'Replace forward slashes with back slashes (URL type to Windows type)
        For ii = 1 To 3 'Loop to see if the tentative LocalWorkbookName is the name of a file that actually exists, if so return the name
            oneDrivePath = Environ(Choose(ii, "OneDriveCommercial", "OneDriveConsumer", "OneDrive")) 'Check possible local paths. "OneDrive" should be the last one
            If 0 < Len(oneDrivePath) Then
                LocalFullName = oneDrivePath & endFilePath
                Exit Function 'Success (i.e. found the correct Environ parameter)
            End If
        Next ii
        'Possibly raise an error here when attempt to convert to a local file name fails - e.g. for "shared with me" files
        LocalFullName = vbNullString
    Else
        LocalFullName = fullPath
    End If
    
End Function
```
