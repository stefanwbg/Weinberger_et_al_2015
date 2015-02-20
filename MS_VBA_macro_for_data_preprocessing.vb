'Script starts automatically when opening the file and opens GUI
Sub Auto_open()
Dim byValue As Byte
'Content message box
byValue = MsgBox("REMO climate data to DYRESM input file" &_
vbCrLf & vbCrLf &_
 "Make sure before starting the script:"&_
vbCrLf & vbCrLf & _
"- Textfiles containing REMO Data are located in folder C:\REMO" &_
 vbCrLf &_
"- Names of Textfiles:" &_
 vbCrLf & _
"    ACLCOV.txt" &_
 vbCrLf &_
"    APRC.txt" &_
 vbCrLf &_
"    APRL.txt" &_
vbCrLf &_
"    SRADS.txt" &_
vbCrLf &_
"    TEMP2.txt" &_
vbCrLf & _
"    WIND10.txt" &_
vbCrLf &_
"- No special characters allowed" &_
vbCrLf & vbCrLf &_
"Meteorological input file for DYRESM will be created in folder C:\REMO"_
, 1, "Please note:")
If byValue = 1 Then
Call REMO_Makro
 'ElseIf byValue = 2 Then
'Application.Quit
End If
End Sub

'Conversion and customisation of REMO data        
Sub REMO_Makro()
'Customisation of separators in MS Excel
With Application
.DecimalSeparator = "."
.ThousandsSeparator = ","
End With
    
'Step 1: Import of text files containing REMO data: total cloud cover (ACLCOV), convective precipitation (APRC), large scale precipitation (APRL), net surface solar radiation (SRADS), air temperature (TEMP) and wind speed (WIND) 
Sheets.Add After:=Sheets(Sheets.Count)
Sheets.Add After:=Sheets(Sheets.Count)
Sheets.Add After:=Sheets(Sheets.Count)
Sheets("Table1").Select
   'REMO data: Example total cloud cover (ACLCOV)
   ‘This step is repeated for all REMO variables  
With ActiveSheet.QueryTables.Add(Connection:=_
"TEXT;C:\REMO\Remo_Excel\ACLCOV.txt",  Destination:=Range("$A$1"))
.Name = "ACLCOV"
.FieldNames = True
.RowNumbers = False
.FillAdjacentFormulas = False
.PreserveFormatting = True
.RefreshOnFileOpen = False
.RefreshStyle = xlInsertDeleteCells
.SavePassword = False
.SaveData = True
.AdjustColumnWidth = True
.RefreshPeriod = 0
.TextFilePromptOnRefresh = False
.TextFilePlatform = 850
.TextFileStartRow = 1
.TextFileParseType = xlDelimited
.TextFileTextQualifier = xlTextQualifierDoubleQuote
.TextFileConsecutiveDelimiter = True
.TextFileTabDelimiter = True
.TextFileSemicolonDelimiter = False
.TextFileCommaDelimiter = False
.TextFileSpaceDelimiter = True
.TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
.TextFileTrailingMinusNumbers = True
.Refresh BackgroundQuery:=False
End With
     
'Step 2: Adaption of columns, example Wind speed
‘This step is repeated for all columns containing imported REMO variables
Sheets("Table6").Select
Columns("A:C").Select
   Selection.Delete Shift:=xlToLeft
   Columns("C:G").Select
   Selection.Delete Shift:=xlToLeft
   Range("C1").Select
   ActiveCell.FormulaR1C1 = "Mean"
   Sheets("Table6").Select
   Sheets("Table6").Name = "WIND"
   
'Step 3: Adaption of date format
'Adaption of number of lines
Dim Rows As Long
   Rows = Cells(Rows.Count,
  1).End(xlUp).Row
   Sheets.Add After:=Sheets(Sheets.Count)
   Sheets("ACLCOV").Select
   Columns("A:B").Select
   Selection.Copy
   Sheets("Table7").Select
   ActiveSheet.Paste
   Columns("B:B").Select
   Application.CutCopyMode = False
   Selection.NumberFormat = "0.000"
   Range("C2").Select
   ActiveCell.FormulaR1C1 = "=YEAR(RC[-2])"
   Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" &  Rows)
Range("D2").Select
ActiveCell.FormulaR1C1 = "=TEXT(RC[-3]-DATE(YEAR(RC[-3]),1,1)+1,""000"")"
Range("D2").Select
Selection.AutoFill Destination:=Range("D2:D" &  Rows)
Range("E2").Select
ActiveCell.FormulaR1C1 = "=(RC[-2]&RC[-1])+RC[-3]"
Columns("E:E").Select
Selection.NumberFormat = "0.000"
Range("E2").Select
Selection.AutoFill Destination:=Range("E2:E" & Rows)
Range("E1").Select
ActiveCell.FormulaR1C1 = "YrDayNum"
Columns("E:E").Select
Selection.Copy
Columns("F:F").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, 
SkipBlanks _:=False, Transpose:=False
Application.CutCopyMode = False
Selection.NumberFormat = "0.000"
Columns("A:E").Select
Selection.Delete Shift:=xlToLeft
Range("F6").Select
Sheets("Table7").Select
Sheets("Table7").Name = "Overall"
    
'Step 4: Adaption of units, calculation of overall precipitation and vapour pressure, assembling in 1 sheet
'Calculation of overall precipitation
Sheets("APRL").Select
Columns("C:C").Select
Selection.Copy
Sheets("APRC").Select
Columns("D:D").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, 
SkipBlanks _:=False, Transpose:=False
Range("E2").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
Range("E2").Select
Selection.AutoFill Destination:=Range("E2:E" & Rows)
ActiveWindow.SmallScroll Down:=-21
ActiveCell.FormulaR1C1 = "=(RC[-2]+RC[-1])/1000"
Range("E2").Select
Selection.AutoFill Destination:=Range("E2:E" & Rows)
Range("E1").Select
ActiveCell.FormulaR1C1 = "Rain [m]"
Range("E1").Select
ActiveCell.FormulaR1C1 = "Rain_[m]"
Columns("E:E").Select
Selection.Copy
Sheets("Overall").Select
Columns("G:G").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, 
SkipBlanks _:=False, Transpose:=False
'Adaption of net surface solar radiation unit
Sheets("SRADS").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "SW_[W/m^2]"
Columns("C:C").Select
Selection.Copy
Sheets("Overall").Select
Columns("B:B").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, 
SkipBlanks _:=False, Transpose:=False
'Adaption of total cloud cover unit
Sheets("ACLCOV").Select
Range("C1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "CloudCover"
Columns("C:C").Select
Selection.Copy
Sheets("Overall").Select
Columns("C:C").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, 
SkipBlanks _:=False, Transpose:=False
'Adaption of 2m temperature unit
Sheets("TEMP").Select
Range("D2").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "=RC[-1]-273.15"
Range("D2").Select
Selection.AutoFill Destination:=Range("D2:D" & Rows)
Range("D1").Select
ActiveCell.FormulaR1C1 = "Tair_[°C]"
Columns("D:D").Select
Selection.Copy
Sheets("Overall").Select
Columns("D:D").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, 
SkipBlanks _:=False, Transpose:=False
'Adaption of wind speed unit
Sheets("WIND").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "Wind_Speed_[m/S]"
Columns("C:C").Select
Selection.Copy
Sheets("Overall").Select
Columns("F:F").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, 
SkipBlanks _:=False, Transpose:=False
'Calculation of vapour pressure and adaption of unit
Sheets("TEMP2").Select
Range("E1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "Vapour_press_[hPA]"
Range("E2").Select
ActiveCell.FormulaR1C1 = "=6.1*10^((7.5*RC[-1])/(RC[-1]+237.2))"
Range("E2").Select
Selection.AutoFill Destination:=Range("E2:E" & Rows)
Columns("E:E").Select
Selection.Copy
Sheets("Overall").Select
Columns("E:E").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, 
SkipBlanks _:=False, Transpose:=False
   
'Step 5: Adaption of format and adding file heading of DYRESM meteorological input file
Columns("A:A").ColumnWidth = 13
Cells.Select
Application.CutCopyMode = False
With Selection
.HorizontalAlignment = xlLeft
.VerticalAlignment = xlBottom
.WrapText = False
.Orientation = 0
.AddIndent = False
.IndentLevel = 0
.ShrinkToFit = False
.ReadingOrder = xlContext
.MergeCells = False
End With
Range("B3").Select
Columns("B:B").ColumnWidth = 14
Columns("C:C").ColumnWidth = 13
Columns("E:E").ColumnWidth = 20
Columns("F:F").ColumnWidth = 20
Rows("1:5").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Range("A1").Select
ActiveCell.FormulaR1C1 = "<#3>"
Range("A2").Select
ActiveCell.FormulaR1C1 = "Meteorology Lake Ammersee"
Range("A3").Select
ActiveCell.FormulaR1C1 = "3600 # Met input data time step"
Range("A4").Select
ActiveCell.FormulaR1C1 = "CLOUD_COVER # longwave radiation indicator (NETT_LW, INCIDENT_LW, CLOUD_COVER)"
Range("A5").Select
ActiveCell.FormulaR1C1 = "FIXED_HT 90 # sensor type (FLOATING, FIXED_HT), height in metres (above water surface, above lake bottom)"
Range("B3").Select
  
'Step 6: Delete redundant sheets
Application.DisplayAlerts = False
Sheets("WIND").Select
ActiveWindow.SelectedSheets.Delete
Sheets("TEMP").Select
ActiveWindow.SelectedSheets.Delete
Sheets("SRADS").Select
ActiveWindow.SelectedSheets.Delete
Sheets("APRL").Select
ActiveWindow.SelectedSheets.Delete
Sheets("APRC").Select
ActiveWindow.SelectedSheets.Delete
Sheets("ACLCOV").Select
ActiveWindow.SelectedSheets.Delete
      
'Step 7: Export file as text file
ChDir "C:\REMO\Remo_Excel"
ActiveWorkbook.SaveAs
Filename:="C:\REMO\Remo_Excel\Meteo.prn",
FileFormat _:=xlTextPrinter, CreateBackup:=False
MsgBox "Data conversion  successful!" & 
strText, 0, "Hinweis"
Application.Visible = False
Workbooks("Meteo.prn").Close SaveChanges:=False 
Application.DisplayAlerts = True
End Sub
 


