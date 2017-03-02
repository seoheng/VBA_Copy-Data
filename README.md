# VBA_Copy-Data
Loop for files to open, and copy data to specified cells within specified sheet



Sub ImportData()

Dim FileRange, FileRange1 As Range
Dim Count, Count1 As Integer 'Count variable
Dim StartRow, StartRow1 As Integer 'Starting row number to place data
Dim r, r1 As Integer
Dim c, c1 As Integer
Dim SourceFile, SourceFile1 As String
Dim CurrentReportingDate As String
Dim path As Variant
  
Application.StatusBar = "Please be patient - clearing the old data"
Application.ScreenUpdating = False

'**********************************************************************************************************************
'Clear Old Data for for Relative Funds
    
    ThisWorkbook.Activate
    Worksheets("Details").Activate
    Range("G15", Range("G15").Offset(3, 30)).Clear
    
'**********************************************************************************************************************
'Loop for Factset Files to be opened and copied for Relative Funds
    
    ThisWorkbook.Activate
    Worksheets("Details").Activate
    Set FileRange = Range("F30", Range("F30").End(xlDown))

    'Initialise variables
    
    StartRow = 30
    Count = 0
            
    'Loop for tabs to clear
    
     For Each d In FileRange
        
        If d.Value <> "" Then
            
           ThisWorkbook.Activate
           Worksheets("Details").Activate
           SourceFile = Cells(StartRow + Count, 6).Value
           r = Cells(StartRow + Count, 9).Value
           c = Cells(StartRow + Count, 10).Value
           
           Worksheets("Print").Activate
           path = Cells(8, 5).Value & "\"
        
           Workbooks.Open FileName:=path & SourceFile, UpdateLinks:=1
           Workbooks(SourceFile).Activate
           
          If InStr(1, ActiveSheet.Cells(8, 1), "Sector", 1) = 0 Then
           
           Range("F9", Range("F9").Offset(0, 1)).Copy
           
           ThisWorkbook.Activate
           Worksheets("Details").Activate
           Cells(r, c).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           Application.CutCopyMode = False
           
           'Close Raw Data file

           Workbooks(SourceFile).Activate
           Workbooks(SourceFile).Close
           
           Count = Count + 1
           
           Else
            Range("G9", Range("G9").Offset(0, 1)).Copy
           
           ThisWorkbook.Activate
           Worksheets("Details").Activate
           Cells(r, c).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           Application.CutCopyMode = False
           
           'Close Raw Data file

           Workbooks(SourceFile).Activate
           Workbooks(SourceFile).Close
           
           Count = Count + 1
           
           End If
                       
        End If
        
      Next d
       
'Cursor Default Location
    
    ThisWorkbook.Activate
    Worksheets("Details").Activate
    Range("G15", Range("G15").Offset(3, 30)).NumberFormat = "#,##0.0000"
         
    
'**********************************************************************************************************************
