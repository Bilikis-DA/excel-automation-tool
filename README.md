# excel-automation-tool

📌 **Excel Automation for Data Formatting**
📖 Overview

Developed an automated Excel workbook to streamline data formatting processes, reduce manual effort, and improve accuracy in handling transaction datasets.

🛠️ **What the Automation Does**

Automatically formats transaction amounts into the correct currency format
Re-arranges dataset columns into the required structure
Processes raw data into a clean, ready-to-use format with minimal manual input


⚙️ **Tools Used**
Microsoft Excel
Excel Automation (Macros / Functions)


🚀 **Impact**
Eliminated repetitive manual formatting tasks
Improved data accuracy and consistency
Reduced processing time significantly

💡 **How It Works**

Instead of manually formatting and restructuring datasets, the user simply runs the dataset through the automated workbook, which instantly generates the required output.

💻 Sample VBA Code

Sub Excel_Automation_WKB()

Windows("Excel_Automation_WKB.xlsm").Activate

 MsgBox "Please Select to upload file to format"

 
    Application.ScreenUpdating = False
 
   Dim FNames As Variant
   Dim Cnt As Long
   Dim Wbk As Workbook
   Dim MstWbk As Workbook
   Dim Ws As Worksheet

Application.ScreenUpdating = False
   Set MstWbk = ThisWorkbook    '(FileFilter:="Excel files (*.xls*), *.xls*", MultiSelect:=True)
   FNames = Application.GetOpenFilename(FileFilter:="Text files,*.csv", MultiSelect:=True)
   If Not IsArray(FNames) Then Exit Sub
   For Cnt = 1 To UBound(FNames)
      Set Wbk = Workbooks.Open(FNames(Cnt))
      Wbk.Sheets(1).Copy Before:=MstWbk.Sheets(1)
      'MstWbk.Sheets(1).Name = Left(Wbk.Name, InStr(1, Wbk.Name, ".") - 1)
      MstWbk.Sheets(1).Cells.UnMerge
      Wbk.Close False
   Next Cnt
   Call Globalsales_data
   End Sub
   
   Sub Globalsales_data()
'
' Globalsales_data Macro
' This is to format the sales data
'

'
    
   Application.ScreenUpdating = False
 
        On Error Resume Next
     Cells.Select
    Selection.Copy
    Sheets("Sales_Report").Visible = True
    Sheets("Sales_Report").Select
    Cells.Select
    ActiveSheet.Paste
 
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Columns("I:N").Select
    Selection.NumberFormat = "$#,##0.00"
    Columns("I:N").EntireColumn.AutoFit
    Columns("D:D").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("E7").Select
    Columns("G:G").Select
    Range("D14").Select
    'On Error GoTo 0
    'Sheets("Dispute_Report").Visible = True
 
 
'Sheets("Sales_Report").Visible = True
 
 
Worksheets(Array("Sales_Report")).Copy
        Workbooks("Excel_Automation_WKB.xlsm").Close SaveChanges:=False
 
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
End Sub
