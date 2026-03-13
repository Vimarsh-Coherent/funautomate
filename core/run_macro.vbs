' VBScript to open an Excel workbook and run a macro
' Usage: cscript run_macro.vbs "C:\path\to\file.xlsm" "MacroName"

If WScript.Arguments.Count < 2 Then
    WScript.Echo "Usage: cscript run_macro.vbs <xlsm_path> <macro_name>"
    WScript.Quit 1
End If

Dim xlsm_path, macro_name
xlsm_path = WScript.Arguments(0)
macro_name = WScript.Arguments(1)

Dim objExcel
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.DisplayAlerts = False
objExcel.AutomationSecurity = 1

On Error Resume Next

WScript.Echo "STATUS: Opening workbook..."
Dim objWorkbook
Set objWorkbook = objExcel.Workbooks.Open(xlsm_path, False, False)

If Err.Number <> 0 Then
    WScript.Echo "ERROR_OPEN: " & Err.Description
    objExcel.Quit
    WScript.Quit 2
End If
Err.Clear

WScript.Echo "STATUS: Running macro..."
objExcel.Run macro_name

If Err.Number <> 0 Then
    WScript.Echo "ERROR_MACRO: " & Err.Description
    Err.Clear
End If

WScript.Echo "STATUS: Saving..."
objWorkbook.Save

If Err.Number <> 0 Then
    WScript.Echo "ERROR_SAVE: " & Err.Description
    Err.Clear
End If

objWorkbook.Close False
objExcel.Quit

WScript.Echo "SUCCESS"
WScript.Quit 0
