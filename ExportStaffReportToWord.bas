Attribute VB_Name = "ExportStaffReportToWord"
Option Compare Database

' Windows API Declarations and Types
Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' Utility Functions
Function GetSavePath() As String
    Dim SaveAsDialog As OPENFILENAME
    Dim lngResult As Long
    
    With SaveAsDialog
        .lStructSize = Len(SaveAsDialog)
        .lpstrFilter = "Word Documents (*.doc)" & Chr(0) & "*.doc" & Chr(0)
        .lpstrFile = String(257, 0)
        .nMaxFile = Len(.lpstrFile) - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = .nMaxFile
        .lpstrInitialDir = "D:\"
        .lpstrTitle = "Select a Location to Save the Report"
        .flags = 0
    End With
    
    lngResult = GetSaveFileName(SaveAsDialog)
    
    If lngResult <> 0 Then
        GetSavePath = TrimNull(SaveAsDialog.lpstrFile)
    End If
End Function

Function TrimNull(ByVal strValue As String) As String
    Dim intPos As Integer
    
    intPos = InStr(strValue, Chr(0))
    If intPos > 0 Then
        TrimNull = Left(strValue, intPos - 1)
    Else
        TrimNull = strValue
    End If
End Function

' Main function to Export the Report to Word DOC format
Sub ExportStaffReportToWord()

    Dim wordApp As Object
    Dim wordDoc As Object
    Dim reportName As String
    Dim tempRTFPath As String
    Dim docPath As String

    ' Set the name of the report to export
    reportName = "staff"

    ' Generate a temporary path for the RTF file
    tempRTFPath = Environ("Temp") & "\" & reportName & ".rtf"

    ' Get the path to save the DOC
    docPath = GetSavePath()
    If docPath = "" Then
        MsgBox "Operation cancelled by user.", vbExclamation
        Exit Sub
    End If

    ' Export the report to RTF
    DoCmd.OutputTo ObjectType:=acOutputReport, _
                    ObjectName:=reportName, _
                    OutputFormat:=acFormatRTF, _
                    OutputFile:=tempRTFPath

    ' Open the RTF in Word and save as DOC
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set wordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    Set wordDoc = wordApp.Documents.Open(tempRTFPath)
    wordDoc.SaveAs2 docPath, FileFormat:=0 ' 0 corresponds to DOC format
    wordDoc.Close
    Kill tempRTFPath ' Delete the temporary RTF file

    ' Show a message to indicate completion
    MsgBox "Report exported successfully as .doc format!", vbInformation

End Sub


