Private Sub Workbook_Open()
 ' LogMacro "Opened QP spreadsheet: " & Range("B2").Value
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
' Author Haris Hassan
' hharis11@hotmail.com


    If SaveAsUI And Not SaveInProgress Then ' Check if Save As dialog is invoked and not already in progress
        Dim fileName As String
        Dim savePath As String
        Dim filePath As Variant
        
        savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray"

        
        If InStr(1, Range("C13").Value, "EMEG") > 0 Then
            savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray\EMEG"
        ElseIf InStr(1, Range("C13").Value, "Preliminary") > 0 Then
            ' Add more subtext conditions as needed
            savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray\PRD's (all)"
        ElseIf InStr(1, Range("C13").Value, "F02") > 0 Then
            ' Add more subtext conditions as needed
            savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray\F02"
        ElseIf InStr(1, Range("C13").Value, "Expert") > 0 Then
            ' Add more subtext conditions as needed
            savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray\Expert Opinion"
        ElseIf InStr(1, Range("C13").Value, "STAD") > 0 Then
            ' Add more subtext conditions as needed
            savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray\EME-EMI-STAD-F01"
        ElseIf InStr(1, Range("C13").Value, "Env") > 0 Then
            ' Add more subtext conditions as needed
            savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray\EME-EMI-STAD-F01"
        ElseIf InStr(1, Range("C13").Value, "EMI") > 0 Then
            ' Add more subtext conditions as needed
            savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray\EME-EMI-STAD-F01"
        ElseIf InStr(1, Range("C13").Value, "Env") > 0 Then
            ' Add more subtext conditions as needed
            savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray\EME-EMI-STAD-F01"
        Else
            ' Default path if no subtext is found
            savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray"
        End If
        
        ' Get the filename from cell B2
        fileName = savePath & "\" & Range("B2").Value & ".pdf"
        
        
        
        ' Set the flag to indicate that Save As is in progress
        SaveInProgress = True

        ' Disable events temporarily
        Application.EnableEvents = False
        
        ' Clear the default path and filename
        ' Application.Dialogs(xlDialogSaveAs).Show savePath & fileName

        ' Application.Dialogs(xlDialogSaveAs).Show CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & savePath, 52
        
        ' Reset the flag after the Save As dialog is closed
        SaveInProgress = False
        ' Re-enable events
        Application.EnableEvents = True

        
        ' Set the default save path
        ' savePath = "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS\1. IN Tray\"

        ' Display the Save As dialog

        filePath = Application.GetSaveAsFilename(InitialFileName:=fileName, FileFilter:="PDF Files (*.pdf), *.pdf", title:="Save As PDF")

        If filePath <> False Then ' Check if a file was selected
            ' Save the workbook as PDF
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePath, Quality:=xlQualityStandard, IncludeDocProperties:=True
            ' Cancel the original save operation
            Cancel = True
        End If
    End If
    LogMacro "Printed QP: " & Replace(fileName, "R:\Central Files\Pending Sites\VIRTUAL WORK TRAYS", "")
End Sub


Sub LogMacro(macroName As String)
    ' Log information about the executed macro

    Dim logFilePath As String
    Dim logFileName As String
    Dim logFileNumber As Integer

    ' Set the path for the log file (adjust the path as needed)
    logFilePath = ThisWorkbook.Path
    logFileName = "QPLog_" & Format(Now, "yyyymmdd") & ".txt"

    ' Open the log file in append mode
    logFileNumber = FreeFile
    Open logFilePath & "\" & logFileName For Append As logFileNumber


   ' Write information to the log file
    Print #logFileNumber, Format(Now(), "hh:mm:ss") & vbTab & Replace(Split(Environ("USERNAME"), ".")(0), "Jennifer", "Jenny") & vbTab & macroName
    'Print #logFileNumber, ""

    ' Close the log file
    Close logFileNumber
End Sub


