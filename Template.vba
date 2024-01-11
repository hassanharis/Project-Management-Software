Sub CreateFolderInPendingSites()
' Author Haris Hassan
' hharis11@hotmail.com

Dim folderPath  As String
Dim folder_1_path, folder_2_path, folder_3_path, folder_4_path As String

    ' create main folder
    folderPath = "R:\Central Files\Pending Sites\" & Sheets("Template").Range("B1").Value
    
    If InStr(1, Sheets("Template").Range("A11").Value, "SSMC") Then
        folderPath = "R:\Central Files\Pending Sites\SSMC TCI RFQ\" & Sheets("Template").Range("B1").Value
    End If
    
    
    ' Check if the folder path is not empty
    If folderPath <> "" Then
        ' Check if the folder already exists
        If FolderExists(folderPath) Then
            ' Open the existing folder in Windows Explorer
            Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
        Else
            ' Prompt the user to create the folder
            'If MsgBox("The folder does not exist. Do you want to create it?", vbQuestion + vbYesNo) = vbYes Then
                ' Create the folder
                MkDir folderPath
                ' Open the newly created folder in Windows Explorer
                Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
           ' End If
        End If
    Else
        ' Display a message if the cell is empty
        MsgBox "Please enter a valid folder path in cell " & Sheets("Template").Range("B1").Value, vbExclamation
    End If

    LogMacro "Created Folder In: " & Replace(folderPath, "R:\Central Files", "")

End Sub

Function FolderExists(folderPath As String) As Boolean
' Author Haris Hassan
' hharis11@hotmail.com

    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function


Private Sub CommandButton1_Click()
' Author Haris Hassan
' hharis11@hotmail.com

    Call CreateFolderInPendingSites
End Sub


Sub CreateFolderinCentralFiles()
' Author Haris Hassan
' hharis11@hotmail.com

    Dim folderPath As String
    Dim cellValue As String

    ' Get the cell value from cell A1
    cellValue = Range("D11").Value

    ' Check if the cell value is not empty
    If cellValue <> "" Then
        ' Determine the folder path based on the first digit
            Dim firstDigit As Integer
            firstDigit = CInt(Left(cellValue, 1))

            ' Determine the folder path based on the first digit
            Select Case firstDigit
                Case 1
                    folderPath = "R:\Central Files\10000 - 19999  ACT\" & cellValue
                Case 2
                    folderPath = "R:\Central Files\20000 - 29999  NSW\" & cellValue
                Case 3
                    folderPath = "R:\Central Files\30000 - 39999  VIC\" & cellValue
                Case 4
                    folderPath = "R:\Central Files\40000 - 49999 QLD\" & cellValue
                Case 5
                    folderPath = "R:\Central Files\50000 - 59999  SA\" & cellValue
                Case 6
                    folderPath = "R:\Central Files\60000 - 69999 WA\" & cellValue
                Case 7
                    folderPath = "R:\Central Files\70000 - 79999  TAS\" & cellValue
                Case 8
                    folderPath = "R:\Central Files\80000 - 89999 NT\" & cellValue
                Case 0
                    Select Case CStr(Left(cellValue, 5))
                        Case "00500"
                            ' Extract the text after the dash in the cell value
                            folderKeyword = " " & Split(cellValue, "-")(1)
                            
                            Dim foundFolder As Boolean
                            Dim basePath As String
                            basePath = "R:\Central Files\00000 - 04999 Other Reports\00500 - NAD\"
                        
                            ' Loop through the folders in the first base path
                            
                            Dim folder As String
                            folder = Dir(basePath, vbDirectory)
                        
                            Do While folder <> ""
                                ' Check if the folder name contains the specified value
                                If InStr(1, folder, folderKeyword, vbTextCompare) > 0 Then
                                    ' Combine the base path and folder name to get the complete folder path
                                    folderPath = basePath & folder
                                    foundFolder = True
                                    Exit Do
                                End If
                                folder = Dir
                            Loop
                            ' Set the search directory
                            ' Check if a matching folder was found
                            If foundFolder Then
                               
                            Else
                                folderPath = basePath & "Antenna Upload" & folderKeyword
                                MsgBox "No matching folder found but I'll create one " & folderPath, vbExclamation
                            End If
                            
                            
                        Case "00150"
                            folderPath = "R:\Central Files\00000 - 04999 Other Reports\" & CStr(Left(cellValue, 5)) & "\" & cellValue & "\"
                        Case "01065"
                            folderPath = "R:\Central Files\00000 - 04999 Other Reports\01065 - Radman Sales"
                        Case Else
                            folderPath = "R:\Central Files\00000 - 04999 Other Reports\" & CStr(Left(cellValue, 5)) & "\"
                    End Select
                ' Add more cases as needed
                Case Else
                    MsgBox "Invalid first digit for determining the folder path.", vbExclamation
                    Exit Sub
            End Select
            
            If FolderExists(folderPath) Then
            ' Open the existing folder in Windows Explorer
                Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
            Else
                ' Prompt the user to create the folder
                'If MsgBox("The folder does not exist. Do you want to create it?", vbQuestion + vbYesNo) = vbYes Then
                    ' Create the folder
                    MkDir folderPath
                    ' Open the newly created folder in Windows Explorer
                    Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
               ' End If
            End If
    Else
        ' Display a message if the cell is empty
        MsgBox "Please enter a valid value in cell A1.", vbExclamation
    End If
    
    LogMacro "Created/opened a Folder in CentralFiles: " & Replace(folderPath, "R:\Central Files", "")
End Sub

Private Sub CommandButton2_Click()
' Author Haris Hassan
' hharis11@hotmail.com

    Call CreateFolderinCentralFiles
End Sub


Sub FindAndMoveQPFolder()

' Author Haris Hassan
' hharis11@hotmail.com

    Dim DfolderPath As String
    Dim cellValue As String

    ' Get the cell value from cell A1
    cellValue = Range("D11").Value

    
    ' Determine the folder path based on the first digit
    If Left(cellValue, 1) Then
        Dim firstDigit As Integer
        firstDigit = CInt(Left(cellValue, 1))
    
        ' Determine the folder path based on the first digit
        Select Case firstDigit
            Case 1
                DfolderPath = "R:\Central Files\10000 - 19999  ACT\"
            Case 2
                DfolderPath = "R:\Central Files\20000 - 29999  NSW\"
            Case 3
                DfolderPath = "R:\Central Files\30000 - 39999  VIC\"
            Case 4
                DfolderPath = "R:\Central Files\40000 - 49999 QLD\"
            Case 5
                DfolderPath = "R:\Central Files\50000 - 59999  SA\"
            Case 6
                DfolderPath = "R:\Central Files\60000 - 69999 WA\"
            Case 7
                DfolderPath = "R:\Central Files\70000 - 79999  TAS\"
            Case 8
                DfolderPath = "R:\Central Files\80000 - 89999 NT\"
            Case 0
                Select Case CStr(Left(cellValue, 5))
                    Case "00500"
                        ' Extract the text after the dash in the cell value
                        folderKeyword = " " & Split(cellValue, "-")(1) & " "
                        
                        ' Set the search directory
                        searchDirectory = "R:\Central Files\00000 - 04999 Other Reports\00500 - NAD\"
                    
                        ' Use the Dir function to find the folder with the matching text
                        DfolderPath = Dir(searchDirectory & "*" & folderKeyword & "*", vbDirectory)
                    Case "00150"
                        DfolderPath = "R:\Central Files\00000 - 04999 Other Reports\" & CStr(Left(cellValue, 5)) & "\" & cellValue & "\"
                    Case "01065"
                        DfolderPath = "R:\Central Files\00000 - 04999 Other Reports\01065 - Radman Sales"
                    Case Else
                        DfolderPath = "R:\Central Files\00000 - 04999 Other Reports\" & CStr(Left(cellValue, 5)) & "\"
                End Select
            ' Add more cases as needed
            Case Else
                MsgBox "Invalid first digit for determining the folder path.", vbExclamation
                Exit Sub
        End Select
       End If
       
    DfolderPath = DfolderPath & cellValue
    
    If FolderExists(DfolderPath) Then
    ' Open the existing folder in Windows Explorer
        ' Call Shell("explorer.exe """ & folderPath & cellValue & """", vbNormalFocus)
    Else
        ' Prompt the user to create the folder
        'If MsgBox("The folder does not exist. Do you want to create it?", vbQuestion + vbYesNo) = vbYes Then
            ' Create the folder
            MkDir DfolderPath
            ' Open the newly created folder in Windows Explorer
            ' Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
       ' End If
    End If
            
    Dim folderName As String
    Dim sourcePath As String
    Dim destinationPath As String
    Dim fs As Object
    Dim sourceFolder As Object
    Dim folderCount As Integer
    Dim newFolderName As String
    Dim oldPath As String
    Dim newPath As String



    ' Get destination path from cell C3
    destinationPath = DfolderPath

    
        ' Specify the base folder paths where you want to search for the folder
    ' Update these paths based on your requirements
    Dim folderNamePart As String
    Dim SfolderPath As String
    Dim basePath1 As String
    Dim basePath2 As String
    basePath1 = "R:\Central Files\Pending Sites\SSMC TCI RFQ\"
    basePath2 = "R:\Central Files\Pending Sites\"

    ' Loop through the folders in the first base path
    Dim foundFolder As Boolean
    Dim folder As String
    'folder = Dir(basePath1, vbDirectory)

    'Do While folder <> ""
        ' Check if the folder name contains the specified value
     '   If InStr(1, folder, cellValue, vbTextCompare) > 0 Then
            ' Combine the base path and folder name to get the complete folder path
     '       SfolderPath = basePath1 & folder
     '       foundFolder = True
     '       Exit Do
     '   End If
     '   folder = Dir
    'Loop

    ' If the folder is not found in the first path, check the second path
   ' If Not foundFolder Then
     '   folder = Dir(basePath2, vbDirectory)
     '   Do While folder <> ""
            ' Check if the folder name contains the specified value
      '      If InStr(1, folder, cellValue, vbTextCompare) > 0 Then
                ' Combine the base path and folder name to get the complete folder path
      '          SfolderPath = basePath2 & folder
       '         foundFolder = True
      '          Exit Do
       '     End If
       '     folder = Dir
      '  Loop
    'End If

    ' Check if a matching folder was found
    
    SfolderPath = "R:\Central Files\Pending Sites\" & Sheets("Template").Range("B1").Value
    
    If InStr(1, Sheets("Template").Range("A11").Value, "SSMC") Then
        SfolderPath = "R:\Central Files\Pending Sites\SSMC TCI RFQ\" & Sheets("Template").Range("B1").Value
    End If
    
    sourcePath = SfolderPath
    
    
    
    If destinationPath <> "" Then
         ' Open the folder using ShellExecute
        ' Create FileSystemObject
        Set fs = CreateObject("Scripting.FileSystemObject")
    
        ' Set the source folder object
        Set sourceFolder = fs.GetFolder(SfolderPath)
    
        ' Get the count of folders in the destination path
        folderCount = fs.GetFolder(destinationPath).SubFolders.Count
    
        ' Generate the new folder name with the count
        newFolderName = CStr(folderCount + 1) & ". " & sourceFolder.Name
    
        ' Build the old and new paths
        oldPath = sourceFolder.Path
        newPath = destinationPath & "\" & newFolderName
    
        ' Move the folder to the destination path
        fs.MoveFolder oldPath, newPath
        
        Call Shell("explorer.exe """ & newPath & """", vbNormalFocus)
        ' Display a success message
        MsgBox "Folder moved successfully! " & newPath, vbInformation

    Else
        ' Display a message if the cell is empty
        MsgBox "No matching folder found for: " & cellValue, vbExclamation
    End If
    
    LogMacro "Moved QP folder from pending sites to: " & Replace(newPath, "R:\Central Files", "")
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

