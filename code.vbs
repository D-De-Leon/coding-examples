Function CheckUserAuthorization() As Boolean
    Dim authorizedUsers As Variant
    Dim currentUser As String
    Dim i As Integer

    ' Define a list of authorized users
    authorizedUsers = Array("authorizedUser1", "authorizedUser2", "authorizedUser3") ' Add actual usernames here

    ' Get the current username
    currentUser = Environ("Username")

    ' Debug: Print the current username
    MsgBox "Current user: " & currentUser

    ' Check if the current user is authorized
    CheckUserAuthorization = False
    For i = LBound(authorizedUsers) To UBound(authorizedUsers)
        If currentUser = authorizedUsers(i) Then
            CheckUserAuthorization = True
            Exit Function
        End If
    Next i
End Function

Sub ProcessRequestForms()
    ' Check user authorization
    If Not CheckUserAuthorization() Then
        MsgBox "You are not authorized to run this script.", vbExclamation
        Exit Sub
    End If

    ' The rest of your code here
    Dim folderPath As String
    Dim inventoryPath As String
    Dim fileName As String
    Dim wbRequest As Workbook
    Dim wsRequest As Worksheet
    Dim wbMaster As Workbook
    Dim wsMaster As Worksheet
    Dim lastRow As Long
    Dim nextRow As Long
    Dim newValue1 As Variant
    Dim newValue2 As Variant
    Dim isDuplicate As Boolean
    Dim cell As Range
    Dim notTransferredFiles As String
    Dim userName As String
    Dim highlightColor As Variant ' Use Variant to handle both Long and -1 values
    Dim newFileName As String
    Dim masterFilePath As String
    Dim baseFileName As String
    Dim fileExtension As String
    Dim fileCounter As Integer

    ' Set the master workbook path
    masterFilePath = "C:\path\to\your\masterfile.xlsx" ' Update with the actual path to your master workbook

    ' Open the master workbook
    Set wbMaster = Workbooks.Open(masterFilePath)
    Set wsMaster = wbMaster.Sheets("Project Request") ' Update with the actual sheet name in the master workbook

    ' Get the OneDrive folder path
    userName = Environ("Username")

    ' Set the folder paths
    folderPath = "C:\Users\" & userName & "\OneDrive - DukeEnergy\Request Form Intake\"
    inventoryPath = "C:\Users\" & userName & "\OneDrive - DukeEnergy\Request For Inventory\"

    ' Check if the folder paths exist
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder path not found: " & folderPath, vbExclamation
        wbMaster.Close SaveChanges:=False
        Exit Sub
    End If

    If Dir(inventoryPath, vbDirectory) = "" Then
        MsgBox "Inventory path not found: " & inventoryPath, vbExclamation
        wbMaster.Close SaveChanges:=False
        Exit Sub
    End If

    ' Initialize not transferred files string
    notTransferredFiles = ""

    ' Initialize the fileName with the first file in the folder
    On Error Resume Next
    fileName = Dir(folderPath & "*.xlsx")
    If Err.Number <> 0 Then
        MsgBox "Error initializing file search: " & Err.Description, vbExclamation
        Err.Clear
        wbMaster.Close SaveChanges:=False
        Exit Sub
    End If
    On Error GoTo 0

    ' Check if there are any files to process
    If fileName = "" Then
        MsgBox "No files found in the folder: " & folderPath, vbExclamation
        wbMaster.Close SaveChanges:=False
        Exit Sub
    End If

    Do While fileName <> ""
        ' Debug: Display current file name
        MsgBox "Processing file: " & fileName

        ' Open the request form
        On Error Resume Next
        Set wbRequest = Workbooks.Open(folderPath & fileName)
        If Err.Number <> 0 Then
            MsgBox "Error opening file: " & fileName, vbExclamation
            notTransferredFiles = notTransferredFiles & fileName & vbCrLf
            Err.Clear
            GoTo NextFile
        End If
        On Error GoTo 0

        Set wsRequest = wbRequest.Sheets(1)

        ' Check sheet 1 for data
        If IsEmpty(wsRequest.Range("B9").Value) And IsEmpty(wsRequest.Range("B10").Value) And IsEmpty(wsRequest.Range("B33").Value) Then
            ' If sheet 1 is empty, check sheet 2
            If wbRequest.Sheets.Count > 1 Then
                Set wsRequest = wbRequest.Sheets(2)
            End If
        End If

        ' Find the next available row in the master sheet
        lastRow = wsMaster.Cells(wsMaster.Rows.Count, "C").End(xlUp).Row
        If lastRow < 7 Then lastRow = 7
        nextRow = lastRow + 1

        ' Check for duplicates in Column C
        newValue1 = wsRequest.Range("B9").Value
        newValue2 = wsRequest.Range("B33").Value
        isDuplicate = False

        For Each cell In wsMaster.Range("C8:C" & lastRow)
            If cell.Value = newValue2 Then
                isDuplicate = True
                Exit For
            End If
        Next cell

        If Not isDuplicate Then
            ' Map the data from request form to master sheet
            wsMaster.Cells(nextRow, "E").Value = wsRequest.Range("B9").Value
            wsMaster.Cells(nextRow, "I").Value = wsRequest.Range("B10").Value
            wsMaster.Cells(nextRow, "J").Value = wsRequest.Range("B11").Value
            wsMaster.Cells(nextRow, "K").Value = wsRequest.Range("B12").Value
            wsMaster.Cells(nextRow, "L").Value = wsRequest.Range("B13").Value
            wsMaster.Cells(nextRow, "O").Value = wsRequest.Range("B14").Value
            wsMaster.Cells(nextRow, "M").Value = wsRequest.Range("B15").Value
            wsMaster.Cells(nextRow, "N").Value = wsRequest.Range("B16").Value
            wsMaster.Cells(nextRow, "P").Value = wsRequest.Range("B17").Value
            wsMaster.Cells(nextRow, "Q").Value = wsRequest.Range("B18").Value
            wsMaster.Cells(nextRow, "R").Value = wsRequest.Range("B19").Value
            wsMaster.Cells(nextRow, "S").Value = wsRequest.Range("B21").Value
            wsMaster.Cells(nextRow, "T").Value = wsRequest.Range("B22").Value
            wsMaster.Cells(nextRow, "U").Value = wsRequest.Range("B23").Value
            wsMaster.Cells(nextRow, "V").Value = wsRequest.Range("B25").Value
            wsMaster.Cells(nextRow, "W").Value = wsRequest.Range("B26").Value
            wsMaster.Cells(nextRow, "X").Value = wsRequest.Range("B27").Value
            wsMaster.Cells(nextRow, "F").Value = wsRequest.Range("B28").Value
            wsMaster.Cells(nextRow, "G").Value = wsRequest.Range("B29").Value
            wsMaster.Cells(nextRow, "Y").Value = wsRequest.Range("B31").Value
            wsMaster.Cells(nextRow, "H").Value = wsRequest.Range("B32").Value
            wsMaster.Cells(nextRow, "C").Value = wsRequest.Range("B33").Value
            wsMaster.Cells(nextRow, "B").Value = wsRequest.Range("B34").Value

            ' Check the value in B34 and set the highlight color
            Select Case wsRequest.Range("B34").Value
                Case "Yellow"
                    highlightColor = RGB(255, 255, 0)
                Case "Orange"
                    highlightColor = RGB(255, 165, 0)
                Case "Red"
                    highlightColor = RGB(255, 0, 0)
                Case Else
                    highlightColor = -1 ' No highlight
            End Select

            ' Apply the highlight color to columns B to Y if applicable
            If highlightColor <> -1 Then
                wsMaster.Range(wsMaster.Cells(nextRow, "B"), wsMaster.Cells(nextRow, "Y")).Interior.Color = highlightColor
            End If

            ' Split the file name into base name and extension
            baseFileName = Left(fileName, InStrRev(fileName, ".") - 1)
            fileExtension = Mid(fileName, InStrRev(fileName, "."))

            ' Check if a file with the same name already exists in the inventory folder and increment the number
            fileCounter = 1
            newFileName = baseFileName & fileExtension
            Do While Dir(inventoryPath & newFileName) <> ""
                newFileName = baseFileName & " " & fileCounter & fileExtension
                fileCounter = fileCounter + 1
            Loop

            ' Debugging information to check the new file name
            MsgBox "New file name: " & newFileName

            ' Close and move the request form to the Inventory folder with the new file name
            On Error Resume Next
            wbRequest.Close SaveChanges:=False

            ' Debugging information to check the paths before moving the file
            MsgBox "Moving file from: " & folderPath & fileName & " to " & inventoryPath & newFileName

            ' Attempt to move the file
            Name folderPath & fileName As inventoryPath & newFileName

            If Err.Number <> 0 Then
                MsgBox "Error moving file: " & fileName & vbCrLf & "Error: " & Err.Description, vbExclamation
                notTransferredFiles = notTransferredFiles & fileName & vbCrLf
                Err.Clear
            End If
        Else
            ' Add to not transferred files list
            notTransferredFiles = notTransferredFiles & fileName & vbCrLf

            ' Close the request form without saving changes
            wbRequest.Close SaveChanges:=False
        End If

NextFile:
        ' Get the next file
        On Error Resume Next
        fileName = Dir
        If fileName <> "" Then
            MsgBox "Next file to process: " & fileName
        Else
            MsgBox "No more files to process."
        End If
        If Err.Number <> 0 Then
            MsgBox "Error retrieving next file: " & Err.Description, vbExclamation
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
    Loop

    ' If there are not transferred files, show a message box
    If notTransferredFiles <> "" Then
        MsgBox "The following files have not been transferred due to duplicates or errors: " & vbCrLf & notTransferredFiles, vbExclamation
    Else
        MsgBox "Data processing complete!"
    End If

    ' Close the master workbook
    wbMaster.Close SaveChanges:=True
End Sub
