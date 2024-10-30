Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private WithEvents olInboxItems As Outlook.Items

Private Sub Application_Startup()
    Dim olNs As Outlook.NameSpace
    Set olNs = Application.GetNamespace("MAPI")
    Set olInboxItems = olNs.GetDefaultFolder(olFolderInbox).Items
    
    Debug.Print "Application Startup Triggered " & Now()
    
End Sub

Private Sub olInboxItems_ItemAdd(ByVal Item As Object)
    
    Debug.Print "New Email Found! - " & Now()
    Dim olMail As Outlook.MailItem
    
    If TypeOf Item Is Outlook.MailItem Then
        Set olMail = Item
        Debug.Print "    Email Subject: " & olMail.Subject
        ' Check if the subject contains "XPTO"
        If InStr(1, olMail.Subject, "CHANGE THIS", vbTextCompare) > 0 Then
            Debug.Print "    Processing Email"
            trigger_function olMail
        Else
            Debug.Print "    Not Processing"
        End If
    End If
End Sub




Private Sub trigger_function(olMail As Outlook.MailItem)
    Dim fso As Object
    Dim jsonDict As Object
    Dim filePath As String
    Dim jsonFileName As String
    Dim attachment As Outlook.attachment
    Dim attachCount As Integer
    Dim attachFilePath As String
    Dim emailDate As String
    Dim attachmentList As Object
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create a dictionary to store email data
    Set jsonDict = CreateObject("Scripting.Dictionary")

    ' Create a collection to store attachment file paths
    Set attachmentList = CreateObject("Scripting.Dictionary")

    ' Format the email received date to use in the file name (YYYYMMDD_HHMMSS)
    emailDate = Format(olMail.ReceivedTime, "YYYYMMDD_HHMMSS")
    
    ' Define the file path where the JSON file will be saved (same directory as this VBA file)
    filePath = "CHANGE THIS FOLDER" & "\" & emailDate
    Debug.Print "        File path: " & filePath
    
    ' Fill the dictionary with email details
    jsonDict.Add "Sender", olMail.SenderName
    jsonDict.Add "Receiver", olMail.To
    jsonDict.Add "CC", olMail.CC
    jsonDict.Add "Subject", olMail.Subject
    Dim cleanBody As String
    cleanBody = olMail.Body
    
    ' Remove trailing newlines from the body
    Do While Len(cleanBody) > 0 And (Right(cleanBody, 2) = vbCrLf Or Right(cleanBody, 1) = vbLf)
        If Right(cleanBody, 2) = vbCrLf Then
            cleanBody = Left(cleanBody, Len(cleanBody) - 2) ' Remove the last vbCrLf
        ElseIf Right(cleanBody, 1) = vbLf Then
            cleanBody = Left(cleanBody, Len(cleanBody) - 1) ' Remove the last vbLf
        End If
    Loop
    
    ' Replace all newlines with \n
    cleanBody = Replace(cleanBody, vbCrLf, "\n") ' Replace vbCrLf with \n
    cleanBody = Replace(cleanBody, vbLf, "\n")   ' Replace vbLf with \n (if needed)
    
    ' Add the cleaned body to the JSON dictionary
    jsonDict.Add "Body", cleanBody

    ' Initialize Attachments field as an empty array
    jsonDict.Add "Attachments", "[]"
    jsonDict.Add "Processed", False
    
    ' If there are attachments, populate the array
    attachCount = 0
    Dim attachmentsArray As String
    attachmentsArray = "["
    For Each attachment In olMail.Attachments
        attachCount = attachCount + 1
        attachFilePath = filePath & "_appendix_" & attachCount & "_" & attachment.FileName
        attachment.SaveAsFile attachFilePath
        attachmentsArray = attachmentsArray & """" & attachFilePath & ""","
        Debug.Print "Attachment saved: " & attachFilePath
    Next attachment
    
    If attachCount > 0 Then
        attachmentsArray = Left(attachmentsArray, Len(attachmentsArray) - 1) ' Remove the last comma
    End If
    attachmentsArray = attachmentsArray & "]"
    jsonDict("Attachments") = attachmentsArray
    
    filePath = filePath & ".json"
    
    ' Convert the dictionary to a JSON string
    Dim jsonString As String
    jsonString = ConvertDictToJson(jsonDict)
    Debug.Print "        JSON string: " & jsonString
    
    ' Write the JSON string to the file
    WriteToFile filePath, jsonString




    ' ExecutePython Script
    ExecutePythonScript filePath



    
    ' Clean up
    Set jsonDict = Nothing
    Set fso = Nothing
    Set attachmentList = Nothing
    
End Sub



Sub ExecutePythonScript(filePath As String)
    Dim pythonExePath As String
    Dim scriptPath As String
    Dim shellCommand As String
    Dim wsh As Object
    Dim exec As Object
    Dim output As String
    Dim errorOutput As String

    ' Define the paths to the Python executable and the script
    pythonExePath = "CHANGE THIS"
    scriptPath = "CHANGE THIS"
    
    ' Combine the Python executable, script path, and the JSON file path
    shellCommand = """" & pythonExePath & """ """ & scriptPath & """ """ & filePath & """"
    Debug.Print "Shell command: " & shellCommand
    
    ' Create a WScript Shell object to run the command
    Set wsh = CreateObject("WScript.Shell")
    
    ' Execute the Python script and capture the output
    Set exec = wsh.exec(shellCommand)
    Do While exec.Status = 0
        DoEvents
    Loop
    
    ' Capture standard output and error
    output = exec.StdOut.ReadAll
    errorOutput = exec.StdErr.ReadAll
    
    ' Print outputs
    Debug.Print "Python Output: " & output
    Debug.Print "Python Error: " & errorOutput
    
    ' Check for errors
    If Len(errorOutput) > 0 Then
        MsgBox "Error executing Python script: " & errorOutput
    Else
        Debug.Print "Python script executed successfully."
    End If
End Sub





' Function to convert a dictionary to a JSON string
Function ConvertDictToJson(dict As Object) As String
    Dim key As Variant
    Dim json As String
    Dim value As Variant
    json = "{"
    For Each key In dict.Keys
        If IsArray(dict(key)) Then
            ' If the value is an array (like the attachments), format it as a JSON array
            json = json & """" & key & """: ["
            For Each value In dict(key)
                json = json & """" & value & ""","
            Next value
            json = Left(json, Len(json) - 1) & "],"
        Else
            json = json & """" & key & """: """ & dict(key) & ""","
        End If
    Next key
    json = Left(json, Len(json) - 1) & "}" ' Remove the last comma and close the JSON
    ConvertDictToJson = json
End Function



' Function to write data to a file
Sub WriteToFile(filePath As String, data As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As fileNum
    Print #fileNum, data
    Close fileNum
    If Err.Number = 0 Then
        Debug.Print "        File written successfully: " & filePath
    Else
        Debug.Print "        Error writing file: " & Err.Description
    End If
End Sub





