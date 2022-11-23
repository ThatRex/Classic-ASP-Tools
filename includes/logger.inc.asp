<%
' Usage:
'   doLogs = True
'   doClearOnStart = False
'   Set l = (New Logger)(doLogs, doClearOnStart)
'   l.write "This is log one"
'   l.write "This is log two"
Class Logger
    Private logFile
    Public Default Function Init(doLogs, doClearOnStart)
        If doLogs Then
            logFolderName = "logs"
            scriptPathAndName = Request.ServerVariables("SCRIPT_NAME") ' Path where class is initiated
            rootPath = Server.MapPath("/")
            ' Builds array of folder names from path
            Set re = New RegExp
            re.Pattern = "([^\/]+?)(\[^\/.])?$" ' Matches file name in path
            scriptPath = re.Replace(scriptPathAndName, "")
            Set re = Nothing 
            logFilePath = logFolderName & Left(scriptPath, Len(scriptPath) - 1)
            folders = Split(logFilePath, "/")
            ' Constructs nonexistent folders
            Set objFSO = Server.createobject("Scripting.FileSystemObject")
            checkedPath = rootPath & "/"
            For Each folder In folders
                checkedPath = checkedPath & folder & "/"
                If Not objFSO.FolderExists(checkedPath) Then objFSO.CreateFolder(checkedPath)
            Next
            ' Sets log file
            logFilePathAndName = rootPath & "/" & logFolderName & scriptPathAndName & ".log"
            mode = 8 ' ForAppending 
            If doClearOnStart Then mode = 2 ' ForWriting 
            Set logFile = objFSO.OpenTextFile(logFilePathAndName, mode, True)
            logFile.WriteLine "## LOGGER START"
        End If
        Set Init = Me
    End Function
    Public Function write(textToLog)
        If Not IsEmpty(logFile) Then logFile.WriteLine Date() & " " & Time() & " | " & Trim(textToLog)
    End Function
    Private Sub Class_Terminate()
        If Not IsEmpty(logFile) Then
            logFile.WriteLine "## LOGGER END" & vbCrLf
            objFSO.Close
            Set objFSO = Nothing
        End If
    End Sub
End Class
%>