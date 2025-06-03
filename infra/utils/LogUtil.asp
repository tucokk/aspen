<%

' @author: tucokk
Public Sub Debug(message)
    Response.Write message
    Response.End
End Sub

Public Sub Log(message, className)
    Const LOG_FOLDER = "/logs"

    Dim fso, folderPath, filePath, logDate, writer
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    folderPath = Server.MapPath(LOG_FOLDER)
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder(folderPath)
    End If

    logDate = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2)
    filePath = folderPath & "\log-" & logDate & ".log"

    Dim logMessage
    logMessage = Now() & " - [" & className & "] - " & message

    Set writer = fso.OpenTextFile(filePath, 8, True)
    writer.WriteLine logMessage
    writer.Close

    Set writer = Nothing
    Set fso = Nothing
End Sub 

%>