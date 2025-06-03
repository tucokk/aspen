<%

' @author: tucokk
Public Function FileExists(path)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
    Set fso = Nothing
End Function

Public Sub CreateFolder(path)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    parts = Split(path, "/")
    currentPath = Server.MapPath(".")
    For i = 0 To UBound(parts)
        currentPath = currentPath & "\" & parts(i)
        If Not fso.FolderExists(currentPath) Then
            fso.CreateFolder(currentPath)
        End If
    Next
    Set fso = Nothing
End Sub

Public Sub CreateFile(path)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    path = Server.MapPath(path)
    If Not FileExists(path) Then
        Set file = fso.CreateTextFile(path, True)
        file.Close
        Set file = Nothing
    End If
    Set fso = Nothing
End Sub

Public Sub WriteToFile(path, content)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    path = Server.MapPath(path)
    
    Set file = fso.OpenTextFile(path, 8, True) ' 2 = ForWriting, True = create if not exists
    file.Write content & vbCrLf 
    file.Close

    Set file = Nothing
    Set fso = Nothing
End Sub

Public Sub ClearFileContent(path)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    path = Server.MapPath(path)

    Set file = fso.OpenTextFile(path, 2, True)
    file.Write "" 
    file.Close

    Set file = Nothing
    Set fso = Nothing
End Sub

Public Function ReadValueFromFileByKey(path, key)
    ReadValueFromFileByKey = Null

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    path = Server.MapPath(path)

    If fso.FileExists(path) Then
        Set file = fso.OpenTextFile(path, 1) ' 1 = ForReading

        Do While Not file.AtEndOfStream
            line = Trim(file.ReadLine)
            If InStr(line, "=") > 0 Then
                parts = Split(line, "=")
                If UBound(parts) = 1 Then
                    If Trim(parts(0)) = key Then
                        ReadValueFromFileByKey = Trim(parts(1))
                        Exit Do
                    End If
                End If
            End If
        Loop

        file.Close
        Set file = Nothing
    End If

    Set fso = Nothing
End Function

%>