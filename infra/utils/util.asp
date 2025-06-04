<%

' @author: tucokk
Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

' @author: tucokk
Public Function IsLocalhost()
	IsLocalhost = false
	' IsLocalhost = InStr(Request.ServerVariables("SERVER_NAME"), "localhost") > 0
End Function

Public Function GetLink(controller, action)
    GetLink = Format("index.asp?controller={0}&action={1}", Array(controller, action))
End Function

%>