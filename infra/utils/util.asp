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

%>