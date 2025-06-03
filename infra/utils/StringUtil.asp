<%

' @author: tucokk
Public Function Title(value)
    If Len(value) = 0 Then
        Title = ""
    Else
        Title = UCase(Left(value, 1)) & LCase(Mid(value, 2))
    End If
End Function

Public Function Format(byVal template, byVal replacements)
    If Not IsArray(replacements) Then
        replacements = Array(replacements)
    End If

    For i = 0 To UBound(replacements)
        If Not IsNull(replacements(i)) Then
            template = Replace(template, "{" & CStr(i) & "}", replacements(i))
        End If
    Next

    Format = template
End Function

Public Function AddTabIndent(value, quantity)
    For i = 0 To quantity
        value = "   " & value    
    Next
    AddTabIndent = value
End Function

%>