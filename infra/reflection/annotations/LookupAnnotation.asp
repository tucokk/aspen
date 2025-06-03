<%

' @author: Arthur Ribeiro
Class LookupAnnotation
    Private p_column
    Private p_alias
    Private p_columnAlias
    Private p_matchMode

    Public Property Get column()
        If p_column = "" Then 
            column = Null
            Exit Property
        End If
        column = p_column
    End Property
    Public Property Let column(value)
        p_column = value
    End Property

    Public Property Get alias()
        If p_alias = "" Then 
            alias = Null
            Exit Property
        End If
        alias = p_alias
    End Property
    Public Property Let alias(value)
        p_alias = value
    End Property

    Public Property Get matchMode()
        If p_matchMode = "" Then 
            matchMode = Null
            Exit Property
        End If
        matchMode = p_matchMode
    End Property
    Public Property Let matchMode(value)
        If IsNull(value) Or value = "" Then
            p_matchMode = "ALIKE"
        Else
            p_matchMode = value
        End If
    End Property

    Public Property Get columnAlias()
        If p_columnAlias = "" Then 
            columnAlias = Null
            Exit Property
        End If
        columnAlias = p_columnAlias
    End Property
    Public Property Let columnAlias(value)
        p_columnAlias = value
    End Property

    Public Sub Class_Initialize()
        p_column = Null
        p_alias  = Null
        p_columnAlias = Null
        p_matchMode = "ALIKE"
    End Sub
End Class

%>