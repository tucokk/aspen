<%

' @author: Arthur Ribeiro
Class FieldAnnotation
    Private p_column
    Private p_alias
    Private p_isPrimaryKey
    Private p_matchMode
    Private p_enableInsertPrimaryKey

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

    Public Property Get isPrimaryKey()
        If p_isPrimaryKey = "" Then 
            isPrimaryKey = False
            Exit Property
        End If
        isPrimaryKey = p_isPrimaryKey
    End Property
    Public Property Let isPrimaryKey(value)
        p_isPrimaryKey = value
    End Property

    Public Property Get enableInsertPrimaryKey()
        If p_enableInsertPrimaryKey = "" Then 
            enableInsertPrimaryKey = False
            Exit Property
        End If
        enableInsertPrimaryKey = p_enableInsertPrimaryKey
    End Property
    Public Property Let enableInsertPrimaryKey(value)
        p_enableInsertPrimaryKey = value
    End Property

    Public Sub Class_Initialize()
        p_enableInsertPrimaryKey = Null
        p_column = Null
        p_alias  = Null
        p_isPrimaryKey = False
        p_matchMode = "ALIKE"
    End Sub
End Class

%>