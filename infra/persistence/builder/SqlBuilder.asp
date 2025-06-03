
<!--#include virtual="/infra/reflection/Reflection.asp"-->

<%

' @author: tucokk
Class SqlBuilder

    Private p_Mirror
    Private p_strSQL

    Public Sub Class_Initialize()
        Set p_Mirror = New Reflection
        p_strSQl = Null
    End Sub

    Public Sub Class_Terminate()
        Set p_Mirror = Nothing
    End Sub

    Public Sub Init(path)   
        p_Mirror.Reflect(path)
    End Sub

    '------------------------------------------------------------
    ' @function RecordsetToDictionary
    ' @description Converte um recordset para um array de dicion�rios populados.
    ' @param {Object} rs - Recordset a ser convertido para o array de dicion�rios
    ' @returns {Object[]} Array de dicion�rios. Caso o query n�o encontre resultados, retorna um array vazio.
    '------------------------------------------------------------
    Public Function RecordsetToDictionary(rs)
        Set list = Server.CreateObject("System.Collections.ArrayList")

        On Error Resume Next
            If Not rs.EOF Then
                Do Until rs.EOF
                    Set obj = Server.CreateObject("Scripting.Dictionary")
                    For Each field In rs.Fields
                        obj(field.Name) = field
                    Next

                    list.Add obj
                    Set obj = Nothing
                    rs.MoveNext
                Loop
            End If
        
        Set RecordsetToDictionary = list
    End Function

    '------------------------------------------------------------
    ' @function RecordsetToObject
    ' @description Converte um recordset para um array de entidades que ele representa.
    ' @param {Object} rs - Recordset a ser convertido para o array de entidades.
    ' @returns {Object[]} Array de entidades populadas. Caso o recordset esteja vazio, retorna um array vazio
    '------------------------------------------------------------
    Public Function RecordsetToObject(rs)
        Set list = Server.CreateObject("System.Collections.ArrayList")

        On Error Resume Next
            If Not rs.EOF Then
                Do Until rs.EOF
                    Set obj = CreateObjectFromClassName(p_Mirror.ClassName)

                    If obj Is Nothing Then
                        Throw "N�o foi poss�vel instanciar a entidade " & p_Mirror.ClassName & ". Verifique se ela encontra-se registrada em ""persistence/register.asp""."
                    End If

                    For Each field In p_Mirror.Fields
                       obj.SetProperty field.alias, rs.Fields.Item(IIf(IsNull(field.alias), field.column, field.alias))
                    Next

                    For Each lookup In p_Mirror.Lookups
                        obj.SetProperty lookup.alias, rs.Fields.Item(IIf(IsNull(lookup.alias), lookup.column, lookup.alias))
                    Next

                    list.Add obj
                    Set obj = Nothing
                    rs.MoveNext
                Loop
            Else
                Set obj = CreateObjectFromClassName(p_Mirror.ClassName)
                list.Add obj
            End If
        
        If Err.Number = 438 Then
            Throw "M�todo SetProperty n�o implementado na entidade."
        End If

        Set RecordsetToObject = list
    End Function

    '------------------------------------------------------------
    ' @function GetDeleteWhere
    ' @description Monta o campo where do query delete.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do where do query delete
    '------------------------------------------------------------
    Public Function GetDeleteWhere(entity)
        strWhere = ""
        For Each field In p_Mirror.Fields
            If field.isPrimaryKey Then
                value = Eval("entity." & field.alias)
                If Not IsNull(value) And Not value = "" Then
                    strWhere = " AND "
                    If IsNumeric(value) Then
                        value = CLng(value)
                    End If
                    Select Case VarType(value)
                        Case 2, 3, 4, 5, 6 ' Integer, Long, Single, Double, Currency
                            strWhere = strWhere & field.column & " = " & sanatizeParameter(value)
                        Case 7 ' Date
                            strWhere = strWhere & "CONVERT(DATETIME, CONVERT(CHAR(10), " & field.column & ", 103), 103) = CONVERT(DATETIME, CONVERT(CHAR(10), '" & sanatizeParameter(value) & "', 103), 103)"
                        Case 8 ' String
                            value = Replace(value, "'", "''")
                            If field.matchMode = "ALIKE" Then
                                strWhere = strWhere & "TRIM(UPPER(" & field.column & "))" & " LIKE 'TRIM(UPPER(%" & sanatizeParameter(value) & "%))'"
                            ElseIf field.matchMode = "EXACT" Then
                                strWhere = strWhere & "TRIM(UPPER(" & field.column & ")) = " & "'TRIM(UPPER(" & sanatizeParameter(value) & "))'"
                            End If
                    End Select
                End If
            End If
        Next

        If strWhere = "" Then
            Throw "N�o foi poss�vel localizar o valor da primary key da entidade."
        End If

        GetDeleteWhere = strWhere
    End Function

    '------------------------------------------------------------
    ' @function GetDeleteFrom
    ' @description Monta o campo from do query delete.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do from do query delete
    '------------------------------------------------------------
    Public Function GetDeleteFrom(entity)
        GetDeleteFrom = p_Mirror.Table.value
    End Function

    '------------------------------------------------------------
    ' @function GetUpdateWhere
    ' @description Monta o campo where do query update.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do where do query update
    '------------------------------------------------------------
    Public Function GetUpdateWhere(entity)
        strWhere = ""
        For Each field In p_Mirror.Fields
            If field.isPrimaryKey Then
                value = Eval("entity." & field.alias)
                If Not IsNull(value) And Not value = "" Then
                    strWhere = " AND "
                    If IsNumeric(value) Then
                        value = CLng(value)
                    End If
                    Select Case VarType(value)
                        Case 2, 3, 4, 5, 6, 14' Integer, Long, Single, Double, Currency, Decimal
                            strWhere = strWhere & field.column & " = " & sanatizeParameter(value)
                        Case 7 ' Date
                            strWhere = strWhere & "CONVERT(DATETIME, CONVERT(CHAR(10), " & field.column & ", 103), 103) = CONVERT(DATETIME, CONVERT(CHAR(10), '" & sanatizeParameter(value) & "', 103), 103)"
                        Case 8 ' String
                            value = Replace(value, "'", "''")
                            value = sanatizeParameter(value)
                            If field.matchMode = "ALIKE" Then
                                strWhere = strWhere & "TRIM(UPPER(" & field.column & "))" & " LIKE TRIM(UPPER('%" & ignoraAcentos(value) & "%'))"
                            ElseIf field.matchMode = "EXACT" Then
                                strWhere = strWhere & "TRIM(UPPER(" & field.column & "))" & " = TRIM(UPPER('" & value & "'))"
                            End If
                    End Select
                End If
            End If
        Next

        If strWhere = "" Then
            Throw "N�o foi poss�vel localizar o valor da primary key da entidade."
        End If

        GetUpdateWhere = strWhere
    End Function

    '------------------------------------------------------------
    ' @function GetUpdateValues
    ' @description Monta o campo set do query update.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do set do query update
    '------------------------------------------------------------
    Public Function GetUpdateValues(entity)
        strValues = ""
        For Each field In p_Mirror.Fields
            value = Eval("entity." & field.alias)
            If Not field.isPrimaryKey And Not IsNull(value) And Not value = "" Then
                strValues = strValues & field.column & " = "
                Select Case VarType(value)
                    Case 2, 3, 4, 5, 6 ' Integer, Long, Single, Double, Currency
                        strValues = strValues & sanatizeParameter(value) & ", "
                    Case 7 ' Date
                        strValues = strValues & "CONVERT(DATETIME, CONVERT(CHAR(10), '" & sanatizeParameter(value) & "', 103), 103), "
                    Case 8 ' String
                        value = Replace(value, "'", "''")
                        strValues = strValues & "'" & sanatizeParameter(value) & "', "
                    Case Else
                End Select
            End If
        Next

        strValues = Left(strValues, Len(strValues) - 2)
        GetUpdateValues = strValues
    End Function

    '------------------------------------------------------------
    ' @function GetUpdateInto
    ' @description Monta o campo into do query update.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do into do query update
    '------------------------------------------------------------
    Public Function GetUpdateInto(entiy)
        GetUpdateInto = p_Mirror.Table.value
    End Function

    '------------------------------------------------------------
    ' @function GetInsertFields
    ' @description Monta o campo fields do query insert.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do fields do query insert
    '------------------------------------------------------------
    Public Function GetInsertFields(entity)
        strFields = ""
        For Each field In p_Mirror.Fields
            value = Eval("entity." & field.alias)
            If field.isPrimaryKey And field.enableInsertPrimaryKey Then
                If Not IsNull(value) Then
                    strFields = strFields & " " & field.column & ","
                End If
            Else
                If Not field.isPrimaryKey And Not IsNull(value) Then
                    strFields = strFields & " " & field.column & ","
                End If
            End If
        Next

        strFields = Left(strFields, Len(strFields) - 1)
        GetInsertFields = strFields
    End Function

    '------------------------------------------------------------
    ' @function GetInsertInto
    ' @description Monta o campo into do query insert.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do into do query insert
    '------------------------------------------------------------
    Public Function GetInsertInto(entiy)
        GetInsertInto = p_Mirror.Table.value
    End Function

    '------------------------------------------------------------
    ' @function GetInsertValues
    ' @description Monta o campo values do query insert.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do values do query insert
    '------------------------------------------------------------
    Public Function GetInsertValues(entity)
        strInsert = ""
        For Each field In p_Mirror.Fields
            value = Eval("entity." & field.alias)
            If field.isPrimaryKey And field.enableInsertPrimaryKey Then
                If Not IsNull(value) And Not value = "" Then      
                    Select Case VarType(value)
                        Case 2, 3, 4, 5, 6 ' Integer, Long, Single, Double, Currency
                            strInsert = strInsert & sanatizeParameter(value) & ", "
                        Case 7 ' Date
                            strInsert = strInsert & "CONVERT(DATETIME, CONVERT(CHAR(10), '" & sanatizeParameter(value) & "', 103), 103), "
                        Case 8 ' String
                            value = Replace(value, "'", "''")
                            strInsert = strInsert & "'" & sanatizeParameter(value) & "', "
                        Case Else
                    End Select
                End If
            Else
                If Not field.isPrimaryKey Then
                    If Not IsNull(value) And Not value = "" Then      
                        Select Case VarType(value)
                            Case 2, 3, 4, 5, 6 ' Integer, Long, Single, Double, Currency
                                strInsert = strInsert & sanatizeParameter(value) & ", "
                            Case 7 ' Date
                                strInsert = strInsert & "CONVERT(DATETIME, CONVERT(CHAR(10), '" & sanatizeParameter(value) & "', 103), 103), "
                            Case 8 ' String
                                value = Replace(value, "'", "''")
                                strInsert = strInsert & "'" & sanatizeParameter(value) & "', "
                            Case Else
                        End Select
                    End If
                End If
            End If
        Next

        strInsert = Left(strInsert, Len(strInsert) - 2)
        GetInsertValues = strInsert
    End Function

    '------------------------------------------------------------
    ' @function GetSelectFields
    ' @description Monta o campo fields do query select.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do fields do query select
    '------------------------------------------------------------
    Public Function GetSelectFields(entity)
        strFields = ""

        For Each field In p_Mirror.Fields
            strFields = strFields & " " & p_Mirror.Table.alias & "." & field.column
            
            If Not IsNull(field.alias) Then
                strFields = strFields & " AS " & field.alias & ","
            Else
                strFields = strFields & " AS " & field.column & ","
            End If
        Next

        For Each lookup In p_Mirror.Lookups
            strFields = strFields & " " & lookup.columnAlias & "." & lookup.column

            If Not IsNull(lookup.alias) Then
                strFields = strFields & " AS " & lookup.alias & ","
            Else
                strFields = strFields & " AS " & lookup.column & ","
            End If
        Next

        strFields = Left(strFields, Len(strFields) - 1)
        GetSelectFields = strFields
    End Function

    '------------------------------------------------------------
    ' @function GetSelectFrom
    ' @description Monta o campo from do query select.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do from do query select
    '------------------------------------------------------------
    Public Function GetSelectFrom(entity)
        strFrom = p_Mirror.Table.value & " " &  p_Mirror.Table.alias
        For Each tableJoin In p_Mirror.Joins
            strFrom = strFrom & " " 
            strFrom = strFrom & tableJoin.joinType & " JOIN " & tableJoin.value & " " & tableJoin.alias 
            strFrom = strFrom & " ON (" & tableJoin.onConditition & ") "
        Next
        GetSelectFrom = strFrom
    End Function

    '------------------------------------------------------------
    ' @function GetSelectPrimaryKeyWhere
    ' @description Monta o campo where do query select de primary key.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do where do query select de primary key
    '------------------------------------------------------------
    Public Function GetSelectPrimaryKeyWhere(entity)
        strWhere = ""
        For Each field In p_Mirror.Fields
            If field.isPrimaryKey Then
                strWhere = " AND "
                value = Eval("entity." & field.alias)
                If IsNumeric(value) Then
                    value = CLng(value)
                End If
                Select Case VarType(value)
                    Case 2, 3, 4, 5, 6 ' Integer, Long, Single, Double, Currency
                        strWhere = strWhere & p_Mirror.Table.alias & "." & field.column & " = " & sanatizeParameter(value)
                    Case 7 ' Date
                        strWhere = strWhere & "CONVERT(DATETIME, CONVERT(CHAR(10), " & p_Mirror.Table.alias & "." & field.column & ", 103), 103) = CONVERT(DATETIME, CONVERT(CHAR(10), '" & sanatizeParameter(value) & "', 103), 103)"
                    Case 8 ' String
                        value = Replace(value, "'", "''")
                        value = sanatizeParameter(value)
                        If field.matchMode = "ALIKE" Then
                            strWhere = strWhere & "TRIM(UPPER(" & p_Mirror.Table.alias & "." & field.column & "))" & " LIKE TRIM(UPPER('%" & ignoraAcentos(value) & "%'))"
                        ElseIf field.matchMode = "EXACT" Then
                            strWhere = strWhere & "TRIM(UPPER(" & p_Mirror.Table.alias & "." & field.column & "))" & " = TRIM(UPPER('" & value & "'))"
                        End If
                End Select
                GetSelectPrimaryKeyWhere = strWhere
                Exit Function
            End If
        Next
    End Function

    '------------------------------------------------------------
    ' @function GetSelectWhere
    ' @description Monta o campo where do query select.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Campos do where do query select
    '------------------------------------------------------------
    Public Function GetSelectWhere(entity)
        strWhere = ""
        For Each field In p_Mirror.Fields
            value = Eval("entity." & field.alias)
            If Not IsNull(value) And Not value = "" Then      
                strWhere = strWhere & " AND "
                If IsNumeric(value) Then
                    value = CLng(value)
                End If
                Select Case VarType(value)
                    Case 2, 3, 4, 5, 6 ' Integer, Long, Single, Double, Currency
                        strWhere = strWhere & p_Mirror.Table.alias & "." & field.column & " = " & sanatizeParameter(value)
                    Case 7 ' Date
                        strWhere = strWhere & "CONVERT(DATETIME, CONVERT(CHAR(10), " & p_Mirror.Table.alias & "." & field.column & ", 103), 103) = CONVERT(DATETIME, CONVERT(CHAR(10), '" & sanatizeParameter(value) & "', 103), 103)"
                    Case 8 ' String
                        value = Replace(value, "'", "''")
                        value = sanatizeParameter(value)
                        If field.matchMode = "ALIKE" Then
                            strWhere = strWhere & "TRIM(UPPER(" & p_Mirror.Table.alias & "." & field.column & ")) LIKE TRIM(UPPER('%" & ignoraAcentos(value) & "%'))"
                        ElseIf field.matchMode = "EXACT" Then
                            strWhere = strWhere & "TRIM(UPPER(" & p_Mirror.Table.alias & "." & field.column & ")) = TRIM(UPPER('" & value & "'))"
                        End If
                    Case 0, 1 ' Empty, Null
                        strWhere = strWhere & p_Mirror.Table.alias & "." & field.column & " IS NULL"
                    Case Else
                End Select
            End If
        Next

        For Each lookup In p_Mirror.Lookups
            value = Eval("entity." & lookup.alias)
            If Not IsNull(value) And Not value = "" Then
                strWhere = strWhere & " AND "
                If IsNumeric(value) Then
                    value = CLng(value)
                End If
                Select Case VarType(value)
                    Case 2, 3, 4, 5, 6 ' Integer, Long, Single, Double, Currency
                        strWhere = strWhere & lookup.columnAlias & "." & lookup.column & " = " & sanatizeParameter(value)
                    Case 7 ' Date
                        strWhere = strWhere & "CONVERT(DATETIME, CONVERT(CHAR(10), " & lookup.columnAlias & "." & lookup.column & ", 103), 103) = CONVERT(DATETIME, CONVERT(CHAR(10), '" & sanatizeParameter(value) & "', 103), 103)"
                    Case 8 ' String
                        value = Replace(value, "'", "''")
                        value = sanatizeParameter(value)
                        If lookup.matchMode = "ALIKE" Then
                            strWhere = strWhere & "TRIM(UPPER(" & lookup.columnAlias & "." & lookup.column & ")) LIKE TRIM(UPPER('%" & ignoraAcentos(value) & "%'))"
                        ElseIf lookup.matchMode = "EXACT" Then
                            strWhere = strWhere & "TRIM(UPPER(" & lookup.columnAlias & "." & lookup.column & ")) = TRIM(UPPER('" & value & "'))"
                        End If
                    Case 0, 1 ' Empty, Null
                        strWhere = strWhere & lookup.columnAlias & "." & lookup.column & " IS NULL"
                    Case Else
                End Select
            End If
        Next

        GetSelectWhere = strWhere
    End Function
End Class

%>