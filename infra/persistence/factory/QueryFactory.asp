
<!--#include virtual="/infra/persistence/db/Connection.asp"-->
<!--#include virtual="/infra/persistence/builder/SqlBuilder.asp"-->

<%

' @author: tucokk
Class QueryFactory
    Private p_Object
    Private p_SqlBuilder
    Private p_DbConnection

    Public Sub Class_Initialize()
        p_Object = Null
        Set p_SqlBuilder = New SqlBuilder
        Set p_DbConnection = New Connection
    End Sub

    Public Sub Class_Terminate()
        Set p_SqlBuilder = Nothing
        Set p_DbConnection = Nothing
    End Sub

    '------------------------------------------------------------
    ' @function Query
    ' @description Executa um query e retorna o resultado (caso houver).
    ' @param {String} strSQL - Query a ser executado no banco de dados
    ' @param {Object} Entidade que representa a tabela no banco de dados
    ' @returns {Object[]} Array de entidades populadas. Caso o query não encontre resultados, retorna um array vazio
    '------------------------------------------------------------
    Public Function Query(strSQL, entity)
        p_SqlBuilder.init(entity.ENTITY_PATH)
        Set p_Object = p_DbConnection.Query(strSQL)
        Set Query = p_SqlBuilder.RecordsetToObject(p_Object)
    End Function

    '------------------------------------------------------------
    ' @function AbstractQuery
    ' @description Executa um query e retorna o resultado (caso houver).
    ' @param {String} strSQL - Query a ser executado no banco de dados
    ' @returns {Object[]} Array de dicionários. Caso o query não encontre resultados, retorna um array vazio.
    '------------------------------------------------------------
    Public Function AbstractQuery(strSQL)
        Set p_Object = p_DbConnection.Query(strSQL)
        Set AbstractQuery = p_SqlBuilder.RecordsetToDictionary(p_Object)
    End Function

    '------------------------------------------------------------
    ' @function GetDelete
    ' @description Monta o query de Delete com base na entidade recebida como parâmetro.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Query a ser executada no banco de dados
    '------------------------------------------------------------
    Public Function GetDelete(entity)
        p_SqlBuilder.init(entity.ENTITY_PATH)

        strFrom  = p_SqlBuilder.GetDeleteFrom(entity)
        strWhere = p_SqlBuilder.GetDeleteWhere(entity)

        GetDelete = " DELETE FROM " & strFrom & " WHERE 1=1 " & strWhere & ";"
    End Function

    '------------------------------------------------------------
    ' @function GetUpdate
    ' @description Monta o query de Update com base na entidade recebida como parâmetro.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Query a ser executada no banco de dados
    '------------------------------------------------------------
    Public Function GetUpdate(entity)
        p_SqlBuilder.init(entity.ENTITY_PATH)

        strInto   = p_SqlBuilder.GetUpdateInto(entity)
        strValues = p_SqlBuilder.GetUpdateValues(entity)
        strWhere  = p_SqlBuilder.GetUpdateWhere(entity)

        GetUpdate = " UPDATE " & strInto & " SET " & strValues & " WHERE 1=1 " & strWhere & ";"
    End Function

    '------------------------------------------------------------
    ' @function GetInsert
    ' @description Monta o query de Insert com base na entidade recebida como parâmetro.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Query a ser executada no banco de dados
    '------------------------------------------------------------
    Public Function GetInsert(entity)
        p_SqlBuilder.init(entity.ENTITY_PATH)

        strFields = p_SqlBuilder.GetInsertFields(entity)
        strInto   = p_SqlBuilder.GetInsertInto(entity)
        strValues = p_SqlBuilder.GetInsertValues(entity)

        GetInsert = " INSERT INTO " & strInto & " ( " & strFields & " ) VALUES (" &  strValues & ");"
    End Function

    '------------------------------------------------------------
    ' @function GetSelectPrimaryKey
    ' @description Monta o query de Select com base na primary key da entidade recebida como parâmetro.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Query a ser executada no banco de dados
    '------------------------------------------------------------
    Public Function GetSelectPrimaryKey(entity)
        p_SqlBuilder.init(entity.ENTITY_PATH)

        strFields = p_SqlBuilder.GetSelectFields(entity)
        strFrom   = p_SqlBuilder.GetSelectFrom(entity)
        strWhere  = p_SqlBuilder.GetSelectPrimaryKeyWhere(entity)

        GetSelectPrimaryKey = " SELECT " & strFields & " FROM " & strFrom & " WHERE 1=1 " & strWhere & ";"
    End Function

    '------------------------------------------------------------
    ' @function GetSelect
    ' @description Monta o query de Select com base na entidade recebida como parâmetro.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {String} Query a ser executada no banco de dados
    '------------------------------------------------------------
    Public Function GetSelect(entity, where, orderBy, limit)
        p_SqlBuilder.init(entity.ENTITY_PATH)
        
        strFields = p_SqlBuilder.GetSelectFields(entity)
        strFrom   = p_SqlBuilder.GetSelectFrom(entity)
        strWhere  = p_SqlBuilder.GetSelectWhere(entity)

        strSQL = " SELECT "

        If Not IsNull(limit) And limit <> "" Then
            strSQL = strSQL & " " & limit
        End If

        strSQL = strSQL & strFields 
        strSQl = strSQL & " FROM " & strFrom
        strSQL = strSQL & " WHERE 1=1 " & strWhere 

        If Not IsNull(where) And where <> "" Then
            strSQL = strSQL & " " & where
        End If
        
        If Not IsNull(orderBy) And orderBy <> "" Then
            strSQL = strSQL & " " & orderBy
        End If

        strSQl = strSQL & ";"

        GetSelect = strSQL  
    End Function

End Class

%>