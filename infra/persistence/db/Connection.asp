<%

' @author: tucokk
Class Connection
    Private p_dbConnection

    Public Sub Class_Initialize()
        Set p_dbConnection = Nothing
        OpenDbConnection()
    End Sub

    Public Sub Class_Terminate()
        On Error Resume Next
            CloseDbConnection()
        On Error Goto 0
        Set p_dbConnection = Nothing
    End Sub

    '------------------------------------------------------------
    ' @function Expose
    ' @description Expõe a conexão do banco de dados publicamente.
    ' @returns {Object} Conexão ao banco de dados (conMSSQL)
    '------------------------------------------------------------
    Public Function Expose()
        Set Expose = p_dbConnection
    End Function

    '------------------------------------------------------------
    ' @function OpenDbConnection
    ' @description Abre a conexão com o banco de dados.
    ' @returns {Tipo} Conexão ao banco de dados (conMSSQL)
    '------------------------------------------------------------
    Private Function OpenDbConnection()
        If p_dbConnection Is Nothing Then
            ' Log "(ID Usuário: " & Session("id_usuario") & ") - Obtendo conexão do pool: " & Application("conMSSQL_ConnectionString") & ";Pooling=true;", "Infra.Persistence.Db.Connection"
            Set p_dbConnection = Server.CreateObject("ADODB.Connection")
            p_dbConnection.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password='xpto2039SW@';Initial Catalog=bkp_insech;Data Source=SERVIDOR_DADOS; Pooling=true;"'Application("conMSSQL_ConnectionString")
            p_dbConnection.Open
        End If
    End Function
    
    '------------------------------------------------------------
    ' @function CloseDbConnection
    ' @description Fecha a conexão com o banco de dados.
    ' @returns {void}
    '------------------------------------------------------------
    Public Function CloseDbConnection()
        If Not p_dbConnection Is Nothing Then
            ' Log "(ID Usuário: " & Session("id_usuario") & ") - Devolvendo conexão ao pool", "Infra.Persistence.Db.Connection"
            p_dbConnection.Close
            Set p_dbConnection = Nothing
        End If
    End Function

    '------------------------------------------------------------
    ' @function SetIsolation
    ' @description Define o isolamento da conexão.
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub SetIsolation(isolation)
        p_dbConnection.IsolationLevel = isolation
    End Sub
    
    '------------------------------------------------------------
    ' @function Query
    ' @description Executa um query no banco de dados e retorna os resultados.
    ' @param {String} strSQL - Query a ser executado
    ' @returns {Object} Resultados do query
    '------------------------------------------------------------
    Public Function Query(strSQL)
        Log "(ID Usuário: " & Session("id_usuario") & ") - Executando query: " & strSQL, "Infra.Persistence.Db.Connection"
        Set Query = p_dbConnection.Execute(strSQL)
    End Function

    '------------------------------------------------------------
    ' @function Execute
    ' @description Executa um query no banco de dados.
    ' @param {String} strSQL - Query a ser executado
    ' @returns {void} 
    '------------------------------------------------------------
    Public Sub Execute(strSQL)
        Log "(ID Usuário: " & Session("id_usuario") & ") - Executando query: " & strSQL, "Infra.Persistence.Db.Connection"
        p_dbConnection.Execute strSQL
    End Sub

    '------------------------------------------------------------
    ' @function StartTransaction
    ' @description Abre uma transaction no banco de dados.
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub StartTransaction()
        Log "(ID Usuário: " & Session("id_usuario") & ") - Abrindo transaction", "Infra.Persistence.Db.Connection"
        p_dbConnection.BeginTrans()
    End Sub

    '------------------------------------------------------------
    ' @function CommitTransaction
    ' @description Commita uma transaction no banco de dados.
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub CommitTransaction()
        Log "(ID Usuário: " & Session("id_usuario") & ") - Commitando transaction", "Infra.Persistence.Db.Connection"
        p_dbConnection.CommitTrans()
    End Sub

    '------------------------------------------------------------
    ' @function RollbackTransaction
    ' @description Rollback (desfaz) uma transaction no banco de dados.
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub RollbackTransaction()
        Log "(ID Usuário: " & Session("id_usuario") & ") - Rollback transaction", "Infra.Persistence.Db.Connection"
        p_dbConnection.RollbackTrans()
    End Sub
End Class

%>