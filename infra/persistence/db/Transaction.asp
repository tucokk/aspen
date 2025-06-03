<%

' @author: tucokk
Class Transaction 
    Private p_DbConnection

    Public Sub Class_Initialize()
    End Sub

    Public Sub Class_Terminate()
    End Sub

    '------------------------------------------------------------
    ' @function StartTransaction
    ' @description Inicia a transaчуo no banco de dados.
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub StartTransaction()
        Set p_DbConnection = New Connection
        p_DbConnection.SetIsolation(1048576)
        p_DbConnection.StartTransaction()
    End Sub

    '------------------------------------------------------------
    ' @function CloseTransaction
    ' @description Fecha (finaliza) a transaчуo e a conexУЃo com o banco de dados.
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub CloseTransaction()
        p_DbConnection.CloseDbConnection()
    End Sub

    '------------------------------------------------------------
    ' @function CommitTransaction
    ' @description Commita a transaчуo no banco de dados.
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub CommitTransaction()
        p_DbConnection.CommitTransaction()
    End Sub

    '------------------------------------------------------------
    ' @function RollbackTransaction
    ' @description Rollback (desfaz) a transaчуo no banco de dados.
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub RollbackTransaction()
        p_DbConnection.RollbackTransaction()
    End Sub

    '------------------------------------------------------------
    ' @function Execute
    ' @description Execute um query no banco de dados.
    ' @param {String} strSQL - Query a ser executado no banco de dados
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub Execute(strSQL)
        p_DbConnection.Execute(strSQL)
    End Sub

    '------------------------------------------------------------
    ' @function Execute
    ' @description Execute um query no banco de dados.
    ' @param {String} strSQL - Query a ser executado no banco de dados
    ' @returns {Object} Resultados do query executado
    '------------------------------------------------------------
    Public Function Query(strSQL)
        Set Query = p_DbConnection.Query(strSQL)
    End Function

    '------------------------------------------------------------
    ' @function Save
    ' @description Insere ou atualiza a entidade no banco de dados.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub Save(entity)
        If EntityExists(entity) Then
            Update(entity)
        Else
            Insert(entity)
        End If
    End Sub

    '------------------------------------------------------------
    ' @function Insert
    ' @description Insere a entidade no banco de dados.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub Insert(entity)
        RunEntityEvent entity, "BeforeInsertEvent"
        Set factory = New QueryFactory
        strSQL = factory.GetInsert(entity)
        Execute(strSQL)
        Set factory = Nothing
        RunEntityEvent entity, "AfterInsertEvent"
    End Sub

    '------------------------------------------------------------
    ' @function Update
    ' @description Atualiza a entidade no banco de dados.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub Update(entity)
        RunEntityEvent entity, "BeforeUpdateEvent"
        Set factory = New QueryFactory
        strSQL = factory.GetUpdate(entity)
        Execute(strSQL)
        Set factory = Nothing
        RunEntityEvent entity, "AfterUpdateEvent"
    End Sub

    '------------------------------------------------------------
    ' @function Delete
    ' @description Remove a entidade no banco de dados.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {void}
    '------------------------------------------------------------
    Public Sub Delete(entity)
        RunEntityEvent entity, "BeforeDeleteEvent"
        Set factory = New QueryFactory
        strSQL = factory.GetDelete(entity)
        Execute(strSQL)
        Set factory = Nothing
        RunEntityEvent entity, "AfterDeleteEvent"
    End Sub

    '------------------------------------------------------------
    ' @function RunEntityEvent
    ' @description Executa o evento indicado na entidade fornecida.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @param {string} entityEvent - Evento que deve ser executado
    ' @returns {void}
    '------------------------------------------------------------
    Private Function RunEntityEvent(entity, entityEvent)
        Set mirror = New Reflection
        mirror.Reflect(entity.ENTITY_PATH)

        For Each item In mirror.Methods
            If InStr(Trim(item.name), entityEvent) Then
                Eval("entity." & Trim(item.name))
            End If
        Next
    End Function

    '------------------------------------------------------------
    ' @function EntityExists
    ' @description Verifica se a entidade jс existe no banco de dados com base na chave primсria.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {Boolean} True se existir, False se nуo existir
    '------------------------------------------------------------
    Private Function EntityExists(entity)
        Set mirror = New Reflection
        mirror.Reflect(entity.ENTITY_PATH)

        primaryKey = Null
        enableInsertPrimaryKey = False

        For Each field In mirror.Fields
            If field.isPrimaryKey Then
                primaryKey = field.alias
                If field.enableInsertPrimaryKey Then
                    enableInsertPrimaryKey = True
                End If
            End If
        Next

        primaryKeyValue = Eval("entity." & primaryKey)

        If IsNull(primaryKeyValue) Then
            EntityExists = False
            Exit Function
        End If

        Set validationEntity = manager.FindByPrimaryKey(entity)
        validationEntityPrimaryKeyValue = Eval("validationEntity." & primaryKey)

        EntityExists = Not (IsNull(validationEntityPrimaryKeyValue) And enableInsertPrimaryKey)

        Set manager = Nothing
        Set validationEntity = Nothing
        Set mirror = Nothing
    End Function
End Class

%>