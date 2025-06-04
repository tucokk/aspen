<!--#include virtual="/infra/reflection/annotations/index.asp"-->

<%

' @author: tucokk
Class Reflection  
    Public Attributes
    Public Methods
    Public Fields
    Public Lookups
    Public Joins
    Public Injects

    Public Service
    Public ClassName
    Public Table

    Private p_Regex
    Private p_oJSON

    Public Sub Class_Initialize()
        Set Attributes = Server.CreateObject("System.Collections.ArrayList")
        Set Fields     = Server.CreateObject("System.Collections.ArrayList")
        Set Lookups    = Server.CreateObject("System.Collections.ArrayList")
        Set Joins      = Server.CreateObject("System.Collections.ArrayList")
        Set Methods    = Server.CreateObject("System.Collections.ArrayList")
        Set Injects    = Server.CreateObject("System.Collections.ArrayList")

        Set p_Regex    = New RegExp
        p_Regex.IgnoreCase = True
        p_Regex.Global     = True
        p_Regex.Pattern    = "(\w+)\s*=\s*""([^""]*)"""

        Set p_oJSON    = New aspJSON

        Table          = Null
        ClassName      = Null
        Service        = Null
    End Sub

    Public Sub Class_Terminate()
        Set Attributes = Nothing
        Set Methods    = Nothing
        Set Fields     = Nothing
        Set Lookups    = Nothing
        Set Joins      = Nothing
        Set Table      = Nothing
        Set p_Regex    = Nothing
        Set p_oJSON    = Nothing
        Set Injects    = Nothing
        Set Service    = Nothing
    End Sub

    '------------------------------------------------------------
    ' @function Reflect
    ' @description Reflete a entidade e extrai todas as características da classe.
    ' @param {String} path - Caminho da entidade que representa a tabela no banco de dados
    ' @returns {void} 
    '------------------------------------------------------------
    Public Function Reflect(path)   
        Attributes.Clear
        Fields.Clear
        Lookups.Clear
        Joins.Clear
        Methods.Clear
        
        If Not IsLocalhost() Then
            isCached = CheckIfCached(path)
            If isCached Then
                LoadFromCache(path)
                Exit Function
            End If
        End If

        Dim fso, file

        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        Set file = fso.OpenTextFile(Server.MapPath(path), 1) ' 1 = For Reading

        Do Until file.AtEndOfStream
            line = Trim(file.ReadLine)

            Set tempTableValue = ManageTable(line)
            If Not tempTableValue Is Nothing Then
                Set Table = tempTableValue
            End If

            Set tempServiceValue = ManageService(line)
            If Not tempServiceValue Is Nothing Then
                Set Service = tempServiceValue
            End If

            Set tempJoinsValue = ManageJoin(line)
            If Not tempJoinsValue Is Nothing Then
                Joins.Add tempJoinsValue
            End If

            Set tempLookupValue = ManageLookup(line)
            If Not tempLookupValue Is Nothing Then
                Lookups.Add tempLookupValue
            End If

            Set tempFieldValue = ManageField(line)
            If Not tempFieldValue Is Nothing Then
                Fields.Add tempFieldValue
            End If

            Set tempInjectValue = ManageInject(line)
            If Not tempInjectValue Is Nothing Then
                Injects.Add tempInjectValue
            End If

            Set tempMethodValue = ManageMethod(line)
            If Not tempMethodValue Is Nothing Then
                Methods.Add tempMethodValue
            End If

            tempAttributeValue = ManageAttribute(line)
            If Not IsNull(tempAttributeValue) Then
                Attributes.Add tempAttributeValue
            End If

            tempClassNameValue = ManageClassName(line)
            If Not IsNull(tempClassNameValue) Then
                ClassName = tempClassNameValue
            End If 
        Loop

        file.Close()

        If Not IsLocalhost() Then
            If Not isCached Then
                CacheMe(path)
            End If
        End If
    End Function

    '------------------------------------------------------------
    ' @function CacheMe
    ' @description Cacheia a reflexão atual.
    ' @param {String} path - Caminho da entidade que representa a tabela no banco de dados
    ' @returns {void}
    '------------------------------------------------------------
    Private Function CacheMe(path)
        cacheKey = GetCacheKey(path)
        
        json = "{"
        
        json = json & SerializeClassName() & ","
        json = json & SerializeAttributesArray() & ","
        json = json & SerializeMethodsArray() & ","
        json = json & SerializeFieldsArray() & ","
        json = json & SerializeLookupsArray() & ","
        json = json & SerializeTable() & ","
        json = json & SerializeService() & ","
        json = json & SerializeJoinsArray() & ","
        json = json & SerializeInjectsArray()

        json = json & "}"

        Log "(ID Usuï¿½rio: " & Session("id_usuario") & ") - Cacheando reflection: " & path & " - " & cacheKey, "Persistence.Core.Db.Connection"

        Application.Lock
        Application(cacheKey) = json
        Application.Unlock
    End Function

    '------------------------------------------------------------
    ' @function SerializeAttributesArray
    ' @description Serializa o array de atributos para JSON.
    ' @returns {String} JSON gerado
    '------------------------------------------------------------
    Private Function SerializeAttributesArray()
        Set json = New aspJSON
        
        With json.data
            .Add 1, json.Collection()
            With .item(1)
                .Add "Attributes", json.Collection()
                With .item("Attributes")
                    For i = 0 To Attributes.Count - 1
                        .Add i, Attributes(i)
                    Next
                End With
            End With
        End With

        result = json.JSONoutput()
        result = Replace(result, "{", "")
        result = Replace(result, "}", "")
        result = Left(result, Len(result) - 1)
        result = Right(result, Len(result) - 10)

        SerializeAttributesArray = result
        Set json = Nothing
    End Function

    '------------------------------------------------------------
    ' @function SerializeMethodsArray
    ' @description Serializa o array de métodos para JSON.
    ' @returns {String} JSON gerado
    '------------------------------------------------------------
    Private Function SerializeMethodsArray()
        Set json = New aspJSON
        
        With json.data
            .Add "Methods", json.Collection()

            With json.data("Methods")
                For i = 0 To Me.Methods.Count - 1
                    .Add CStr(i), json.Collection()
                    With .item(CStr(i))
                        .Add "name", Me.Methods(i).name
                        .Add "methodType", Me.Methods(i).methodType
                    End With
                Next
            End With    
        End With

        result = json.JSONoutput()
        result = Left(result, Len(result) - 1)
        result = Right(result, Len(result) - 1)

        SerializeMethodsArray = result
        Set json = Nothing
    End Function

    Private Function SerializeInjectsArray()
        Set json = New aspJSON
        
        json.data.Add "Injects", json.Collection()
        Set injectsArray = json.data.Item("Injects")

        For i = 0 To Me.Injects.Count - 1
            injectsArray.Add i, Me.Injects(i).interface
        Next

        result = json.JSONoutput()
        result = Left(result, Len(result) - 1)
        result = Right(result, Len(result) - 1)

        SerializeInjectsArray = result
        Set injectsArray = Nothing
        Set json = Nothing
    End Function

    '------------------------------------------------------------
    ' @function SerializeJoinsArray
    ' @description Serializa o array de joins para JSON.
    ' @returns {String} JSON gerado
    '------------------------------------------------------------
    Private Function SerializeJoinsArray()
        Set json = New aspJSON
        
        With json.data
            .Add "Joins", json.Collection()

            With json.data("Joins")
                For i = 0 To Me.Joins.Count - 1
                    .Add CStr(i), json.Collection()
                    With .item(CStr(i))
                        .Add "value", Me.Joins(i).value
                        .Add "alias", Me.Joins(i).alias
                        .Add "onConditition", Replace(Me.Joins(i).onConditition, " ", "\n")
                        .Add "joinType", Me.Joins(i).joinType
                    End With
                Next
            End With    
        End With

        result = json.JSONoutput()
        result = Left(result, Len(result) - 1)
        result = Right(result, Len(result) - 1)

        SerializeJoinsArray = result
        Set json = Nothing
    End Function

    '------------------------------------------------------------
    ' @function SerializeLookupsArray
    ' @description Serializa o array de lookups para JSON.
    ' @returns {String} JSON gerado
    '------------------------------------------------------------
    Private Function SerializeLookupsArray()
        Set json = New aspJSON
        
        With json.data
            .Add "Lookups", json.Collection()

            With json.data("Lookups")
                For i = 0 To Me.Lookups.Count - 1
                    .Add CStr(i), json.Collection()
                    With .item(CStr(i))
                        .Add "column", Me.Lookups(i).column
                        .Add "alias", Me.Lookups(i).alias
                        .Add "columnAlias", Me.Lookups(i).columnAlias
                        .Add "matchMode", Me.Lookups(i).matchMode
                    End With
                Next
            End With    
        End With

        result = json.JSONoutput()
        result = Left(result, Len(result) - 1)
        result = Right(result, Len(result) - 1)

        SerializeLookupsArray = result
        Set json = Nothing
    End Function

    '------------------------------------------------------------
    ' @function SerializeFieldsArray
    ' @description Serializa o array de fields para JSON.
    ' @returns {String} JSON gerado
    '------------------------------------------------------------
    Private Function SerializeFieldsArray()
        Set json = New aspJSON
        
        With json.data
            .Add "Fields", json.Collection()

            With json.data("Fields")
                For i = 0 To Me.Fields.Count - 1
                    .Add CStr(i), json.Collection()
                    With .item(CStr(i))
                        .Add "column", Me.Fields(i).column
                        .Add "alias", Me.Fields(i).alias
                        .Add "isPrimaryKey", Me.Fields(i).isPrimaryKey
                        .Add "enableInsertPrimaryKey", Me.Fields(i).enableInsertPrimaryKey
                        .Add "matchMode", Me.Fields(i).matchMode
                    End With
                Next
            End With    
        End With

        result = json.JSONoutput()
        result = Left(result, Len(result) - 1)
        result = Right(result, Len(result) - 1)

        SerializeFieldsArray = result
        Set json = Nothing
    End Function

    '------------------------------------------------------------
    ' @function SerializeService
    ' @description Serializa o objeto service para JSON.
    ' @returns {String} JSON gerado
    '------------------------------------------------------------
    Private Function SerializeService()
        Set json = New aspJSON
        With json.data
            .Add "Service", json.Collection()
            If IsNull(Service) Then
                With .item("Service")
                    .Add "interface", Null
                End With
            Else
                With .item("Service")
                    .Add "interface", Me.Service.interface
                End With
            End If
        End With

        result = json.JSONoutput()
        result = Left(result, Len(result) - 1)
        result = Right(result, Len(result) - 1)

        SerializeService = result
        Set json = Nothing
    End Function

    '------------------------------------------------------------
    ' @function SerializeTable
    ' @description Serializa o objeto table para JSON.
    ' @returns {String} JSON gerado
    '------------------------------------------------------------
    Private Function SerializeTable()
        Set json = New aspJSON

        If IsNull(Table) Then
            With json.data
                .Add "Table", Null
            End With
        Else
            With json.data
                .Add "Table", json.Collection()
                    With .item("Table")
                        .Add "value", Me.Table.value
                        .Add "alias", Me.Table.alias
                    End With
            End With
        End If
        
        result = json.JSONoutput()
        result = Left(result, Len(result) - 1)
        result = Right(result, Len(result) - 1)

        SerializeTable = result
        Set json = Nothing
    End Function

    '------------------------------------------------------------
    ' @function SerializeClassName
    ' @description Serializa a classname para JSON.
    ' @returns {String} JSON gerado
    '------------------------------------------------------------
    Private Function SerializeClassName()
        Set json = New aspJSON

        With json.data
            .Add "ClassName", Me.ClassName
        End With

        result = json.JSONoutput()
        result = Left(result, Len(result) - 1)
        result = Right(result, Len(result) - 1)

        SerializeClassName = result
        Set json = Nothing
    End Function

    '------------------------------------------------------------
    ' @function LoadFromCache
    ' @description Carrega as informações de reflexão da entidade atual do cache.
    ' @param {String} path - Caminho da entidade que representa a tabela no banco de dados
    ' @returns {void} 
    '------------------------------------------------------------
    Private Function LoadFromCache(path)
        Dim cacheKey, cache, i, jsonArray
        cacheKey = GetCacheKey(path)

        cache = Replace(Application(cacheKey), " ", "")

        Set p_oJSON = New aspJSON
        p_oJSON.loadJSON(cache)

        ' ClassName
        ClassName = p_oJSON.data("ClassName")
        
        ' Table
        If Not IsNull(p_oJSON.data("Table")) Then
            Dim vTable : Set vTable = New TableAnnotation
            vTable.value = p_oJSON.data("Table")("value")
            vtable.alias = p_oJSON.data("Table")("alias")
            Set Table = vTable
            Set vTable = Nothing
        End If

        ' Service
        value = p_oJSON.data("Service")("interface")
        If Not IsNull(value) Then
            Dim vService : Set vService = New ServiceAnnotation
            vService.interface = value
            Set Service = vService
            Set vService = Nothing
        End If

        ' Attributes
        If Not p_oJSON.data("Attributes").Count = 0 Then
            For i = 0 To p_oJSON.data("Attributes").Count
                Attributes.Add p_oJSON.data("Attributes")(i)
            Next
        End If

        ' Methods
        Dim vMethod
        For i = 0 To p_oJSON.data("Methods").Count - 1
            Set vMethod = New Method
            vMethod.name = p_oJSON.data("Methods")(CStr(i))("name")
            vMethod.methodType = p_oJSON.data("Methods")(CStr(i))("methodType")
            Methods.Add vMethod
            Set vMethod = Nothing 
        Next
        
        ' Joins
        If p_oJSON.data("Joins").Count - 1 > 0 Then
            Dim vJoin
            For i = 0 To p_oJSON.data("Joins").Count - 1
                Set vJoin = New JoinAnnotation
                vJoin.value = p_oJSON.data("Joins")(CStr(i))("value")
                vJoin.alias = p_oJSON.data("Joins")(CStr(i))("alias")
                vJoin.onConditition = Replace(p_oJSON.data("Joins")(CStr(i))("onConditition"), "\n", " ")
                vJoin.joinType = p_oJSON.data("Joins")(CStr(i))("joinType")
                Joins.Add vJoin
                Set vJoin = Nothing
            Next
        End If 

        ' Fields
        If p_oJSON.data("Fields").Count - 1 > 0 Then
            Dim vField
            For i = 0 To p_oJSON.data("Fields").Count - 1
                Set vField = New FieldAnnotation
                vField.column = p_oJSON.data("Fields")(CStr(i))("column")
                vField.alias = p_oJSON.data("Fields")(CStr(i))("alias")
                vField.isPrimaryKey = p_oJSON.data("Fields")(CStr(i))("isPrimaryKey")
                vField.enableInsertPrimaryKey = p_oJSON.data("Fields")(CStr(i))("enableInsertPrimaryKey")
                vField.matchMode = p_oJSON.data("Fields")(CStr(i))("matchMode")
                Fields.Add vField
                Set vField = Nothing
            Next
        End If

        ' Injects
        If Not p_oJSON.data("Injects").Count = 0 Then
            For i = 0 To p_oJSON.data("Injects").Count - 1
                Set vInject = New InjectAnnotation
                vInject.interface = p_oJSON.data("Injects")(i)
                Injects.Add vInject
                Set vInject = Nothing
            Next
        End If

        ' Lookups
        If p_oJSON.data("Lookups").Count - 1 > 0 Then
            Dim vLookup
            For i = 0 To p_oJSON.data("Lookups").Count - 1
                Set vLookup = New LookupAnnotation
                vLookup.column = p_oJSON.data("Lookups")(CStr(i))("column")
                vLookup.alias = p_oJSON.data("Lookups")(CStr(i))("alias")
                vLookup.columnAlias = p_oJSON.data("Lookups")(CStr(i))("columnAlias")
                vLookup.matchMode = p_oJSON.data("Lookups")(CStr(i))("matchMode")
                Lookups.Add vLookup
                Set vLookup = Nothing
            Next
        End If
    End Function

    '------------------------------------------------------------
    ' @function GetCacheKey
    ' @description Retorna a chave do cache.
    ' @param {String} path - Caminho da entidade que representa a tabela no banco de dados
    ' @returns {String} Chave do cache
    '------------------------------------------------------------
    Private Function GetCacheKey(path)
        GetCacheKey = "CACHE_" & UCase(path) & "_METADATA_MIRROR"
    End Function

    '------------------------------------------------------------
    ' @function CheckIfCached
    ' @description Verifica se a reflexão da entidade já está cacheada.
    ' @param {String} path - Caminho da entidade que representa a tabela no banco de dados
    ' @returns {Boolean} Está cacheado
    '------------------------------------------------------------
    Private Function CheckIfCached(path)
        CheckIfCached = False
        cacheKey = GetCacheKey(path)
        cache = Application(cacheKey)   
        If VarType(cache) = 8 Then ' String
            If InStr(cache, "{") Then
                CheckIfCached = True
            End If
        End If
    End Function

    '------------------------------------------------------------
    ' @function ManageMethod
    ' @description Gerencia a existência de métodos na linha recebida.
    ' @param {String} line - Linha da leitura do stream da classe (entidade)
    ' @returns {Object | Nothing} O método contido na linha recebida
    '------------------------------------------------------------
    Private Function ManageMethod(line)
        Dim vMethod
        
        If InStr(line, "Sub") Then    
            parts = Split(Trim(line), " ")
            index = Iif(InStr(line, "Public") Or InStr(line, "Private"), 2, 1)      

            Set vMethod = New Method
            vMethod.name = parts(index)
            vMethod.methodType = "Sub"

            Set ManageMethod = vMethod
            Set vMethod = Nothing

            Exit Function
        End If

        If InStr(line, "Function") Then
            parts = Split(Trim(line), " ")
            index = Iif(InStr(line, "Public") Or InStr(line, "Private"), 2, 1)      

            Set vMethod = New Method
            vMethod.name = parts(index)
            vMethod.methodType = "Function"

            Set ManageMethod = vMethod
            Set vMethod = Nothing

            Exit Function
        End If

        Set ManageMethod = Nothing
        Set vMethod = Nothing
    End Function

    '------------------------------------------------------------
    ' @function ManageInject
    ' @description Gerencia a existência de injects na linha recebida.
    ' @param {String} line - Linha da leitura do stream da classe (entidade)
    ' @returns {Object | Nothing} O inject contido na linha recebida
    '------------------------------------------------------------
    Private Function ManageInject(line)
        If InStr(line, "@Inject") Then
            Dim matches : Set matches = p_Regex.Execute(line)

            interfaceValue = Null

            For Each match In matches
                If Trim(match.SubMatches(0)) = "interface" Then
                    interfaceValue = match.SubMatches(1)
                End If
            Next

            Dim vInject : Set vInject = New InjectAnnotation
            vInject.interface = interfaceValue

            Set ManageInject = vInject
            Set vInject = Nothing

            Exit Function
        End If

        Set ManageInject = Nothing
    End Function
    
    '------------------------------------------------------------
    ' @function ManageAttribute
    ' @description Gerencia a existência de atributos na linha recebida.
    ' @param {String} line - Linha da leitura do stream da classe (entidade)
    ' @returns {String | Null} O atributo contido na linha recebida
    '------------------------------------------------------------
    Private Function ManageAttribute(line)
        If InStr(line, "Public ") Then
            parts = Split(Trim(line), " ")
            If UBound(parts) >= 1 Then
                tempValue = parts(1) ' Public id_aluno -> 1st idx is the attribute
                If tempValue <> "Function" And tempValue <> "Sub" And tempValue <> "Property" And tempValue <> "ENTITY_PATH" Then
                    ManageAttribute = Trim(tempValue)
                    Exit Function
                End If
            End If
        End If

        ManageAttribute = Null
    End Function

    '------------------------------------------------------------
    ' @function ManageField
    ' @description Gerencia a existência de fields na linha recebida.
    ' @param {String} line - Linha da leitura do stream da classe (entidade)
    ' @returns {Object | Nothing} O field contido na linha recebida
    '------------------------------------------------------------
    Private Function ManageField(line)
        If InStr(line, "@Field") Then
            Dim matches : Set matches = p_Regex.Execute(line)

            enableInsertPrimaryKeyValue = False
            isPrimaryKeyValue = False
            matchModeValue = "ALIKE"
            columnValue = Null
            aliasValue = Null

            For Each match In matches
                If Trim(match.SubMatches(0)) = "column" Then
                    columnValue = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "alias" Then
                    aliasValue = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "matchMode" Then
                    matchMode = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "isPrimaryKey" Then
                    If match.SubMatches(1) = "true" Then
                        isPrimaryKeyValue = True 
                    End If
                ElseIf Trim(match.SubMatches(0)) = "enableInsertPrimaryKey" Then
                    If match.SubMatches(1) = "true" Then
                        enableInsertPrimaryKeyValue = True 
                    End If
                End If
            Next

            Dim vField : Set vField = New FieldAnnotation
            vField.column = columnValue
            vField.alias = aliasValue
            vField.isPrimaryKey = isPrimaryKeyValue
            vField.matchMode = matchModeValue
            vField.enableInsertPrimaryKey = enableInsertPrimaryKeyValue

            Set ManageField = vField
            Set vField = Nothing

            Exit Function
        End If

        Set ManageField = Nothing
    End Function

    '------------------------------------------------------------
    ' @function ManageLookup
    ' @description Gerencia a existência de lookups na linha recebida.
    ' @param {String} line - Linha da leitura do stream da classe (entidade)
    ' @returns {Object | Nothing} O lookup contido na linha recebida
    '------------------------------------------------------------
    Private Function ManageLookup(line)
        If InStr(line, "@Lookup") Then
            Dim matches : Set matches = p_Regex.Execute(line)

            matchModeValue   = "ALIKE"
            columnAliasValue = Null
            columnValue      = Null
            aliasValue       = Null

            For Each match In matches
                If Trim(match.SubMatches(0)) = "column" Then
                    columnValue = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "alias" Then
                    aliasValue = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "matchMode" Then
                    matchModeValue = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "columnAlias" Then
                    columnAliasValue = match.SubMatches(1)
                End If
            Next

            Dim vLookup : Set vLookup = New LookupAnnotation
            vLookup.column = columnValue
            vLookup.alias = aliasValue
            vLookup.matchMode = matchModeValue
            vLookup.columnAlias = columnAliasValue

            Set ManageLookup = vLookup
            Set vLookup = Nothing

            Exit Function
        End If

        Set ManageLookup = Nothing
    End Function

    '------------------------------------------------------------
    ' @function ManageJoin
    ' @description Gerencia a existência de joins na linha recebida.
    ' @param {String} line - Linha da leitura do stream da classe (entidade)
    ' @returns {Object | Nothing} O join contido na linha recebida
    '------------------------------------------------------------
    Private Function ManageJoin(line)
        If InStr(line, "@Join") Then
            Dim matches : Set matches = p_Regex.Execute(line)

            tableValue = Null
            aliasValue = Null
            typeValue  = Null
            onValue    = Null

            For Each match In matches
                If Trim(match.SubMatches(0)) = "value" Then
                    tableValue = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "alias" Then
                    aliasValue = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "on" Then
                    onValue = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "type" Then
                    typeValue = match.SubMatches(1)
                End If
            Next

            Dim vJoin : Set vJoin = New JoinAnnotation
            vJoin.onConditition = onValue
            vJoin.joinType = typeValue
            vJoin.value = tableValue
            vJoin.alias = aliasValue

            Set ManageJoin = vJoin
            Set vJoin = Nothing

            Exit Function
        End If

        Set ManageJoin = Nothing
    End Function

    '------------------------------------------------------------
    ' @function ManageService
    ' @description Gerencia a existência de service na linha recebida.
    ' @param {String} line - Linha da leitura do stream da classe (entidade)
    ' @returns {Object | Nothing} O service contido na linha recebida
    '------------------------------------------------------------
    Private Function ManageService(line)
        If InStr(line, "@Service") Then
            Dim matches : Set matches = p_Regex.Execute(line)

            interfaceValue = Null

            For Each match In matches
                If Trim(match.SubMatches(0)) = "interface" Then
                    interfaceValue = match.SubMatches(1)
                End If
            Next

            Dim vService : Set vService = New ServiceAnnotation
            vService.interface = interfaceValue

            Set ManageService = vService
            Set vService = Nothing

            Exit Function
        End If

        Set ManageService = Nothing
    End Function

    '------------------------------------------------------------
    ' @function ManageTable
    ' @description Gerencia a existência de table na linha recebida.
    ' @param {String} line - Linha da leitura do stream da classe (entidade)
    ' @returns {Object | Nothing} A table contida na linha recebida
    '------------------------------------------------------------
    Private Function ManageTable(line)
        If InStr(line, "@Table") Then
            Dim matches : Set matches = p_Regex.Execute(line)

            tableValue = Null
            aliasValue = Null

            For Each match In matches
                If Trim(match.SubMatches(0)) = "value" Then
                    tableValue = match.SubMatches(1)
                ElseIf Trim(match.SubMatches(0)) = "alias" Then
                    aliasValue = match.SubMatches(1)
                End If
            Next

            Dim vTable : Set vTable = New TableAnnotation
            vTable.value = tableValue
            vTable.alias = aliasValue

            Set ManageTable = vTable
            Set vTable = Nothing

            Exit Function
        End If

        Set ManageTable = Nothing
    End Function

    '------------------------------------------------------------
    ' @function ManageClassName
    ' @description Gerencia o nome da classe da entidade.
    ' @param {String} line - Linha da leitura do stream da classe (entidade)
    ' @returns {String | Null} O nome da classe da entidade
    '------------------------------------------------------------
    Private Function ManageClassName(line)
        If InStr(line, "Class ") Then
            parts = Split(Trim(line), " ")
            ManageClassName = parts(1)
            Exit Function
        End If

        ManageClassName = Null
    End Function
End Class

Class Method
    Public name
    Public methodType

    Public Sub Class_Initialize()
        methodType = Null
        name = Null
    End Sub
End Class

%>