<!--#include virtual="/infra/persistence/factory/QueryFactory.asp"-->
<!--#include virtual="/infra/persistence/db/Transaction.asp"-->
<!--#include virtual="/infra/persistence/register.asp"-->

<%

Response.Charset = "ISO-8859-1" 
Response.ContentType = "text/html"

' @author: tucokk
Class ObjectManager
    Private p_QueryFactory
    Public strSQL

    Public Sub Class_Initialize()
        Set p_QueryFactory = New QueryFactory
    End Sub

    Public Sub Class_Terminate()
        Set p_QueryFactory = Nothing
    End Sub

    '------------------------------------------------------------
    ' @function FindByPrimaryKey
    ' @description Realiza a busca no banco de dados com base na chave primária definida na entidade.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {Object} Entidade populada (ou não, caso o query não encontre resultados)
    '------------------------------------------------------------
    Public Function FindByPrimaryKey(entity)
        strSQL = p_QueryFactory.GetSelectPrimaryKey(entity)
        Set vector = p_QueryFactory.Query(strSQL, entity)
        Set FindByPrimaryKey = vector(0)
    End Function

    '------------------------------------------------------------
    ' @function FindByOtherFields
    ' @description Realiza a busca no banco de dados com base em outros atributos da entidade.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @param {String} where - Comando where a ser adicionado no query de forma extra (Ex.: "AND a.id_aluno = 12")
    ' @param {String} orderBy - Comando orderBy a ser adicionado no query de forma extra (Ex.: "ORDER BY a.nome")
    ' @param {String} limit - Comando limit a ser adicionado no query de forma extra (Ex.: "TOP 1")
    ' @returns {Object[]} Array de entidades populadas. Caso o query não encontre resultados, retorna um array vazio
    '------------------------------------------------------------
    Public Function FindByOtherFields(entity, where, orderBy, limit) 
        strSQL = p_QueryFactory.GetSelect(entity, where, orderBy, limit)
        Set FindByOtherFields = p_QueryFactory.Query(strSQL, entity)
    End Function

    '------------------------------------------------------------
    ' @function FindFirstdByOtherFields
    ' @description Realiza a busca no banco de dados com base em outros atributos da entidade.
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @param {String} where - Comando where a ser adicionado no query de forma extra (Ex.: "AND a.id_aluno = 12")
    ' @param {String} orderBy - Comando orderBy a ser adicionado no query de forma extra (Ex.: "ORDER BY a.nome")
    ' @returns {Object} Entidade populada (ou não, caso o query não encontre resultados)
    '------------------------------------------------------------
    Public Function FindFirstdByOtherFields(entity, where, orderBy)
        strSQL = p_QueryFactory.GetSelect(entity, where, orderBy, "TOP 1")
        Set vector = p_QueryFactory.Query(strSQL, entity)
        Set FindFirstdByOtherFields = vector(0)
    End Function

    '------------------------------------------------------------
    ' @function FindByQuery
    ' @description Realiza a busca no banco de dados com base no query informado.
    ' @param {String} strSQL - Query que será executado no banco de dados
    ' @param {Object} entity - Entidade que representa a tabela no banco de dados
    ' @returns {Object[]} Array de entidades populadas. Caso o query não encontre resultados, retorna um array vazio
    '------------------------------------------------------------
    Public Function FindByQuery(strSQL, entity)
        Set FindByQuery = p_QueryFactory.Query(strSQL, entity)
    End Function

    '------------------------------------------------------------
    ' @function SqlQuery
    ' @description Realiza a busca no banco de dados com base no query informado.
    ' @param {String} strSQL - Query que será executado no banco de dados
    ' @returns {Object[]} Array de dicionários. Caso o query não encontre resultados, retorna um array vazio.
    '------------------------------------------------------------
    Public Function SqlQuery(strSQL)
        Set SqlQuery = p_QueryFactory.AbstractQuery(strSQL)
    End Function
End Class

%>