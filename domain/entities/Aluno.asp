<%

' @author: Arthur Ribeiro

' @Table(value="aluno", alias="a")
' @Join(type="LEFT" value="cidade" alias="c" on="c.id_cidade = a.id_cidade")
' @Join(type="LEFT" value="cidade" alias="cn" on="cn.id_cidade = a.id_cidade_nasc")
' @Join(type="LEFT" value="cidade" alias="cm" on="cm.id_cidade = a.id_cidade_medio")
Class Aluno
    ' @Field(column="id_aluno", alias="id_aluno", isPrimaryKey="true")
    Public id_aluno
    ' @Field(column="nome", alias="nome")
    Public nome
    ' @Field(column="e_mail", alias="email")
    Public email
    ' @Field(column="id_cargo", alias="id_cargo")
    Public id_cargo
    ' @Field(column="fone_residencial", alias="telefone_residencial")
    Public telefone_residencial
    ' @Field(column="fone_celular", alias="telefone_celular")
    Public telefone_celular
    ' @Field(column="nascimento", alias="data_nascimento")
    Public data_nascimento
    ' @Field(column="sexo", alias="sexo")
    Public sexo
    ' @Field(column="nome_pai", alias="nome_pai")
    Public nome_pai
    ' @Field(column="nome_mae", alias="nome_mae")
    Public nome_mae
    ' @Field(column="num_cpf", alias="cpf")
    Public cpf
    ' @Field(column="num_rg", alias="rg")
    Public rg
    ' @Field(column="escolaridade", alias="escolaridade")
    Public escolaridade
    ' @Field(column="nome_curso", alias="nome_curso")
    Public nome_curso
    ' @Field(column="nome_escola", alias="nome_escola")
    Public nome_escola
    ' @Field(column="tipo_logradouro", alias="tipo_logradouro")
    Public tipo_logradouro
    ' @Field(column="nome_logradouro", alias="nome_logradouro")
    Public nome_logradouro
    ' @Field(column="numero", alias="numero")
    Public numero
    ' @Field(column="complemento", alias="complemento")
    Public complemento
    ' @Field(column="id_cidade", alias="id_cidade")
    Public id_cidade
    ' @Lookup(columnAlias="c", column="cidade", alias="m_cidade")
    Public m_cidade
    ' @Field(column="cep", alias="cep")
    Public cep
    ' @Field(column="bairro", alias="bairro")
    Public bairro
    ' @Field(column="end_cobranca", alias="endereco_cobranca")
    Public endereco_cobranca
    ' @Field(column="end_comercial", alias="endereco_comercial")
    Public endereco_comercial
    ' @Field(column="id_antigo", alias="id_antigo")
    Public id_antigo
    ' @Field(column="titulo_num", alias="numero_titulo_eleitor")
    Public numero_titulo_eleitor
    ' @Field(column="titulo_zona", alias="zona_titulo_eleitor")
    Public zona_titulo_eleitor
    ' @Field(column="titulo_sessao", alias="sessao_titulo_eleitor")
    Public sessao_titulo_eleitor
    ' @Field(column="militar_doc", alias="militar_documento")
    Public militar_documento
    ' @Field(column="militar_ano", alias="militar_ano")
    Public militar_ano
    ' @Field(column="nr_sere", alias="numero_sere")
    Public numero_sere
    ' @Field(column="nacionalidade", alias="nacionalidade")
    Public nacionalidade
    ' @Field(column="uf_rg", alias="uf_rg")
    Public uf_rg
    ' @Field(column="escola_medio", alias="escola_medio")
    Public escola_medio
    ' @Field(column="id_cidade_medio", alias="id_cidade_medio")
    Public id_cidade_medio
    ' @Field(column="ano_medio", alias="ano_medio")
    Public ano_medio
    ' @Field(column="recebe_email", alias="recebe_email")
    Public recebe_email
    ' @Field(column="id_aluno_anterior", alias="id_aluno_anterior")
    Public id_aluno_anterior
    ' @Field(column="id_cidade_nasc", alias="id_cidade_nascimento")
    Public id_cidade_nascimento
    ' @Lookup(columnAlias="cn", column="cidade", alias="m_cidade_nascimento")
    Public m_cidade_nascimento
    ' @Lookup(columnAlias="cm", column="cidade", alias="m_cidade_medio")
    Public m_cidade_medio
    ' @Field(column="id_estado_civil", alias="id_estado_civil")
    Public id_estado_civil
    ' @Field(column="id_usuario_cad", alias="id_usuario_cadastro")
    Public id_usuario_cadastro
    ' @Field(column="data_cadastro", alias="data_cadastro")
    Public data_cadastro
    ' @Field(column="foto", alias="foto")
    Public foto
    ' @Field(column="anoConclusaoFund", alias="ano_conclusao_fundamental")
    Public ano_conclusao_fundamental
    ' @Field(column="nome_arq_importacao", alias="nome_arquivo_importacao")
    Public nome_arquivo_importacao
    ' @Field(column="data_expedicao_rg", alias="data_expedicao_rg")
    Public data_expedicao_rg
    ' @Field(column="orgao_expedidor_rg", alias="orgao_expedidor_rg")
    Public orgao_expedidor_rg
    ' @Field(column="nome_social", alias="nome_social")
    Public nome_social
    ' @Field(column="e_mail2", alias="email_2")
    Public email_2
    ' @Field(column="ultima_atualizacao", alias="ultima_atualizacao")
    Public ultima_atualizacao

    Public ENTITY_PATH

    Public Sub Class_Initialize()
        id_aluno = Null
        nome = Null
        email = Null
        id_cargo = Null
        telefone_residencial = Null
        telefone_celular = Null
        data_nascimento = Null
        sexo = Null
        nome_pai = Null
        nome_mae = Null
        cpf = Null
        rg = Null
        esolaridade = Null
        nome_curso = Null
        nome_escola = Null
        tipo_logradouro = Null
        nome_logradouro = Null
        numero = Null
        complemento = Null
        id_cidade = Null
        cep = Null
        bairro = Null
        endereco_cobranca = Null
        endereco_comercial = Null
        id_antigo = Null
        numero_titulo_eleitor = Null
        zona_titulo_eleitor = Null
        sessao_titulo_eleitor = Null
        militar_documento = Null
        militar_ano = Null
        numero_sere = Null
        nacionalidade = Null
        uf_rg = Null
        escola_medio = Null
        id_cidade_medio = Null
        ano_medio = Null
        recebe_email = Null
        id_aluno_anterior = Null
        id_cidade_nascimento = Null
        id_estado_civil = Null
        id_usuario_cadastro = Null
        data_cadastro = Null
        foto = Null
        ano_conclusao_fundamental = Null
        nome_arquivo_importacao = Null
        data_expedicao_rg = Null
        orgao_expedidor_rg = Null
        nome_social = Null
        email_2 = Null
        ultima_atualizacao = Null
        m_cidade_nascimento = Null
        m_cidade = Null
        m_cidade_medio = Null

        Me.ENTITY_PATH = "/domain/entities/Aluno.asp"
    End Sub

    Public Sub SetProperty(propName, value)
        On Error Resume Next
            Execute propName & " = value"
            If Err.Number <> 0 Then
                Throw "Erro ao definir " & propName & ": " & Err.Description
            End If
        On Error Goto 0
    End Sub

End Class

%>