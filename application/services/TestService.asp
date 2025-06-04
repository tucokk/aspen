<!--#include virtual="/infra/mvc/mvc-dependency-injection-only.asp"-->

<%

' @author: tucokk

' @Service(interface="ITestService")
Class TestService
    Public Function FindAluno()
        Dim vAluno : Set vAluno = New Aluno
        vAluno.id_aluno = 12
        Set FindAluno = SINGLETONS("Manager").FindByPrimaryKey(vAluno)
    End Function
End Class

%>