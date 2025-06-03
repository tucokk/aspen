<%

' @author: tucokk

' @Service(interface="ITestService")
Class TestService
    Public Function FindAluno()
        Dim vAluno : Set vAluno = New Aluno
        vAluno.id_aluno = 12
        Set FindAluno = manager.FindByPrimaryKey(vAluno)
    End Function
End Class

%>