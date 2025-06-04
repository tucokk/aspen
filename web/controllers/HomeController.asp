<!--#include virtual="/infra/mvc/mvc-dependencies.asp"-->

<%

' @author: tucokk
Class HomeController 

    ' @Inject(interface="ITestService")
    Public ITestService

    Public Sub Index()
        Set vAluno = ITestService.FindAluno()
        debug vALuno.nome
    End Sub
End Class

Const FILE_PATH = "/web/controllers/HomeController.asp"
Const CLASSNAME = "HomeController"

SINGLETONS("MVC").Resolve FILE_PATH, CLASSNAME

%>