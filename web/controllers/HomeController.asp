<!--#include virtual="/infra/mvc/mvc-dependencies-initialize.asp"-->

<%

Const FILE_PATH = "/web/controllers/HomeController.asp"
Const CLASSNAME = "HomeController"

' @author: tucokk
Class HomeController
    
    ' @Inject(interface="ITestService")
    Public ITestService

    Public Sub Index()
        Set teste = ITestService.FindAluno()
        response.write teste.nome
    End Sub
End Class

%>

<!--#include virtual="/infra/mvc/mvc-dependencies-terminate.asp"-->
