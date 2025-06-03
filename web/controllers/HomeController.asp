<!--#include virtual="/infra/mvc/mvc-dependencies-initialize.asp"-->

<%

Const FILE_PATH = "/web/controllers/HomeController.asp"
Const CLASSNAME = "HomeController"

' @author: tucokk
Class HomeController
    
    ' @Inject(interface="IHomeService")
    Public IHomeService
    ' @Inject(interface="ITestService")
    Public ITestService

    Public Sub Index()
        Set teste = ITestService.FindAluno()
        response.write teste.nome
        test()
    End Sub

    Public Sub Test()
        IHomeService.teste()
    End Sub
End Class

%>

<!--#include virtual="/infra/mvc/mvc-dependencies-terminate.asp"-->
