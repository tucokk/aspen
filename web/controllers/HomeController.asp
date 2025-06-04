<!--#include virtual="/infra/mvc/mvc-dependencies-initialize.asp"-->

<%

' @author: tucokk
Class HomeController 

    ' @Inject(interface="ITestService")
    Public ITestService

    Public Sub Index()
        Set vAluno = ITestService.FindAluno()
        Engine.RenderView "home/index", vAluno
    End Sub
End Class

Const FILE_PATH = "/web/controllers/HomeController.asp"
Const CLASSNAME = "HomeController"

%>

<!--#include virtual="/infra/mvc/mvc-dependencies-terminate.asp"-->
