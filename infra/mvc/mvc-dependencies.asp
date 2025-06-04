<% '@author: tucokk %>

<!-- Dependencies -->
<!--#include virtual="/infra/utils/index.asp"-->
<!--#include virtual="/application/services/index.asp"-->

<!-- Singletons -->
<!--#include virtual="/infra/mvc/core/WebApplication.asp"-->
<!--#include virtual="/infra/mvc/core/MvcEngine.asp"-->
<!--#include virtual="/infra/mvc/core/DependencyInjection.asp"-->
<!--#include virtual="/infra/persistence/ObjectManager.asp"-->

<%

ExecuteGlobal "Set SINGLETONS = Server.CreateObject(""Scripting.Dictionary"")"

SINGLETONS.Add "Manager", New ObjectManager
SINGLETONS.Add "MVC", New MvcEngine

%>