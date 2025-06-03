<!--#include virtual="/infra/mvc/mvc-dependencies-initialize.asp"-->

<% 

DI.StartServicesReflectionCaching()
App.route Request("controller"), Request("action") 

%>

<!--#include virtual="/infra/mvc/mvc-dependencies-terminate.asp"-->