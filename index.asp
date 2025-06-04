<!--#include virtual="/infra/mvc/mvc-dependencies.asp"-->

<% 

Set DI = New DependencyInjection
DI.StartServicesReflectionCaching()
Set Di = Nothing

Set App = New WebApplication
App.route Request("controller"), Request("action") 
Set App = Nothing

SINGLETONS("MVC").TerminateSingletons()

%>
