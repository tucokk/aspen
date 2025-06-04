<% '@author: tucokk %>

<!-- Dependencies -->
<!--#include virtual="/infra/utils/index.asp"-->
<!--#include virtual="/application/services/index.asp"-->

<!-- Singletons -->
<!--#include virtual="/infra/mvc/core/WebApplication.asp"-->
<!--#include virtual="/infra/mvc/core/ViewEngine.asp"-->
<!--#include virtual="/infra/mvc/core/DependencyInjection.asp"-->
<!--#include virtual="/infra/persistence/ObjectManager.asp"-->

<%

Dim App     : Set App     = New WebApplication
Dim DI      : Set DI      = New DependencyInjection
Dim Manager : Set Manager = New ObjectManager
Dim Engine  : Set Engine  = New ViewEngine

%>