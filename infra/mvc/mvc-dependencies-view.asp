<!--#include virtual="/infra/utils/index.asp"-->

<%

Dim ViewBag 

If Not Session("ViewBag") Is Nothing Then
    Set ViewBag = Session("ViewBag")
Else
    ViewBag = Null
End If

%>