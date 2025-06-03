<%

' @author: tucokk
Class WebApplication
    Private p_Mirror
    Private p_DependencyInjection

    Public Sub Class_Initialize()
        Set p_Mirror = New Reflection
        Set p_DependencyInjection = New DependencyInjection
    End Sub

    Public Sub Class_Terminate()
        Set p_Mirror = Nothing
    End Sub

    Public Sub Route(controllerName, actionName)
        If controllerName = "" Then controllerName = "home"
        If actionName     = "" Then actionName = "index"
        
        controllerPath = GetControllerPath(controllerName, actionName)
        If Not ControllerExists(controllerPath) Then 
            Exit Sub
        End If

        Session("action") = Title(actionName)

        Server.Execute controllerPath   
    End Sub    

    Private Function ControllerExists(path)
        ControllerExists = FileExists(Server.MapPath(path))
    End Function

    Private Function GetControllerPath(controllerName, actionName)
        GetControllerPath = Format("web/controllers/{0}controller.asp", controllerName)
    End Function    
End Class

%>