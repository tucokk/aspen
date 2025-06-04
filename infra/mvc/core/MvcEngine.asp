<%

' @author: tucokk
Class MvcEngine
    Private p_Instance, p_DependencyInjection
    Private p_className, p_classPath

    Public Sub Resolve(classPath, className)
        Set p_Instance = CreateInstance(className)
        p_className = Trim(className)
        p_classPath = Trim(classPath)
    
        Call RunMvcClass()
    End Sub

    Public Sub RenderView(viewPath, viewBag)
        If Not IsNull(viewBag) And Not viewBag Is Nothing Then
            Set Session("ViewBag") = viewBag
        Else
            Set Session("ViewBag") = Nothing
        End If

        fullPath = Format("/web/views/{0}.asp", viewPath)
        If FileExists(Server.MapPath(fullPath)) Then
            Server.Execute(fullPath)
        End If
    End Sub

    Public Sub TerminateSingletons()
        For Each key In SINGLETONS.Keys
            Set obj = SINGLETONS(key)
            If IsObject(obj) Then
                Set obj = Nothing
            End If
        Next
        SINGLETONS.RemoveAll
    End Sub

    Private Sub CallActions()
        If InStr(p_className, "Controller") Then
            action = Request("action")
            strCommand = Format("p_Instance.{0}", action)
            Execute(strCommand)
        End If
    End Sub 

    Private Function RunMvcClass()
        Call InjectDependencies()
        Call CallActions()
        Call TerminateSingletons()
    End Function

    Private Sub InjectDependencies()
        Set p_DependencyInjection = New DependencyInjection
        Set injections = p_DependencyInjection.GetInjectedServices(p_classPath)
        
        Log injections("log"), "Infra.Mvc.Core.MvcEngine"
        For Each command In injections("services")
            Execute(command)
        Next

        Set injections = Nothing
        Set p_DependencyInjection = Nothing
    End Sub

    Private Function CreateInstance(className)
        strCommand = Format("Set temp = New {0}", className)
        ExecuteGlobal strCommand
        Set CreateInstance = temp
    End Function
End Class

%>