<%

Set instance = MvcCreateController(className, path)

Public Sub SingletonsTerminate()
    Set App = Nothing
    Set Manager = Nothing
    Set DI = Nothing
    Set Engine = Nothing
    Set instance = Nothing
End Sub

Public Sub MvcRunClass(className, path, action)
    Call MvcInjectDependencies(instance, path)

    If InStr(className, "Controller") Then
        ExecuteGlobal Format("instance.{0}()", action)
    End If

    Call SingletonsTerminate()
End Sub

Public Function MvcCreateController(className, path)
    Dim strCommand
    strCommand = "Set temp = New " & className
    ExecuteGlobal strCommand
    Set MvcCreateController = temp
End Function

Public Sub MvcInjectDependencies(instance, path)
    strLog = Format("Injecting dependencies: {0}", path)
    Set injections = DI.GetInjectedServices(path)
    For Each interface In injections.Keys
        strLog = strLog & chr(13) & AddTabIndent(Format("-> Injecting dependency [{0} -> {1}]", Array(interface, injections(interface))), 20)
        ExecuteGlobal Format("Set instance.{0} = New {1}", Array(interface, injections(interface)))
    Next
    Log strLog, "Infra.Mvc.MvcDependenciesTerminate"
End Sub

Call MvcRunClass(CLASSNAME, FILE_PATH, Session("action"))
Response.End

%>