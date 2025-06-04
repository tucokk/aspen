<%

' @author: tucokk
Class DependencyInjection
    Private p_Mirror
    Private p_configFilePath
    Private p_configPath

    Public Sub Class_Initialize()
        Set p_Mirror = New Reflection
        p_configFilePath = ".aspen/.di/di.properties"
        p_configPath = ".aspen/.di"
    End Sub

    Public Sub Class_Terminate()
        Set p_Mirror = Nothing
    End Sub

    Public Function GetInjectedServices(classPath)
        Set res = Server.CreateObject("Scripting.Dictionary")
        p_Mirror.Reflect(classPath)

        If p_Mirror.Injects.Count > 0 Then
            Set services = Server.CreateObject("System.Collections.ArrayList")
            strLog = Format("Injecting dependencies: {0}", classPath)

            For Each injection In p_mirror.Injects
                service = GetServiceByInterface(injection.interface)
                strLog = strLog & chr(13) & AddTabIndent(Format("-> Injecting dependency [{0} -> {1}]", Array(injection.interface, service)), 16)
                strCommand = Format("Set p_Instance.{0} = New {1}", Array(injection.interface, service))
                services.Add strCommand
            Next

            res.Add "services", services
            res.Add "log", strLog
        End If

        Set GetInjectedServices = res
    End Function

    Public Function GetServiceByInterface(interface)
        path = ".aspen/.di/di.properties"
        value = ReadValueFromFileByKey(path, interface)
        GetServiceByInterface = value
    End Function

    Public Sub StartServicesReflectionCaching()
        filePath = ".aspen/.di/di.properties"
        path = ".aspen/.di"

        CreateFolder(path)
        CreateFile(filePath)
        ClearFileContent(filePath)

        path = Server.MapPath("/application/services")
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        Set folder = fso.GetFolder(path)
        For Each file in folder.Files
            If InStr(file, "index") Then  
            Else
                parts = Split(file, "\")
                filePath = Format("/{0}/{1}/{2}", Array(parts(UBound(parts) - 2), parts(UBound(parts) - 1), parts(UBound(parts))))

                Set mirror = New Reflection
                mirror.Reflect(filePath)   
                
                If Not IsNull(mirror.Service) Then
                    strCommand = Format("{0}={1}", Array(mirror.Service.interface, mirror.ClassName))
                    WriteToFile ".aspen/.di/di.properties", strCommand
                End If
            End If
        Next
    End Sub
End Class

%>