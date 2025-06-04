<%

' @author: tucokk
Class ViewEngine
    Public Sub RenderView(viewPath, model)
        If Not IsNull(model) Then
            Set Session("ViewBag") = model
        Else
            Set Session("ViewBag") = Nothing
        End If

        fullPath = Format("/web/views/{0}.asp", viewPath)
        If FileExists(Server.MapPath(fullPath)) Then
            Server.Execute(fullPath)
        End If

        Response.End
    End Sub
End Class

%>