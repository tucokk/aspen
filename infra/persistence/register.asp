<% 

' @author: tucokk
'
' As novas entidades devem ser registradas neste arquivo para funcionarem no framework de persist�ncia.
' Deve ser adicionada na function abaixo uma linha seguindo o mesmo padr�o das demais, mas utilizando a nova entidade.
' Tamb�m deve ser realizado o include da mesma ao final do arquivo, conforme as demais j� encontram-se.

Public Function CreateObjectFromClassName(className)
    Select Case className
        Case "Aluno"
            Set CreateObjectFromClassName = New Aluno
        Case Else
            Set CreateObjectFromClassName = Nothing
    End Select
End Function

%>

<!--#include virtual="/domain/entities/Aluno.asp"-->