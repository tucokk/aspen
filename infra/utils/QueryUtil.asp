<%

' @author: tucokk
Function sanatizeParameter(str)
    Dim badChars, pattern, i
    Dim newChars : newChars = str
    Dim regEx : Set regEx = New RegExp

    ' Padrões maliciosos comuns em injeções SQL
    badChars = Array( _
        "select(.*)(from|with|by){1}", _
        "insert(.*)(into|values){1}", _
        "update(.*)set", _
        "delete(.*)(from|with){1}", _
        "drop(.*)(from|aggre|role|assem|key|cert|cont|credential|data|endpoint|event|fulltext|function|index|login|type|schema|procedure|que|remote|role|route|sign|stat|syno|table|trigger|user|view|xml){1}", _
        "alter(.*)(application|assem|key|author|cert|credential|data|endpoint|fulltext|function|index|login|type|schema|procedure|que|remote|role|route|serv|table|user|view|xml){1}", _
        "xp_", "sp_", "restore\s", "grant\s", "revoke\s", _
        "dbcc", "dump", "use\s", "set\s", "truncate\s", _
        "backup\s", "load\s", "save\s", "shutdown", _
        "cast(.*)\(", "convert(.*)\(", "execute\s", _
        "updatetext", "writetext", "reconfigure", _
        "/\*", "\*/", ";", "\-\-", "\[", "\]", _
        "char(.*)\(", "nchar(.*)\(" _
    )

    ' Remover os padrões maliciosos usando expressão regular
    For i = 0 To UBound(badChars)
        regEx.Pattern = badChars(i)
        regEx.IgnoreCase = True
        regEx.Global = True
        newChars = regEx.Replace(newChars, "")
    Next

    ' Escapar aspas simples
    newChars = Replace(newChars, "'", "''")

    Set regEx = Nothing
    sanatizeParameter = newChars
End Function


%>