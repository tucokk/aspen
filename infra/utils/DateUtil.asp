<%

' @author: tucokk
Function ddMmYyyyToDate(value)
    Dim vDay, vMonth, vYear

	value = Replace(value, "/", "")
	value = Replace(value, "-", "")
	value = Replace(value, "'", "")

    vDay = Left(value, 2)
    vMonth = Mid(value, 3, 2)
    vYear = Right(value, 4) 

    ddMmYyyyToDate = DateSerial(vYear, vMonth, vDay)
End Function

%>