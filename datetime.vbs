Option Explicit

' Import string.vbs

Public Function JSTToUTC(ByVal dtmValue)
    JSTToUTC = DateAdd("h", -9, dtmValue)
End Function

' -> YYYY-MM-DDTHH:MM:SSZ
Public Function ToISO8601(ByVal dtmValue)
    Dim strDate, strTime, strTimeZone
    strDate = ToISO8601Date(dtmValue)
    strTime = ToISO8601Time(dtmValue)
    strTimeZone = "Z" ' FIXME
    ToISO8601 = strDate & "T" & strTime & strTimeZone
End Function

' -> YYYY-MM-DD
Public Function ToISO8601Date(ByVal dtmValue)
    Dim strYear, strMonth, strDay
    strYear = CStr(Year(dtmValue))
    strMonth = PadLeft(CStr(Month(dtmValue)), 2, "0")
    strDay = PadLeft(CStr(Day(dtmValue)), 2, "0")
    ToISO8601Date = strYear & "-" & strMonth & "-" & strDay
End Function

' -> HH:MM:SS
Public Function ToISO8601Time(ByVal dtmValue)
    Dim strHour, strMinute, strSecond
    strHour = PadLeft(CStr(Hour(dtmValue)), 2, "0")
    strMinute = PadLeft(CStr(Minute(dtmValue)), 2, "0")
    strSecond = PadLeft(CStr(Second(dtmValue)), 2, "0")
    ToISO8601Time = strHour & ":" & strMinute & ":" & strSecond
End Function

