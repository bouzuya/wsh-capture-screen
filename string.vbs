Option Explicit

Public Function PadLeft(ByVal strValue, ByVal intTotalLength, ByVal strPadding)
    Dim strResult
    strResult = strValue
    While Len(strResult) < intTotalLength
        strResult = strPadding & strResult
    Wend
    PadLeft = strResult
End Function

Public Function PadRight(ByVal strValue, ByVal intTotalLength, ByVal strPadding)
    Dim strResult
    strResult = strValue
    While Len(strResult) < intTotalLength
        strResult = strResult & strPadding
    Wend
    PadRight = strResult
End Function

