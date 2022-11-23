<%

Function RandomInt(min, max)
    If Not isNumeric(min) Then min = 0
    If Not isNumeric(max) Then min = 1000
    Randomize
    RandomInt = Int((max-min+1)*Rnd+min)
End Function

Function GUID()
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    GUID = Mid(TypeLib.Guid, 2, 36)
    Set TypeLib = Nothing
End Function

%>