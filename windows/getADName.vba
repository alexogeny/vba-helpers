Function getADName()
    Set objAD = CreateObject("ADSystemInfo")
    Set objUser = GetObject("LDAP://" & objAD.UserName)
    strDisplayName = objUser.DisplayName
    strDisplayName = VBA.Split(strDisplayName, ", ")
    getADName = strDisplayName(1) & " " & strDisplayName(0)
End Function
