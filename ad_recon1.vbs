Set objRootDSE = GetObject("LDAP://RootDSE")
strDomain = objRootDSE.Get("defaultNamingContext")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADSDSOObject;"
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection

objCommand.CommandText = "<LDAP://" & strDomain & ">;(objectCategory=User);samAccountName,servicePrincipalName;subtree"
Set objRecordSet = objCommand.Execute

On Error Resume Next
While Not objRecordSet.EOF
    strUser = objRecordSet.Fields("samAccountName").Value
    spnField = objRecordSet.Fields("servicePrincipalName").Value
    
    strSPN = ""
    If IsArray(spnField) Then
        For Each spn In spnField
            strSPN = strSPN & spn & " "
        Next
    Else
        strSPN = spnField
    End If
    
    Wscript.Echo "User: " & strUser & " | SPN: " & strSPN
    objRecordSet.MoveNext
Wend
