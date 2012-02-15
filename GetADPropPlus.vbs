Function GetADPropPlus(ByVal SearchField As String, ByVal ObjectType As String, ByVal SearchString As String, ByVal ReturnField As String) As String
    'Get the domain string ("dc=domain, dc=local")
    Dim strDomain As String
    strDomain = GetObject("LDAP://rootDSE").Get("defaultNamingContext")
   
    'ADODB Connection to AD
    Dim objConnection As ADODB.Connection
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"
       
    'Connection
    Dim objCommand As ADODB.Command
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection
       
    'Search the AD recursively, starting at root of the domain
    objCommand.CommandText = _
        "<LDAP://" & strDomain & ">;(&(objectCategory=" & ObjectType & ")" & _
        "(" & SearchField & "=" & SearchString & "));" & SearchField & "," & ReturnField & ";subtree"
    'Recordset
    Dim objRecordSet As ADODB.Recordset
    Set objRecordSet = objCommand.Execute
         
   
    If objRecordSet.RecordCount = 0 Then
        GetADPropPlus = "not found"  'no records returned
    Else
        'GetADPropPlus = objRecordSet.Fields(ReturnField)  'return value
        tempVal = VarType(objRecordSet.Fields(ReturnField).Value)
        If (tempVal = 8) Then
            GetADPropPlus = objRecordSet.Fields(ReturnField).Value
        Else
            GetADPropPlus = objRecordSet.Fields(ReturnField).Value(0)
        End If
    End If
     
    'Close connection
    objConnection.Close
   
    'Cleanup
    Set objRecordSet = Nothing
    Set objCommand = Nothing
    Set objConnection = Nothing
End Function

