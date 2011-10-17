Attribute VB_Name = "Module1"
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                                    Bluehat TCP/IP Server                               *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                                  v1.0 - (Unstable Version)                              *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                           VIGNESH RAJAGOPALAN and Srinath PB                 *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                                    (Suggestions welcome)                               *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                               vignesh.rajagopalan@yahoo.com                       *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
' [Module1 - Sub routines]



Public Function datapassthru(ByVal data As String) As Integer
    '
    'A subroutine to check the nature of the client-code which will return
    'an integer corresponding to the type of the client-code.
    '
    data = Left(data, 16) 'Extract the code from the data.
    
        If data = "$***|\/|\/|***$P" Then
            datapassthru = 0
        ElseIf data = "$***|\/|\/|***$C" Then
            datapassthru = 1
        ElseIf data = "$***|\/|\/|***$N" Then
            datapassthru = 2
        Else 'no code, return -1.
            datapassthru = -1
        End If
End Function

Public Function authenticate(ByVal data As String) As Boolean
    '
    'A subroutine which is called when the client-code type is 0.
    'This extracts the Username and Password from the data using
    'getuser() and getpass() and sends the information to the
    'database for authentication and returns true or false.
    '
    Dim Username As String, Password As String
    Dim cn As ADODB.Connection, rs As ADODB.Recordset
    Dim st As String
    
    Username = getuser(data) 'Extract the Username.
    Password = getpass(data) 'Extract the Password
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    'Define the query
    st = "select * from users where username = '" & Username & "'"
    
    'Connect to the database
    If cn.State > 1 Then cn.Close
    cn.Open "Provider=oraoledb.oracle.1;user id=srinath;password=srinath;"
    cn.Execute st
    
    'Open the recordset
    If rs.State > 1 Then rs.Close
    rs.Open st, cn, adOpenDynamic
      
    'Match the input with information in the database
    'and return T or F accordingly.
     
    If rs.EOF Then
        authenticate = False
    ElseIf rs!Password = Password Then
        authenticate = True
    Else
        authenticate = False
    End If
    
    'Close the recordset and connection.
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
     
End Function

Public Function getuser(ByVal data As String) As String

    'Sub routine to extract the Username from the data.
    'Used in the authenticate subroutine
    
    Dim noch As Integer
    
    data = remseccode(data)
    noch = InStr(data, "+")
    noch = (Len(data) - (Len(data) - noch))
    getuser = Left(data, noch - 1)

End Function

Public Function remseccode(data As String) As String

    'Sub routine to remove the client code from the data

    data = Right(data, (Len(data) - 16))
    remseccode = data
End Function

Public Function getpass(data As String) As String
    
    'Sub routine to extract the Password from the data.
    'Used in the authenticate subroutine
    
    Dim noch As Integer
    
    noch = InStr(data, "+")
    noch = Len(data) - noch
    getpass = Right(data, noch)

End Function
Public Function checkC(ByVal data As String) As Boolean
    '
    'A subroutine which is called when the client-code type is 1.
    'This extracts the Computer name from the data using
    'remseccode() and sends the information to the
    'database for authentication and returns true or false.
    '

    
    Dim cn As ADODB.Connection, rs As ADODB.Recordset
    Dim st As String
    
    data = remseccode(data) 'Extract the Computer name
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    'Define the query
    st = "select * from users where pcname = '" & data & "'"
    
    'Open the connection.
    If cn.State > 1 Then cn.Close
    cn.Open "Provider=oraoledb.oracle.1;user id=srinath;password=srinath;"
    
    'Open the recordset.
    If rs.State > 1 Then rs.Close
    rs.Open st, cn, adOpenDynamic
    

    'Match the user's input with the information and return
    'T or F accordingly.
    
    If rs.EOF Then
        checkC = True
    Else
        checkC = False
    End If
    
    'Close the recordset and connection.
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

End Function

Public Function CreateNew(data As String)
    '
    'A subroutine which is called when the client-code type is 2.
    'This extracts all the details from the data one by one using
    'a series of sub routines and sends the information to the
    'database for storage.
    '
    Dim cn As ADODB.Connection, rs As ADODB.Recordset
    Dim st As String
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    'Open the connection
    If cn.State > 1 Then cn.Close
    cn.Open "Provider=oraoledb.oracle.1;user id=srinath;password=srinath;"
    
    'Open the recordset
    rs.Open "select * from users", cn, adOpenKeyset, adLockOptimistic
    
    data = remseccode(data) 'Remove the client-code
    Name = getname(data) 'Extract name
    age = getage(data) 'Extract age
    phone = getphone(data) 'Extract phone
    email = getemail(data) 'Extract email
    ip = getip(data) 'Extract IP
    user = getusername(data) 'Extract Username
    pass = getpassword(data) 'Extract Password
    pc = getpcname(data) 'Extract PC name
    
    rs.AddNew 'Add all the extracted info to the recordset
    rs!Username = user
    rs!Password = pass
    rs!Name = Name
    rs!age = age
    rs!phone = phone
    rs!email = email
    rs!ip = ip
    rs!pcname = pc
    rs.Update 'Update the recordset
    
    'CLose the recordset and connection.
    
    rs.Close
    cn.Close

    Set cn = Nothing
    Set rs = Nothing
      
      
End Function

Public Function getname(data As String) As String
    noch = InStr(data, "!")
    getname = Left(data, noch - 1)
End Function
Public Function getage(data As String) As Integer
    noch = InStr(data, "!")
    noch1 = InStr(data, "@")
    lenage = (noch1 - noch) - 1
    getage = Mid(data, noch + 1, lenage)
End Function

Public Function getphone(data As String) As String
    noch = InStr(data, "@")
    noch1 = InStr(data, "#")
    lenphone = (noch1 - noch) - 1
    getphone = Mid(data, noch + 1, lenphone)
    
End Function
Public Function getemail(data As String) As String
    noch = InStr(data, "#")
    noch1 = InStr(data, "$")
    lenemail = (noch1 - noch) - 1
    getemail = Mid(data, noch + 1, lenemail)
    
    
End Function
Public Function getip(data As String) As String
    noch = InStr(data, "$")
    noch1 = InStr(data, "%")
    lenip = (noch1 - noch) - 1
    getip = Mid(data, noch + 1, lenip)
End Function
Public Function getusername(data As String) As String

    noch = InStr(data, "%")
    noch1 = InStr(data, "^")
    lenuser = (noch1 - noch) - 1
    getusername = Mid(data, noch + 1, lenuser)
End Function
Public Function getpassword(data As String) As String

    noch = InStr(data, "^")
    noch1 = InStr(data, "&")
    lenpass = (noch1 - noch) - 1
    getpassword = Mid(data, noch + 1, lenpass)
End Function
Public Function getpcname(data As String) As String

    noch = InStr(data, "&")
    getpcname = Right(data, (Len(data) - noch))

End Function

