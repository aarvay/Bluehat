Attribute VB_Name = "Module1"
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                Bluehat TCP/IP Client              *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*              v1.0 - (Unstable Version)            *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*          VIGNESH RAJAGOPALAN and Srinath PB       *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*               (Suggestions welcome)               *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*            vignesh.rajagopalan@yahoo.com          *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
' [Module1 - Sub routines.]

Public Function datapassthru(ByVal data As String) As Integer
    
    'A subroutine to check the nature of the server-code which will return
    'an integer corresponding to the type of the server-code.
    
    data = Left(data, 11) 'Extract the code from the data.
    
    ' Check for its type.
    
    If data = "***|\/|***P" Then
        datapassthru = 0
    ElseIf data = "***|\/|***C" Then
        datapassthru = 1
    ElseIf data = "***|\/|***N" Then
        datapassthru = 2
    ElseIf data = "***|\/|***R" Then
        datapassthru = 3
    ElseIf data = "***|\/|***O" Then
        datapassthru = 4
    ElseIf data = "***|\/|***D" Then
        datapassthru = 5
    ElseIf data = "***|\/|***F" Then
        datapassthru = 6
    ElseIf data = "***|\/|***M" Then
        datapassthru = 7
    Else    'no code, return -1
        datapassthru = -1
    End If
    
End Function
Public Function reqresult(data As String) As Integer

    'A subroutine called when the server code type is 0 or 1 which will
    'return the result(integer) 1 - true and 2 - false.

    i = Mid(data, 12, 1)
    If i >= 0 And i <= 5 Then
        reqresult = i
    Else
        reqresult = 9
    End If
    
End Function

Public Function givemess(data As String) As String

    'A subroutine which will extract the logical message from the data arrived.

    data = Right(data, (Len(data) - 13))
    givemess = data
    
End Function


