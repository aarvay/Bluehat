VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Login"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5670
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2640
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5640
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Status  :"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5640
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Password  :"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Username  :"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Log-in to the Bluehat server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   3840
      Picture         =   "frmLogin.frx":0E42
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                Bluehat TCP/IP Client              *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*              v1.0 - (Unstable Version)            *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*          VIGNESH RAJAGOPALAN and Srinath PB       *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*               (Suggestions welcome)               *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*            vignesh.rajagopalan@yahoo.com          *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
' [This Form - Login Form]

Dim message As String
Private eod As String, buffer As String

Private Sub Command2_Click() 'Login Button
    '
    'Check for the necessary fields.
    '
    If Text1.Text = "" Then
        MsgBox "Please enter your username.", , "Error"
        Exit Sub
    End If
    If Text2.Text = "" Then
        MsgBox "Please enter your password.", , "Error"
        Exit Sub
    End If
    
    Label5.Caption = ""
    
    'Send the User's input(username and password) to the server for authentication.
    If Winsock1.State = sckConnected Then
        Winsock1.SendData ("$***|\/|\/|***$P" & Text1.Text & "+" & Text2.Text & "~`~")
    End If
    
End Sub

Private Sub Command3_Click() 'Cancel Button
    '
    'If connected send to the server a signal of disconnection.
    '
    If frmLogin.Winsock1.State = 7 Then
        frmLogin.Winsock1.SendData ("*.*" & "~`~")
        DoEvents
        End
    End If
End Sub

Private Sub Form_Load()

    'Center the Form.
    
    With frmLogin
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    
    eod = "~`~"
End Sub

Private Sub Form_Terminate()

    'If connected send to the server a signal of disconnection.

    If frmLogin.Winsock1.State = 7 Then
        frmLogin.Winsock1.SendData ("*.*" & "~`~")
        DoEvents
        End
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text1.Text = "" Then
            MsgBox "Please enter your username.", , "Error"
            Exit Sub
        End If
        If Text2.Text = "" Then
            MsgBox "Please enter your password.", , "Error"
            Exit Sub
        End If
    
    Label5.Caption = ""
        If Winsock1.State = sckConnected Then
            Winsock1.SendData ("$***|\/|\/|***$P" & Text1.Text & "+" & Text2.Text & "~`~")
        End If
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text1.Text = "" Then
            MsgBox "Please enter your username.", , "Error"
            Exit Sub
        End If
        If Text2.Text = "" Then
            MsgBox "Please enter your password.", , "Error"
            Exit Sub
        End If
    
    Label5.Caption = ""
        If Winsock1.State = sckConnected Then
            Winsock1.SendData ("$***|\/|\/|***$P" & Text1.Text & "+" & Text2.Text & "~`~")
        End If
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim Record As String 'Net message recieved.
    Dim EORPos As Long 'Pos of the eod mark.
    Dim Finished As Boolean ' ..processing finished or not ?

    'Get the data into a variable - message.
    Winsock1.GetData message
    
    'Append the message to the buffer.
    buffer = buffer & message
    
    Do
        'Check for the position of the eod mark
        EORPos = InStr(buffer, eod)
        
        If EORPos > 0 Then 'If eod is present then consider it a logical record.
            
            'Remove the eod mark and store it in record.
            Record = Mid(buffer, 1, EORPos - 1)
    
                            '********PROCESSING PART*********'
            
            If datapassthru(Record) = 0 Then
                If reqresult(Record) = 0 Then
                    Label5.Caption = givemess(Record)
                ElseIf reqresult(Record) = 1 Then
                    frmLogin.Hide
                    MsgBox "Connected to " & Winsock1.RemoteHost
                    Load Form1
                    Form1.Show
                End If
            
            ElseIf datapassthru(Record) = 1 Then
                If reqresult(Record) = 0 Then
                    frmCreate.Show
                    Unload frmIP
                ElseIf reqresult(Record) = 1 Then
                    frmLogin.Show
                    Unload frmIP
                End If
            
            ElseIf datapassthru(Record) = 2 Then
                MsgBox givemess(Record)
            Else
                
                pos = InStr(Record, ":")
                username = Mid(Record, 1, pos - 2)
                
                If username = Text1.Text Then
                    Form1.Text1.Text = Form1.Text1.Text & Record & vbCrLf
    
                    With Form1.Text1
                        .SelStart = .Find(username)
                        .SelLength = Len(username)
                        .SelBold = True
                    End With

                Else
                    Form1.Text1.Text = Form1.Text1.Text & Record & vbCrLf
                    
                    With Form1.Text1
                        .SelStart = .Find(username)
                        .SelLength = Len(username)
                        .SelColor = vbBlue
                        .SelBold = True
                    End With
                End If
            End If
                        
                         '******END OF PROCESSING PART*****'
                        
            'Check if there is anything left in the management.
            If EORPos + Len(eod) < Len(buffer) Then
                '
                'If so move it to the front and continue the loop.
                buffer = Mid(buffer, EORPos + Len(eod))
            
            Else 'If all message has been processed then flush the buffer and exit
                 'the loop.
                 
                buffer = ""
                Finished = True
                
            End If
        
        Else '..you didn't find the eor mark. Exit the loop and wait for the
             'remaining message.
            
            Finished = True
            
        End If
        
    Loop Until Finished = True 'Exit when finished = true

End Sub

'****************************************************************************
'***************Logic Behind my Winsock Buffer Management********************
'****************************************************************************
'1. Get the message and append it to the Buffer.
'2. Look for an Eor mark.(will tell the server that the whole data has been
'   recieved). If so, process the data.
'3. After processing check if there is anything in the buffer. If so then
'   move the remaining data to the front and look again(loop again)
'   for a complete record.
'4. If you did not find any remaining data in the buffer, flush it and exit
'   the loop.
'5. If you did no find an eor mark in the buffer then exit the loop then
'   wait for the remaining message.
'****************************************************************************
