VERSION 5.00
Begin VB.Form frmIP 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter IP"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   480
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your server's IP Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Enter here  : "
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
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmIP"
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
' [This Form - Connection Form]

Private Sub Command1_Click() 'Connect Buttin
    
    'Load frmlogin bcoz we are using the winsock located in that form
    'Assign values for RemotePort and RemoteHost and Connect to the server.
    
    Load frmLogin
    
    'Check for the field
    If Text1.Text = "" Then
        MsgBox "Please enter server's IP", vbInformation, "Enter"
        Exit Sub
    End If
    
    frmLogin.Winsock1.RemotePort = 1234
    frmLogin.Winsock1.RemoteHost = Text1.Text
    frmLogin.Winsock1.Connect
    Timer1.Enabled = True
    Timer2.Enabled = True
    Command1.Enabled = False

End Sub

Private Sub Form_Load()

    'Center the Form
    
    With frmIP
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    
    eod = "~`~"
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Load frmLogin
        If Text1.Text = "" Then
            MsgBox "Please enter server's IP", vbInformation, "Enter"
            Exit Sub
        End If
        frmLogin.Winsock1.RemotePort = 1234
        frmLogin.Winsock1.RemoteHost = Text1.Text
        frmLogin.Winsock1.Connect
        Timer1.Enabled = True
        Timer2.Enabled = True
        Command1.Enabled = False
    End If
    
End Sub

Private Sub Timer1_Timer() 'Timer for sending the computer's name.
    If frmLogin.Winsock1.State = sckConnected Then
        frmLogin.Winsock1.SendData "$***|\/|\/|***$C" & Environ("ComputerName") & "~`~"
    End If
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer() 'Connection timeout check.

    'Wait for 5 sec and if winsock's state <> 7 then flag an error.

    If frmLogin.Winsock1.State <> 7 Then
        MsgBox "Connection timeout. Server does not exist", vbCritical, "ERROR"
        frmLogin.Winsock1.Close
        Command1.Enabled = True
        Timer2.Enabled = False
    End If
End Sub
