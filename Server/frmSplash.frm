VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3360
      X2              =   6240
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      FillColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright --------- Vignesh Rajagopalan and Srinath PB"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "BLUEHAT TCP NETWORK SERVER"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   975
      Left            =   3360
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   360
      Picture         =   "frmSplash.frx":0E42
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                Bluehat TCP/IP Server              *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*              v1.0 - (Unstable Version)            *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*          VIGNESH RAJAGOPALAN and Srinath PB       *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*               (Suggestions welcome)               *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*            vignesh.rajagopalan@yahoo.com          *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
' [This Form - Splash screen]

Private Sub Form_Load()

    'Center the screen

    With frmSplash
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    
End Sub

Private Sub Timer1_Timer()

    Static cnt As Integer 'Define an interval
    
    Select Case cnt '....on each interval...
        
        Case 0
            Image1.Visible = True
        Case 1
            Label1.Visible = True
        Case 2
            Line1.Visible = True
        Case 3
            Label2.Visible = True
            Timer1.Interval = 2000
        Case 4
            
        Case 5
            Timer1.Enabled = False
            Unload frmSplash
            Load frmMain
            frmMain.Show
    End Select
    
     cnt = cnt + 1
    
End Sub
