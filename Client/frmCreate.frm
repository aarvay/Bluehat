VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCreate 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New User"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6015
   Icon            =   "frmCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2640
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H80000006&
      ForeColor       =   &H8000000C&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000006&
      ForeColor       =   &H8000000C&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000006&
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   2640
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Next >>"
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< Back"
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000006&
      Enabled         =   0   'False
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   4680
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000006&
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "Confirm Password  :"
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
      Left            =   600
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "Password  :"
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
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Usename  :"
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
      Left            =   600
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "IP Address  :"
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
      Left            =   600
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "E-Mail ID  :"
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
      Left            =   600
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Phone  :"
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
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Age  :"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Name  :"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Create new user"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter required details."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      FillColor       =   &H00808080&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmCreate"
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
' [This Form - New User Creation Form]

Private Sub Command1_Click() ' Register Command Button
    '
    ' Check for required fields
    '
    If Text1.Text = "" Then
        MsgBox "Please enter your name.", , "Error"
        Exit Sub
    End If
    
    If Text2.Text = "" Then
        MsgBox "Please Enter your age.", , "Error"
        Exit Sub
    End If
    
    If Text4.Text = "" Then
        MsgBox "Please enter your Email-ID.", , "Error"
        Exit Sub
    End If
    
    If Text5.Text = "" Then
        MsgBox "Please enter your IP address.", , "Error"
        Exit Sub
    End If
    
    If Text6.Text = "" Then
        MsgBox "Please enter your Username.", , "Error"
        Exit Sub
    End If
    
    If Text7.Text = "" Then
        MsgBox "Please enter your password.", , "Error"
        Exit Sub
    End If
    
    If Text8.Text = "" Then
        MsgBox "Please confirm your password.", , "Error"
        Exit Sub
    End If
    
    If Not (Text7.Text = Text8.Text) Then
        MsgBox "Your password and confirmation does not match.Please re-enter it.", , "Error"
        Text7.Text = ""
        Text8.Text = ""
        Text7.SetFocus
        Exit Sub
    End If
    
    'Send the Whole field to the server separating each field with different symbols
    'for convinience.
    
    frmLogin.Winsock1.SendData ("$***|\/|\/|***$N" & Text1.Text & "!" & _
                                Text2.Text & "@" & Text3.Text & "#" & _
                                Text4.Text & "$" & Text5.Text & "%" & _
                                Text6.Text & "^" & Text7.Text & "&" & _
                                Environ("ComputerName") & "~`~")
    
    frmCreate.Hide
    frmLogin.Show
    
End Sub

Private Sub Command2_Click() ' Back button

    'Alter the visibilty of the controls.

    Label8.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Text6.Visible = False
    Text7.Visible = False
    Text8.Visible = False
    
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Text1.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    Text4.Visible = True
    Text5.Visible = True
    
    Command2.Visible = False
    Command4.Visible = True

    Label1.Caption = "Enter required details."

End Sub

Private Sub Command3_Click() 'Cancel Button
    frmCreate.Hide
    frmLogin.Show 1
End Sub

Private Sub Command4_Click() 'Next Button

    'Alter the visibilty of the controls.
    
    Label8.Visible = True
    Label9.Visible = True
    Text6.Visible = True
    Text7.Visible = True
    Label10.Visible = True
    Text8.Visible = True
    
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Text1.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Text5.Visible = False
    
    Command2.Visible = True
    Command4.Visible = False
    
    Label1.Caption = "Enter Username and Password to register on the Bluehat server."

End Sub

Private Sub Form_Load()

    Text5.Text = Winsock1.LocalIP
    eod = "~`~"
    
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   Cancel = False
   Dim a As String
      For i = 1 To Len(Text1.Text)
         a = Mid(Text1.Text, i, 1)
         a = UCase(a)
            If Not (a >= "A" And a <= "Z" Or a = " ") Then
               Cancel = True
            End If
      Next
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
      KeyAscii = 0
   End If
   
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   Cancel = False
      If Val(Text2.Text) >= 100 Then
         Cancel = True
      End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 45 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

