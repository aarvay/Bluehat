VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Management"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9030
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9030
   Begin VB.TextBox Label13 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Label10 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   1111
      ButtonWidth     =   1693
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "See Users"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User Details"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2580
      Left            =   2640
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label9 
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
      Left            =   2640
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "E-mail  ID  :"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
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
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
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
      Left            =   2640
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
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
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "See registered users, and their details..!!!"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "User Manager"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Registered Users"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                                  Bluehat TCP/IP Server                                                     *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                                v1.0 - (Unstable Version)            *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                      VIGNESH RAJAGOPALAN and Srinath PB       *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                                  (Suggestions welcome)               *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                          vignesh.rajagopalan@yahoo.com          *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
' [This Form - User Manager Form]

Dim cn As ADODB.Connection, rs As ADODB.Recordset
Private Sub Form_Load()
    With frmUsers
        .Height = 6255
        .Width = 9120
    End With
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    If cn.State > 0 Then cn.Close
    cn.Open "provider =oraoledb.oracle.1;user id = srinath;password = srinath;"
    If rs.State > 0 Then rs.Close
    rs.Open "select * from users", cn, adOpenDynamic, adLockOptimistic
    
    rs.MoveFirst
    Do While Not rs.EOF
        List2.AddItem rs!Username
        rs.MoveNext
    Loop
        
End Sub

Private Sub List2_Click()
    If cn.State > 0 Then cn.Close
    cn.Open "provider = oraoledb.oracle.1;user id = srinath;password = srinath;"
    If rs.State > 0 Then rs.Close
    rs.Open "select * from users where username ='" & List2.Text & "'", cn, adOpenDynamic, adLockOptimistic
    Label10.Text = rs!Name
    Label11.Caption = rs!age
    Label12.Caption = rs!phone
    Label13.Text = rs!email
    Label14.Caption = rs!ip
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
    
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False
    
        Label2.Visible = True

        List2.Visible = True

    End If
    
    If Button.Index = 2 Then

        Label2.Visible = False

        List2.Visible = False

        
        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = True
        Label12.Visible = True
        Label13.Visible = True
        Label14.Visible = True
    End If
    
End Sub
