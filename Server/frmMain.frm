VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H80000007&
   Caption         =   "MDIForm1"
   ClientHeight    =   3105
   ClientLeft      =   2685
   ClientTop       =   2055
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   5640
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2610
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2593
            Text            =   "WELCOME TO BLUEHAT TCP NETWORK SERVER"
            TextSave        =   "WELCOME TO BLUEHAT TCP NETWORK SERVER"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/11/2009"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:17 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   2
      Top             =   1605
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1535
      ButtonWidth     =   1958
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User Manager"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox Shape1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   1605
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12720
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "Computer Name  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   11760
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12720
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   "Server IP  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11760
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "BLUEHAT TCP NETWORK SERVER "
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         Caption         =   "Copyright - Vignesh Rajagopalan and Srinath PB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   11760
         TabIndex        =   1
         Top             =   1080
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                              Bluehat TCP/IP Server                                          *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                          v1.0 - (Unstable Version)                                        *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                  VIGNESH RAJAGOPALAN and Srinath PB                           *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                              (Suggestions welcome)                                          *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*                     vignesh.rajagopalan@yahoo.com                                  *~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
'*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
' [This Form - Main Form - MDI]


Private eod As String 'Eod mark
Private ind As Integer 'Winsock array count
Private Buffer() As String 'Array of buffer
Dim auth As Boolean, message As String, constat As Boolean
Dim listind As Integer 'index of the listbox

Private Sub MDIForm_Load()
    With frmMain
        .Height = Screen.Height
        .Width = Screen.Width
        .Top = -60
        .Left = 0
    End With
    
    'Fill the necessary labels.
    Label4.Caption = Winsock1(0).LocalIP
    Label6.Caption = Environ("ComputerName")
    
    'Keep the port 1234 listening for the clients.
    Winsock1(0).LocalPort = 1234
    Winsock1(0).Listen
    eod = "~`~"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        frmUsers.Show
    End If
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    '
    'The winsock(0) is the control that is only listening and only accepts
    'the connections.
    '
    'Check if the Index is 0
    If Index = 0 Then
        
        'check all the existing winsock's state
        'If closed then accept the connection in this itself and flush its buffer and
        'client information.
        
        For i = 1 To ind
            If Winsock1(i).State = sckClosed Then
                Winsock1(i).Accept requestID
                
                Buffer(i) = ""

                Exit Sub
            End If
        Next
        
        'Did not find any winsock that is closed. so increment the ind by 1.
        ind = ind + 1
        
        'Load the Winsock
        Load Winsock1(ind)
        
        'Create a new buffer for the new winsock with the same index number.
        ReDim Preserve Buffer(ind)
        

        
        'Accept the connection.
        Winsock1(ind).Accept requestID
        

    
    End If
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Record As String 'Net data recieved
    Dim EORPos As Long 'Pos of eod mark
    Dim Finished As Boolean '..finished processing or not?

    'Get the message in a str Variable.
    Winsock1(Index).GetData message
    
    'Append the data to the particular client's buffer.
    Buffer(Index) = Buffer(Index) & message
    
    Do
        'Check wheather the buffer contains the eod mark
        EORPos = InStr(Buffer(Index), eod)
        
        'If so then...
        If EORPos > 0 Then
            
            'remove the eod mark and store it in the record.
            Record = Mid(Buffer(Index), 1, EORPos - 1)
            
                            
                            '******PROCESSING PART*******'
            
            
            If datapassthru(Record) = 0 Then
                auth = authenticate(Record)
                
                If auth = False Then
                    Winsock1(Index).SendData "***|\/|***P" & "0 " & "Invalid Username/Password. Connection Forbidden!" & eod
                Else
                    Winsock1(Index).SendData "***|\/|***P" & "1 " & "Connected to " & Winsock1(0).LocalIP & eod
                End If
                   
            ElseIf datapassthru(Record) = 1 Then
                If checkC(Record) = True Then
                    Winsock1(Index).SendData "***|\/|***C" & "0" & eod
                Else
                    Winsock1(Index).SendData "***|\/|***C" & "1" & eod
                End If
            ElseIf datapassthru(Record) = 2 Then
                CreateNew Record
                Winsock1(Index).SendData "***|\/|***N" & "Your ID has been succesfully registered in the Bluehat server." & vbCrLf & _
                                                      "You can now login with your username and password." & eod
            ElseIf Record = "*.*" Then
                Winsock1(Index).Close
                Buffer(Index) = ""

            Else
                For i = 1 To ind
                    If Winsock1(i).State = sckConnected Then
                        Winsock1(i).SendData Record & eod
                    End If
                Next
            End If
            
            
                            '********END OF PROCESSING PART*******'
            
            'Check wheather if there is any remaining data
            If EORPos + Len(eod) < Len(Buffer(Index)) Then
                
                'then move it to the front of the buffer and loop again.
                Buffer(Index) = Mid(Buffer(Index), EORPos + Len(eod))
            
            Else '..no remaining data.
                            
                'then flush the buffer and exit the loop.
                Buffer(Index) = ""
                Finished = True
                
            End If
        
        Else 'You did not find any eor mark
            
            'Exit the loop and wait for the remaining data.
            Finished = True
            
        End If
    
    Loop Until Finished = True 'End loop when finished = true

End Sub


'****************************************************************************
'*************** Logic Behind my Winsock Buffer Management **********************
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
