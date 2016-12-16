VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmTCPConnection 
   Caption         =   "TCP Connection"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLookupIP 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "close"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   4560
      Width           =   2055
   End
   Begin VB.PictureBox pctConnection 
      Height          =   5415
      Left            =   3960
      ScaleHeight     =   5355
      ScaleWidth      =   7515
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmdSendHTTPHeader 
         Caption         =   "Send HTTP Header"
         Height          =   495
         Left            =   5520
         TabIndex        =   19
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdCloseConnection 
         Caption         =   "Close Connection"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   4800
         Width           =   1935
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   5520
         TabIndex        =   12
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox txtReceive 
         Height          =   3375
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmTCPConnection.frx":0000
         Top             =   1320
         Width           =   5175
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   375
         Left            =   5520
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtSend 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label lblConnectionInformation 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtLocalPort 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   11040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label lblMyIPAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Local Port"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Remote Port"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "RemoteHost IP"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "frmTCPConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMessageReceived As Boolean
Private mMessageSent As Boolean
Const TIMEOUT = 10

Private Sub cmdClear_Click()
  txtReceive.Text = ""
End Sub

Private Sub cmdCloseConnection_Click()
  Winsock1.Close
  pctConnection.Visible = False
  lblStatus.Caption = ""
End Sub

Private Sub cmdConnect_Click()
  On Error GoTo cmdConnect_Error
  Dim T1 As Long
  If MousePointer = vbHourglass Then Exit Sub
  
  MousePointer = vbHourglass
  If Winsock1.State <> sckClosed Then Winsock1.Close
  
  'Winsock1.LocalPort = txtLocalPort.Text
  Winsock1.RemoteHost = txtRemoteHost.Text
  Winsock1.RemotePort = txtRemotePort.Text
  
  Winsock1.Connect
  lblStatus.Caption = ""
  
  T1 = Timer
  Do
    DoEvents
  Loop Until Winsock1.State = sckConnected Or Winsock1.State = sckError Or Timer - T1 > TIMEOUT
  
  If Winsock1.State = sckConnected Then
    pctConnection.Visible = True
  Else
    pctConnection.Visible = False
  End If
  
cmdConnect_End:
  MousePointer = vbDefault
  Exit Sub
cmdConnect_Error:
  Call MsgBox(Err.Description)
  Resume cmdConnect_End
  Resume Next
End Sub

Private Sub cmdListen_Click()
  On Error GoTo cmdListen_Error
  Dim T1 As Long
  If MousePointer = vbHourglass Then Exit Sub
  
  'don't know why i had to put this here...
  If Winsock1.State = sckClosing Then
    Winsock1.Close
  End If
  
  If Winsock1.State <> sckClosed Then
    Call MsgBox("Cannot listen while a connection already exists.")
    Exit Sub
  End If
  MousePointer = vbHourglass
  
  Winsock1.LocalPort = txtLocalPort.Text
  Winsock1.Listen
  T1 = Timer
  Do
    DoEvents
  Loop Until Winsock1.State = sckListening Or Winsock1.State = sckError Or Timer - T1 > TIMEOUT
  
  If Winsock1.State = sckListening Then
    lblStatus.Caption = "I am listening on port " & Winsock1.LocalPort
  Else
    lblStatus.Caption = ""
  End If
cmdListen_End:
  MousePointer = vbDefault
  Exit Sub
cmdListen_Error:
  Call MsgBox(Err.Description)
  Resume cmdListen_End
  Resume Next
End Sub

Private Sub cmdLookupIP_Click()
  Dim Answer As String
  
  Answer = InputBox("lwhat is the DNS name to resolve?", "ex: www.?????.com")
  If Answer <> "" Then
    Dim cResolve As DNSResolve
    Set cResolve = New DNSResolve
    txtRemoteHost.Text = cResolve.GetIPFromHostName(Answer)
  End If
End Sub

Private Sub cmdSend_Click()
  Dim T1 As Long
  If MousePointer = vbHourglass Then Exit Sub
  MousePointer = vbHourglass
  Winsock1.SendData txtSend.Text & vbCrLf
  mMessageSent = False

  T1 = Timer
  Do
    DoEvents
  Loop Until mMessageSent Or Timer - T1 > TIMEOUT
  MousePointer = vbDefault

  If mMessageSent Then
    Call AddMessage(Winsock1.LocalIP, txtSend.Text)
    txtSend.Text = ""
  End If
  
End Sub

Private Sub cmdSendHTTPHeader_Click()
  Dim T1 As Long
  Dim Header As String
  
  If MousePointer = vbHourglass Then Exit Sub
  MousePointer = vbHourglass
  
  Header = "GET / HTTP/1.1" & vbCrLf
  Header = Header & "Host: 192.168.1.2:80" & vbCrLf
  Header = Header & "Connection: keep -alive" & vbCrLf
  Header = Header & "Cache -Control: Max -age = 0" & vbCrLf
  Header = Header & "Upgrade-Insecure-Requests: 1" & vbCrLf
  Header = Header & "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.71 Safari/537.36" & vbCrLf
  Header = Header & "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8" & vbCrLf
  Header = Header & "Accept -Encoding: gzip , deflate, sdch" & vbCrLf
  Header = Header & "Accept-Language: en-US,en;q=0.8" & vbCrLf
  
  Winsock1.SendData Header
  mMessageSent = False

  T1 = Timer
  Do
    DoEvents
  Loop Until mMessageSent Or Timer - T1 > TIMEOUT
  MousePointer = vbDefault

  If mMessageSent Then
    Call AddMessage(Winsock1.LocalIP, txtSend.Text)
    txtSend.Text = ""
  End If
End Sub

Private Sub Command1_Click()
  Winsock1.Close
  lblStatus.Caption = ""
  pctConnection.Visible = False
End Sub

Private Sub Form_Load()
  lblMyIPAddress.Caption = "My IP Address: " & Winsock1.LocalIP
  
  txtRemoteHost.Text = Winsock1.LocalIP
End Sub

Private Sub txtSend_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    cmdSend_Click
  End If
End Sub

Private Sub Winsock1_Close()
  pctConnection.Visible = False
  lblStatus.Caption = ""
End Sub

Private Sub Winsock1_Connect()
  txtSend.Text = ""
  txtReceive.Text = ""
  pctConnection.Visible = True
  lblStatus.Caption = "Connected to " & Winsock1.RemoteHostIP
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
  ' Accept the request with the requestID parameter.
  If Winsock1.State = sckListening Then
    Winsock1.Close
    Winsock1.Accept requestID
    lblConnectionInformation.Caption = "RemoteHostIP: " & Winsock1.RemoteHostIP & " RemotePort: " & Winsock1.RemotePort
    txtSend.Text = ""
    txtReceive.Text = ""
    pctConnection.Visible = True
    lblStatus.Caption = "Connected to " & Winsock1.RemoteHostIP
    
  Else
    Call MsgBox("i just got a request for a connection when I am not even listening.")
  End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  Dim SomeText As String
  Winsock1.GetData SomeText
  mMessageReceived = True
  Call AddMessage(Winsock1.RemoteHostIP, SomeText)
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Call MsgBox(Description, , Number)
  Winsock1.Close
  pctConnection.Visible = False
  lblStatus.Caption = ""
End Sub

Private Sub Winsock1_SendComplete()
  mMessageSent = True
End Sub

Private Sub AddMessage(From As String, SomeText As String)
  If Len(txtReceive.Text) > 0 Then
    txtReceive.Text = txtReceive.Text & vbCrLf
  End If
  txtReceive.Text = txtReceive.Text & From & ": " & SomeText
End Sub

