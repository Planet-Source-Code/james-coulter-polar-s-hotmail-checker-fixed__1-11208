VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmhotmail 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotmail Messages"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmhotmail.frx":0000
   ScaleHeight     =   2310
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin SocketWrenchCtrl.Socket Socket 
      Left            =   0
      Top             =   2160
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton cmdconnect 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sign-In"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6495
      TabIndex        =   4
      Top             =   1080
      Width           =   1065
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4425
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1095
      Width           =   1980
   End
   Begin VB.TextBox txtlogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4425
      TabIndex        =   1
      Top             =   495
      Width           =   1980
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4305
      X2              =   8160
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Label lblhotmail 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "@ hotmail.com"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6465
      TabIndex        =   6
      Top             =   495
      Width           =   1695
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Sign-In and Password."
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4125
      TabIndex        =   5
      Top             =   1965
      Width           =   4290
   End
   Begin VB.Label lblpassword 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4425
      TabIndex        =   2
      Top             =   855
      Width           =   1215
   End
   Begin VB.Label lblsignin 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sign-In Name: "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4425
      TabIndex        =   0
      Top             =   255
      Width           =   1545
   End
End
Attribute VB_Name = "frmhotmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'Initialize Socket
Socket.AddressFamily = AF_INET
Socket.Binary = False
Socket.Blocking = False
Socket.BufferSize = 5000
Socket.Protocol = IPPROTO_IP
Socket.SocketType = SOCK_STREAM
Socket.RemotePort = 80
    
End Sub

Private Sub SOCKET_CONNECT()
Dim str As String ' holds data to be sent to server

Select Case BatchNumber
Case 0
lbl.Caption = "2. Sending Login Data..."
str$ = MakeString(0) ' make first batch of data to send
Case 1
lbl.Caption = "4. Requesting Mailbox..."
str$ = MakeString(1) ' make second batch of data
Case 2
    str = MakeString(2)
End Select

' send data to server
Socket.SendLen = Len(str$)
Socket.SendData = str$
End Sub

Private Sub Socket_Read(DataLength As Integer, IsUrgent As Integer)

Dim NewData As String ' holds the data we receive from hotmail server

Socket.RecvLen = DataLength
NewData = Socket.RecvData ' get data

Select Case BatchNumber ' depending on which batch of data we receive, we will do different actions
Case 0
    If InStr(1, NewData, "Location:") <> 0 Then ' in first batch, if login and password is correct, server directs you to a new server and new url
        Dim temp As String
        temp$ = Right(NewData, Len(NewData) - Len("Location: "))
        temp$ = Left(temp, Len(temp) - 2)
        NewHost = Mid(temp, 8, (Len(temp) - 8) - (Len(temp) - InStr(8, temp, "/"))) ' here we get the new server address
        NewUrl = Right(temp, Len(temp) - InStr(8, temp, "/")) ' and here we get the new url to request
        BatchNumber = 1
        lbl.Caption = "3. Finding Mailbox Server..."
        ' disconnect and reconnect to new server to send data
        Socket.Action = SOCKET_DISCONNECT
        Socket.HostName = NewHost
        Socket.Action = 2 ' once we connect, we'll request the new page (NewUrl)
    End If
    If InStr(1, NewData, "reauthhead.asp") <> 0 Then
        Socket.Action = SOCKET_DISCONNECT
        lbl.Caption = "Error: Invalid Login or Password"
        Call ResetAll
    End If
Case 1
    If InStr(1, NewData, "Set-Cookie:") <> 0 Then ' now that we've succesfully sent the correct data to the new server, it sends cookies to be re-sent when we request the mailbox
        Cookies(CurrentCookie) = Mid(NewData, InStr(1, NewData, "Set-Cookie:") + 12, Len(NewData) - (InStr(1, NewData, "Set-Cookie:") + 12) - (Len(NewData) - InStr(1, NewData, ";"))) ' store cookies in array
        CurrentCookie = CurrentCookie + 1
    End If
    If InStr(1, NewData, "Refresh") <> 0 Then 'after the server sends all the cookies, it tell us to refresh to the actual mailbox, therefore demanding us to send the cookies back to the server
        NewUrl = Mid(NewData, InStr(1, NewData, "content=") + 16, Len(NewData) - (InStr(1, NewData, "content=") + 16) - 3) ' the url of the final mailbox
        'BatchNumber = 2
        'Dim str As String
        'str$ = MakeString(2) ' compile the final data to be sent, containing the url of the mailbox, and all of the cookies received
        ' now send the data
        'Socket.SendLen = Len(str$)
        'Socket.SendData = str$ ' send final data
    End If
    If InStr(1, NewData, "</html>") <> 0 Then
        BatchNumber = 2
        Socket.Action = SOCKET_DISCONNECT
        Socket.Action = 2
    End If
Case 2 ' if all correct data was send correctly, on the third time we begin to receive the mailbox data
    lbl.Caption = "5. Processing Mailbox..."
    If InStr(1, NewData, "title.asp") <> 0 Then ' here is where the number of new messages can be read
        ReadBox = True ' begin storing incoming batches (pages) into the variable 'BoxData'
        BoxBatch = 0 ' we will only store 10 batches of data, as that is all we need to find the new messages
        MailData = NewData
    End If
    If ReadBox = True Then
        BoxBatch = BoxBatch + 1
        MailData = MailData & NewData
        If BoxBatch = 10 Then GoTo 1
    Exit Sub
1: ' we now have all the crucial mailbox source stored, and we are ready to extract the number of new messages from it.
   ' By storing more batches, you can also extract other information that you want. this is just shown as an example.
    Socket.Action = SOCKET_DISCONNECT
    Dim NewMessages As String
    Dim Location As Integer, Offset As Integer, Length As Integer
    Location = InStr(1, MailData, "new")
    Offset = InStr(Location - 5, MailData, ">") + 1
    Length = Location - (Location - Offset)
    NewMessages = Mid(MailData, Length, Len(MailData) - Offset - (Len(MailData) - Location) - 1) ' store value of new messages
    If Int(NewMessages) = 1 Then
    lbl.Caption = "You have: " & NewMessages & " new message."
    Else
    lbl.Caption = "You have: " & NewMessages & " new messages." ' whew! all done! a lot of work for such little information, huh?
    End If
    Call ResetAll
    ' Because you have accessed the entire source of the mailbox, to retrieve other useful information such as who sent you what, the subject line of the new mail,
    ' and even check any messages in your mailbox, all you need to do is store more batches of the data, and find this information from the source, or find the url containing
    ' the new messages, and request for that page as shown in the function MakeString(). If you have any questions, comments, or suggestions, please feel free to
    ' email me at:   nmjblue@hotmail.com
    End If
End Select
End Sub

Private Sub cmdconnect_Click()
cmdconnect.Enabled = False
Call ConnectToHotmail ' begin connection
End Sub

Private Sub ConnectToHotmail()
lbl.Caption = "1. Connecting to Hotmail..."
StrLogin$ = Trim$(txtlogin.Text)
StrPass$ = Trim$(txtpass.Text)
Socket.HostName = "lc5.law5.hotmail.passport.com"
Socket.Action = 2
End Sub

Private Sub ResetAll()
On Error Resume Next
Socket.Action = SOCKET_DISCONNECT
BatchNumber = 0
cmdconnect.Enabled = True
For i = 0 To 5
Cookies(i) = ""
Next
CurrentCookie = 0
ReadBox = False
MailData = ""
NewHost = ""
NewUrl = ""
End Sub
