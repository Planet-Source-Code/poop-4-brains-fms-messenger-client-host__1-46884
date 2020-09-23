VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "FMS Server"
   ClientHeight    =   5280
   ClientLeft      =   645
   ClientTop       =   1215
   ClientWidth     =   9855
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   9855
   Begin VB.ListBox lstChat 
      Height          =   2205
      Left            =   2760
      TabIndex        =   8
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox lstUB 
      Height          =   2205
      Left            =   7920
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox lstP 
      Height          =   2205
      Left            =   6840
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox lstU 
      Height          =   2205
      Left            =   5520
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtLog 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   9615
   End
   Begin MSWinsockLib.Winsock wsServer 
      Left            =   1440
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5333
   End
   Begin MSWinsockLib.Winsock wsConnect 
      Index           =   0
      Left            =   1440
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstUsers 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblChat 
      BackStyle       =   0  'Transparent
      Caption         =   "Chatrooms"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblReg 
      BackStyle       =   0  'Transparent
      Caption         =   "Registered Users/Passwords/Buddy Lists (raw)"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblLog 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Log"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label lblUsers 
      BackStyle       =   0  'Transparent
      Caption         =   "Online Users"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu mnuServer 
      Caption         =   "Server"
      Begin VB.Menu mnuStart 
         Caption         =   "Start/Stop Server"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide Window"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "Users"
      Begin VB.Menu mnuKick 
         Caption         =   "Kick"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim buf1, buf2, buf3

lstU.Clear
lstP.Clear
lstUB.Clear
Open App.Path & "\ulist.dat" For Input As #1
Do Until EOF(1)
    Input #1, buf1, buf2, buf3
    lstU.AddItem buf1
    lstP.AddItem buf2
    lstUB.AddItem buf3
Loop
Close #1
lstChat.Clear
Open App.Path & "\clist.dat" For Input As #1
Do Until EOF(1)
    Input #1, buf1, buf2
    MakeChat buf1, buf2, "SERVER"
Loop
Close #1
mnuStart_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim I As Long

Kill App.Path & "/ulist.dat"

Open App.Path & "/ulist.dat" For Output As #1
For I = 0 To lstU.ListCount - 1
    Write #1, lstU.List(I), lstP.List(I), lstUB.List(I)
Next I
Close #1

Kill App.Path & "/clist.dat"

Open App.Path & "/clist.dat" For Output As #1
For I = 0 To lstChat.ListCount - 1
    Write #1, lstChat.List(I), Val(Rm(lstChat.ItemData(I)).lblMax)
Next I
Close #1

For I = 0 To 19
    Unload Rm(I)
Next I

If wsServer.State <> sckClosed Then
    wsServer.Close
    
    If lstUsers.ListCount = 0 Then Exit Sub
    For I = 0 To lstUsers.ListCount
        If wsConnect(lstUsers.ItemData(I)).State <> sckClosed Then
            wsConnect(lstUsers.ItemData(I)).Close
        End If
    Next I
End If
End Sub

Function AddBuddy(user, auser)
Dim I As Long, ia As Long
If CheckExist(auser) = False Then
    wsConnect(FindUserPort(user)).SendData "ERR|Ñ|User does not exist"
    Exit Function
End If
For I = 0 To lstU.ListCount - 1
    If lstU.List(I) = user Then
        For ia = 0 To UBound(Split(lstUB.List(I), "|¿|"))
            If Split(lstUB.List(I), "|¿|")(ia) = auser Then
                wsConnect(FindUserPort(user)).SendData "ERR|Ñ|User already on buddy list"
                Exit Function
            End If
        Next ia
        lstUB.List(I) = lstUB.List(I) & auser & "|¿|"
        Exit For
    End If
Next I
GetBuddy user
End Function

Function GetBuddy(user)
Dim I As Long, bdlist
For I = 0 To lstU.ListCount - 1
    If lstU.List(I) = user Then
        wsConnect(FindUserPort(user)).SendData "BUD|Ñ|" & lstUB.List(I)
        Exit For
    End If
Next I
End Function

Function RemoveBuddy(user, ruser)
Dim I As Long, ia As Long
For I = 0 To lstU.ListCount - 1
    If lstU.List(I) = user Then
        For ia = 0 To UBound(Split(lstUB.List(I), "|¿|"))
            If Split(lstUB.List(I), "|¿|")(ia) = ruser Then
                lstUB.List(I) = Replace(lstUB.List(I), ruser, "")
                Exit Function
            End If
        Next ia
        Exit For
    End If
Next I
GetBuddy user
End Function

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHide_Click()
Me.Hide
End Sub

Private Sub mnuStart_Click()
If wsServer.State = sckListening Then
    wsServer.Close
Else
    wsServer.Listen
End If
End Sub

Function LoginVerify(user, pass)
Dim I As Long

LoginVerify = False

For I = 0 To lstU.ListCount - 1
    If lstU.List(I) = user And lstP.List(I) = pass Then
        LoginVerify = True
    End If
Next I
End Function

Private Sub wsConnect_Close(Index As Integer)
Dim I As Long, tuser As String
tuser = FindPortUser(Index)
If tuser = "" Then GoTo Ok

For I = 0 To lstUsers.ListCount - 1
    If lstUsers.List(I) = tuser Then
        BroadCast "OFF|Ñ|" & lstUsers.List(I)
        lstUsers.RemoveItem I
        Exit For
    End If
Next I

Ok:

Unload wsConnect(Index)
End Sub

Private Sub wsConnect_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, I

wsConnect(Index).GetData Data

txtLog.Text = txtLog.Text & Data & vbCrLf
txtLog.SelStart = Len(txtLog.Text)

Select Case Split(Data, "|Ñ|")(0)
Case "NUSER" 'create new user
If CheckExist(Split(Data, "|Ñ|")(1)) = True Then
    wsConnect(Index).SendData "ERR|Ñ|This username is already in use"
    Exit Sub
End If
lstU.AddItem Split(Data, "|Ñ|")(1)
lstP.AddItem Split(Data, "|Ñ|")(2)
lstUB.AddItem ""
wsConnect(Index).SendData "USERCREATE"
Case "LOGIN" 'user login
If LoginVerify(Split(Data, "|Ñ|")(1), Split(Data, "|Ñ|")(2)) = True Then
    BroadCast "ON|Ñ|" & Split(Data, "|Ñ|")(1)
    DoEvents
    lstUsers.AddItem Split(Data, "|Ñ|")(1)
    lstUsers.ItemData(lstUsers.NewIndex) = Index
    wsConnect(Index).SendData "LOGINCONFIRM|Ñ|" & Split(Data, "|Ñ|")(1)
Else
    wsConnect(Index).SendData "ERR|Ñ|User does not exist or your password is incorrect"
End If
Case "PMSG" 'private message
wsConnect(FindUserPort(Split(Data, "|Ñ|")(1))).SendData "PMSG|Ñ|" & FindPortUser(Index) & "|Ñ|" & Split(Data, "|Ñ|")(2)
Case "ADDB" 'add buddy
AddBuddy FindPortUser(Index), Split(Data, "|Ñ|")(1)
Case "REMB" 'remove buddy
AddBuddy FindPortUser(Index), Split(Data, "|Ñ|")(1)
Case "BUD" 'request online users
GetBuddy FindPortUser(Index)
Case "RON" 'request online users
GetOnline FindPortUser(Index)
Case "GETIP" 'request ip address of a user
Dim gIP As String
gIP = wsConnect(FindUserPort(Split(Data, "|Ñ|")(1))).RemoteHostIP
wsConnect(Index).SendData "SHIP|Ñ|" & Split(Data, "|Ñ|")(1) & "|Ñ|" & gIP
'*************** chat commands
Case "CJOIN"
LoginChat FindPortUser(Index), Split(Data, "|Ñ|")(1)
Case "COUT"
LogoutChat FindPortUser(Index), Split(Data, "|Ñ|")(1)
Case "CMSG"
MessageChat FindPortUser(Index), Split(Data, "|Ñ|")(1), Split(Data, "|Ñ|")(2)
Case "CMK"
MakeChat Split(Data, "|Ñ|")(1), Split(Data, "|Ñ|")(2), FindPortUser(Index)
Case "QCHAT"
Dim Mx, Cu
Mx = Rm(GetChatI(Split(Data, "|Ñ|")(1))).lblMax.Caption
Cu = Rm(GetChatI(Split(Data, "|Ñ|")(1))).lstUsers.ListCount
wsConnect(Index).SendData "QCHAT|Ñ|" & Cu & "|Ñ|" & Mx & "|Ñ|" & Split(Data, "|Ñ|")(1)
Case "CLIST"
Dim Cl
For I = 0 To lstChat.ListCount - 1
    Cl = Cl & lstChat.List(I) & "|¿|"
Next I
wsConnect(Index).SendData "CLIST|Ñ|" & Cl
Case "RLIST"
Dim Rl, R
R = Split(Data, "|Ñ|")(1)
For I = 0 To Rm(GetChatI(R)).lstUsers.ListCount - 1
    Rl = Rl & Rm(GetChatI(R)).lstUsers.List(I) & "|¿|"
Next I
wsConnect(Index).SendData "RLIST|Ñ|" & Rl
End Select
End Sub

Function GetChatI(name) As Long
Dim I As Long
For I = 0 To lstChat.ListCount - 1
    If lstChat.List(I) = name Then
        GetChatI = lstChat.ItemData(I)
        Exit For
    End If
Next I
End Function

Function MessageChat(user, message, room)
Rm(GetChatI(room)).MessageChat user, message
End Function

Function LogoutChat(user, room)
Dim I As Long
For I = 0 To Rm(GetChatI(room)).lstUsers.ListCount - 1
    If Rm(GetChatI(room)).lstUsers.List(I) = user Then
        Rm(GetChatI(room)).lstUsers.RemoveItem I
        MessageChat "SERVER", "<font face=arial size=3 color=red>" & user & " has left the room</font>", room
        Exit For
    End If
Next I
End Function

Function LoginChat(user, room)
Dim I As Long
If Rm(GetChatI(room)).lstUsers.ListCount >= Val(Rm(GetChatI(room)).lblMax.Caption) Then
    wsConnect(FindUserPort(user)).SendData "ERR|Ñ|Chat room is full"
    Exit Function
End If
For I = 0 To Rm(GetChatI(room)).lstUsers.ListCount - 1
    If Rm(GetChatI(room)).lstUsers.List(I) = user Then
        Exit Function
    End If
Next I
wsConnect(FindUserPort(user)).SendData "CC|Ñ|" & room
DoEvents
Rm(GetChatI(room)).lstUsers.AddItem user
MessageChat "SERVER", "<font face=""arial"" size=3 color=green>" & user & " has entered the room</font>", room
End Function

Function ChatUserList(room) As String

End Function

Function MakeChat(name, max, user)
Dim I As Long
If max < 2 Then max = 2
If max > 20 Then max = 20
For I = 0 To lstChat.ListCount - 1
    If lstChat.List(I) = name Then
        If user <> "SERVER" Then
            wsConnect(FindUserPort(user)).SendData "ERR|Ñ|Chatroom already exists"
        End If
        Exit Function
    End If
Next I
For I = 0 To 19
    If Rm(I).Caption = "CR" Then
        Rm(I).Caption = name
        Rm(I).lstUsers.Clear
        Rm(I).lblMax.Caption = max
        lstChat.AddItem name
        lstChat.ItemData(lstChat.NewIndex) = I
        If user <> "SERVER" Then
            LoginChat user, name
        End If
        Exit For
    End If
Next I
End Function

Function BroadCast(Msg)
On Error Resume Next
Dim I As Long
For I = 0 To lstUsers.ListCount - 1
    wsConnect(lstUsers.ItemData(I)).SendData Msg
Next I
End Function

Function GetOnline(user)
Dim I As Long, buf As String
For I = 0 To lstUsers.ListCount - 1
    buf = buf & lstUsers.List(I) & "|¿|"
Next I
wsConnect(FindUserPort(user)).SendData "RON|Ñ|" & buf
End Function

Function CheckOnline(user)
Dim I As Long
For I = 0 To lstUsers.ListCount - 1
    If lstUsers.List(I) = user Then
        CheckOnline = True
    End If
Next I
End Function

Function CheckExist(user)
Dim I As Long
For I = 0 To lstU.ListCount - 1
    If lstU.List(I) = user Then
        CheckExist = True
    End If
Next I
End Function

Function FindUserPort(user) As Long
Dim I As Long
For I = 0 To lstUsers.ListCount - 1
    If lstUsers.List(I) = user Then
        FindUserPort = lstUsers.ItemData(I)
        Exit For
    End If
Next I
End Function

Function FindPortUser(port) As String
Dim I As Long
For I = 0 To lstUsers.ListCount - 1
    If lstUsers.ItemData(I) = port Then
        FindPortUser = lstUsers.List(I)
        Exit For
    End If
Next I
End Function

Private Sub wsServer_ConnectionRequest(ByVal requestID As Long)
wsI = wsI + 1
Load wsConnect(wsI)
wsConnect(wsI).Accept requestID
End Sub

