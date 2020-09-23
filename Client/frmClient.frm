VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "FMS Client"
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2955
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   2955
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   5333
   End
   Begin MSComctlLib.TreeView lstUsers 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2566
      _Version        =   393217
      Style           =   7
      ImageList       =   "imgStatus"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgStatus 
      Left            =   2520
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":0F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":12CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTools 
      Left            =   1680
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":161E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":1EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":27D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":30AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4995
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1535
      ButtonWidth     =   1296
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "imgTools"
      DisabledImageList=   "imgTools"
      HotImageList    =   "imgTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Login"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Message"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Chat"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Login"
      Begin VB.Menu mnuLogin 
         Caption         =   "Login As..."
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrefs 
         Caption         =   "Preferences"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "Make new user"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMessage 
      Caption         =   "Message"
      Begin VB.Menu mnuIM 
         Caption         =   "Start Instant Message"
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "Users"
      Begin VB.Menu mnuAddb 
         Caption         =   "Add"
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "_list"
      Begin VB.Menu mnuPMSG 
         Caption         =   "Message Buddy"
      End
      Begin VB.Menu mnuREMB 
         Caption         =   "Remove Buddy"
      End
      Begin VB.Menu mnuVP 
         Caption         =   "View Profile"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
On Error Resume Next
lstUsers.Top = tlBar.Height
lstUsers.Left = 0
lstUsers.Width = Me.ScaleWidth
lstUsers.Height = (Me.ScaleHeight - tlBar.Height) - stBar.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstUsers_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub lstUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    PopupMenu mnuList
End If
End Sub

Private Sub mnuAddb_Click()
If wsClient.State = sckConnected Then
    Dim SN As String
    SN = InputBox("What is the users name you want to add?", "Enter Username", "")
    If SN = "" Then Exit Sub
    wsClient.SendData "ADDB|Ñ|" & SN
End If
End Sub

Private Sub mnuDisconnect_Click()
If wsClient.State <> sckClosed Then
    wsClient.Close
End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Function PopIM(User, msg)
Dim I As Long
RE:
For I = 0 To 20
    If IM(I).Caption = "IM - " & User Then
        If IM(I).Visible = False Then
            IM(I).Form_Load
        End If
        If Len(msg) > 0 Then
            IM(I).AddText User, msg
        End If
        Exit Function
    End If
Next I

For I = 0 To 20
    If Len(IM(I).Caption) = 2 Then
        IM(I).Show
        IM(I).Caption = "IM - " & User
        Exit For
    End If
Next I
GoTo RE
End Function

Private Sub mnuIM_Click()
If cUser <> "" Then
    Dim SN As String
    User = InputBox("Who do you want to instant message?", "Enter SN")
    If User = "" Then Exit Sub
    PopIM User, ""
End If
End Sub

Private Sub mnuLogin_Click()
If wsClient.State <> sckClosed Then
    wsClient.Close
End If

wsClient.Connect
End Sub

Private Sub mnuNew_Click()
Load frmNewUser
frmNewUser.Show
End Sub

Private Sub mnuPMSG_Click()
On Error GoTo Err:
If lstUsers.SelectedItem.Text <> "" Then
    PopIM lstUsers.SelectedItem.Text, ""
End If
Exit Sub
Err:
If Err.Number <> 91 Then
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "FMS BuddyListr"
End If
End Sub

Private Sub mnuREMB_Click()
On Error GoTo Err:
If lstUsers.SelectedItem.Text <> "" Then
    If MsgBox("Are you sure you want to remove " & lstUsers.SelectedItem.Text, vbYesNo, "FMS BuddyList") = vbYes Then
        
    End If
End If
Exit Sub
Err:
If Err.Number <> 91 Then
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "FMS BuddyListr"
End If
End Sub

Private Sub tlBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    mnuLogin_Click
Case 2
    mnuIM_Click
Case 3
    mnuAddb_Click
Case 4
    If frmChat.Visible = True Then
        frmChat.SetFocus
    Else
        If wsClient.State = sckConnected Then
            wsClient.SendData "CLIST|Ñ|"
        End If
    End If
End Select
End Sub

Private Sub wsClient_Close()
cUser = ""
Me.Caption = "FMS Client"
stBar.SimpleText = "Not logged in"
lstUsers.Nodes.Clear
If frmChat.Visible = True Then
    Unload frmChat
End If
End Sub

Private Sub wsClient_Connect()
frmLogin.Show
End Sub

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)
Dim Data As String

wsClient.GetData Data

Debug.Print Data

Select Case Split(Data, "|Ñ|")(0)
Case "LOGINCONFIRM"
    cUser = Split(Data, "|Ñ|")(1)
    stBar.SimpleText = "Logged in as " & cUser
    Me.Caption = "FMS Client - " & cUser
    wsClient.SendData "BUD|Ñ|"
Case "PMSG"
    PopIM Split(Data, "|Ñ|")(1), Split(Data, "|Ñ|")(2)
Case "ERR"
    MsgBox Split(Data, "|Ñ|")(1), vbCritical, "Server Error"
Case "BUD"
    ShowBuddyList Split(Data, "|Ñ|")(1)
Case "RON"
    ShowOnline Split(Data, "|Ñ|")(1)
Case "ON"
    Dim I3 As Long
    For I3 = 1 To lstUsers.Nodes.Count
        If lstUsers.Nodes.Item(I3).Text = Split(Data, "|Ñ|")(1) Then
            lstUsers.Nodes.Item(I3).Image = 1
        End If
    Next I3
Case "OFF"
    Dim I2 As Long
    For I2 = 1 To lstUsers.Nodes.Count
        If lstUsers.Nodes.Item(I2).Text = Split(Data, "|Ñ|")(1) Then
            lstUsers.Nodes.Item(I2).Image = 2
            Exit For
        End If
    Next I2
Case "CLIST"
    Dim J
    frmCMenu.Show
    frmCMenu.lstChat.Clear
    For J = 0 To UBound(Split(Split(Data, "|Ñ|")(1), "|¿|")) - 1
        frmCMenu.lstChat.AddItem Split(Split(Data, "|Ñ|")(1), "|¿|")(J)
    Next J
Case "QCHAT"
    frmCMenu.lblCUsers.Caption = Split(Data, "|Ñ|")(1) & " / " & Split(Data, "|Ñ|")(2)
    frmCMenu.lblCName.Caption = Split(Data, "|Ñ|")(3)
Case "CMSG"
    frmChat.AddText Split(Data, "|Ñ|")(1), Split(Data, "|Ñ|")(2)
Case "CC"
    Load frmChat
    frmChat.Show
    frmChat.Caption = "Chat Room - " & Split(Data, "|Ñ|")(1)
Case "RLIST"
    Dim K
    frmChat.lstUsers.Clear
    For K = 0 To UBound(Split(Split(Data, "|Ñ|")(1), "|¿|")) - 1
        frmChat.lstUsers.AddItem Split(Split(Data, "|Ñ|")(1), "|¿|")(K)
    Next K
Case "SHIP"
    MsgBox Split(Data, "|Ñ|")(1) & "'s ip address is " & Split(Data, "|Ñ|")(2)
End Select
'*** all purpose splitter ! Split(Data, "|Ñ|")(1)
End Sub

Function ShowBuddyList(bdlist)
Dim I As Long
lstUsers.Nodes.Clear
For I = 0 To UBound(Split(bdlist, "|¿|")) - 1
    lstUsers.Nodes.Add , , , Split(bdlist, "|¿|")(I)
Next I
wsClient.SendData "RON|Ñ|"
End Function

Function ShowOnline(bdlist)
Dim I As Long, ia As Long
For I = 1 To lstUsers.Nodes.Count
    For ia = 0 To UBound(Split(bdlist, "|¿|")) - 1
        If Split(bdlist, "|¿|")(ia) = lstUsers.Nodes.Item(I).Text Then
            lstUsers.Nodes.Item(I).Image = 1
        Else
            lstUsers.Nodes.Item(I).Image = 2
        End If
    Next ia
Next I
End Function

Private Sub wsClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Select Case Number
Case 10061
    MsgBox "Could not connect to the server" & vbCrLf & "The server may be down, or you may not be connected to the internet", vbCritical, "FMS Messenger"
Case Else
    MsgBox Number & " - " & Description, vbCritical, "UNEXPECTED ERROR! - FMS Messenger"
End Select
End Sub
