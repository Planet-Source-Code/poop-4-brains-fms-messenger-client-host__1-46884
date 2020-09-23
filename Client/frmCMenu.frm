VERSION 5.00
Begin VB.Form frmCMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chat Menu"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmCMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMake 
      Caption         =   "Make Chatroom"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join Chatroom"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.ListBox lstChat 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblCUsers 
      BackStyle       =   0  'Transparent
      Caption         =   "0 / 2"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblCName 
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMax 
      BackStyle       =   0  'Transparent
      Caption         =   "Chatroom Stats:"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmCMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdJoin_Click()
If lstChat.ListIndex > -1 Then
    frmClient.wsClient.SendData "CJOIN|Ñ|" & lstChat.List(lstChat.ListIndex)
    Unload Me
End If
End Sub

Private Sub cmdMake_Click()
Dim Nm, Mx
Nm = InputBox("What is the name of the chatroom?", "Name of chatroom")
Mx = Val(InputBox("What is the max ammount of users? (2-20)", "Max Ammount of users", "2"))
If Mx < 2 Then Mx = 2
If Mx > 20 Then Mx = 20
If Len(Nm) < 1 Then Exit Sub
frmClient.wsClient.SendData "CMK|Ñ|" & Nm & "|Ñ|" & Mx
Unload Me
End Sub

Private Sub lstChat_Click()
If lstChat.ListIndex > -1 And lblCName.Caption <> lstChat.List(lstChat.ListIndex) Then
    frmClient.wsClient.SendData "QCHAT|Ñ|" & lstChat.List(lstChat.ListIndex)
End If
End Sub
