VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1575
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4095
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930.562
   ScaleMode       =   0  'User
   ScaleWidth      =   3844.983
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New User"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   135
      Width           =   2445
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2445
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1200
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
frmClient.wsClient.Close
Unload Me
End Sub

Private Sub cmdNew_Click()
Unload Me
frmClient.wsClient.Close
Load frmNewUser
frmNewUser.Show
End Sub

Private Sub cmdOK_Click()
If Len(txtUserName.Text) < 4 Or Len(txtUserName.Text) > 16 Or Len(txtPassword.Text) < 4 Or Len(txtPassword.Text) > 16 Then
    MsgBox "Invalid username or password!", vbCritical, ""
End If

If frmClient.wsClient.State = sckConnected Then
    frmClient.wsClient.SendData "LOGIN|Ñ|" & txtUserName.Text & "|Ñ|" & txtPassword.Text
End If

Unload Me
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK_Click
End Sub
