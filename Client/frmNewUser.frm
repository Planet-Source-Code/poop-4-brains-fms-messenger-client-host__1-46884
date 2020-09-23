VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmNewUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make New FMS User"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsNew 
      Left            =   3720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   5333
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create New User"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblConfirm 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password (4-16 chars)"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Username (4-16 chars)"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()
If Len(txtUsername.Text) < 4 Or Len(txtUsername.Text) > 16 Then
    MsgBox "Username must be 4-16 characters long", vbCritical, "Cannot Create New User"
    Exit Sub
End If
If Len(txtPassword.Text) < 4 Or Len(txtPassword.Text) > 16 Then
    MsgBox "Password must be 4-16 characters long", vbCritical, "Cannot Create New User"
    Exit Sub
End If
If txtPassword.Text <> txtConfirm.Text Then
    MsgBox "Passwords do not match", vbCritical, "Cannot Create New User"
    Exit Sub
End If

wsNew.Connect
End Sub

Private Sub Form_Load()
wsNew.RemoteHost = frmClient.wsClient.RemoteHost
End Sub

Private Sub wsNew_Connect()
wsNew.SendData "NUSER|Ñ|" & txtUsername.Text & "|Ñ|" & txtPassword.Text
End Sub

Private Sub wsNew_DataArrival(ByVal bytesTotal As Long)
Dim Data As String

wsNew.GetData Data

If Data = "USERCREATE" Then
    MsgBox "User created successfully", vbinformatio, "FMS Messenger"
    wsNew.Close
    Unload Me
End If

If Split(Data, "|Ñ|")(0) = "ERR" Then
    MsgBox Split(Data, "|Ñ|")(1), vbCritical, "Server Error"
End If
End Sub
