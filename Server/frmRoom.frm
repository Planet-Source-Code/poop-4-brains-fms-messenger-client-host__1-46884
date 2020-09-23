VERSION 5.00
Begin VB.Form frmRoom 
   Caption         =   "CR"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2055
   Icon            =   "frmRoom.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   2055
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrClose 
      Interval        =   1000
      Left            =   720
      Top             =   720
   End
   Begin VB.ListBox lstUsers 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblMax 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblUsers 
      BackStyle       =   0  'Transparent
      Caption         =   "Max Users:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function MessageChat(user, message)
Dim I As Long
For I = 0 To lstUsers.ListCount - 1
    DoEvents
    frmServer.wsConnect(frmServer.FindUserPort(lstUsers.List(I))).SendData "CMSG|Ñ|" & user & "|Ñ|" & message
Next I
End Function

Private Sub Form_Unload(Cancel As Integer)
Dim I As Long
For I = 0 To frmServer.lstChat.ListCount - 1
    If frmServer.lstChat.List(I) = Me.Caption Then
        frmServer.lstChat.RemoveItem I
        Exit For
    End If
Next I

Me.Caption = "CR"
lstUsers.Clear
End Sub
