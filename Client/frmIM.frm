VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIM 
   Caption         =   "IM"
   ClientHeight    =   3255
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5580
   Icon            =   "frmIM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   372
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmnColor 
      Left            =   3720
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cmnFont 
      Left            =   3120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
   End
   Begin MSComctlLib.ImageList imgFont 
      Left            =   5040
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM.frx":0F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM.frx":1514
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM.frx":1AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM.frx":2048
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM.frx":25E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlFont 
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgFont"
      DisabledImageList=   "imgFont"
      HotImageList    =   "imgFont"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      MaxLength       =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin SHDocVwCtl.WebBrowser wbIM 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function AddText(User, msg)
msg = Replace(msg, ">:(", "<img src=""" & App.Path & "/Emo/icon_twisted.gif"">")
msg = Replace(msg, ":))", "<img src=""" & App.Path & "/Emo/icon_lol.gif"">")
msg = Replace(msg, ":((", "<img src=""" & App.Path & "/Emo/icon_cry.gif"">")
msg = Replace(msg, ":)", "<img src=""" & App.Path & "/Emo/icon_smile.gif"">")
msg = Replace(msg, ":(", "<img src=""" & App.Path & "/Emo/icon_sad.gif"">")
msg = Replace(msg, ":|", "<img src=""" & App.Path & "/Emo/icon_neutral.gif"">")
msg = Replace(msg, ":o", "<img src=""" & App.Path & "/Emo/icon_surprised.gif"">")
msg = Replace(msg, ":x", "<img src=""" & App.Path & "/Emo/icon_puke.gif"">")
msg = Replace(msg, ":S", "<img src=""" & App.Path & "/Emo/icon_confused.gif"">")
msg = Replace(msg, ":D", "<img src=""" & App.Path & "/Emo/icon_biggrin.gif"">")
msg = Replace(msg, "B)", "<img src=""" & App.Path & "/Emo/icon_coolgif"">")
msg = Replace(msg, "8O", "<img src=""" & App.Path & "/Emo/icon_eek.gif"">")
msg = Replace(msg, "8)", "<img src=""" & App.Path & "/Emo/icon_razz.gif"">")
msg = Replace(msg, ":%", "<img src=""" & App.Path & "/Emo/icon_redface.gif"">")
msg = Replace(msg, "@)", "<img src=""" & App.Path & "/Emo/icon_rolleyes.gif"">")

wbIM.Document.body.innerHTML = wbIM.Document.body.innerHTML & "<B>" & User & "</B>: " & msg & "<br>"
wbIM.Document.body.scrolltop = CLng(Len(wbIM.Document.body.innerHTML)) * 100
End Function

Private Sub cmdSend_Click()
Dim nSize
If Len(txtSend.Text) > 0 And frmClient.wsClient.State = sckConnected Then
    nSize = txtSend.FontSize / 3
    txtSend.Text = "<font face=" & txtSend.FontName & " size=" & nSize & "pt color=" & DectoWebCol(txtSend.ForeColor) & ">" & IIf(txtSend.FontBold, "<B>", "") & IIf(txtSend.FontItalic, "<I>", "") & txtSend.Text & IIf(txtSend.FontBold, "</B>", "") & IIf(txtSend.FontItalic, "</I>", "") & "</font>"
    AddText cUser, txtSend.Text
    frmClient.wsClient.SendData "PMSG|Ñ|" & Mid(Me.Caption, 6) & "|Ñ|" & txtSend.Text
    txtSend.Text = ""
End If
End Sub

Sub Form_Load()
On Error Resume Next
wbIM.Navigate "about:blank"
Do While wbIM.ReadyState <> READYSTATE_COMPLETE
    DoEvents
Loop
DoEvents
wbIM.Document.body.innerHTML = ""
txtSend.Font = "Arial"
txtSend.FontBold = False
txtSend.FontItalic = False
txtSend.ForeColor = vbBlack
End Sub

Private Sub Form_Resize()
On Error Resume Next
wbIM.Left = 5
wbIM.Top = 5
wbIM.Height = ((Me.ScaleHeight - txtSend.Height) - tlFont.Height) - 15
wbIM.Width = Me.ScaleWidth - 10
tlFont.Left = 5
tlFont.Top = wbIM.Height + 5
tlFont.Width = Me.ScaleWidth - 10

txtSend.Top = wbIM.Height + tlFont.Height + 10
txtSend.Left = 5
txtSend.Width = Me.ScaleWidth - cmdSend.Width - 15

cmdSend.Height = txtSend.Height
cmdSend.Left = txtSend.Width + 10
cmdSend.Top = txtSend.Top
End Sub

Private Sub tlFont_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Select Case Button.Index
Case 1
    cmnFont.FontName = txtSend.Font
    cmnFont.FontBold = txtSend.FontBold
    cmnFont.FontItalic = txtSend.FontItalic
    cmnFont.FontSize = txtSend.FontSize
    cmnFont.FontStrikethru = txtSend.FontStrikethru
    cmnFont.FontUnderline = txtSend.FontUnderline
    cmnFont.ShowFont
    txtSend.Font = cmnFont.FontName
    txtSend.FontBold = cmnFont.FontBold
    txtSend.FontItalic = cmnFont.FontItalic
    txtSend.FontSize = cmnFont.FontSize
    txtSend.FontStrikethru = cmnFont.FontStrikethru
    txtSend.FontUnderline = cmnFont.FontUnderline
Case 2
    cmnColor.Color = txtSend.ForeColor
    cmnColor.ShowColor
    txtSend.ForeColor = cmnColor.Color
Case 4
    Dim URL As String, Cap As String, buf
    URL = InputBox("What is the url of the hyperlink you want?", "HyperLink URL")
    Cap = InputBox("What is the caption of the hyperlink you want?", "Hyperlink Caption")
    If URL = "" Then Exit Sub
    If Cap = "" Then Exit Sub
    buf = "<a href=" & URL & " target=_new>" & Cap & "</a>"
    txtSend.Text = txtSend.Text & buf
Case 5
    If txtSend.FontBold = True Then
        txtSend.FontBold = False
        Exit Sub
    Else
        txtSend.FontBold = True
        Exit Sub
    End If
Case 6
    If txtSend.FontItalic = True Then
        txtSend.FontItalic = False
        Exit Sub
    Else
        txtSend.FontItalic = True
        Exit Sub
    End If
End Select
Exit Sub
Err:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Unexpected Error"
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSend_Click
End If
End Sub
