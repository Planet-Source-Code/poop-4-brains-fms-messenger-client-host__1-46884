Attribute VB_Name = "modPublic"
Public IM(20) As New frmIM
Public cUser As String
Public trSys As New clsTrayIcon

'**** taken from pscchat 5 sourcecode... thanks tim and carsten
Public Function DectoWebCol(lngColour As Long) As String
    Dim strColour As String
    'Convert decimal colour to hex
    strColour = Hex(lngColour)
    'Add leading zero's


    Do While Len(strColour) < 6
        strColour = "0" & strColour
    Loop
    'Reverse the bgr string pairs to rgb
    DectoWebCol = "#" & Right$(strColour, 2) & _
    Mid$(strColour, 3, 2) & _
    Left$(strColour, 2)
End Function

