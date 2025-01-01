Attribute VB_Name = "modInput"
Option Explicit

Public pad1(0 To 7) As Long, pad1net(0 To 7) As Long
Public pad1bit As Long
Public pad2(0 To 7) As Long, pad2net(0 To 7) As Long
Public pad2bit As Long
Public zapper_trigger As Long, zapper_level As Long, zapper_x As Single, zapper_y As Single

Public Sub padstrobe()
    pad1bit = 0
    pad2bit = 0
End Sub

Public Function pad1read() As Long
    Select Case netmode
        Case NET_SERVER_PLAYING, NET_CLIENT_PLAYING
            If pad1bit < 8 Then pad1read = pad1net(pad1bit): pad1bit = pad1bit + 1 Else pad1read = 1
        Case Else
            If pad1bit < 8 Then pad1read = pad1(pad1bit): pad1bit = pad1bit + 1 Else pad1read = 1
    End Select
End Function

Public Function pad2read() As Long
    Select Case netmode
        Case NET_SERVER_PLAYING, NET_CLIENT_PLAYING
            If pad2bit < 8 Then pad2read = pad2net(pad2bit): pad2bit = pad2bit + 1 Else pad2read = 1
        Case Else
            If pad2bit < 8 Then pad2read = pad2(pad2bit): pad2bit = pad2bit + 1 Else pad2read = 1
    End Select
    
    'If zapper_trigger = 1 Then pad2read = pad2read Or 16: frmMain.pic_MouseMove 0, 0, zapper_x, zapper_y
    'If zapper_level < &HD0 Then pad2read = pad2read Or 8
End Function

