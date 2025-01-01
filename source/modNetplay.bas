Attribute VB_Name = "modNetplay"
Option Explicit

Public netmode As Long
Public clientVer(0 To 2) As Long
Public netROM As String
Public netgotinput As Long

Public Const NET_SERVER_LISTEN = 1
Public Const NET_SERVER_HANDSHAKE = 2
Public Const NET_SERVER_SENDROM = 3
Public Const NET_SERVER_WAITGAME = 4
Public Const NET_SERVER_PLAYING = 5
Public Const NET_CLIENT_LISTEN = 6
Public Const NET_CLIENT_HANDSHAKE = 7
Public Const NET_CLIENT_SENDROM = 8
Public Const NET_CLIENT_SENDROM2 = 9
Public Const NET_CLIENT_WAITGAME = 10
Public Const NET_CLIENT_PLAYING = 11

Public Sub netConnect()
    netmode = NET_SERVER_HANDSHAKE
    netSendData Chr$(App.Major) + Chr$(App.Minor) + Chr$(App.Revision)
End Sub

Public Sub netSendData(ByVal data As String)
    frmMain.ws.SendData data
End Sub

Private Sub netError()
    netmode = 0
    MsgBox "Network error!"
End Sub

Public Sub netSendROM()
    Dim romchunk As String
    Dim getlen As Long
    
    getlen = LOF(7) - Loc(7)
    If getlen > 1024 Then getlen = 1024
    romchunk = Space$(getlen)
    Get #7, , romchunk
    frmMain.Caption = CStr(Loc(7))
    If Len(romchunk) < 1024 Then
        Close #7
        If Len(romchunk) = 0 Then romchunk = Chr$(255)
        netmode = NET_SERVER_WAITGAME
        gameready = 1
    End If
    netSendData Chr$(getlen \ 256&) + Chr$(getlen And 255) + romchunk
End Sub

Public Sub netGetData(ByVal data As String)
    Select Case netmode
        'server
        Case NET_SERVER_HANDSHAKE, NET_CLIENT_HANDSHAKE
            'If Len(data) <> 3 Then netError
            totalframes = 0
            clientVer(0) = Asc(Left$(data, 1)): clientVer(1) = Asc(Mid$(data, 2, 1)): clientVer(2) = Asc(Right$(data, 1))
            If netmode = NET_CLIENT_HANDSHAKE Then netmode = NET_CLIENT_SENDROM: netSendData Chr$(App.Major) + Chr$(App.Minor) + Chr$(App.Revision): netROM = ""
            'MsgBox "Remote vbNES version: " + CStr(clientVer(0)) + "." + CStr(clientVer(1)) + "." + CStr(clientVer(2))
            If netmode = NET_SERVER_HANDSHAKE Then netmode = NET_SERVER_SENDROM: Open romfile For Binary As #7: netSendROM
            
        Case NET_SERVER_SENDROM
            netSendROM
        
        Case NET_SERVER_WAITGAME
            netmode = NET_SERVER_PLAYING
        
        Case NET_SERVER_PLAYING
            If (Asc(data) And 1) Then pad2net(0) = 1 Else pad2net(0) = 0
            If (Asc(data) And 2) Then pad2net(1) = 1 Else pad2net(1) = 0
            If (Asc(data) And 4) Then pad2net(2) = 1 Else pad2net(2) = 0
            If (Asc(data) And 8) Then pad2net(3) = 1 Else pad2net(3) = 0
            If (Asc(data) And 16) Then pad2net(4) = 1 Else pad2net(4) = 0
            If (Asc(data) And 32) Then pad2net(5) = 1 Else pad2net(5) = 0
            If (Asc(data) And 64) Then pad2net(6) = 1 Else pad2net(6) = 0
            If (Asc(data) And 128) Then pad2net(7) = 1 Else pad2net(7) = 0
            netgotinput = 1
        
        
        'client
        Case NET_CLIENT_SENDROM2
            netROM = netROM + data
            'MsgBox "got data"
            If Len(data) < 1024 Then
                If Dir$(App.Path + "\netrom.tmp") <> "" Then Kill App.Path + "\netrom.tmp"
                Open App.Path + "\netrom.tmp" For Binary As #7
                Put #7, , netROM
                Close #7
                romfile = App.Path + "\netrom.tmp"
                gameready = 1
                netmode = NET_CLIENT_WAITGAME
            Else
                netmode = NET_CLIENT_SENDROM
                netSendData Chr$(255)
            End If
        
        Case NET_CLIENT_WAITGAME
            netmode = NET_CLIENT_PLAYING
        
        Case NET_CLIENT_PLAYING
            If (Asc(data) And 1) Then pad1net(0) = 1 Else pad1net(0) = 0
            If (Asc(data) And 2) Then pad1net(1) = 1 Else pad1net(1) = 0
            If (Asc(data) And 4) Then pad1net(2) = 1 Else pad1net(2) = 0
            If (Asc(data) And 8) Then pad1net(3) = 1 Else pad1net(3) = 0
            If (Asc(data) And 16) Then pad1net(4) = 1 Else pad1net(4) = 0
            If (Asc(data) And 32) Then pad1net(5) = 1 Else pad1net(5) = 0
            If (Asc(data) And 64) Then pad1net(6) = 1 Else pad1net(6) = 0
            If (Asc(data) And 128) Then pad1net(7) = 1 Else pad1net(7) = 0
            netgotinput = 1
        
    End Select
End Sub

