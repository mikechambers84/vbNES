Attribute VB_Name = "modCart"
Option Explicit

Public Type typeHdr
    match(0 To 3) As Byte
    PRGsize As Long
    CHRsize As Long
    flags(0 To 3) As Byte
    zerobytes(0 To 4) As Byte
    mapper As Long
    mirroring As Long
End Type

Public hdr As typeHdr
Public tempromfile As String

Public Sub loadROM(ByVal romfile As String)
    Dim addr As Long
    Dim n As Long
    Dim d$
    
    addr = 0
    If gamegenie = 1 Then
        If Dir$(App.Path + "\genie.nes") <> "" Then
            Open App.Path + "\genie.nes" For Binary As #1
            tempromfile = romfile
        Else
            MsgBox "GENIE.NES not found in application path!" + vbCrLf + "Opening ROM without Game Genie...", vbInformation Or vbOKOnly, "Game Genie"
            gamegenie = 0
        End If
    End If
    If gamegenie = 2 Then romfile = tempromfile
    If gamegenie <> 1 Then Open romfile For Binary As #1
    For n = 0 To 3
        d$ = Space$(1): Get #1, , d$: hdr.match(n) = Asc(d$)
    Next n
    d$ = Space$(1): Get #1, , d$: hdr.PRGsize = Asc(d$)
    d$ = Space$(1): Get #1, , d$: hdr.CHRsize = Asc(d$)
    For n = 0 To 3
        d$ = Space$(1): Get #1, , d$: hdr.flags(n) = Asc(d$)
    Next n
    For n = 0 To 4
        d$ = Space$(1): Get #1, , d$: hdr.zerobytes(n) = Asc(d$)
    Next n
    Seek #1, 17
    For n = 0 To (16384& * hdr.PRGsize) - 1
        d$ = Space$(1): Get #1, , d$: PRGbin(addr \ 1024&, addr And 1023&) = Asc(d$)
        'If hdr.PRGsize = 1 Then RAM(addr + 16384&) = RAM(addr)
        addr = addr + 1
    Next n
    addr = 0
    For n = 0 To (8192& * hdr.CHRsize) - 1
        d$ = Space$(1): Get #1, , d$: CHRbin(addr \ 1024&, addr And 1023&) = Asc(d$)
        addr = addr + 1
    Next n
    Close #1
    
    hdr.PRGsize = hdr.PRGsize * 16&
    hdr.CHRsize = hdr.CHRsize * 8&
    If (Asc(hdr.flags(0)) And 8) Then hdr.mirroring = 2 Else hdr.mirroring = Asc(hdr.flags(0)) And 1
    Select Case hdr.mirroring
        Case 0 'horizontal
            PPU.ntmap(0) = &H2000
            PPU.ntmap(1) = &H2000
            PPU.ntmap(2) = &H2400
            PPU.ntmap(3) = &H2400
        Case 1 'vertical
            PPU.ntmap(0) = &H2000
            PPU.ntmap(1) = &H2400
            PPU.ntmap(2) = &H2000
            PPU.ntmap(3) = &H2400
        Case 2 'four-screen
            PPU.ntmap(0) = &H2000
            PPU.ntmap(1) = &H2400
            PPU.ntmap(2) = &H2800
            PPU.ntmap(3) = &H2C00
    End Select
    
    If hdr.flags(1) = Asc("D") Then hdr.flags(1) = 0 'DISKDUDE fix
    hdr.mapper = (hdr.flags(0) \ 16) Or (hdr.flags(1) And 240)
End Sub
