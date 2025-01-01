Attribute VB_Name = "modMemory"
Option Explicit

Public ram(0 To &H7FF) As Byte
Public PRGbin(0 To 1023, 0 To 1023) As Byte
Public CHRbin(0 To 1023, 0 To 1023) As Byte
Public PRGbank(0 To 1023) As Long
Public CHRbank(0 To 1023) As Long
Public batRAM(0 To 8191) As Long
Dim n As Long

Public Sub write6502(ByVal addr As Long, ByVal value As Long)
    Dim templ As Long
    Select Case addr
        Case Is < &H2000
            ram(addr And &H7FF) = value
        Case &H2000 To &H3FFF
            writePPUregs &H2000 Or (addr And 7), value
        Case &H4014
            For n = 0 To 255
                OAM.ram((OAM.addr + n) And 255) = read6502(value * 256 + n)
            Next n
        Case &H4000 To &H4017
            If (addr = &H4016) Then padstrobe Else writeAPU addr, value
        Case &H6000 To &H7FFF
            Select Case hdr.mapper
                Case 69
                    If (map69.ram = True) And (map69.ramenable = True) Then
                        templ = (map69.prg6000 * 8192&) + (addr - &H6000&)
                        CHRbin(templ \ 1024&, templ And 1023&) = value
                    End If
                Case Else
                    batRAM(addr And 8191) = value
            End Select
        Case Is >= 32768
            If gamegenie = 1 Then geniewrite addr, value Else mapperwrite addr, value
    End Select
End Sub

Public Function read6502(ByVal addr As Long) As Long
    Dim retval As Long, cnum As Long, templ As Long
    
    Select Case addr
        Case Is < &H2000
            retval = ram(addr And &H7FF)
        Case &H2000 To &H3FFF
            retval = readPPUregs(&H2000 Or (addr And 7))
        Case &H4015
            retval = readAPU(addr)
        Case &H4016
            retval = pad1read
        Case &H4017
            retval = pad2read
        Case &H6000 To &H7FFF
            Select Case hdr.mapper
                Case 69
                    templ = (map69.prg6000 * 8192&) + (addr - &H6000&)
                    retval = CHRbin(templ \ 1024&, templ And 1023&)
                Case Else
                    retval = batRAM(addr And 8191)
            End Select
        Case Is >= 32768
            addr = addr - 32768
            retval = PRGbin(PRGbank(addr \ 1024&), addr And 1023&)
        Case Else
            retval = 0
    End Select
    If (gamegenie = 2) Then
        For cnum = 0 To 2
            If ((cheat(cnum).valid) And (addr = cheat(cnum).addr)) Then
                If (cheat(cnum).docompare = 1) Then
                    If (retval = cheat(cnum).compare) Then retval = cheat(cnum).replace
                Else
                    retval = cheat(cnum).replace
                End If
            End If
        Next cnum
    End If

    read6502 = retval
End Function

