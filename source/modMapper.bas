Attribute VB_Name = "modMapper"
Option Explicit

Type map1type
    map1reg(0 To 3) As Long
    map1bitpos As Long
    map1accum As Long
End Type

Type map9type
    latch0_fd As Long
    latch0_fe As Long
    latch1_fd As Long
    latch1_fe As Long
    latch1 As Long
    latch2 As Long
End Type

Type map4type
    command As Long
    vromsize As Long
    chraddrselect As Long
    irqcounter As Long
    irqlatch As Long
    irqenable As Long
    swap As Long
    prgswitch1 As Long
    prgswitch2 As Long
    prgaddr As Long
End Type

Type map69type
    command As Long
    prg6000 As Long
    RAM As Boolean
    ramenable As Boolean
End Type

Public map1 As map1type
Public map4 As map4type
Public map9 As map9type
Public map69 As map69type

Public Sub PRGswap(ByVal banknum As Long, ByVal newbank As Long, ByVal banksize As Long)
     Dim tmpstart As Long, tmplen As Long, tmpcur As Long, tmpcur2 As Long
     tmpstart = (newbank * banksize) \ 1024&
     tmplen = (banksize \ 1024&) - 1
     tmpcur2 = (banknum * banksize) \ 1024&
     For tmpcur = 0 To tmplen
       PRGbank(tmpcur2) = tmpstart Mod hdr.PRGsize
       tmpcur2 = tmpcur2 + 1&
       tmpstart = tmpstart + 1&
     Next tmpcur
End Sub

Public Sub PRGswap2(ByVal banknum As Long, ByVal newbank As Long, ByVal banksize As Long)
     Dim tmpstart As Long, tmplen As Long, tmpcur As Long, tmpcur2 As Long
     tmpstart = newbank
     tmplen = (banksize \ 1024&) - 1
     tmpcur2 = banknum \ 1024&
     For tmpcur = 0 To tmplen
       PRGbank(tmpcur2) = tmpstart Mod hdr.PRGsize
       tmpcur2 = tmpcur2 + 1&
       tmpstart = tmpstart + 1&
     Next tmpcur
End Sub

Public Sub CHRswap(ByVal banknum As Long, ByVal newbank As Long, ByVal banksize As Long)
     Dim tmpstart As Long, tmplen As Long, tmpcur As Long, tmpcur2 As Long, useCHRsize As Long
     tmpstart = (newbank * banksize) \ 1024&
     tmplen = (banksize \ 1024&) - 1
     tmpcur2 = (banknum * banksize) \ 1024&
     If hdr.CHRsize = 0 Then useCHRsize = 1024 Else useCHRsize = hdr.CHRsize
     For tmpcur = 0 To tmplen
       CHRbank(tmpcur2) = tmpstart Mod useCHRsize
       tmpcur2 = tmpcur2 + 1&
       tmpstart = tmpstart + 1&
     Next tmpcur
End Sub

Public Sub initMapper()
    Select Case hdr.mapper
        Case 0, 3 'no mapper hardware or CNROM
            If (hdr.PRGsize = 16) Then
                PRGswap 0, 0, 16384
                PRGswap 1, 0, 16384
            Else
                PRGswap 0, 0, 32768
            End If
            CHRswap 0, 0, 8192
        
        Case 1 'MMC1
            PRGswap 0, 0, 16384
            PRGswap 1, (hdr.PRGsize \ 16&) - 1, 16384
            CHRswap 0, 0, 8192
            write6502 32768, 12
        
        Case 2 'UxROM
            PRGswap 0, 0, 16384
            PRGswap 1, (hdr.PRGsize \ 16&) - 1, 16384
            CHRswap 0, 0, 8192
        
        Case 4, 64, 158 'MMC3
            PRGswap 0, 0, 16384
            PRGswap 1, hdr.PRGsize - 1, 16384
            CHRswap 0, 0, 8192
            map4.irqenable = 0
            map4.irqcounter = 255
            map4.irqlatch = 255
            PPU.bgtable = 0
            PPU.sprtable = &H1000
        
        Case 7 'AxROM
            PRGswap 0, 0, 32768
            CHRswap 0, 0, 8192
            PPU.ntmap(0) = &H2000
            PPU.ntmap(1) = &H2000
            PPU.ntmap(2) = &H2000
            PPU.ntmap(3) = &H2000
        
        Case 9, 10 'MMC2 and MMC4
            PRGswap 0, 0, 8192
            PRGswap 1, (hdr.PRGsize \ 2&) - 3, 8192
            PRGswap 2, (hdr.PRGsize \ 2&) - 2, 8192
            PRGswap 3, (hdr.PRGsize \ 2&) - 1, 8192
            map9.latch1 = &HFE&
            map9.latch2 = &HFE&
        
        Case 66 'GNROM/MHROM
            PRGswap 0, (hdr.PRGsize * 2&) - 1, 32768
            CHRswap 0, 0, 8192
        
        Case 69 'Sunsoft FME-7
            PRGswap 0, 0, 8192
            PRGswap 1, (hdr.PRGsize \ 2&) - 3, 8192
            PRGswap 2, (hdr.PRGsize \ 2&) - 2, 8192
            PRGswap 3, (hdr.PRGsize \ 2&) - 1, 8192
            CHRswap 0, 0, 1024
            CHRswap 1, 1, 1024
            CHRswap 2, 2, 1024
            CHRswap 3, 3, 1024
            CHRswap 4, 4, 1024
            CHRswap 5, 5, 1024
            CHRswap 6, 6, 1024
            CHRswap 7, 7, 1024
        
        Case 75, 151 'VRC1
            PRGswap 0, (hdr.PRGsize \ 2&) - 4, 8192
            PRGswap 1, (hdr.PRGsize \ 2&) - 3, 8192
            PRGswap 2, (hdr.PRGsize \ 2&) - 2, 8192
            PRGswap 3, (hdr.PRGsize \ 2&) - 1, 8192
            CHRswap 0, 0, 4096
            CHRswap 1, 1, 4096

        Case 99 'simple VS mapper
            PRGswap 0, (hdr.PRGsize \ 2&) - 4, 8192
            PRGswap 1, (hdr.PRGsize \ 2&) - 3, 8192
            PRGswap 2, (hdr.PRGsize \ 2&) - 2, 8192
            PRGswap 3, (hdr.PRGsize \ 2&) - 1, 8192
            CHRswap 0, 0, 8192
            
        Case 105 'NWC 1990
            PRGswap 0, 0, 32768
            CHRswap 0, 0, 8192
        
        Case 228 'Action 52 / Cheetahmen II
            mapperwrite &H8000, 0

        Case Else
            MsgBox "Mapper " + CStr(hdr.mapper) + " not supported!"
            running = 0
    End Select
End Sub

Public Sub mapperwrite(ByVal addr As Long, ByVal value As Long)
    Dim chrval As Long
    Dim prgval As Long
    Dim chipval As Long
    
    Select Case hdr.mapper
        Case 1 'MMC1
              If (value And 128) Then
                map1.map1reg(0) = (map1.map1reg(0) And &HF3) + &HC 'bits 2, 3 set - others unchanged
                map1.map1bitpos = 0
                map1.map1accum = 0
                Exit Sub
              End If
              map1.map1accum = map1.map1accum Or ((value And 1) * (2 ^ map1.map1bitpos))
              If (map1.map1bitpos = 4) Then
                Select Case addr
                    Case Is >= 57344 '0xE000
                        map1.map1reg(3) = map1.map1accum
                    Case Is >= 49152 '0xC000
                        map1.map1reg(2) = map1.map1accum
                    Case Is >= 40960 '0xA000
                        map1.map1reg(1) = map1.map1accum
                    Case Else
                        map1.map1reg(0) = map1.map1accum
                End Select
                map1calc
                map1.map1bitpos = 0
                map1.map1accum = 0
                Exit Sub
              End If
              map1.map1bitpos = (map1.map1bitpos + 1) Mod 5

        Case 2 'UxROM
            PRGswap 0, value, 16384
        
        Case 3 'CNROM
            CHRswap 0, value And 3, 8192
        
        Case 4, 64, 158 'MMC3
            If ((addr >= 32768) And (addr < 40960)) Then
                If (addr And 1) Then
                    Select Case map4.command
                        Case 0 'select two 1 KB VROM pages at PPU $0000
                            CHRswap map4.chraddrselect \ 1024&, value, 1024
                            CHRswap ((1 * 1024&) Xor map4.chraddrselect) \ 1024&, value + 1, 1024
                        Case 1: 'select two 1 KB VROM pages at PPU $0800
                            CHRswap ((2 * 1024&) Xor map4.chraddrselect) \ 1024&, value, 1024
                            CHRswap ((3 * 1024&) Xor map4.chraddrselect) \ 1024&, value + 1, 1024
                        Case 2 'select 1 KB VROM page at PPU $1000
                            CHRswap ((4 * 1024&) Xor map4.chraddrselect) \ 1024&, value, 1024
                        Case 3 'select 1 KB VROM page at PPU $1400
                            CHRswap ((5 * 1024&) Xor map4.chraddrselect) \ 1024&, value, 1024
                        Case 4 'select 1 KB VROM page at PPU $1800
                            CHRswap ((6 * 1024&) Xor map4.chraddrselect) \ 1024&, value, 1024
                        Case 5 'select 1 KB VROM page at PPU $1C00
                            CHRswap ((7 * 1024&) Xor map4.chraddrselect) \ 1024&, value, 1024
                        Case 6 'select first switchable ROM page
                            map4.prgswitch1 = value 'And ((hdr.PRGsize * 2&) - 1)
                        Case 7 'select second switchable ROM page
                            map4.prgswitch2 = value 'And ((hdr.PRGsize * 2&) - 1)
                    End Select
                    If ((map4.command = 6) Or (map4.command = 7)) Then
                        If (map4.prgaddr) Then
                            PRGswap 0, (hdr.PRGsize * 2&) - 2, 8192&
                            PRGswap 1, map4.prgswitch2, 8192&
                            PRGswap 2, map4.prgswitch1, 8192&
                            PRGswap 3, (hdr.PRGsize * 2&) - 1, 8192&
                        Else
                            PRGswap 0, map4.prgswitch1, 8192&
                            PRGswap 1, map4.prgswitch2, 8192&
                            PRGswap 2, (hdr.PRGsize * 2&) - 2, 8192&
                            PRGswap 3, (hdr.PRGsize * 2&) - 1, 8192&
                        End If
                        Exit Sub
                    End If
                Else
                    map4.command = value And 7
                    map4.prgaddr = value And &H40
                    If (value And &H80) Then map4.chraddrselect = &H1000 Else map4.chraddrselect = &H0
                    Exit Sub
                End If
            ElseIf ((addr >= 40960) And (addr < 49152)) Then
                If (addr And 1) = 0 Then
                    hdr.mirroring = value And 1
                    If (hdr.mirroring) Then
                        PPU.ntmap(0) = &H2000
                        PPU.ntmap(1) = &H2000
                        PPU.ntmap(2) = &H2400
                        PPU.ntmap(3) = &H2400
                    Else
                        PPU.ntmap(0) = &H2000
                        PPU.ntmap(1) = &H2400
                        PPU.ntmap(2) = &H2000
                        PPU.ntmap(3) = &H2400
                    End If
                End If
            ElseIf ((addr >= 49152) And (addr < 57344)) Then
                If (addr And 1) Then map4.irqcounter = map4.irqlatch Else map4.irqlatch = value
            Else
                If (addr And 1) Then map4.irqenable = 1 Else map4.irqenable = 0
            End If

        Case 7 'AxROM
            PRGswap 0, value And &HF, 32768
            If (value And 16) Then
                PPU.ntmap(0) = &H2400
                PPU.ntmap(1) = &H2400
                PPU.ntmap(2) = &H2400
                PPU.ntmap(3) = &H2400
            Else
                PPU.ntmap(0) = &H2000
                PPU.ntmap(1) = &H2000
                PPU.ntmap(2) = &H2000
                PPU.ntmap(3) = &H2000
            End If
        
        Case 9, 10 'MMC2 and MMC4
            Select Case addr
                Case 40969 To 45055 'A000 to AFFF
                    PRGswap 0, value And 15, 8192
                Case 45056 To 49151 'B000 to BFFF
                    map9.latch0_fd = value
                    If (map9.latch1 = &HFD&) Then CHRswap 0, map9.latch0_fd, 4096
                Case 49152 To 53247 'C000 to CFFF
                    map9.latch0_fe = value
                    If (map9.latch1 = &HFE&) Then CHRswap 0, map9.latch0_fe, 4096
                Case 53248 To 57343 'D000 to DFFF
                    map9.latch1_fd = value
                    If (map9.latch2 = &HFD&) Then CHRswap 1, map9.latch1_fd, 4096
                Case 57344 To 61439 'E000 to EFFF
                    map9.latch1_fe = value
                    If (map9.latch2 = &HFE&) Then CHRswap 1, map9.latch1_fe, 4096
                Case Is >= 61440 'F000
                    If (value And 1) Then 'horizontal
                        PPU.ntmap(0) = &H2000&
                        PPU.ntmap(1) = &H2000&
                        PPU.ntmap(2) = &H2400&
                        PPU.ntmap(3) = &H2400&
                    Else
                        PPU.ntmap(0) = &H2000&
                        PPU.ntmap(1) = &H2400&
                        PPU.ntmap(2) = &H2000&
                        PPU.ntmap(3) = &H2400&
                    End If
            End Select
        
        Case 66 'GNROM/MHROM
            PRGswap 0, (value \ 16) And 3, 32768
            CHRswap 0, value And 3, 8192
        
        Case 69 'Sunsoft FME-7
            If addr < 40960 Then
                map69.command = value And 15
            Else
                Select Case map69.command
                    Case 0 To 7
                        CHRswap map69.command, value, 1024
                    Case 8
                        If (value And 128) Then map69.ramenable = True Else map69.ramenable = False
                        If (value And 64) Then map69.RAM = True Else map69.RAM = False
                        map69.prg6000 = value And 63
                    Case 9 To &HB
                        PRGswap map69.command - 9, value And 63, 8192
                    Case &HC
                        Select Case (value And 3)
                            Case 0 'horizontal
                                PPU.ntmap(0) = &H2000&
                                PPU.ntmap(1) = &H2000&
                                PPU.ntmap(2) = &H2400&
                                PPU.ntmap(3) = &H2400&
                            Case 1 'vertical
                                PPU.ntmap(0) = &H2000&
                                PPU.ntmap(1) = &H2400&
                                PPU.ntmap(2) = &H2000&
                                PPU.ntmap(3) = &H2400&
                            Case 2 'one-screen 0x2000
                                PPU.ntmap(0) = &H2000&
                                PPU.ntmap(1) = &H2000&
                                PPU.ntmap(2) = &H2000&
                                PPU.ntmap(3) = &H2000&
                            Case 3 'one-screen 0x2400
                                PPU.ntmap(0) = &H2400&
                                PPU.ntmap(1) = &H2400&
                                PPU.ntmap(2) = &H2400&
                                PPU.ntmap(3) = &H2400&
                        End Select
                End Select
            End If
        
        Case 75, 151 'VRC1
            Select Case addr
                Case &H8000 To &H8FFF
                    PRGswap 0, value And 15, 8192
                Case &H9000 To &H9FFF
                    Select Case (value And 1)
                        Case 0 'horizontal
                            PPU.ntmap(0) = &H2000&
                            PPU.ntmap(1) = &H2000&
                            PPU.ntmap(2) = &H2400&
                            PPU.ntmap(3) = &H2400&
                        Case 1 'vertical
                            PPU.ntmap(0) = &H2000&
                            PPU.ntmap(1) = &H2400&
                            PPU.ntmap(2) = &H2000&
                            PPU.ntmap(3) = &H2400&
                    End Select
                Case &HA000 To &HAFFF
                    PRGswap 1, value And 15, 8192
                Case &HC000 To &HCFFF
                    PRGswap 2, value And 15, 8192
                Case &HE000 To &HEFFF
                    CHRswap 0, value And 15, 4096
                Case Is >= &HF000
                    CHRswap 1, value And 15, 4096
            End Select

        Case 99 'simple VS mapper
            
        Case 228 'Action 52 / Cheetahmen II
            chrval = (addr And 15) * 4& + (value And 3)
            prgval = (addr \ 64&) And 31
            chipval = (addr \ 2048&) And 3
            If chipval > 0 Then chipval = chipval - 1
            If (addr & 32) Then
                PRGswap2 0, (prgval * 16&) + (chipval * 512&), 32768
                'MsgBox "PRGval " + CStr(prgval * 16&) + ", chip val " + CStr(chipval) + ", addr mode 1"
            Else
                PRGswap2 0, (prgval * 16&) + (chipval * 512&), 16384
                PRGswap2 16384, (prgval * 16&) + (chipval * 512&), 16384
                'MsgBox "PRGval " + CStr(prgval * 16&) + ", chip val " + CStr(chipval) + ", addr mode 0"
            End If
            CHRswap 0, chrval, 8192
            
            Select Case ((addr \ 8192&) And 1)
                Case 1 'horizontal
                    PPU.ntmap(0) = &H2000&
                    PPU.ntmap(1) = &H2000&
                    PPU.ntmap(2) = &H2400&
                    PPU.ntmap(3) = &H2400&
                Case 0 'vertical
                    PPU.ntmap(0) = &H2000&
                    PPU.ntmap(1) = &H2400&
                    PPU.ntmap(2) = &H2000&
                    PPU.ntmap(3) = &H2400&
            End Select
    End Select
End Sub

Sub map1calc()
    Select Case (map1.map1reg(0) And 3)
        Case 0:
            PPU.ntmap(0) = &H2000
            PPU.ntmap(1) = &H2000
            PPU.ntmap(2) = &H2000
            PPU.ntmap(3) = &H2000
        Case 1:
            PPU.ntmap(0) = &H2400
            PPU.ntmap(1) = &H2400
            PPU.ntmap(2) = &H2400
            PPU.ntmap(3) = &H2400
        Case 2:
            PPU.ntmap(0) = &H2000
            PPU.ntmap(1) = &H2400
            PPU.ntmap(2) = &H2000
            PPU.ntmap(3) = &H2400
        Case 3:
            PPU.ntmap(0) = &H2000
            PPU.ntmap(1) = &H2000
            PPU.ntmap(2) = &H2400
            PPU.ntmap(3) = &H2400
    End Select
  If (map1.map1reg(0) And 8) Then
    If (map1.map1reg(0) And 4) Then
        PRGswap 0, map1.map1reg(3) And 15, 16384
        PRGswap 1, (hdr.PRGsize \ 16&) - 1, 16384
    Else
        PRGswap 0, 0, 16384
        PRGswap 1, map1.map1reg(3) And 15, 16384
    End If
  Else
    PRGswap 0, (map1.map1reg(3) And 15) \ 2&, 32768
  End If
  If (map1.map1reg(0) And 16) Then
    CHRswap 0, map1.map1reg(1), 4096
    CHRswap 1, map1.map1reg(2), 4096
  Else
    CHRswap 0, map1.map1reg(1) \ 2&, 8192
  End If
End Sub

Public Sub map4irqdecrement()
    If map4.irqcounter > 0 Then
        map4.irqcounter = map4.irqcounter - 1
    Else
        If map4.irqenable Then doirq = 1 ': MsgBox "MMC3 IRQ on scanline " + CStr(curscan)
        map4.irqcounter = map4.irqlatch
    End If
End Sub
