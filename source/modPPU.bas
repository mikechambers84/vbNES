Attribute VB_Name = "modPPU"
Option Explicit

Public Type typePPU
    ntx As Long
    nty As Long
    ntmap(0 To 3) As Long
    addr As Long
    tempaddr As Long
    addrinc As Long
    addrlatch As Long
    nametable As Long
    r2006(0 To 1) As Long
    
    sprtable As Long
    bgtable As Long
    sprvisible As Long
    bgvisible As Long
    sprsize As Long
    nmivblank As Long
    
    xscroll As Long
    yscroll As Long
    tempx As Long
    tempy As Long
    scrolllatch As Long
    
    greyscale As Long
    bgclip As Long
    sprclip As Long
    
    regs(0 To 7) As Long
    sprzero As Long
    sprover As Long
    vblank As Long
    
    bytebuffer As Long
End Type

Public Type typeOAM
    addr As Long
    RAM(0 To 255) As Long
    buf(0 To 255) As Long
    valid(0 To 7) As Long
End Type

Public PPU As typePPU
Public OAM As typeOAM
Public VRAM(0 To 16383) As Long
Public CHRdata(0 To 8191) As Long
Public backgnd(0 To 255) As Long, backgndpixel(0 To 255) As Long
Public sprfront(0 To 255) As Long, sprback(0 To 255) As Long, sprzero(0 To 255) As Long, frontidx(0 To 255) As Long, backidx(0 To 255) As Long

Public NESpal(0 To 63) As Long
Dim lastwritten As Long
Dim lastaddr As Long

Public Sub initPPU()
    Dim n As Long
    For n = 0 To 255
        OAM.RAM(n) = 0
    Next n
    For n = 0 To 16383
        VRAM(n) = 0
    Next n
    PPU.sprsize = 8

   initPalette
End Sub

Public Function readPPU(ByVal addr As Long, ByVal external As Long) As Long
    Dim tempbyte As Long
    Dim oldbyte As Long
    addr = addr And &H3FFF
    
    If hdr.mapper = 4 Then
        If ((lastaddr And &H1000) = 0) And ((addr And &H1000) = &H1000) Then map4irqdecrement
        lastaddr = addr
    End If
    
    If (addr >= &H3F00) Then addr = &H3F00 Or (addr And &H1F)
    If (addr > &H2FFF) And (addr < &H3F00) Then addr = &H3000 Or (addr And &HEFF)
    Select Case addr
        Case &H2000 To &H23FF
            addr = (addr - &H2000) + PPU.ntmap(0)
        Case &H2400 To &H27FF
            addr = (addr - &H2400) + PPU.ntmap(1)
        Case &H2800 To &H2BFF
            addr = (addr - &H2800) + PPU.ntmap(2)
        Case &H2C00 To &H2FFF
            addr = (addr - &H2C00) + PPU.ntmap(3)
    End Select
    If hdr.mapper = 9 Then
        Select Case addr
            Case &HFD8&
                CHRswap 0, map9.latch0_fd, 4096
            Case &HFE8&
                CHRswap 0, map9.latch0_fe, 4096
            Case &H1FD8& To &H1FDF&
                CHRswap 1, map9.latch1_fd, 4096
            Case &H1FE8& To &H1FEF&
                CHRswap 1, map9.latch1_fe, 4096
        End Select
    End If
    If (addr < &H2000) Then
        If external = 1 Then
            tempbyte = CHRbin(CHRbank(addr \ 1024&), addr And 1023&)
            oldbyte = PPU.bytebuffer
            PPU.bytebuffer = tempbyte
            readPPU = oldbyte
        Else
            readPPU = CHRbin(CHRbank(addr \ 1024&), addr And 1023&)
        End If
        Exit Function
    End If
    If (addr >= &H3F00) Or (external = 0) Then tempbyte = VRAM(addr) Else tempbyte = PPU.bytebuffer
    PPU.bytebuffer = VRAM(addr)
    readPPU = tempbyte
End Function

Public Sub writePPU(ByVal addr As Long, ByVal value As Long)
    addr = addr And &H3FFF
    If (addr >= &H3F00) Then addr = &H3F00 Or (addr And &H1F)
    If (addr = &H3F00) Or (addr = &H3F10) Then VRAM(&H3F00) = value: VRAM(&H3F10) = value: Exit Sub
    If (addr = &H3F04) Or (addr = &H3F14) Then VRAM(&H3F04) = value: VRAM(&H3F14) = value: Exit Sub
    If (addr = &H3F08) Or (addr = &H3F18) Then VRAM(&H3F08) = value: VRAM(&H3F18) = value: Exit Sub
    If (addr = &H3F0C) Or (addr = &H3F1C) Then VRAM(&H3F0C) = value: VRAM(&H3F1C) = value: Exit Sub
    If (addr > &H2FFF) And (addr < &H3F00) Then addr = &H3000 Or (addr And &HEFF)
    Select Case addr
        Case &H2000 To &H23FF
            addr = (addr - &H2000) + PPU.ntmap(0)
        Case &H2400 To &H27FF
            addr = (addr - &H2400) + PPU.ntmap(1)
        Case &H2800 To &H2BFF
            addr = (addr - &H2800) + PPU.ntmap(2)
        Case &H2C00 To &H2FFF
            addr = (addr - &H2C00) + PPU.ntmap(3)
    End Select
    If (addr >= &H2000) Then
        VRAM(addr) = value
        'If hdr.mapper = 9 Then
        '    Select Case (addr And &H3FF0&)
        '        Case &HFD0& To &HFDF&: map9.latch1 = &HFD&: CHRswap 0, map9.latch0_fd, 4096
        '        Case &HFE0& To &HFEF&: map9.latch1 = &HFE&: CHRswap 0, map9.latch0_fe, 4096
        '        Case &H1FD0&: map9.latch2 = &HFD&: CHRswap 0, map9.latch1_fd, 4096
        '        Case &H1FE0&: map9.latch2 = &HFE&: CHRswap 0, map9.latch1_fe, 4096
        '    End Select
        'End If
    Else
        If hdr.CHRsize = 0 Then CHRbin(CHRbank(addr \ 1024&), addr And 1023&) = value
    End If
End Sub

Public Function readPPUregs(ByVal addr As Long)
    Dim tempbyte As Long
    'MsgBox "Read PPU reg " + Hex$(addr)
    Select Case addr
        Case &H2002
            tempbyte = (PPU.vblank * 128) Or (PPU.sprzero * 64) Or (PPU.sprover * 32) Or (lastwritten And &H1F)
            PPU.vblank = 0
            PPU.addrlatch = 0
            PPU.scrolllatch = 0
            readPPUregs = tempbyte
            'MsgBox "return " + Hex$(tempbyte)
            If (TRACEPPU = 1) And (PPU.sprzero = 1) Then
                Print #200, "    Read $2002 with sprite 0 hit set"
            End If
            Exit Function
        Case &H2004
            readPPUregs = OAM.RAM(OAM.addr)
            Exit Function
        Case &H2007
            tempbyte = readPPU(PPU.addr, 1)
            PPU.addr = (PPU.addr + PPU.addrinc) And &H3FFF
            readPPUregs = tempbyte
            Exit Function
    End Select
    readPPUregs = 0
End Function

Public Sub writePPUregs(ByVal addr As Long, ByVal value As Long)
    'MsgBox "Write PPU regs " + Hex$(addr) + " = " + Hex$(value)
    PPU.regs(addr And 7) = value
    lastwritten = value
    Select Case addr
        Case &H2000
            If (value And 128) Then PPU.nmivblank = 1 Else PPU.nmivblank = 0
            If (value And 32) Then PPU.sprsize = 16 Else PPU.sprsize = 8
            If (value And 16) Then PPU.bgtable = &H1000 Else PPU.bgtable = 0
            If (value And 8) Then PPU.sprtable = &H1000 Else PPU.sprtable = 0
            If (value And 4) Then PPU.addrinc = 32 Else PPU.addrinc = 1
            PPU.nametable = value And 3
            PPU.tempaddr = (PPU.tempaddr And 62463) Or ((value And 3) * 1024&)
        Case &H2001
            If (value And 16) Then PPU.sprvisible = 1 Else PPU.sprvisible = 0
            If (value And 8) Then PPU.bgvisible = 1 Else PPU.bgvisible = 0
            If (value And 4) Then PPU.sprclip = 0 Else PPU.sprclip = 1
            If (value And 2) Then PPU.bgclip = 0 Else PPU.bgclip = 1
        Case &H2003
            OAM.addr = value
        Case &H2004
            OAM.RAM(OAM.addr) = value
            OAM.addr = (OAM.addr + 1) And 255
        Case &H2005
            If TRACEPPU = 1 Then
                Print #200, "    Write $2005 = $" + Hex$(value)
            End If
            If (PPU.addrlatch = 0) Then
                PPU.tempx = value And 7
                PPU.tempaddr = (PPU.tempaddr And 65504) Or (value \ 8)
                PPU.addrlatch = 1
            Else
                PPU.tempy = value And 7
                PPU.tempaddr = (PPU.tempaddr And 64543) Or ((value And 248) * 4&)
                PPU.addrlatch = 0
            End If
        Case &H2006
            If TRACEPPU = 1 Then
                Print #200, "    Write $2006 = $" + Hex$(value)
            End If
            If (PPU.addrlatch = 0) Then
                PPU.tempaddr = (PPU.tempaddr And 255) Or ((value And 63) * 256&)
                PPU.addrlatch = 1
            Else
                PPU.tempaddr = (PPU.tempaddr And 65280) Or value
                PPU.addr = PPU.tempaddr
                PPU.addrlatch = 0
            End If
        Case &H2007
            writePPU PPU.addr, value
            'MsgBox "writePPU " + Hex$(PPU.addr) + " <- " + Hex$(value)
            PPU.addr = (PPU.addr + PPU.addrinc) And &H3FFF
    End Select
End Sub

Public Function Shr(ByVal orig As Long, ByVal Shift As Long)
    Dim outval As Long
    Dim n As Long
    
    outval = orig
    For n = 1 To Shift
    outval = outval \ 2
    Next n
    Shr = outval
End Function

Public Sub rendersprites(ByVal scanline As Long)
    Dim OAMptr As Long, attr As Long, sprx As Long, spry As Long, table As Long, tile As Long, flipx As Long, flipy As Long, x As Long, startx As Long, plotx As Long, calcx As Long, calcy As Long, patoffset As Long
    Dim curpixel As Long, palette As Long, priority As Long
    Dim valid(0 To 7) As Long
    Dim spridx As Long, drawcount As Long
    Dim n As Long
    Dim patbank As Long
    Dim dummy As Long
    
    spridx = 0
    drawcount = 0
    
    If scanline = 0 Then Exit Sub
    
    For n = 0 To 255
        backidx(n) = 255
        frontidx(n) = 255
    Next n
    OAMptr = 0
    
    If PPU.sprclip Then startx = 8 Else startx = 0
    If PPU.sprsize = 8 Then table = PPU.sprtable
    
    dummy = readPPU(PPU.sprtable, 0) 'this is needed to make MMC3 work when no sprites are visible on the scanline, the PPU reads data from the table anyway on the real NES
    
    For OAMptr = 252 To 0 Step -4
        If ((scanline >= OAM.buf(OAMptr)) And (scanline < (OAM.buf(OAMptr) + PPU.sprsize))) Then
            spry = OAM.buf(OAMptr)
            spry = scanline - spry
            tile = OAM.buf(OAMptr + 1)
            attr = OAM.buf(OAMptr + 2)
            sprx = OAM.buf(OAMptr + 3)
            palette = (attr And 3) * 4
            priority = (attr \ 32) And 1
            flipx = (attr \ 64) And 1
            flipy = (attr \ 128) And 1
            
            drawcount = drawcount + 1
            If drawcount > 8 Then
                PPU.sprover = 1
                'Exit For
            End If

            If (PPU.sprsize = 16) Then
                table = (tile And 1) * 4096&
                tile = tile And &HFE
                If (flipy) Then calcy = (Not spry) And 15 Else calcy = spry And 15
                If (calcy > 7) Then tile = tile + 1
                calcy = calcy And 7
            Else
                If (flipy) Then calcy = (Not spry) And 7 Else calcy = spry And 7
            End If

            For x = 0 To 7
                If (flipx) Then calcx = (Not x) And 7 Else calcx = x
                plotx = sprx + x
                If ((plotx >= startx) And (plotx < 255)) Then
                    patoffset = table + (tile * 16&) + calcy
                    patbank = patoffset \ 1024&
                    'curpixel = Shr(CHRbin(CHRbank(patoffset \ 1024&), patoffset And 1023&), ((Not calcx) And 7)) And 1
                    curpixel = Shr(readPPU(patoffset, 0), ((Not calcx) And 7)) And 1
                    patoffset = patoffset + 8
                    'curpixel = curpixel Or ((Shr(CHRbin(CHRbank(patoffset \ 1024&), patoffset And 1023&), (Not calcx) And 7) And 1) * 2)
                    curpixel = curpixel Or ((Shr(readPPU(patoffset, 0), (Not calcx) And 7) And 1) * 2)
                    If (curpixel > 0) Then
                        If (OAMptr = 0) And (backgndpixel(plotx) > 0) And (scanline < 239) Then
                            If (TRACEPPU = 1) And (PPU.sprzero = 0) Then
                                Print #200, "    sprite 0 hit X = " + CStr(plotx)
                            End If
                            PPU.sprzero = 1
                        End If
                        curpixel = curpixel Or palette
                        If (priority) Then
                            sprback(plotx) = &H3F10 + curpixel
                            backidx(plotx) = OAMptr
                        Else
                            sprfront(plotx) = &H3F10 + curpixel
                            frontidx(plotx) = OAMptr
                        End If
                    End If
                    'If (hdr.mapper = 9) Then 'And (calcy = 0) Then
                    '    If tile = &HFD& Then
                    '        If patbank > 3 Then CHRswap 1, map9.latch1_fd, 4096 Else CHRswap 0, map9.latch0_fd, 4096
                    '    ElseIf tile = &HFE& Then
                    '        If patbank > 3 Then CHRswap 1, map9.latch1_fe, 4096 Else CHRswap 0, map9.latch0_fe, 4096
                    '    End If
                    'End If
                End If
            Next x
        End If
    Next OAMptr
End Sub

Public Sub renderbackground(ByVal scanline As Long)
    'On Error Resume Next
    Dim lastx As Long
    Dim lasty As Long
    Dim startx As Long
    Dim x As Long
    Dim calcx As Long
    Dim calcy As Long
    Dim usent As Long
    Dim tile As Long
    Dim patoffset As Long
    Dim curattrib As Long
    Dim curpixel As Long
    Dim ntbase As Long
    Dim curcolor As Long
    Dim tempval As Long
    Dim patbank As Long
    
    lastx = 65535
    lasty = 65535
    If (PPU.bgclip = 1) Then startx = 8 Else startx = 0
    For x = 0 To 263
      If (skipdraw = 0) Or ((scanline >= OAM.buf(0)) And (scanline < (OAM.buf(0) + PPU.sprsize))) Then
        calcx = ((PPU.addr * 8&) And 255&) Or PPU.xscroll
        calcy = ((PPU.addr \ 4&) And 248&) Or PPU.yscroll
        usent = (PPU.addr \ 1024&) And 3&
        
        'ntbase = &H2000 + &H400& * usent 'nametable base address
        ntbase = PPU.ntmap(usent) 'nametable base address
        tile = VRAM(ntbase + ((calcy And 248) * 4&) + (calcx \ 8&)) 'calculate tile offset based on X,Y coords
        'tile = readPPU(ntbase + ((calcy And 248) * 4&) + (calcx \ 8&), 0) 'calculate tile offset based on X,Y coords
        patoffset = PPU.bgtable + (tile * 16&) + (calcy And 7) 'then turn that into the byte offset in the nametable array
        curattrib = VRAM(ntbase + &H3C0& + ((calcy And 224&) \ 4&) + (calcx \ 32&))
        'curattrib = readPPU(ntbase + &H3C0& + ((calcy And 224&) \ 4&) + (calcx \ 32&), 0)
        Select Case ((calcy \ 8&) And 2) Or ((calcx \ 16&) And 1)
            Case 0: curattrib = curattrib And 3
            Case 1: curattrib = (curattrib \ 4&) And 3
            Case 2: curattrib = (curattrib \ 16&) And 3
            Case 3: curattrib = (curattrib \ 64&) And 3
        End Select
        curattrib = curattrib * 4&
        patbank = patoffset \ 1024&

        'curpixel = Shr(CHRbin(CHRbank(patoffset \ 1024&), patoffset And 1023&), (Not calcx) And 7) And 1
        curpixel = Shr(readPPU(patoffset, 0), (Not calcx) And 7) And 1
        patoffset = patoffset + 8
        'curpixel = curpixel Or ((Shr(CHRbin(CHRbank(patoffset \ 1024&), patoffset And 1023&), (Not calcx) And 7) And 1) * 2)
        curpixel = curpixel Or ((Shr(readPPU(patoffset, 0), (Not calcx) And 7) And 1) * 2)
        
        If (curpixel > 0) And (x >= startx) And (x < 256) Then
            backgndpixel(x) = curpixel
            curcolor = curpixel Or curattrib
            'frmMain.pic.PSet (x, scanline), NESpal(curcolor)
            backgnd(x) = &H3F00 + curcolor
        End If
            
        'If hdr.mapper = 9 Then
            'patoffset = patoffset - 8
            'If tile = &HFD& Then
            '    If patbank > 3 Then CHRswap 1, map9.latch1_fd, 4096 Else CHRswap 0, map9.latch0_fd, 4096
            'ElseIf tile = &HFE& Then
            '    If patbank > 3 Then CHRswap 1, map9.latch1_fe, 4096 Else CHRswap 0, map9.latch0_fe, 4096
            'End If
        'End If
      End If
        PPU.xscroll = PPU.xscroll + 1
        If PPU.xscroll = 8 Then
            PPU.xscroll = 0
            tempval = (PPU.addr And &H1F) + 1
            If (tempval = 32) Then PPU.addr = PPU.addr Xor 1024
            PPU.addr = (PPU.addr And 65504) Or (tempval And &H1F)
        End If
    Next x
    PPU.yscroll = PPU.yscroll + 1
    If (PPU.yscroll = 8) Then
        PPU.yscroll = 0
        tempval = ((PPU.addr \ 32) And &H1F) + 1
        If (tempval = 30) Then
            tempval = 0
            PPU.addr = PPU.addr Xor 2048
        End If
        If (tempval = 32) Then tempval = 0
        PPU.addr = (PPU.addr And 64543) Or ((tempval And &H1F) * 32)
    End If
End Sub

Public Sub renderscanline(ByVal scanline As Long)
    Dim n As Long
    Dim tmpx As Long
    
    PPU.sprover = 0

    For n = 0 To 255
        backgnd(n) = 0
        backgndpixel(n) = 0
        sprback(n) = 0
        sprfront(n) = 0
        sprzero(n) = 0
    Next n
    
    For n = 0 To 255
        OAM.buf(n) = OAM.RAM(n)
    Next n
    If (PPU.bgvisible) Then renderbackground scanline
    exec6502 85: If doirq = 1 Then irq6502: PPU.yscroll = 0: doirq = 0
    If skipdraw = 0 Then
        If (PPU.sprvisible) Then rendersprites scanline - 1
        If doirq = 1 Then irq6502: PPU.yscroll = 0: doirq = 0
        For tmpx = 0 To 255
            If (sprback(tmpx) = 0) Then outputNES(scanline, tmpx) = VRAM(&H3F10) And &H3F Else outputNES(scanline, tmpx) = VRAM(sprback(tmpx)) And &H3F
            If (backgnd(tmpx) > 0) Then
                outputNES(scanline, tmpx) = VRAM(backgnd(tmpx)) And &H3F
                'If (sprzero(tmpx)) Then PPU.sprzero = 1
            End If
            If ((sprfront(tmpx) > 0) And (frontidx(tmpx) < backidx(tmpx))) Then outputNES(scanline, tmpx) = VRAM(sprfront(tmpx)) And &H3F
        Next tmpx
    Else
        If (scanline >= OAM.buf(0)) And (scanline < (OAM.buf(0) + PPU.sprsize)) Then
            If (PPU.sprvisible) Then rendersprites scanline
            If doirq = 1 Then irq6502: PPU.yscroll = 0: doirq = 0
        End If
    End If
    If (scanline Mod 3) = 0 Then exec6502 28 Else exec6502 29
End Sub

