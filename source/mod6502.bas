Attribute VB_Name = "mod6502"
Option Explicit

Public Const DBG = 0

Const UNDOCUMENTED = 1
Const NES_CPU = 1

Const FLAG_CARRY = 1
Const FLAG_ZERO = 2
Const FLAG_INTERRUPT = 4
Const FLAG_DECIMAL = 8
Const FLAG_BREAK = 16
Const FLAG_CONSTANT = 32
Const FLAG_OVERFLOW = 64
Const FLAG_SIGN = 128

Const BASE_STACK = 256

Dim ticktable(0 To 255) As Long

'6502 CPU registers
Public pc As Long
Public sp As Long
Public a As Long
Public x As Long
Public y As Long
Public status As Long


'helper variables
Public instructions As Long 'keep track of total instructions executed
Dim clockticks6502 As Long
Dim clockgoal6502 As Long
Public realticks6502 As Long
Dim oldpc As Long
Dim ea As Long
Dim reladdr As Long
Dim value As Long
Dim result As Long
Dim opcode As Byte
Dim oldstatus As Byte
Dim penaltyop As Byte
Dim penaltyaddr As Byte
Dim isacc As Byte
Dim temp16 As Long
Dim eahelp As Long
Dim eahelp2 As Long
Dim startpage As Long

Private Sub saveaccum(ByVal n As Byte)
    a = n
End Sub


'flag modifiers
Private Sub setcarry()
    status = status Or FLAG_CARRY
End Sub

Private Sub clearcarry()
    status = status And (Not FLAG_CARRY)
End Sub

Private Sub setzero()
    status = status Or FLAG_ZERO
End Sub

Private Sub clearzero()
    status = status And (Not FLAG_ZERO)
End Sub

Private Sub setinterrupt()
    status = status Or FLAG_INTERRUPT
End Sub

Private Sub clearinterrupt()
    status = status And (Not FLAG_INTERRUPT)
End Sub

Private Sub setdecimal()
    status = status Or FLAG_DECIMAL
End Sub

Private Sub cleardecimal()
    status = status And (Not FLAG_DECIMAL)
End Sub

Private Sub setoverflow()
    status = status Or FLAG_OVERFLOW
End Sub

Private Sub clearoverflow()
    status = status And (Not FLAG_OVERFLOW)
End Sub

Private Sub setsign()
    status = status Or FLAG_SIGN
End Sub

Private Sub clearsign()
    status = status And (Not FLAG_SIGN)
End Sub


'flag calculations
Private Sub zerocalc(ByVal n As Long)
    n = n And 255
    If n > 0 Then clearzero Else setzero
End Sub

Private Sub signcalc(ByVal n As Long)
    If (n And 128) Then setsign Else clearsign
End Sub

Private Sub carrycalc(ByVal n As Long)
    If (n And 65280) Then setcarry Else clearcarry
End Sub

Private Sub overflowcalc(ByVal n As Long, ByVal m As Long, ByVal o As Long)  'n = result, m = accumulator, o = memory
    If ((n Xor m) And (n Xor o) And 128) Then setoverflow Else clearoverflow
End Sub

'a few general functions used by various other functions
Private Sub push16(ByVal pushval As Long)
    write6502 BASE_STACK + sp, pushval \ 256
    write6502 BASE_STACK + ((sp - 1) And 255), pushval And 255
    sp = (sp - 2) And 255
End Sub

Private Sub push8(ByVal pushval As Byte)
    write6502 BASE_STACK + sp, pushval
    sp = (sp - 1) And 255
End Sub

Private Function pull16() As Long
    temp16 = read6502(BASE_STACK + ((sp + 1) And 255)) Or (read6502(BASE_STACK + ((sp + 2) And 255)) * 256&)
    sp = (sp + 2) And 255
    pull16 = temp16
End Function

Private Function pull8() As Byte
    sp = (sp + 1) And 255
    pull8 = read6502(BASE_STACK + sp)
End Function

Public Sub reset6502()
    pc = read6502(65532) Or (read6502(65533) * 256&)
    a = 0
    x = 0
    y = 0
    sp = 253
    status = status Or FLAG_CONSTANT
    initticks
End Sub

'addressing mode functions, calculates effective addresses
Private Sub acc()
    isacc = 1
End Sub

Private Sub impl()
End Sub

Private Sub imm()  'immediate
    ea = pc
    pc = (pc + 1) And 65535
End Sub

Private Sub zp() 'zero-page
    ea = read6502(pc)
    pc = (pc + 1) And 65535
End Sub

Private Sub zpx() 'zero-page,X
    ea = (read6502(pc) + x) And 255 'zero-page wraparound
    pc = (pc + 1) And 65535
End Sub

Private Sub zpy() 'zero-page,Y
    ea = (read6502(pc) + y) And 255 'zero-page wraparound
    pc = (pc + 1) And 65535
End Sub

Private Sub rel() 'relative for branch ops (8-bit immediate value, sign-extended)
    reladdr = read6502(pc)
    pc = (pc + 1) And 65535
    If (reladdr And 128) Then reladdr = reladdr Or 65280
End Sub

Private Sub abso() 'absolute
    ea = read6502(pc) Or (read6502(pc + 1) * 256&)
    pc = (pc + 2) And 65535
End Sub

Private Sub absx() 'absolute,X
    ea = (read6502(pc) Or (read6502(pc + 1) * 256&))
    startpage = ea And 65280
    ea = (ea + x) And 65535

    If (startpage <> (ea And 65280)) Then 'one cycle penlty for page-crossing on some opcodes
        penaltyaddr = 1
    End If

    pc = (pc + 2) And 65535
End Sub

Private Sub absy() 'absolute,Y
    ea = (read6502(pc) Or (read6502(pc + 1) * 256&))
    startpage = ea And 65280
    ea = (ea + y) And 65535

    If (startpage <> (ea And 65280)) Then 'one cycle penlty for page-crossing on some opcodes
        penaltyaddr = 1
    End If

    pc = (pc + 2) And 65535
End Sub

Private Sub ind() 'indirect
    eahelp = read6502(pc) Or (read6502(pc + 1) * 256&)
    eahelp2 = (eahelp And 65280) Or ((eahelp + 1) And 255)   'replicate 6502 page-boundary wraparound bug
    ea = read6502(eahelp) Or (read6502(eahelp2) * 256&)
    pc = (pc + 2) And 65535
End Sub

Private Sub indx() ' (indirect,X)
    eahelp = ((read6502(pc) + x) And 255) 'zero-page wraparound for table pointer
    ea = read6502(eahelp And 255) Or (read6502((eahelp + 1) And 255) * 256&)
    pc = (pc + 1) And 65535
End Sub

Private Sub indy() ' (indirect),Y
    eahelp = read6502(pc)
    eahelp2 = (eahelp And 65280) Or ((eahelp + 1) And 255) 'zero-page wraparound
    ea = read6502(eahelp) Or (read6502(eahelp2) * 256&)
    startpage = ea And 65280
    ea = (ea + y) And 65535

    If (startpage <> (ea And 65280)) Then 'one cycle penlty for page-crossing on some opcodes
        penaltyaddr = 1
    End If
    
    pc = (pc + 1) And 65535
End Sub

Private Function getvalue() As Byte
    If (isacc = 1) Then getvalue = a Else getvalue = read6502(ea)
End Function

Private Function getvalue16() As Long
    getvalue16 = read6502(ea) Or (read6502(ea + 1) * 256&)
End Function

Private Sub putvalue(ByVal saveval As Byte)
    If (isacc = 1) Then a = saveval Else write6502 ea, saveval
End Sub


'instruction handler functions
Private Sub adc()
    penaltyop = 1
    value = getvalue()
    result = a + value + (status And FLAG_CARRY)
   
    carrycalc result
    zerocalc result
    overflowcalc result, a, value
    signcalc result
    
    If NES_CPU = 0 Then
    If (status And FLAG_DECIMAL) Then
        clearcarry
        
        If ((a And 15) > 9) Then
            a = (a + 6) And 255
        End If
        If ((a And 240) > 144) Then
            a = (a + 96) And 255
            setcarry
        End If
        
        clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1
    End If
    End If
   
    saveaccum result And 255
End Sub

Private Sub opand()
    penaltyop = 1
    value = getvalue()
    result = a And value
   
    zerocalc result
    signcalc result
   
    saveaccum result
End Sub

Private Sub asl()
    value = getvalue()
    result = value * 2

    carrycalc result
    zerocalc result
    signcalc result
   
    putvalue result And 255
End Sub

Private Sub bcc()
    If ((status And FLAG_CARRY) = 0) Then
        oldpc = pc
        pc = (pc + reladdr) And 65535
        If ((oldpc And 65280) <> (pc And 65280)) Then clockticks6502 = clockticks6502 + 2: realticks6502 = realticks6502 + 1 Else clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1 'check if jump crossed a page boundary
    End If
End Sub

Private Sub bcs()
    If ((status And FLAG_CARRY) = FLAG_CARRY) Then
        oldpc = pc
        pc = (pc + reladdr) And 65535
        If ((oldpc And 65280) <> (pc And 65280)) Then clockticks6502 = clockticks6502 + 2: realticks6502 = realticks6502 + 1 Else clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1 'check if jump crossed a page boundary
    End If
End Sub

Private Sub beq()
    If ((status And FLAG_ZERO) = FLAG_ZERO) Then
        oldpc = pc
        pc = (pc + reladdr) And 65535
        If ((oldpc And 65280) <> (pc And 65280)) Then clockticks6502 = clockticks6502 + 2: realticks6502 = realticks6502 + 1 Else clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1 'check if jump crossed a page boundary
    End If
End Sub

Private Sub bit()
    value = getvalue()
    result = a And value
   
    zerocalc result
    status = (status And &H3F) Or (value And &HC0)
End Sub

Private Sub bmi()
    If ((status And FLAG_SIGN) = FLAG_SIGN) Then
        oldpc = pc
        pc = (pc + reladdr) And 65535
        If ((oldpc And 65280) <> (pc And 65280)) Then clockticks6502 = clockticks6502 + 2: realticks6502 = realticks6502 + 1 Else clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1 'check if jump crossed a page boundary
    End If
End Sub

Private Sub bne()
    If ((status And FLAG_ZERO) = 0) Then
        oldpc = pc
        pc = (pc + reladdr) And 65535
        If ((oldpc And 65280) <> (pc And 65280)) Then clockticks6502 = clockticks6502 + 2: realticks6502 = realticks6502 + 1 Else clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1 'check if jump crossed a page boundary
    End If
End Sub

Private Sub bpl()
    If ((status And FLAG_SIGN) = 0) Then
        oldpc = pc
        pc = (pc + reladdr) And 65535
        If ((oldpc And 65280) <> (pc And 65280)) Then clockticks6502 = clockticks6502 + 2: realticks6502 = realticks6502 + 1 Else clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1 'check if jump crossed a page boundary
    End If
End Sub

Private Sub brk()
    pc = (pc + 1) And 65535
    push16 pc 'push next instruction address onto stack
    push8 status Or FLAG_BREAK 'push CPU status to stack
    setinterrupt 'set interrupt flag
    pc = read6502(65534) Or (read6502(65535) * 256&)
End Sub

Private Sub bvc()
    If ((status And FLAG_OVERFLOW) = 0) Then
        oldpc = pc
        pc = (pc + reladdr) And 65535
        If ((oldpc And 65280) <> (pc And 65280)) Then clockticks6502 = clockticks6502 + 2: realticks6502 = realticks6502 + 1 Else clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1 'check if jump crossed a page boundary
    End If
End Sub

Private Sub bvs()
    If ((status And FLAG_OVERFLOW) = FLAG_OVERFLOW) Then
        oldpc = pc
        pc = (pc + reladdr) And 65535
        If ((oldpc And 65280) <> (pc And 65280)) Then clockticks6502 = clockticks6502 + 2: realticks6502 = realticks6502 + 1 Else clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1 'check if jump crossed a page boundary
    End If
End Sub

Private Sub clc()
    clearcarry
End Sub

Private Sub cld()
    cleardecimal
End Sub

Private Sub cli()
    clearinterrupt
End Sub

Private Sub clv()
    clearoverflow
End Sub

Private Sub cmp()
    penaltyop = 1
    value = getvalue()
    result = (a - value) And 65535
   
    If (a >= (value And 255)) Then setcarry Else clearcarry
    If (a = (value And 255)) Then setzero Else clearzero
    signcalc result And 255
End Sub

Private Sub cpx()
    value = getvalue()
    result = (x - value) And 65535
   
    If (x >= (value And 255)) Then setcarry Else clearcarry
    If (x = (value And 255)) Then setzero Else clearzero
    signcalc result And 255
End Sub

Private Sub cpy()
    value = getvalue()
    result = (y - value) And 65535
   
    If (y >= (value And 255)) Then setcarry Else clearcarry
    If (y = (value And 255)) Then setzero Else clearzero
    signcalc result And 255
End Sub

Private Sub dec()
    value = getvalue()
    result = (value - 1) And 65535
   
    zerocalc result
    signcalc result
   
    putvalue result And 255
End Sub

Private Sub dex()
    x = (x - 1) And 255
   
    zerocalc x
    signcalc x
End Sub

Private Sub dey()
    y = (y - 1) And 255
   
    zerocalc y
    signcalc y
End Sub

Private Sub eor()
    penaltyop = 1
    value = getvalue()
    result = a Xor value
   
    zerocalc result
    signcalc result
   
    saveaccum result
End Sub

Private Sub inc()
    value = getvalue()
    result = (value + 1) And 65535
   
    zerocalc result
    signcalc result
   
    putvalue result And 255
End Sub

Private Sub inx()
    x = (x + 1) And 255
   
    zerocalc x
    signcalc x
End Sub

Private Sub iny()
    y = (y + 1) And 255
   
    zerocalc y
    signcalc y
End Sub

Private Sub jmp()
    pc = ea
End Sub

Private Sub jsr()
    push16 (pc - 1) And 65535
    pc = ea
End Sub

Private Sub lda()
    penaltyop = 1
    value = getvalue()
    a = value
   
    zerocalc a
    signcalc a
End Sub

Private Sub ldx()
    penaltyop = 1
    value = getvalue()
    x = value
   
    zerocalc x
    signcalc x
End Sub

Private Sub ldy()
    penaltyop = 1
    value = getvalue()
    y = value
   
    zerocalc y
    signcalc y
End Sub

Private Sub lsr()
    value = getvalue()
    result = value \ 2
   
    If (value And 1) Then setcarry Else clearcarry
    zerocalc result
    signcalc result
   
    putvalue result
End Sub

Private Sub nop()
    Select Case opcode
        Case &H1C, &H3C, &H5C, &H7C, &HDC, &HFC
            penaltyop = 1
    End Select
End Sub

Private Sub ora()
    penaltyop = 1
    value = getvalue()
    result = a Or value
   
    zerocalc result
    signcalc result
   
    saveaccum result
End Sub

Private Sub pha()
    push8 a
End Sub

Private Sub php()
    push8 status Or FLAG_BREAK
End Sub

Private Sub pla()
    a = pull8()
   
    zerocalc a
    signcalc a
End Sub

Private Sub plp()
    status = pull8() Or FLAG_CONSTANT
End Sub

Private Sub rol()
    value = getvalue()
    result = ((value * 2) Or (status And FLAG_CARRY)) And 65535
   
    carrycalc result
    zerocalc result
    signcalc result
   
    putvalue result And 255
End Sub

Private Sub ror()
    value = getvalue()
    result = (value \ 2) Or ((status And FLAG_CARRY) * 128)
   
    If (value And 1) Then setcarry Else clearcarry
    zerocalc result
    signcalc result
   
    putvalue result And 255
End Sub

Private Sub rti()
    status = pull8()
    value = pull16()
    pc = value
End Sub

Private Sub rts()
    value = pull16()
    pc = (value + 1) And 65535
End Sub

Private Sub sbc()
    penaltyop = 1
    value = getvalue() Xor 255
    result = (a + value + (status And FLAG_CARRY)) And 65535
   
    carrycalc result
    zerocalc result
    overflowcalc result, a, value
    signcalc result

    If NES_CPU = 0 Then
        If (status And FLAG_DECIMAL) Then
            clearcarry
        
            a = (a - 102) And 255
            If ((a And 15) > 9) Then
                a = (a + 6) And 255
            End If
            If ((a And &HF0) > &H90) Then
                a = (a + 96) And 255
                setcarry
            End If
        
            clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1
        End If
    End If
   
    saveaccum result And 255
End Sub

Private Sub sec()
    setcarry
End Sub

Private Sub sed()
    setdecimal
End Sub

Private Sub sei()
    setinterrupt
End Sub

Private Sub sta()
    putvalue a
End Sub

Private Sub stx()
    putvalue x
End Sub

Private Sub sty()
    putvalue y
End Sub

Private Sub tax()
    x = a
   
    zerocalc x
    signcalc x
End Sub

Private Sub tay()
    y = a
   
    zerocalc y
    signcalc y
End Sub

Private Sub tsx()
    x = sp
   
    zerocalc x
    signcalc x
End Sub

Private Sub txa()
    a = x
   
    zerocalc a
    signcalc a
End Sub

Private Sub txs()
    sp = x
End Sub

Private Sub tya()
    a = y
   
    zerocalc a
    signcalc a
End Sub

'undefined ops
Private Sub lax()
    lda
    ldx
End Sub

Private Sub sax()
    sta
    stx
    putvalue a And x
    If (penaltyop And penaltyaddr) Then clockticks6502 = clockticks6502 - 1: realticks6502 = realticks6502 - 1
End Sub

Private Sub dcp()
    dec
    cmp
    If (penaltyop And penaltyaddr) Then clockticks6502 = clockticks6502 - 1: realticks6502 = realticks6502 - 1
End Sub

Private Sub isb()
    inc
    sbc
    If (penaltyop And penaltyaddr) Then clockticks6502 = clockticks6502 - 1: realticks6502 = realticks6502 - 1
End Sub

Private Sub slo()
    asl
    ora
    If (penaltyop And penaltyaddr) Then clockticks6502 = clockticks6502 - 1: realticks6502 = realticks6502 - 1
End Sub

Private Sub rla()
    rol
    opand
    If (penaltyop And penaltyaddr) Then clockticks6502 = clockticks6502 - 1: realticks6502 = realticks6502 - 1
End Sub

Private Sub sre()
    lsr
    eor
    If (penaltyop And penaltyaddr) Then clockticks6502 = clockticks6502 - 1: realticks6502 = realticks6502 - 1
End Sub

Private Sub rra()
    ror
    adc
    If (penaltyop And penaltyaddr) Then clockticks6502 = clockticks6502 - 1: realticks6502 = realticks6502 - 1
End Sub

Public Sub nmi6502()
    status = status And (Not FLAG_BREAK)
    push16 pc
    push8 status
    status = status Or FLAG_INTERRUPT
    pc = read6502(65530) Or (read6502(65531) * 256&)
End Sub

Public Sub irq6502()
    If (status And FLAG_INTERRUPT) = 0 Then
        status = status And (Not FLAG_BREAK)
        push16 pc
        push8 status
        status = status Or FLAG_INTERRUPT
        pc = read6502(65534) Or (read6502(65535) * 256&)
    End If
End Sub

Public Sub exec6502(ByVal tickcount As Long)
    While (clockticks6502 < tickcount) '(clockticks6502 < clockgoal6502)
        pc = pc And 65535
                
        opcode = read6502(pc)
        If DBG = 1 Then
            Print #2, CStr(pc) + " " + CStr(opcode) + " " + CStr(a) + " " + CStr(x) + " " + CStr(y) + " " + CStr(sp) + " " + CStr(status)
        End If
        
        pc = (pc + 1) And 65535
        status = status Or FLAG_CONSTANT

        penaltyop = 0
        penaltyaddr = 0
        isacc = 0
                
        'If a < 0 Then a = 256 - a
        'If x < 0 Then x = 256 - x
        'If y < 0 Then y = 256 - y
        'If sp < 0 Then sp = 256 - sp
        'a = a And 255
        'x = x And 255
        'y = y And 255
        'sp = sp And 255
        
        Select Case opcode
            Case &H0
                impl
                brk
            Case &H1
                indx
                ora
            Case &H2
                impl
                nop
            Case &H3
                indx
                slo
            Case &H4
                zp
                nop
            Case &H5
                zp
                ora
            Case &H6
                zp
                asl
            Case &H7
                zp
                slo
            Case &H8
                impl
                php
            Case &H9
                imm
                ora
            Case &HA
                acc
                asl
            Case &HB
                imm
                nop
            Case &HC
                abso
                nop
            Case &HD
                abso
                ora
            Case &HE
                abso
                asl
            Case &HF
                abso
                slo
            Case &H10
                rel
                bpl
            Case &H11
                indy
                ora
            Case &H12
                impl
                nop
            Case &H13
                indy
                slo
            Case &H14
                zpx
                nop
            Case &H15
                zpx
                ora
            Case &H16
                zpx
                asl
            Case &H17
                zpx
                slo
            Case &H18
                impl
                clc
            Case &H19
                absy
                ora
            Case &H1A
                impl
                nop
            Case &H1B
                absy
                slo
            Case &H1C
                absx
                nop
            Case &H1D
                absx
                ora
            Case &H1E
                absx
                asl
            Case &H1F
                absx
                slo
            Case &H20
                abso
                jsr
            Case &H21
                indx
                opand
            Case &H22
                impl
                nop
            Case &H23
                indx
                rla
            Case &H24
                zp
                bit
            Case &H25
                zp
                opand
            Case &H26
                zp
                rol
            Case &H27
                zp
                rla
            Case &H28
                impl
                plp
            Case &H29
                imm
                opand
            Case &H2A
                acc
                rol
            Case &H2B
                imm
                nop
            Case &H2C
                abso
                bit
            Case &H2D
                abso
                opand
            Case &H2E
                abso
                rol
            Case &H2F
                abso
                rla
            Case &H30
                rel
                bmi
            Case &H31
                indy
                opand
            Case &H32
                impl
                nop
            Case &H33
                indy
                rla
            Case &H34
                zpx
                nop
            Case &H35
                zpx
                opand
            Case &H36
                zpx
                rol
            Case &H37
                zpx
                rla
            Case &H38
                impl
                sec
            Case &H39
                absy
                opand
            Case &H3A
                impl
                nop
            Case &H3B
                absy
                rla
            Case &H3C
                absx
                nop
            Case &H3D
                absx
                opand
            Case &H3E
                absx
                rol
            Case &H3F
                absx
                rla
            Case &H40
                impl
                rti
            Case &H41
                indx
                eor
            Case &H42
                impl
                nop
            Case &H43
                indx
                sre
            Case &H44
                zp
                nop
            Case &H45
                zp
                eor
            Case &H46
                zp
                lsr
            Case &H47
                zp
                sre
            Case &H48
                impl
                pha
            Case &H49
                imm
                eor
            Case &H4A
                acc
                lsr
            Case &H4B
                imm
                nop
            Case &H4C
                abso
                jmp
            Case &H4D
                abso
                eor
            Case &H4E
                abso
                lsr
            Case &H4F
                abso
                sre
            Case &H50
                rel
                bvc
            Case &H51
                indy
                eor
            Case &H52
                impl
                nop
            Case &H53
               indy
               sre
            Case &H54
                zpx
                nop
            Case &H55
                zpx
                eor
            Case &H56
                zpx
                lsr
            Case &H57
                zpx
                sre
            Case &H58
                impl
                cli
            Case &H59
                absy
                eor
            Case &H5A
                impl
                nop
            Case &H5B
                absy
                sre
            Case &H5C
                absx
                nop
            Case &H5D
                absx
                eor
            Case &H5E
                absx
                lsr
            Case &H5F
                absx
                sre
            Case &H60
                impl
                rts
            Case &H61
                indx
                adc
            Case &H62
                impl
                nop
            Case &H63
                indx
                rra
            Case &H64
                zp
                nop
            Case &H65
                zp
                adc
            Case &H66
                zp
                ror
            Case &H67
                zp
                rra
            Case &H68
                impl
                pla
            Case &H69
                imm
                adc
            Case &H6A
                acc
                ror
            Case &H6B
                imm
                nop
            Case &H6C
                ind
                jmp
            Case &H6D
                abso
                adc
            Case &H6E
                abso
                ror
            Case &H6F
                abso
                rra
            Case &H70
                rel
                bvs
            Case &H71
                indy
                adc
            Case &H72
                impl
                nop
            Case &H73
                indy
                rra
            Case &H74
                zpx
                nop
            Case &H75
                zpx
                adc
            Case &H76
                zpx
                ror
            Case &H77
                zpx
                rra
            Case &H78
                impl
                sei
            Case &H79
                absy
                adc
            Case &H7A
                impl
                nop
            Case &H7B
                absy
                rra
            Case &H7C
                absx
                nop
            Case &H7D
                absx
                adc
            Case &H7E
                absx
                ror
            Case &H7F
                absx
                rra
            Case &H80
                imm
                nop
            Case &H81
                indx
                sta
            Case &H82
                imm
                nop
            Case &H83
                indx
                sax
            Case &H84
                zp
                sty
            Case &H85
                zp
                sta
            Case &H86
                zp
                stx
            Case &H87
               zp
               sax
            Case &H88
                impl
                dey
            Case &H89
                imm
                nop
            Case &H8A
                impl
                txa
            Case &H8B
                imm
                nop
            Case &H8C
                abso
                sty
            Case &H8D
                abso
                sta
            Case &H8E
                abso
                stx
            Case &H8F
                abso
                sax
            Case &H90
                rel
                bcc
            Case &H91
                indy
                sta
            Case &H92
                impl
                nop
            Case &H93
                indy
                nop
            Case &H94
                zpx
                sty
            Case &H95
                zpx
                sta
            Case &H96
                zpy
                stx
            Case &H97
                zpy
                sax
            Case &H98
                impl
                tya
            Case &H99
                absy
                sta
            Case &H9A
                impl
                txs
            Case &H9B
                absy
                nop
            Case &H9C
                absx
                nop
            Case &H9D
                absx
                sta
            Case &H9E
                absy
                nop
            Case &H9F
                absy
                nop
            Case &HA0
                imm
                ldy
            Case &HA1
                indx
                lda
            Case &HA2
                imm
                ldx
            Case &HA3
                indx
                lax
            Case &HA4
                zp
                ldy
            Case &HA5
                zp
                lda
            Case &HA6
                zp
                ldx
            Case &HA7
                zp
                lax
            Case &HA8
                impl
                tay
            Case &HA9
                imm
                lda
            Case &HAA
                impl
                tax
            Case &HAB
                imm
                nop
            Case &HAC
                abso
                ldy
            Case &HAD
                abso
                lda
            Case &HAE
                abso
                ldx
            Case &HAF
                abso
                lax
            Case &HB0
                rel
                bcs
            Case &HB1
                indy
                lda
            Case &HB2
                impl
                nop
            Case &HB3
                indy
                lax
            Case &HB4
                zpx
                ldy
            Case &HB5
                zpx
                lda
            Case &HB6
                zpy
                ldx
            Case &HB7
                zpy
                lax
            Case &HB8
                impl
                clv
            Case &HB9
                absy
                lda
            Case &HBA
                impl
                tsx
            Case &HBB
                absy
                lax
            Case &HBC
                absx
                ldy
            Case &HBD
                absx
                lda
            Case &HBE
                absy
                ldx
            Case &HBF
               absy
               lax
            Case &HC0
                imm
                cpy
            Case &HC1
                indx
                cmp
            Case &HC2
                imm
                nop
            Case &HC3
                indx
                dcp
            Case &HC4
                zp
                cpy
            Case &HC5
                zp
                cmp
            Case &HC6
                zp
                dec
            Case &HC7
                zp
                dcp
            Case &HC8
                impl
                iny
            Case &HC9
                imm
                cmp
            Case &HCA
                impl
                dex
            Case &HCB
                imm
                nop
            Case &HCC
                abso
                cpy
            Case &HCD
                abso
                cmp
            Case &HCE
                abso
                dec
            Case &HCF
               abso
               dcp
            Case &HD0
                rel
                bne
            Case &HD1
                indy
                cmp
            Case &HD2
                impl
                nop
            Case &HD3
                indy
                dcp
            Case &HD4
                zpx
                nop
            Case &HD5
                zpx
                cmp
            Case &HD6
                zpx
                dec
            Case &HD7
                zpx
                dcp
            Case &HD8
                impl
                cld
            Case &HD9
                absy
                cmp
            Case &HDA
                impl
                nop
            Case &HDB
               absy
               dcp
            Case &HDC
                absx
                nop
            Case &HDD
                absx
                cmp
            Case &HDE
                absx
                dec
            Case &HDF
                absx
                dcp
            Case &HE0
                imm
                cpx
            Case &HE1
                indx
                sbc
            Case &HE2
                imm
                nop
            Case &HE3
                indx
                isb
            Case &HE4
                zp
                cpx
            Case &HE5
                zp
                sbc
            Case &HE6
                zp
                inc
            Case &HE7
                zp
                isb
            Case &HE8
                impl
                inx
            Case &HE9
                imm
                sbc
            Case &HEA
                impl
                nop
            Case &HEB
                imm
                sbc
            Case &HEC
                abso
                cpx
            Case &HED
                abso
                sbc
            Case &HEE
                abso
                inc
            Case &HEF
                abso
                isb
            Case &HF0
                rel
                beq
            Case &HF1
                indy
                sbc
            Case &HF2
                impl
                nop
            Case &HF3
                indy
                isb
            Case &HF4
                zpx
                nop
            Case &HF5
                zpx
                sbc
            Case &HF6
                zpx
                inc
            Case &HF7
                zpx
                isb
            Case &HF8
                impl
                sed
            Case &HF9
                absy
                sbc
            Case &HFA
                impl
                nop
            Case &HFB
                absy
                isb
            Case &HFC
                absx
                nop
            Case &HFD
                absx
                sbc
            Case &HFE
                absx
                inc
            Case &HFF
                absx
                isb
        End Select
        
        clockticks6502 = clockticks6502 + ticktable(opcode)
        realticks6502 = realticks6502 + ticktable(opcode)
        If (penaltyop And penaltyaddr) Then clockticks6502 = clockticks6502 + 1: realticks6502 = realticks6502 + 1

        instructions = instructions + 1
        
        tickchannelsAPU
    Wend
    clockticks6502 = clockticks6502 - tickcount
End Sub

Private Sub initticks()
    ticktable(&H0) = 7
    ticktable(&H1) = 6
    ticktable(&H2) = 2
    ticktable(&H3) = 8
    ticktable(&H4) = 3
    ticktable(&H5) = 3
    ticktable(&H6) = 5
    ticktable(&H7) = 5
    ticktable(&H8) = 3
    ticktable(&H9) = 2
    ticktable(&HA) = 2
    ticktable(&HB) = 2
    ticktable(&HC) = 4
    ticktable(&HD) = 4
    ticktable(&HE) = 6
    ticktable(&HF) = 6
    ticktable(&H10) = 2
    ticktable(&H11) = 5
    ticktable(&H12) = 2
    ticktable(&H13) = 8
    ticktable(&H14) = 4
    ticktable(&H15) = 4
    ticktable(&H16) = 6
    ticktable(&H17) = 6
    ticktable(&H18) = 2
    ticktable(&H19) = 4
    ticktable(&H1A) = 2
    ticktable(&H1B) = 7
    ticktable(&H1C) = 4
    ticktable(&H1D) = 4
    ticktable(&H1E) = 7
    ticktable(&H1F) = 7
    ticktable(&H20) = 6
    ticktable(&H21) = 6
    ticktable(&H22) = 2
    ticktable(&H23) = 8
    ticktable(&H24) = 3
    ticktable(&H25) = 3
    ticktable(&H26) = 5
    ticktable(&H27) = 5
    ticktable(&H28) = 4
    ticktable(&H29) = 2
    ticktable(&H2A) = 2
    ticktable(&H2B) = 2
    ticktable(&H2C) = 4
    ticktable(&H2D) = 4
    ticktable(&H2E) = 6
    ticktable(&H2F) = 6
    ticktable(&H30) = 2
    ticktable(&H31) = 5
    ticktable(&H32) = 2
    ticktable(&H33) = 8
    ticktable(&H34) = 4
    ticktable(&H35) = 4
    ticktable(&H36) = 6
    ticktable(&H37) = 6
    ticktable(&H38) = 2
    ticktable(&H39) = 4
    ticktable(&H3A) = 2
    ticktable(&H3B) = 7
    ticktable(&H3C) = 4
    ticktable(&H3D) = 4
    ticktable(&H3E) = 7
    ticktable(&H3F) = 7
    ticktable(&H40) = 6
    ticktable(&H41) = 6
    ticktable(&H42) = 2
    ticktable(&H43) = 8
    ticktable(&H44) = 3
    ticktable(&H45) = 3
    ticktable(&H46) = 5
    ticktable(&H47) = 5
    ticktable(&H48) = 3
    ticktable(&H49) = 2
    ticktable(&H4A) = 2
    ticktable(&H4B) = 2
    ticktable(&H4C) = 3
    ticktable(&H4D) = 4
    ticktable(&H4E) = 6
    ticktable(&H4F) = 6
    ticktable(&H50) = 2
    ticktable(&H51) = 5
    ticktable(&H52) = 2
    ticktable(&H53) = 8
    ticktable(&H54) = 4
    ticktable(&H55) = 4
    ticktable(&H56) = 6
    ticktable(&H57) = 6
    ticktable(&H58) = 2
    ticktable(&H59) = 4
    ticktable(&H5A) = 2
    ticktable(&H5B) = 7
    ticktable(&H5C) = 4
    ticktable(&H5D) = 4
    ticktable(&H5E) = 7
    ticktable(&H5F) = 7
    ticktable(&H60) = 6
    ticktable(&H61) = 6
    ticktable(&H62) = 2
    ticktable(&H63) = 8
    ticktable(&H64) = 3
    ticktable(&H65) = 3
    ticktable(&H66) = 5
    ticktable(&H67) = 5
    ticktable(&H68) = 4
    ticktable(&H69) = 2
    ticktable(&H6A) = 2
    ticktable(&H6B) = 2
    ticktable(&H6C) = 5
    ticktable(&H6D) = 4
    ticktable(&H6E) = 6
    ticktable(&H6F) = 6
    ticktable(&H70) = 2
    ticktable(&H71) = 5
    ticktable(&H72) = 2
    ticktable(&H73) = 8
    ticktable(&H74) = 4
    ticktable(&H75) = 4
    ticktable(&H76) = 6
    ticktable(&H77) = 6
    ticktable(&H78) = 2
    ticktable(&H79) = 4
    ticktable(&H7A) = 2
    ticktable(&H7B) = 7
    ticktable(&H7C) = 4
    ticktable(&H7D) = 4
    ticktable(&H7E) = 7
    ticktable(&H7F) = 7
    ticktable(&H80) = 2
    ticktable(&H81) = 6
    ticktable(&H82) = 2
    ticktable(&H83) = 6
    ticktable(&H84) = 3
    ticktable(&H85) = 3
    ticktable(&H86) = 3
    ticktable(&H87) = 3
    ticktable(&H88) = 2
    ticktable(&H89) = 2
    ticktable(&H8A) = 2
    ticktable(&H8B) = 2
    ticktable(&H8C) = 4
    ticktable(&H8D) = 4
    ticktable(&H8E) = 4
    ticktable(&H8F) = 4
    ticktable(&H90) = 2
    ticktable(&H91) = 6
    ticktable(&H92) = 2
    ticktable(&H93) = 6
    ticktable(&H94) = 4
    ticktable(&H95) = 4
    ticktable(&H96) = 4
    ticktable(&H97) = 4
    ticktable(&H98) = 2
    ticktable(&H99) = 5
    ticktable(&H9A) = 2
    ticktable(&H9B) = 5
    ticktable(&H9C) = 5
    ticktable(&H9D) = 5
    ticktable(&H9E) = 5
    ticktable(&H9F) = 5
    ticktable(&HA0) = 2
    ticktable(&HA1) = 6
    ticktable(&HA2) = 2
    ticktable(&HA3) = 6
    ticktable(&HA4) = 3
    ticktable(&HA5) = 3
    ticktable(&HA6) = 3
    ticktable(&HA7) = 3
    ticktable(&HA8) = 2
    ticktable(&HA9) = 2
    ticktable(&HAA) = 2
    ticktable(&HAB) = 2
    ticktable(&HAC) = 4
    ticktable(&HAD) = 4
    ticktable(&HAE) = 4
    ticktable(&HAF) = 4
    ticktable(&HB0) = 2
    ticktable(&HB1) = 5
    ticktable(&HB2) = 2
    ticktable(&HB3) = 5
    ticktable(&HB4) = 4
    ticktable(&HB5) = 4
    ticktable(&HB6) = 4
    ticktable(&HB7) = 4
    ticktable(&HB8) = 2
    ticktable(&HB9) = 4
    ticktable(&HBA) = 2
    ticktable(&HBB) = 4
    ticktable(&HBC) = 4
    ticktable(&HBD) = 4
    ticktable(&HBE) = 4
    ticktable(&HBF) = 4
    ticktable(&HC0) = 2
    ticktable(&HC1) = 6
    ticktable(&HC2) = 2
    ticktable(&HC3) = 8
    ticktable(&HC4) = 3
    ticktable(&HC5) = 3
    ticktable(&HC6) = 5
    ticktable(&HC7) = 5
    ticktable(&HC8) = 2
    ticktable(&HC9) = 2
    ticktable(&HCA) = 2
    ticktable(&HCB) = 2
    ticktable(&HCC) = 4
    ticktable(&HCD) = 4
    ticktable(&HCE) = 6
    ticktable(&HCF) = 6
    ticktable(&HD0) = 2
    ticktable(&HD1) = 5
    ticktable(&HD2) = 2
    ticktable(&HD3) = 8
    ticktable(&HD4) = 4
    ticktable(&HD5) = 4
    ticktable(&HD6) = 6
    ticktable(&HD7) = 6
    ticktable(&HD8) = 2
    ticktable(&HD9) = 4
    ticktable(&HDA) = 2
    ticktable(&HDB) = 7
    ticktable(&HDC) = 4
    ticktable(&HDD) = 4
    ticktable(&HDE) = 7
    ticktable(&HDF) = 7
    ticktable(&HE0) = 2
    ticktable(&HE1) = 6
    ticktable(&HE2) = 2
    ticktable(&HE3) = 8
    ticktable(&HE4) = 3
    ticktable(&HE5) = 3
    ticktable(&HE6) = 5
    ticktable(&HE7) = 5
    ticktable(&HE8) = 2
    ticktable(&HE9) = 2
    ticktable(&HEA) = 2
    ticktable(&HEB) = 2
    ticktable(&HEC) = 4
    ticktable(&HED) = 4
    ticktable(&HEE) = 6
    ticktable(&HEF) = 6
    ticktable(&HF0) = 2
    ticktable(&HF1) = 5
    ticktable(&HF2) = 2
    ticktable(&HF3) = 8
    ticktable(&HF4) = 4
    ticktable(&HF5) = 4
    ticktable(&HF6) = 6
    ticktable(&HF7) = 6
    ticktable(&HF8) = 2
    ticktable(&HF9) = 4
    ticktable(&HFA) = 2
    ticktable(&HFB) = 7
    ticktable(&HFC) = 4
    ticktable(&HFD) = 4
    ticktable(&HFE) = 7
    ticktable(&HFF) = 7
End Sub
