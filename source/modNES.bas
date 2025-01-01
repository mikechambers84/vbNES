Attribute VB_Name = "modNES"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
'Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
'Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Public hThread As Long, hThreadID As Long

Public Const TRACEPPU = 0

Public gameready As Long
Public running As Long
Public totalframes As Long
Public curtile As Long
Public skipdraw As Long
Public donextframe As Long
Public outputNES(0 To 239, 0 To 255) As Long
Public romfile As String
Public curscan As Long
Public scalefactor As Long, scalew As Long, scaleh As Long
Public doirq As Long

Public Sub execframe()
    Dim sprtablesave As Long
    Dim scanline As Long
    Dim X As Long, Y As Long, DX As Long, dy As Long
    Dim r As Byte, g As Byte, b As Byte, curcolor As Long
    
    If TRACEPPU = 1 Then
        Print #200, ""
        Print #200, "=================="
        Print #200, "EXEC FRAME " + CStr(totalframes)
    End If
    
    PPU.vblank = 0
    PPU.sprzero = 0
    PPU.sprover = 0
    sprtablesave = PPU.sprtable
    
    If (PPU.bgvisible = 1) Then
        exec6502 101
        PPU.addr = PPU.tempaddr
        PPU.yscroll = PPU.tempy
        exec6502 13
    Else
        exec6502 114
    End If
    
    For scanline = 0 To 239
        curscan = scanline
        renderscanline scanline
        If TRACEPPU = 1 Then
            Print #200, "Scanline " + CStr(scanline) + ": loopy Y = " + CStr((PPU.addr \ 32) And 31) + ", loopy H = " + CStr(PPU.addr And 31)
        End If
        If (PPU.bgvisible = 1) Then PPU.addr = (PPU.addr And 64480) Or (PPU.tempaddr And &H41F&): PPU.xscroll = PPU.tempx
        'If doirq = 1 Then irq6502: PPU.yscroll = 0: doirq = 0
'        If (scanline Mod 3) = 0 Then exec6502 28 Else exec6502 29
    Next scanline
    
    exec6502 340
    PPU.vblank = 1
    If PPU.nmivblank Then nmi6502
    
    exec6502 2289
    If (totalframes And 1) = 1 Then exec6502 1
    
    If skipdraw = 0 Then
        For Y = 0 To scaleh - 1
            For X = 0 To scalew - 1
                curcolor = NESpal(outputNES(239 - Y \ scalefactor, X \ scalefactor) And 63)
                r = curcolor And 255
                g = (curcolor \ 256&) And 255
                b = curcolor \ 65536
                drawscreen(2, X, Y) = r
                drawscreen(1, X, Y) = g
                drawscreen(0, X, Y) = b
            Next X
        Next Y
        SetImageData frmMain.pic, drawscreen()
    End If
    
    totalframes = totalframes + 1
End Sub

