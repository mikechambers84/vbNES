Attribute VB_Name = "modAudio"
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public dospeedup As Long

Const SAVE_AUDIO = 0

Type triangletype
    lasttick As Long
    lengthcounter As Long
    lengthindex As Long
    lengthhalt As Long 'also control flag
    period As Long
    linearcounter As Long
    linearreload As Long
    seqstep As Long
    enabled As Long
End Type

Type squaretype
    lasttick As Long
    lengthcounter As Long
    lengthindex As Long
    envelope As Long
    constant As Long
    lengthhalt As Long
    sweep As Long
    period As Long
    negate As Long
    Shift As Long
    sweepenable As Long
    sweepresult As Long
    sweepreload As Long
    sweepperiod As Long
    duty As Long
    dutypos As Long
    wrotebit4 As Long
    envloop As Long
    enabled As Long
End Type

Type noisetype
    lasttick As Long
    lengthcounter As Long
    lengthindex As Long
    envelope As Long
    lengthhalt As Long
    period As Long
    Shift As Long
    enabled As Long
    constant As Long
    envreload As Long
    wrotebit4 As Long
    mode As Long
End Type

Type dmctype
    lasttick As Long
    period As Long
    samplebuffer As Long
    bufferbit As Long
    sampleempty As Long
    bytesremain As Long
    lengthreg As Long
    address As Long
    addressreg As Long
    loopmode As Long
    enabled As Long
End Type

Const CPUNTSC = 1789773
Const CPUPAL = 1662607

Public audbuf(0 To 399) As Byte

Public lengthlookup(0 To 1, 0 To 15) As Long
Public noiselookup(0 To 15) As Long
Public trianglestep(0 To 31) As Long
Public dmcperiod(0 To 15) As Long
Public squareduty(0 To 3, 0 To 7) As Long

Dim buffersize As Long, cursamplerate As Long
Public buf(0 To 11999) As Long

Dim seqtickgap As Long, seqticknext As Long
Dim sampleticks As Long, fastsampleticks As Long, nextsamplecycle As Long, nextseqcycle As Long
Public cursampleticks As Long

Public seqstep As Long, seqmode As Long, interruptAPU As Long
Public square(0 To 1) As squaretype
Public triangle As triangletype
Public noise As noisetype
Public dmc As dmctype
Const allowsweeps = 1 'sweep code is completely buggy right now, and really breaks a lot of stuff. allow sweeps if you want, but be prepared for extreme fail.

Dim channels(0 To 4) As Long 'square 1, square 2, triangle, noise, DMC

Public bufferpos As Long, buffersync As Long, audiosync As Long, unpausebufpos As Long, last As Long, copyhalf As Long
Public didstart As Long

Public Function readAPU(ByVal addr As Long) As Long
    Dim outbyte As Long
    If (addr = &H4015) Then
        If (square(0).enabled <> 0) And (square(0).lengthcounter <> 0) Then outbyte = 1 Else outbyte = 0
        If (square(1).enabled <> 0) And (square(1).lengthcounter <> 0) Then outbyte = outbyte Or 2
        If (triangle.enabled <> 0) And (triangle.lengthcounter <> 0) Then outbyte = outbyte Or 4
        If (noise.enabled <> 0) And (noise.lengthcounter <> 0) Then outbyte = outbyte Or 8

        readAPU = outbyte
        Exit Function
    End If
    readAPU = 0
End Function

Public Sub writeAPU(ByVal addr As Long, ByVal value As Long)
    Select Case addr
        Case &H4000 'square 1
            square(0).duty = value \ 64
            If (value And &H20) Then square(0).lengthhalt = 1 Else square(0).lengthhalt = 0
            If (value And &H10) Then square(0).constant = 1 Else square(0).constant = 0
            square(0).envelope = value And &HF
        Case &H4001
            If (value And &H80) Then square(0).sweepenable = 1: square(0).sweepresult = square(0).period Else square(0).sweepenable = 0: square(0).sweepresult = square(0).period
            square(0).sweepperiod = ((value \ 16) And 7) '+ 1
            If (value And &H8) Then square(0).negate = 1 Else square(0).negate = 0
            square(0).sweep = value And 7
            square(0).sweepreload = 1
        Case &H4002
            square(0).period = square(0).period And 65280
            square(0).period = square(0).period + value + 1
            square(0).sweepresult = square(0).period
        Case &H4003
            square(0).lengthindex = value \ 8
            square(0).period = square(0).period And 255
            square(0).period = square(0).period + ((value And 7) * 256&)
            square(0).dutypos = 0
            squarereload (0)
            square(0).wrotebit4 = 1
            square(0).sweepresult = square(0).period

        Case &H4004 'square 2
            square(1).duty = value \ 64
            If (value And &H20) Then square(1).lengthhalt = 1 Else square(1).lengthhalt = 0
            If (value And &H10) Then square(1).constant = 1 Else square(1).constant = 0
            square(1).envelope = value And &HF
        Case &H4005
            If (value And &H80) Then square(1).sweepenable = 1: square(1).sweepresult = square(1).period Else square(1).sweepenable = 0: square(1).sweepresult = square(1).period
            square(1).sweepperiod = ((value \ 16) And 7) '+ 1
            If (value And &H8) Then square(1).negate = 1 Else square(1).negate = 0
            square(1).sweep = value And 7
            square(1).sweepreload = 1
        Case &H4006
            square(1).period = square(1).period And 65280
            square(1).period = square(1).period + value
            square(1).sweepresult = square(1).period
        Case &H4007
            square(1).lengthindex = value \ 8
            square(1).period = square(1).period And 255
            square(1).period = square(1).period + ((value And 7) * 256&)
            square(1).dutypos = 0
            squarereload (1)
            square(1).wrotebit4 = 1
            square(1).sweepresult = square(1).period

        Case &H4008 'triangle
            If (value And &H80) Then triangle.lengthhalt = 1 Else triangle.lengthhalt = 0
            triangle.linearreload = value And &H7F
        Case &H400A
            triangle.period = triangle.period And 65280
            triangle.period = triangle.period + value + 1
        Case &H400B
            triangle.lengthindex = value \ 8
            triangle.period = triangle.period And 255
            triangle.period = triangle.period + ((value And 7) * 256&)
            'if (triangle.lengthhalt = 0)
            triangle.lengthcounter = lengthlookup(triangle.lengthindex And 1, triangle.lengthindex \ 2)

        Case &H400C 'noise
            If (value And &H20) Then noise.lengthhalt = 1 Else noise.lengthhalt = 0
            If (value And &H10) Then noise.constant = 1 Else noise.constant = 0
            noise.envreload = value And &HF
            noise.envelope = noise.envreload
        Case &H400E
            noise.mode = value \ 128
            noise.period = noiselookup(value And &HF)
        Case &H400F
            noise.lengthindex = value \ 8
            noise.lengthcounter = lengthlookup(noise.lengthindex And 1, noise.lengthindex \ 2)
            noise.wrotebit4 = 1

        Case &H4010 'DMC
            If (value And &H40) Then dmc.loopmode = 1 Else dmc.loopmode = 0
            dmc.period = dmcperiod(value And &HF) \ 8
        Case &H4011
            'channels(4) = value And &H7F
        Case &H4012
            dmc.addressreg = value
            dmc.address = (value * 64&) Or 49152
        Case &H4013
            dmc.lengthreg = value
            dmc.bytesremain = (value * 16&) + 1

        Case &H4015
            If (value And 1) Then
                square(0).enabled = 1
            Else
                square(0).enabled = 0
                square(0).lengthcounter = 0
            End If
            If (value And 2) Then
                square(1).enabled = 1
            Else
                square(1).enabled = 0
                square(1).lengthcounter = 0
            End If
            If (value And 4) Then
                triangle.enabled = 1
            Else
                triangle.enabled = 0
                triangle.lengthcounter = 0
            End If
            If (value And 8) Then
                noise.enabled = 1
            Else
                noise.enabled = 0
            End If
            If (value And 16) Then
                dmc.enabled = 1
                dmc.lasttick = realticks6502
            Else
                dmc.enabled = 0
            End If

        Case &H4017
            If (value And &H80) Then
                seqmode = 5
                lengthclock
            Else
                seqmode = 4
            End If
    End Select
End Sub

Private Sub squarereload(ByVal channel As Long)
    If (square(channel).enabled = 0) Then Exit Sub
    square(channel).lengthcounter = lengthlookup(square(channel).lengthindex And 1, square(channel).lengthindex \ 2)
End Sub

Private Sub envclock()
    If (square(0).constant = 0) Then
        If (square(0).wrotebit4) Then
            square(0).envelope = 15
            square(0).wrotebit4 = 0
        Else
            If (square(0).envelope > 0) Then square(0).envelope = square(0).envelope - 1
        End If
        If ((square(0).envelope = 0) And (square(0).envloop <> 0)) Then
            square(0).envelope = 15
        End If
    End If

    If (square(1).constant = 0) Then
        If (square(1).wrotebit4) Then
            square(1).envelope = 15
            square(1).wrotebit4 = 0
        Else
            If (square(1).envelope > 0) Then square(1).envelope = square(1).envelope - 1
        End If
        If ((square(1).envelope = 0) And (square(1).envloop <> 0)) Then
            square(1).envelope = 15
        End If
    End If

    If (noise.constant = 0) Then
        If (noise.wrotebit4) Then
            noise.envelope = 15
            noise.wrotebit4 = 0
        Else
            If (noise.envelope > 0) Then noise.envelope = noise.envelope - 1
        End If
        If ((noise.envelope = 0) And (noise.lengthhalt <> 0)) Then
            noise.envelope = 15
        End If
    End If
End Sub

Private Sub linearclock()
    If (triangle.lengthhalt = 0) Then
        If (triangle.linearcounter > 0) Then triangle.linearcounter = triangle.linearcounter - 1 Else If (triangle.enabled = 1) Then triangle.linearcounter = triangle.linearreload
    Else
        triangle.linearcounter = triangle.linearreload
    End If
End Sub

Private Sub lengthclock()
    If (square(0).lengthhalt = 0) And (square(0).enabled = 1) Then
        If (square(0).lengthcounter > 0) Then square(0).lengthcounter = square(0).lengthcounter - 1
    End If

    If (square(1).lengthhalt = 0) And (square(1).enabled = 1) Then
        If (square(1).lengthcounter > 0) Then square(1).lengthcounter = square(1).lengthcounter - 1
    End If

    If (triangle.lengthhalt = 0) And (triangle.enabled = 1) Then
        If (triangle.lengthcounter > 0) Then triangle.lengthcounter = triangle.lengthcounter - 1
    End If

    If (noise.lengthhalt = 0) And (noise.enabled = 1) Then
        If (noise.lengthcounter > 0) Then noise.lengthcounter = noise.lengthcounter - 1
    End If
End Sub

Private Sub sweepclock()
    Dim s As Long, wl As Long, d(0 To 3) As Long, chan As Long, origsweep As Long
    
    If allowsweeps = 0 Then Exit Sub

    For chan = 0 To 1
     'Else If square(chan).sweepreload Then square(chan).sweep = square(chan).sweepperiod: square(chan).sweepreload = 0
        If (square(chan).sweepperiod > 0) Then
            square(chan).sweepperiod = square(chan).sweepperiod - 1
        ElseIf square(chan).sweepenable Then
            wl = square(chan).period
            s = wl \ (2 ^ square(chan).sweep)
            If square(chan).negate Then s = (Not s) + (chan Xor 1)
            wl = square(chan).period + s
            If (((Not square(chan).negate) And (wl < &H800&)) And (wl > 7)) Then square(chan).period = wl Else square(chan).period = 0
        End If
    Next chan
End Sub

Private Sub setinterruptAPU()
End Sub

Private Sub frameseqAPU()
    If (realticks6502 < seqticknext) Then Exit Sub Else seqticknext = seqticknext + seqtickgap

        If (seqmode = 4) Then
            Select Case seqstep
                Case 0
                    envclock
                    linearclock
                Case 1
                    lengthclock
                    sweepclock
                    envclock
                    linearclock
                Case 2:
                    envclock
                    linearclock
                Case 3:
                    setinterruptAPU
                    lengthclock
                    sweepclock
                    envclock
                    linearclock
            End Select
            seqstep = (seqstep + 1) Mod seqmode
        Else 'seqmode = 5
            Select Case seqstep
                Case 0
                    lengthclock
                    sweepclock
                    envclock
                    linearclock
                Case 1
                    envclock
                    linearclock
                Case 2
                    lengthclock
                    sweepclock
                    envclock
                    linearclock
                Case 3
                    envclock
                    linearclock
                Case 4
            End Select
            seqstep = (seqstep + 1) Mod seqmode
        End If
End Sub

Private Sub mixerinput(ByVal channel As Long, ByVal value As Long)
    channels(channel) = value
End Sub

Private Function mixeroutAPU() As Long
    Dim outbyte As Double, squareout As Double, tndout As Double
    
    squareout = 0
    If ((square(0).enabled = 1) And (square(0).period > 7) And (square(0).lengthcounter > 0)) Then squareout = squareout + channels(0)
    If ((square(1).enabled = 1) And (square(1).period > 7) And (square(1).lengthcounter > 0)) Then squareout = squareout + channels(1)
    squareout = squareout * 0.00752
    
    tndout = 0
    If ((triangle.enabled = 1) And (triangle.period > 7) And (triangle.lengthcounter > 0) And (triangle.linearcounter > 0)) Then tndout = tndout + 0.00851 * channels(2)

    tndout = tndout + 0.00494 * channels(3) + 0.00335 * channels(4)
    outbyte = 127 * squareout
    outbyte = outbyte + (127 * tndout)

    mixeroutAPU = (CLng(outbyte) + 128) And 255
End Function

Public Sub buffersampleAPU()
    Dim n As Long
    Dim a As String
    
    If bufferpos = 10000 Then didstart = 1
    If (bufferpos < buffersize) Then
        buf(bufferpos) = mixeroutAPU()
        If SAVE_AUDIO = 1 Then
            a = Chr$(buf(bufferpos) And 255)
            Put #14, , a
        End If
        bufferpos = bufferpos + 1
    End If
End Sub

Public Sub tickchannelsAPU()
    Dim tmpbit As Long, feedback As Long, squaretarget As Long

    If (realticks6502 >= nextsamplecycle) Then
        buffersampleAPU
        nextsamplecycle = nextsamplecycle + cursampleticks
    End If

    If (bufferpos >= buffersize) Then 'if sample buffer is full, postpone channel ticks
        square(0).lasttick = realticks6502
        square(1).lasttick = realticks6502
        triangle.lasttick = realticks6502
        noise.lasttick = realticks6502
        dmc.lasttick = realticks6502
        Exit Sub
    End If

    If (realticks6502 >= nextseqcycle) Then
        frameseqAPU
        nextseqcycle = nextseqcycle + seqtickgap
    End If

    'If (square(0).sweepenable) Then squaretarget = square(0).sweepresult Else
    squaretarget = square(0).period
    If ((realticks6502 - square(0).lasttick) >= (squaretarget * 2&)) Then
        If (square(0).enabled And (squaretarget > 7) And (square(0).lengthcounter > 0)) Then
            channels(0) = square(0).envelope * squareduty(square(0).duty, square(0).dutypos)
            square(0).dutypos = (square(0).dutypos + 1) And 7
        Else
            channels(0) = 0
        End If
        square(0).lasttick = realticks6502 - ((realticks6502 - square(0).lasttick) - (squaretarget * 2&))
    End If

    'If (square(1).sweepenable) Then squaretarget = square(1).sweepresult Else
    squaretarget = square(1).period
    If ((realticks6502 - square(1).lasttick) >= (squaretarget * 2&)) Then
        If (square(1).enabled And (squaretarget > 7) And (square(1).lengthcounter > 0)) Then
            channels(1) = square(1).envelope * squareduty(square(1).duty, square(1).dutypos)
            square(1).dutypos = (square(1).dutypos + 1) And 7
        Else
            channels(1) = 0
        End If
        square(1).lasttick = realticks6502 - ((realticks6502 - square(1).lasttick) - (squaretarget * 2&))
    End If

    If ((realticks6502 - triangle.lasttick) >= triangle.period) Then
        If (triangle.enabled And (triangle.period > 7) And (triangle.lengthcounter > 0) And (triangle.linearcounter > 0)) Then
            channels(2) = trianglestep(triangle.seqstep) - 8
            triangle.seqstep = (triangle.seqstep + 1) And 31
        Else
            channels(2) = 0
        End If
        triangle.lasttick = realticks6502 - ((realticks6502 - triangle.lasttick) - triangle.period)
    End If

    If ((realticks6502 - noise.lasttick) >= noise.period) Then
        If (noise.enabled) Then
            If (noise.mode) Then
                If (noise.Shift And &H40) Then tmpbit = 1 Else tmpbit = 0
            Else
                If (noise.Shift And &H2) Then tmpbit = 1 Else tmpbit = 0
            End If
            feedback = (noise.Shift And 1) Xor tmpbit
            noise.Shift = (noise.Shift \ 2&) Or (feedback * 8192&)
            If ((noise.lengthcounter > 0) And ((noise.Shift And 1) = 0) And (noise.period > 7)) Then
                channels(3) = noise.envelope
            Else
                channels(3) = 0
            End If
        Else
            channels(3) = 0
        End If
        noise.lasttick = realticks6502 - ((realticks6502 - noise.lasttick) - noise.period)
    End If

    If ((realticks6502 - dmc.lasttick) >= (dmc.period * 8&)) Then
        If (dmc.enabled) Then
            If (dmc.sampleempty) Then
                dmc.sampleempty = 0
                dmc.bufferbit = 0
                dmc.samplebuffer = read6502(dmc.address): dmc.address = dmc.address + 1
                If (dmc.address < &H8000) Then dmc.address = 32768
                dmc.bytesremain = dmc.bytesremain - 1
                If (dmc.bytesremain = 0) Then
                    If (dmc.loopmode = 0) Then
                        dmc.enabled = 0
                    Else
                        dmc.address = (dmc.addressreg * 64&) Or 49152
                        dmc.bytesremain = (dmc.lengthreg * 16&) + 1
                    End If
                End If
            End If

            If (dmc.sampleempty = 0) Then
                If ((dmc.samplebuffer \ (2 ^ (dmc.bufferbit And 7))) And 1) Then
                    If (channels(4) <= &H7D) Then channels(4) = channels(4) + 2
                Else
                    If (channels(4) >= 2) Then channels(4) = channels(4) - 2
                End If
                dmc.bufferbit = dmc.bufferbit + 1
                If (dmc.bufferbit = 8) Then
                    dmc.sampleempty = 1
                End If
            End If
        End If

        dmc.lasttick = realticks6502 - ((realticks6502 - dmc.lasttick) - (dmc.period * 8&))
    End If
End Sub

Public Sub initAPU()
    initAPUarrays
    ZeroMemory square(0), Len(square(0))
    ZeroMemory square(1), Len(square(1))
    ZeroMemory triangle, Len(triangle)
    ZeroMemory noise, Len(noise)
    ZeroMemory dmc, Len(dmc)
    
    nextsamplecycle = &HFFFFFFFF
    nextseqcycle = &HFFFFFFFF

    seqtickgap = CPUNTSC \ 240
    sampleticks = CPUNTSC \ 48000
    cursampleticks = sampleticks

    seqticknext = realticks6502 + seqtickgap
    seqmode = 4
    seqstep = 0
    triangle.seqstep = 0
    noise.Shift = 1
    
    cursamplerate = 48000
    buffersize = 12000
    copyhalf = 0
    
    If SAVE_AUDIO = 1 Then
        If Dir$("audio.raw") <> "" Then Kill "audio.raw"
        Open "audio.raw" For Binary As #14
    End If
End Sub
