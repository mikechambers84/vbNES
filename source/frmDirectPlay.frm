VERSION 5.00
Begin VB.Form frmDirectPlay 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmDirectPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Credit for original DirectSound code goes to Michael Ciurescu
' It has been somewhat modified from his sources

Option Explicit

' Direct Sound objects
Private DX As New DirectX8
Private SEnum As DirectSoundEnum8
Private DIS As DirectSound8

' buffer, and buffer description
Private Buff As DirectSoundSecondaryBuffer8
Private BuffDesc As DSBUFFERDESC

' For the events
Private EventsNotify() As DSBPOSITIONNOTIFY
Private EndEvent As Long, MidEvent As Long, StartEvent As Long

' to know the buffer size
Private BuffLen As Long, HalfBuffLen As Long

Implements DirectXEvent8

Public Event NeedWavData(ByRef data() As Byte)

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
    Dim WaveBuffer() As Byte
    Dim n As Long
    
    ' make sure that Buff object is actually initialized to a buffer instance
    If Not (Buff Is Nothing) Then
        If eventid = StartEvent Or eventid = MidEvent Then
            ReDim WaveBuffer(HalfBuffLen - 1)
            
            RaiseEvent NeedWavData(WaveBuffer)
        End If
        
        
        If didstart = 1 Then
            If HalfBuffLen >= bufferpos Then dospeedup = 1
            For n = 0 To HalfBuffLen - 1
                WaveBuffer(n) = buf(n)
            Next n
            For n = HalfBuffLen To 11999
                buf(n - HalfBuffLen) = buf(n)
            Next n
            For n = (12000 - HalfBuffLen) To 11999
                buf(n) = 128
            Next n
            bufferpos = bufferpos - HalfBuffLen
            If bufferpos < 0 Then bufferpos = 0
        Else
            For n = 0 To HalfBuffLen - 1
                WaveBuffer(n) = 128
            Next n
        End If
        
        ' check for Buff object in case you call the
        ' uninitialize sub in the NeedWavData event
        If Not (Buff Is Nothing) Then
            Select Case eventid
            Case StartEvent
                ' we got the event that the read cursor is at the beginning of the buffer
                ' therefore write from the middle of the buffer to the end
                Buff.WriteBuffer HalfBuffLen, HalfBuffLen, WaveBuffer(0), DSBLOCK_DEFAULT
            Case MidEvent
                ' we got an event that the read cursor is at the middle of the buffer
                ' threfore write from the beginning of the buffer to the middle
                Buff.WriteBuffer 0, HalfBuffLen, WaveBuffer(0), DSBLOCK_DEFAULT
            Case EndEvent
                ' not used
            End Select
        End If
    End If
End Sub

Public Function Initialize(Optional ByVal SamplesPerSec As Long = 44100, _
                            Optional ByVal BitsPerSample As Integer = 16, _
                            Optional ByVal channels As Integer = 2, _
                            Optional ByVal HalfBufferLen As Long = 0, _
                            Optional ByVal GUID As String = "") As String
    
    On Error GoTo ReturnError
    Set SEnum = DX.GetDSEnum
    If Len(GUID) = 0 Then GUID = SEnum.GetGuid(1)
    
    Set DIS = DX.DirectSoundCreate(GUID)
    
    With BuffDesc.fxFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = channels
        .nBitsPerSample = BitsPerSample
        .lSamplesPerSec = SamplesPerSec
        
        .nBlockAlign = (.nBitsPerSample * .nChannels) \ 8
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
        
        If HalfBufferLen <= 0 Then
            HalfBuffLen = .lAvgBytesPerSec / 10
        Else
            HalfBuffLen = HalfBufferLen
        End If
        
        HalfBuffLen = HalfBuffLen - (HalfBuffLen Mod .nBlockAlign)
    End With
    
    BuffLen = HalfBuffLen * 2
    
    BuffDesc.lBufferBytes = BuffLen
    BuffDesc.lFlags = DSBCAPS_CTRLPOSITIONNOTIFY Or DSBCAPS_STICKYFOCUS
    
    DIS.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
    Set Buff = DIS.CreateSoundBuffer(BuffDesc)
    
    ReDim EventsNotify(0 To 2) As DSBPOSITIONNOTIFY
    
    ' create event to signal that DirectSound read cursor
    ' is at the beginning of the buffer
    StartEvent = DX.CreateEvent(Me)
    EventsNotify(0).hEventNotify = StartEvent
    EventsNotify(0).lOffset = 1
    
    ' create event to signal that DirectSound read cursor
    ' is at half of the buffer
    MidEvent = DX.CreateEvent(Me)
    EventsNotify(1).hEventNotify = MidEvent
    EventsNotify(1).lOffset = HalfBuffLen
    
    ' create the event to signal the sound has stopped
    EndEvent = DX.CreateEvent(Me)
    EventsNotify(2).hEventNotify = EndEvent
    EventsNotify(2).lOffset = DSBPN_OFFSETSTOP
    
    ' Assign the notification points to the buffer
    Buff.SetNotificationPositions 3, EventsNotify()
    
    Initialize = ""
    Exit Function
ReturnError:
    ' return error number, description and source
    Initialize = "Error: " & Err.Number & vbNewLine & _
        "Desription: " & Err.Description & vbNewLine & _
        "Source: " & Err.Source
    MsgBox Initialize
    
    Err.Clear
    UninitializeSound
    Exit Function
End Function

Public Sub UninitializeSound()
    On Error Resume Next
    If UBound(EventsNotify) > 0 Then
        If Err.Number = 0 Then
            ' distroy all events
            DX.DestroyEvent EventsNotify(0).hEventNotify
            DX.DestroyEvent EventsNotify(1).hEventNotify
            DX.DestroyEvent EventsNotify(2).hEventNotify
            
            Erase EventsNotify
        End If
    End If
    
    Set Buff = Nothing
    Set DIS = Nothing
    Set SEnum = Nothing
End Sub

Public Function SoundPlay() As Boolean
    On Error GoTo ReturnError
    
    If Not Buff Is Nothing Then Buff.Play DSBPLAY_LOOPING
    
    SoundPlay = True
    Exit Function
ReturnError:
    SoundPlay = False
    Err.Clear
End Function

Public Function SoundStop() As Boolean
    On Error GoTo ReturnError
    
    If Not Buff Is Nothing Then Buff.Stop
    
    SoundStop = True
    Exit Function
ReturnError:
    SoundStop = False
    Err.Clear
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UninitializeSound
End Sub
