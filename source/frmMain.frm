VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vbNES"
   ClientHeight    =   7275
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7740
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ws 
      Left            =   1560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1983
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   7260
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   0
      Width           =   7740
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu itmOpenROM 
         Caption         =   "&Open ROM..."
      End
      Begin VB.Menu itmOpenGenie 
         Caption         =   "Open ROM with &Game Genie..."
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu itmSaveState 
         Caption         =   "&Save state..."
      End
      Begin VB.Menu itmLoadState 
         Caption         =   "&Load state..."
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu itmROMinfo 
         Caption         =   "ROM &Information"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu itmExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuVideo 
      Caption         =   "Video"
      Begin VB.Menu mnuZoom 
         Caption         =   "&Zoom"
         Begin VB.Menu itm1x 
            Caption         =   "&1x"
         End
         Begin VB.Menu itm2x 
            Caption         =   "&2x"
         End
         Begin VB.Menu itm3x 
            Caption         =   "&3x"
         End
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu itmFrameskip 
         Caption         =   "Enable frame&skip"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuInput 
      Caption         =   "Input"
      Begin VB.Menu itmController1 
         Caption         =   "Configure controller 1..."
      End
      Begin VB.Menu itmController2 
         Caption         =   "Configure controller 2..."
      End
   End
   Begin VB.Menu mnuNetplay 
      Caption         =   "&Netplay"
      Begin VB.Menu itmHost 
         Caption         =   "&Host game..."
      End
      Begin VB.Menu itmConnect 
         Caption         =   "&Connect to host..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public dosave As Long, doload As Long
Dim chunklen As Long

Public Sub saveconfig()
    If Dir$(App.Path + "\vbnes.cfg") <> "" Then Kill App.Path + "\vbnes.cfg"
    Open App.Path + "\vbnes.cfg" For Binary As #6
    Put #6, , keymap
    Close #6
End Sub

Private Sub loadconfig()
    If Dir$(App.Path + "\vbnes.cfg") <> "" Then
        Open App.Path + "\vbnes.cfg" For Binary As #6
        Get #6, , keymap
        Close #6
    Else
        MsgBox "This appears to be the first time you have run vbNES," + vbCrLf + "or the configuration file has been deleted. Creating a new one" + vbCrLf + "with default keymap.", vbInformation Or vbOKOnly, "Configuration"
        keymap(0) = &H58
        keymap(1) = &H5A
        keymap(2) = &H10
        keymap(3) = &HD
        keymap(4) = &H26
        keymap(5) = &H28
        keymap(6) = &H25
        keymap(7) = &H27
        saveconfig
    End If
End Sub

Private Sub Form_Load()
    Load frmDirectPlay
    frmMain.Caption = "vbNES v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    frmMain.Show
    itmSaveState.enabled = False
    itmLoadState.enabled = False
    If Dir$(App.Path + "\slot1.vbn") <> "" Then Kill App.Path + "\slot1.vbn"
    If Dir$(App.Path + "\slot2.vbn") <> "" Then Kill App.Path + "\slot2.vbn"
    If Dir$(App.Path + "\slot3.vbn") <> "" Then Kill App.Path + "\slot3.vbn"
    If Dir$(App.Path + "\slot4.vbn") <> "" Then Kill App.Path + "\slot4.vbn"
    loadconfig
    itm2x_Click
    Do
        Do Until gameready = 1
            DoEvents
        Loop
        gameready = 0
        startgame
        running = 0
        If gamegenie = 2 Then loadROM "": gameready = 1
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim res As Integer
    res = MsgBox("Are you sure you want to exit vbNES?", vbQuestion Or vbYesNo, "vbNES")
    If res = vbNo Then Cancel = True: Exit Sub
    'TerminateThread hThread, 0&
    frmDirectPlay.SoundStop
    End
End Sub

Private Sub itm1x_Click()
    scalefactor = 1
    scalew = 256
    scaleh = 240
    pic.Width = 3900
    pic.Height = 3660
    frmMain.Width = pic.Width + 90
    frmMain.Height = pic.Height + 675
    ReDim drawscreen(0 To 2, 0 To (scalew - 1), 0 To (scaleh - 1))
End Sub

Private Sub itm2x_Click()
    scalefactor = 2
    scalew = 512
    scaleh = 480
    pic.Width = 7740
    pic.Height = 7260
    frmMain.Width = pic.Width + 90
    frmMain.Height = pic.Height + 675
    ReDim drawscreen(0 To 2, 0 To (scalew - 1), 0 To (scaleh - 1))
End Sub

Private Sub itm3x_Click()
    scalefactor = 3
    scalew = 768
    scaleh = 720
    pic.Width = 11580
    pic.Height = 10860
    frmMain.Width = pic.Width + 90
    frmMain.Height = pic.Height + 675
    ReDim drawscreen(0 To 2, 0 To (scalew - 1), 0 To (scaleh - 1))
End Sub

Private Sub itmHost_Click()
    ws.Close
    running = 0
    ws.LocalPort = Val(InputBox("TCP port to host on:", "Host game", "1983"))
    
    cd.Filter = "NES ROM files (*.nes)|*.nes"
    cd.ShowOpen
    If cd.FileName = "" Then Exit Sub
    If Dir$(cd.FileName) = "" Then Exit Sub
    romfile = cd.FileName

    MsgBox "You are now configured to host a game." + vbCrLf + "Click OK and wait for client to connect!", vbInformation Or vbOKOnly, "Host game"
    netmode = NET_SERVER_LISTEN
    ws.Listen
End Sub

Private Sub itmConnect_Click()
    Dim remoteHost As String, remotePort As Integer
    
    ws.Close
    running = 0
    remoteHost = InputBox("Remote hostname or IP:", "Connect to host")
    remotePort = Val(InputBox("Remote port to host on:", "Connect to host", "1983"))

    netmode = NET_CLIENT_HANDSHAKE
    ws.Connect remoteHost, remotePort
End Sub

Private Sub itmController1_Click()
    frmConfigInput.Show
    frmConfigInput.startkeyconfig1
End Sub

Private Sub itmController2_Click()
    frmConfigInput.Show
    frmConfigInput.startkeyconfig2
End Sub

Private Sub itmExit_Click()
    frmDirectPlay.SoundStop
    End
End Sub

Private Sub itmFrameskip_Click()
    If itmFrameskip.Checked = True Then
        itmFrameskip.Checked = False
    Else
        itmFrameskip.Checked = True
    End If
End Sub

Private Sub itmLoadState_Click()
    cd.Filter = "vbNES state files (*.vbn)|*.vbn"
    cd.ShowOpen
    If cd.FileName = "" Then Exit Sub
    statefile = cd.FileName
    doload = 1
End Sub

Private Sub itmOpenGenie_Click()
    gamegenie = 1
    opengame
End Sub

Private Sub itmOpenROM_Click()
    gamegenie = 0
    opengame
End Sub

Private Sub opengame()
    cd.Filter = "NES ROM files (*.nes)|*.nes"
    cd.ShowOpen
    If cd.FileName = "" Then Exit Sub
    If Dir$(cd.FileName) = "" Then Exit Sub
    romfile = cd.FileName
    gameready = 1
End Sub

Public Sub startgame()
    Dim freq As Currency, nextframetime As Currency, curtime As Currency, framelength As Currency
    Dim netout As Long
    Dim indata As String
        
    dosave = 0
    doload = 0
    loadROM romfile
    running = 1
    initMapper
    initPPU
    initAPU
    reset6502
    frmDirectPlay.Initialize 48000, 8, 1, 3000 '6000
    frmDirectPlay.SoundPlay
    
    If TRACEPPU = 1 Then
        Open "o:\traceppu.txt" For Output As #200
    End If
    
    QueryPerformanceFrequency freq
    QueryPerformanceCounter curtime
    framelength = freq / 60@
    nextframetime = curtime + framelength
    itmSaveState.enabled = True
    itmLoadState.enabled = True
    If netmode > 0 Then
        netSendData Chr$(255)
        Do
            DoEvents
            If ws.BytesReceived > 0 Then
                ws.GetData indata, vbString, 1
                netGetData indata
            End If
        Loop Until (netmode = NET_SERVER_PLAYING) Or (netmode = NET_CLIENT_PLAYING)
    End If
    donextframe = 0
    Do Until (running = 0) Or (gameready = 1)
        If dosave = 1 Then savestate: dosave = 0
        If doload = 1 Then loadstate: doload = 0
        If dospeedup = 1 Then framelength = framelength * 0.99: dospeedup = 0
        QueryPerformanceCounter curtime
        If (itmFrameskip.Checked = True) And (curtime > nextframetime) Then skipdraw = 1 Else skipdraw = 0
        Do
            QueryPerformanceCounter curtime
            If curtime >= nextframetime Then Exit Do
            DoEvents
            'Sleep 1
        Loop
        'nextframetime = nextframetime + framelength
        nextframetime = curtime + framelength
        QueryPerformanceCounter curtime
        If (curtime > (nextframetime + (framelength * 5@))) Then nextframetime = curtime + framelength: skipdraw = 0
        
        If (netmode > 0) And ((totalframes Mod 3) = 0) Then
            If netmode = NET_SERVER_PLAYING Then
                pad1net(0) = pad1(0): If pad1net(0) Then netout = 1 Else netout = 0
                pad1net(1) = pad1(1): If pad1net(1) Then netout = netout Or 2
                pad1net(2) = pad1(2): If pad1net(2) Then netout = netout Or 4
                pad1net(3) = pad1(3): If pad1net(3) Then netout = netout Or 8
                pad1net(4) = pad1(4): If pad1net(4) Then netout = netout Or 16
                pad1net(5) = pad1(5): If pad1net(5) Then netout = netout Or 32
                pad1net(6) = pad1(6): If pad1net(6) Then netout = netout Or 64
                pad1net(7) = pad1(7): If pad1net(7) Then netout = netout Or 128
            ElseIf netmode = NET_CLIENT_PLAYING Then
                pad2net(0) = pad1(0): If pad2net(0) Then netout = 1 Else netout = 0
                pad2net(1) = pad1(1): If pad2net(1) Then netout = netout Or 2
                pad2net(2) = pad1(2): If pad2net(2) Then netout = netout Or 4
                pad2net(3) = pad1(3): If pad2net(3) Then netout = netout Or 8
                pad2net(4) = pad1(4): If pad2net(4) Then netout = netout Or 16
                pad2net(5) = pad1(5): If pad2net(5) Then netout = netout Or 32
                pad2net(6) = pad1(6): If pad2net(6) Then netout = netout Or 64
                pad2net(7) = pad1(7): If pad2net(7) Then netout = netout Or 128
            End If
            netSendData Chr$(netout)
            Do Until netgotinput = 1
                If ws.State <> sckConnected Then
                    netmode = 0
                    ws.Close
                    MsgBox "Network connection closed!"
                    Exit Do
                End If
                If ws.BytesReceived > 0 Then
                    ws.GetData indata, vbString, 1
                    netGetData indata
                End If
                DoEvents
            Loop
            netgotinput = 0
        End If
        
        execframe
        DoEvents
    Loop
    itmSaveState.enabled = False
    itmLoadState.enabled = False
    
    frmDirectPlay.SoundStop
End Sub

Private Sub itmROMinfo_Click()
    MsgBox "Filename: " + romfile + vbCrLf + "PRG size: " + CStr(hdr.PRGsize * 16&) + vbCrLf + "CHR size: " + CStr(hdr.CHRsize * 8&) + " KB" + vbCrLf + "Mapper: " + CStr(hdr.mapper)
End Sub

Private Sub itmSaveState_Click()
    cd.Filter = "vbNES state files (*.vbn)|*.vbn"
    cd.ShowSave
    If cd.FileName = "" Then Exit Sub
    statefile = cd.FileName
    If Right$(LCase$(statefile), 4) <> ".vbn" Then statefile = statefile + ".vbn"
    dosave = 1
End Sub

Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        'controller 1
        Case keymap(0) 'A
            pad1(0) = 1
        Case keymap(1) 'B
            pad1(1) = 1
        Case keymap(2) 'select
            pad1(2) = 1
        Case keymap(3) 'start
            pad1(3) = 1
        Case keymap(4) 'up
            pad1(4) = 1
        Case keymap(5) 'down
            pad1(5) = 1
        Case keymap(6) 'left
            pad1(6) = 1
        Case keymap(7) 'right
            pad1(7) = 1
            
        'controller 2
        Case keymap(8) 'A
            pad2(0) = 1
        Case keymap(9) 'B
            pad2(1) = 1
        Case keymap(10) 'select
            pad2(2) = 1
        Case keymap(11) 'start
            pad2(3) = 1
        Case keymap(12) 'up
            pad2(4) = 1
        Case keymap(13) 'down
            pad2(5) = 1
        Case keymap(14) 'left
            pad2(6) = 1
        Case keymap(15) 'right
            pad2(7) = 1
        
        Case &H70 'F1
            statefile = App.Path + "\slot1.vbn"
            dosave = 1
        Case &H71 'F2
            statefile = App.Path + "\slot2.vbn"
            dosave = 1
        Case &H72 'F3
            statefile = App.Path + "\slot3.vbn"
            dosave = 1
        Case &H73 'F4
            statefile = App.Path + "\slot4.vbn"
            dosave = 1
        Case &H74 'F5
            statefile = App.Path + "\slot1.vbn"
            doload = 1
        Case &H75 'F6
            statefile = App.Path + "\slot2.vbn"
            doload = 1
        Case &H76 'F7
            statefile = App.Path + "\slot3.vbn"
            doload = 1
        Case &H77 'F8
            statefile = App.Path + "\slot4.vbn"
            doload = 1
    End Select
End Sub

Private Sub pic_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        'controller 1
        Case keymap(0) 'A
            pad1(0) = 0
        Case keymap(1) 'B
            pad1(1) = 0
        Case keymap(2) 'select
            pad1(2) = 0
        Case keymap(3) 'start
            pad1(3) = 0
        Case keymap(4) 'up
            pad1(4) = 0
        Case keymap(5) 'down
            pad1(5) = 0
        Case keymap(6) 'left
            pad1(6) = 0
        Case keymap(7) 'right
            pad1(7) = 0
    
        'controller 2
        Case keymap(8) 'A
            pad2(0) = 0
        Case keymap(9) 'B
            pad2(1) = 0
        Case keymap(10) 'select
            pad2(2) = 0
        Case keymap(11) 'start
            pad2(3) = 0
        Case keymap(12) 'up
            pad2(4) = 0
        Case keymap(13) 'down
            pad2(5) = 0
        Case keymap(14) 'left
            pad2(6) = 0
        Case keymap(15) 'right
            pad2(7) = 0
    End Select
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then zapper_trigger = 1: zapper_x = x: zapper_y = y
End Sub

Public Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    zapper_level = (((pic.Point(x, y) \ &H10000) And &HFF) + ((pic.Point(x, y) \ &H100) And &HFF) + (pic.Point(x, y) And &HFF)) \ 3
    zapper_x = x
    zapper_y = y
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then zapper_trigger = 0: zapper_x = x: zapper_y = y
End Sub

Private Sub ws_Close()
    'netmode = 0
    'ws.Close
    'MsgBox "Network connection closed!"
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
    If netmode <> NET_SERVER_LISTEN Then Exit Sub
    ws.Close
    ws.Accept requestID
    netConnect
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim indata As String
    
    Do Until ws.BytesReceived = 0
        Select Case netmode
            Case NET_SERVER_HANDSHAKE, NET_CLIENT_HANDSHAKE
                If ws.BytesReceived >= 3 Then
                    ws.GetData indata, vbString, 3
                    netGetData indata
                End If
            Case NET_CLIENT_SENDROM
                If ws.BytesReceived >= 2 Then
                    ws.GetData indata, vbString, 2
                    netGetData indata
                    chunklen = Asc(Left$(indata, 1)) * 256& + Asc(Right$(indata, 1))
                    netmode = NET_CLIENT_SENDROM2
                End If
            Case NET_CLIENT_SENDROM2
                If ws.BytesReceived >= chunklen Then
                    ws.GetData indata, vbString, chunklen
                    netGetData indata
                End If
            Case NET_SERVER_SENDROM
                ws.GetData indata, vbString, 1
                netGetData indata
            Case NET_SERVER_WAITGAME, NET_CLIENT_WAITGAME, NET_SERVER_PLAYING, NET_CLIENT_PLAYING
                Exit Sub
        End Select
    Loop
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'netmode = 0
    'ws.Close
    'MsgBox "Network connection closed!"
End Sub
