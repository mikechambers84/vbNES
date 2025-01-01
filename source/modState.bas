Attribute VB_Name = "modState"
Option Explicit

Public statefile As String

Public Sub savestate()
    If statefile = "" Then Exit Sub
    If Dir$(statefile) <> "" Then Kill statefile
    Open statefile For Binary As #5
    Put #5, , RAM
    Put #5, , PPU
    Put #5, , OAM
    Put #5, , VRAM
    Put #5, , map1
    Put #5, , map4
    Put #5, , map9
    Put #5, , PRGbank
    Put #5, , CHRbank
    Put #5, , pc
    Put #5, , sp
    Put #5, , a
    Put #5, , x
    Put #5, , y
    Put #5, , status
    Put #5, , square(0)
    Put #5, , square(1)
    Put #5, , triangle
    Put #5, , noise
    Put #5, , dmc
    Put #5, , seqstep
    Put #5, , seqmode
    Put #5, , interruptAPU
    Put #5, , batRAM
    If hdr.CHRsize = 0 Then Put #5, , CHRbin
    Close #5
    statefile = ""
End Sub

Public Sub loadstate()
    If netmode <> 0 Then Exit Sub
    If statefile = "" Then Exit Sub
    If Dir$(statefile) = "" Then Exit Sub
    Open statefile For Binary As #5
    Get #5, , RAM
    Get #5, , PPU
    Get #5, , OAM
    Get #5, , VRAM
    Get #5, , map1
    Get #5, , map4
    Get #5, , map9
    Get #5, , PRGbank
    Get #5, , CHRbank
    Get #5, , pc
    Get #5, , sp
    Get #5, , a
    Get #5, , x
    Get #5, , y
    Get #5, , status
    Get #5, , square(0)
    Get #5, , square(1)
    Get #5, , triangle
    Get #5, , noise
    Get #5, , dmc
    Get #5, , seqstep
    Get #5, , seqmode
    Get #5, , interruptAPU
    Get #5, , batRAM
    If hdr.CHRsize = 0 Then Get #5, , CHRbin
    Close #5
    statefile = ""
End Sub
