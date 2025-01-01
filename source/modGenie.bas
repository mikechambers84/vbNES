Attribute VB_Name = "modGenie"
Option Explicit

Type cheattype
    replace As Long
    compare As Long
    docompare As Long
    valid As Long
    addr As Long
End Type

Public cheat(0 To 2) As cheattype
Public gamegenie As Long

Public Sub geniewrite(ByVal addr As Long, ByVal value As Long)
    Dim cnum As Long
    
    addr = addr - 32768
    cnum = ((addr And 15) - 1) \ 4&
    Select Case addr
        Case &H0
            If (value = 0) Then
                gamegenie = 2
                running = 0
            Else
                If (value And 2) Then cheat(0).docompare = 1
                If (value And 4) Then cheat(1).docompare = 1
                If (value And 8) Then cheat(2).docompare = 1
                If (value And 16) Then cheat(0).valid = 0 Else cheat(0).valid = 1
                If (value And 32) Then cheat(1).valid = 0 Else cheat(1).valid = 1
                If (value And 64) Then cheat(2).valid = 0 Else cheat(2).valid = 1
            End If

        Case &H1
            cheat(0).addr = (value * 256&) Or (cheat(0).addr And &HFF)
        Case &H2
            cheat(0).addr = value Or (cheat(0).addr And 65280)
        Case &H3
            cheat(0).compare = value
        Case &H4
            cheat(0).replace = value

        Case &H5
            cheat(1).addr = (value * 256&) Or (cheat(1).addr And &HFF)
        Case &H6
            cheat(1).addr = value Or (cheat(1).addr And 65280)
        Case &H7
            cheat(1).compare = value
        Case &H8
            cheat(1).replace = value

        Case &H9
            cheat(2).addr = (value * 256&) Or (cheat(2).addr And &HFF)
        Case &HA
            cheat(2).addr = value Or (cheat(2).addr And 65280)
        Case &HB
            cheat(2).compare = value
        Case &HC
            cheat(2).replace = value
    End Select
End Sub
