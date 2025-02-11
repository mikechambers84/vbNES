Attribute VB_Name = "modArrays"
Public keymap(0 To 15) As Long 'can't put this on a form

Public Sub initPalette()
    NESpal(0) = &H808080
    NESpal(1) = &HBB0000
    NESpal(2) = &HBF0037
    NESpal(3) = &HA60084
    NESpal(4) = &H6A00BB
    NESpal(5) = &H1E00B7
    NESpal(6) = &HB3
    NESpal(7) = &H2691
    NESpal(8) = &H2B7B
    NESpal(9) = &H3E00
    NESpal(10) = &HD4800
    NESpal(11) = &H223C00
    NESpal(12) = &H662F00
    NESpal(13) = &H0
    NESpal(14) = &H50505
    NESpal(15) = &H50505
    NESpal(16) = &HC8C8C8
    NESpal(17) = &HFF5900
    NESpal(18) = &HFF3C44
    NESpal(19) = &HCC33B7
    NESpal(20) = &HAA33FF
    NESpal(21) = &H5E37FF
    NESpal(22) = &H1A37FF
    NESpal(23) = &H4BD5
    NESpal(24) = &H62C4
    NESpal(25) = &H7B3C
    NESpal(26) = &H15841E
    NESpal(27) = &H669500
    NESpal(28) = &HC48400
    NESpal(29) = &H111111
    NESpal(30) = &H90909
    NESpal(31) = &H90909
    NESpal(32) = &HFFFFFF
    NESpal(33) = &HFF9500
    NESpal(34) = &HFF846F
    NESpal(35) = &HFF6FD5
    NESpal(36) = &HCC77FF
    NESpal(37) = &H996FFF
    NESpal(38) = &H597BFF
    NESpal(39) = &H5F91FF
    NESpal(40) = &H33A2FF
    NESpal(41) = &HBFA6
    NESpal(42) = &H6AD951
    NESpal(43) = &HAED54D
    NESpal(44) = &HFFD900
    NESpal(45) = &H666666
    NESpal(46) = &HD0D0D
    NESpal(47) = &HD0D0D
    NESpal(48) = &HFFFFFF
    NESpal(49) = &HFFBF84
    NESpal(50) = &HFFBBBB
    NESpal(51) = &HFFBBD0
    NESpal(52) = &HEABFFF
    NESpal(53) = &HCCBFFF
    NESpal(54) = &HB7C4FF
    NESpal(55) = &HAECCFF
    NESpal(56) = &HA2D9FF
    NESpal(57) = &H99E1CC
    NESpal(58) = &HB7EEAE
    NESpal(59) = &HEEF7AA
    NESpal(60) = &HFFEEB3
    NESpal(61) = &HDDDDDD
    NESpal(62) = &H111111
    NESpal(63) = &H111111
End Sub

Public Sub initAPUarrays()
    lengthlookup(0, 0) = &HA
    lengthlookup(0, 1) = &H14
    lengthlookup(0, 2) = &H28
    lengthlookup(0, 3) = &H50
    lengthlookup(0, 4) = &HA0
    lengthlookup(0, 5) = &H3C
    lengthlookup(0, 6) = &HE
    lengthlookup(0, 7) = &H1A
    lengthlookup(0, 8) = &HC
    lengthlookup(0, 9) = &H18
    lengthlookup(0, 10) = &H30
    lengthlookup(0, 11) = &H60
    lengthlookup(0, 12) = &HC0
    lengthlookup(0, 13) = &H48
    lengthlookup(0, 14) = &H10
    lengthlookup(0, 15) = &H20
    lengthlookup(1, 0) = &HFE
    lengthlookup(1, 1) = &H2
    lengthlookup(1, 2) = &H4
    lengthlookup(1, 3) = &H6
    lengthlookup(1, 4) = &H8
    lengthlookup(1, 5) = &HA
    lengthlookup(1, 6) = &HC
    lengthlookup(1, 7) = &HE
    lengthlookup(1, 8) = &H10
    lengthlookup(1, 9) = &H12
    lengthlookup(1, 10) = &H14
    lengthlookup(1, 11) = &H16
    lengthlookup(1, 12) = &H18
    lengthlookup(1, 13) = &H1A
    lengthlookup(1, 14) = &H1C
    lengthlookup(1, 15) = &H1E
    
    noiselookup(0) = &H4
    noiselookup(1) = &H8
    noiselookup(2) = &H10
    noiselookup(3) = &H20
    noiselookup(4) = &H40
    noiselookup(5) = &H60
    noiselookup(6) = &H80
    noiselookup(7) = &HA0
    noiselookup(8) = &HCA
    noiselookup(9) = &HFE
    noiselookup(10) = &H17C
    noiselookup(11) = &H1FC
    noiselookup(12) = &H2FA
    noiselookup(13) = &H3F8
    noiselookup(14) = &H7F2
    noiselookup(15) = &HFE4

    trianglestep(0) = 15
    trianglestep(1) = 14
    trianglestep(2) = 13
    trianglestep(3) = 12
    trianglestep(4) = 11
    trianglestep(5) = 10
    trianglestep(6) = 9
    trianglestep(7) = 8
    trianglestep(8) = 7
    trianglestep(9) = 6
    trianglestep(10) = 5
    trianglestep(11) = 4
    trianglestep(12) = 3
    trianglestep(13) = 2
    trianglestep(14) = 1
    trianglestep(15) = 0
    trianglestep(16) = 0
    trianglestep(17) = 1
    trianglestep(18) = 2
    trianglestep(19) = 3
    trianglestep(20) = 4
    trianglestep(21) = 5
    trianglestep(22) = 6
    trianglestep(23) = 7
    trianglestep(24) = 8
    trianglestep(25) = 9
    trianglestep(26) = 10
    trianglestep(27) = 11
    trianglestep(28) = 12
    trianglestep(29) = 13
    trianglestep(30) = 14
    trianglestep(31) = 15

    dmcperiod(0) = &H1AC
    dmcperiod(1) = &H17C
    dmcperiod(2) = &H154
    dmcperiod(3) = &H140
    dmcperiod(4) = &H11E
    dmcperiod(5) = &HFE
    dmcperiod(6) = &HE2
    dmcperiod(7) = &HD6
    dmcperiod(8) = &HBE
    dmcperiod(9) = &HA0
    dmcperiod(10) = &H8E
    dmcperiod(11) = &H80
    dmcperiod(12) = &H6A
    dmcperiod(13) = &H54
    dmcperiod(14) = &H48
    dmcperiod(15) = &H36

    squareduty(0, 0) = 0
    squareduty(0, 1) = 1
    squareduty(0, 2) = 0
    squareduty(0, 3) = 0
    squareduty(0, 4) = 0
    squareduty(0, 5) = 0
    squareduty(0, 6) = 0
    squareduty(0, 7) = 0
    squareduty(1, 0) = 0
    squareduty(1, 1) = 1
    squareduty(1, 2) = 1
    squareduty(1, 3) = 0
    squareduty(1, 4) = 0
    squareduty(1, 5) = 0
    squareduty(1, 6) = 0
    squareduty(1, 7) = 0
    squareduty(2, 0) = 0
    squareduty(2, 1) = 1
    squareduty(2, 2) = 1
    squareduty(2, 3) = 1
    squareduty(2, 4) = 1
    squareduty(2, 5) = 0
    squareduty(2, 6) = 0
    squareduty(2, 7) = 0
    squareduty(3, 0) = 1
    squareduty(3, 1) = 0
    squareduty(3, 2) = 0
    squareduty(3, 3) = 1
    squareduty(3, 4) = 1
    squareduty(3, 5) = 1
    squareduty(3, 6) = 1
    squareduty(3, 7) = 1
End Sub
