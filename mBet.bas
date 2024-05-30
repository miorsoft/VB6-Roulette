Attribute VB_Name = "mBet"
Option Explicit


Public BetMouseX#
Public BetMouseY#

Public BetPosX    As Long
Public BetPosY    As Long


Public BetActive  As Boolean

Public FichesOUTAnim As Boolean
Public FichesOUTX As Double
Public FichesOUTY As Double
Public FichesOUTAmount As Long

Public FichesINAnim As Boolean
Public FichesINX  As Double
Public FichesINY  As Double
Public FichesINAmount As Long


Public TableW     As Double
Attribute TableW.VB_VarUserMemId = 1073741827
Public TableH     As Double
Attribute TableH.VB_VarUserMemId = 1073741828
Public TableScreenX As Double
Attribute TableScreenX.VB_VarUserMemId = 1073741829
Public TableScreenY As Double
Attribute TableScreenY.VB_VarUserMemId = 1073741830
Public TableTopBorder As Double
Attribute TableTopBorder.VB_VarUserMemId = 1073741831
Public TableCellW As Double
Attribute TableCellW.VB_VarUserMemId = 1073741832
Public TableCellH As Double
Attribute TableCellH.VB_VarUserMemId = 1073741833

Public FichesPlacedAt() As Double
Public WINTABLEMultiplier() As Double
Public WINTABLENumbersList() As String


Public Const FicheRadius As Double = 12

Public BUDGET     As Long

Public TotalBet   As Long
Public TotalWin   As Long


Public Function BetPosInsideBounds(ByVal X&, ByVal Y&) As Boolean
    BetPosInsideBounds = True
    If X < 3 Then BetPosInsideBounds = False: Exit Function
    If X > 30 Then BetPosInsideBounds = False: Exit Function
    If Y < 0 Then BetPosInsideBounds = False: Exit Function
    If Y > 10 Then BetPosInsideBounds = False: Exit Function
End Function

Public Sub SETUPWINTABLE()
    Dim X&, Y&
    Dim X2&, Y2&
    Dim S         As String
    Dim I         As Long


    Dim N&
    ReDim WINTABLEMultiplier(30, 10)
    ReDim WINTABLENumbersList(30, 10)

    N = 0
    For X = 5 To 27 Step 2
        For Y = 5 To 1 Step -2
            WINTABLEMultiplier(X, Y) = 36    ' 1 NUMBER
            N = N + 1
            WINTABLENumbersList(X, Y) = N
        Next
    Next

    For X = 5 To 27 Step 2
        S = ""
        For Y = 4 To 2 Step -2
            X2 = (X - 5) / 2 * 3
            Y2 = (-Y + 6) / 2
            N = X2 + Y2
            S = N & "," & N + 1
            'MsgBox X2 & ", " & Y2 & "  " & S
            WINTABLEMultiplier(X, Y) = 18    '2 NUMBERS vert
            WINTABLENumbersList(X, Y) = S
        Next
    Next


    For X = 6 To 26 Step 2
        For Y = 5 To 1 Step -2
            X2 = (X - 6) / 2 * 3
            Y2 = (-Y + 7) / 2
            N = X2 + Y2
            S = N & "," & N + 3
            'MsgBox X2 & ", " & Y2 & "  " & S
            WINTABLEMultiplier(X, Y) = 18    '2 NUMBERS hor
            WINTABLENumbersList(X, Y) = S
        Next
    Next


    For X = 5 To 27 Step 2
        WINTABLEMultiplier(X, 6) = 12    '3 Numbers vert
        WINTABLEMultiplier(X, 0) = 12    '3 Numbers vert (SAME)
        N = (X - 5) / 2 * 3 + 1
        S = N & "," & N + 1 & "," & N + 2
        WINTABLENumbersList(X, 0) = S
        WINTABLENumbersList(X, 6) = S

    Next

    For X = 6 To 26 Step 2
        For Y = 2 To 4 Step 2
            X2 = (X - 6) / 2 * 3
            Y2 = (-Y + 4) / 2
            N = X2 + Y2 + 1
            S = N & "," & N + 1 & "," & N + 3 & "," & N + 4
            'MsgBox S
            WINTABLEMultiplier(X, Y) = 9    '4 NUMBERS
            WINTABLENumbersList(X, Y) = S
        Next
    Next


    For X = 6 To 26 Step 2
        WINTABLEMultiplier(X, 6) = 6    '6 Numbers vert
        WINTABLEMultiplier(X, 0) = 6    '6 Numbers vert (SAME)
        X2 = (X - 6) / 2 * 3
        N = X2 + 1
        S = N
        For I = N + 1 To N + 5
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLENumbersList(X, 6) = S
        WINTABLENumbersList(X, 0) = S
    Next

    X = 29
    For Y = 5 To 1 Step -2
        WINTABLEMultiplier(X, Y) = 3    '12 Numbers vert
        N = (5 - Y) / 2 + 1
        S = N
        For I = 1 To 11
            S = S & "," & N + I * 3
        Next
        'MsgBox S
        WINTABLENumbersList(X, Y) = S
    Next

    ' ZERO and 00
    WINTABLEMultiplier(3, 1) = 36
    WINTABLEMultiplier(3, 2) = 36
    WINTABLEMultiplier(3, 4) = 36
    WINTABLEMultiplier(3, 5) = 36

    WINTABLENumbersList(3, 1) = "-1"
    WINTABLENumbersList(3, 2) = "-1"
    WINTABLENumbersList(3, 4) = "0"
    WINTABLENumbersList(3, 5) = "0"


    '---------------------------

    For X = 5 To 11
        WINTABLEMultiplier(X, 7) = 3    ' 1st 12
        S = "1"
        For I = 2 To 12
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLENumbersList(X, 7) = S
    Next

    For X = 13 To 19
        WINTABLEMultiplier(X, 7) = 3    ' 2nd 12

        S = "13"
        For I = 14 To 24
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLENumbersList(X, 7) = S

    Next

    For X = 21 To 27
        WINTABLEMultiplier(X, 7) = 3    ' 3rd 12
        S = "25"
        For I = 26 To 36
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLENumbersList(X, 7) = S

    Next
    '---------------------------
    For X = 5 To 7
        WINTABLEMultiplier(X, 9) = 2    ' Manque

        S = "1"
        For I = 2 To 18
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLENumbersList(X, 9) = S
    Next
    '---------------------------
    For X = 9 To 11
        WINTABLEMultiplier(X, 9) = 2    'Pair
        S = "2"
        For I = 4 To 36 Step 2
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLENumbersList(X, 9) = S
    Next
    '---------------------------
    For X = 13 To 15
        WINTABLEMultiplier(X, 9) = 2    'Red
        WINTABLENumbersList(X, 9) = "1,3,5,7,9,12,14,16,18,19,21,23,25,27,30,32,34,36"
    Next

    '---------------------------
    For X = 17 To 19
        WINTABLEMultiplier(X, 9) = 2    'Black
        WINTABLENumbersList(X, 9) = "2,4,6,8,10,11,13,15,17,20,22,24,26,28,29,31,33,35"
    Next
    '---------------------------

    For X = 21 To 23
        WINTABLEMultiplier(X, 9) = 2    'Impair
        S = "1"
        For I = 3 To 35 Step 2
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLENumbersList(X, 9) = S
    Next
    '---------------------------
    For X = 25 To 27
        WINTABLEMultiplier(X, 9) = 2    'pass
        S = "19"
        For I = 20 To 36
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLENumbersList(X, 9) = S
    Next
    '---------------------------



    BUDGET = 1000

End Sub



Public Sub BET()
    Dim T#

    ReDim FichesPlacedAt(30, 10)

    BetActive = True

    fMain.lBudget = "Budget: " & BUDGET

    T = Timer
    Do

        DoEvents

        DRAWALL True
        Sleep 20

        If Timer > T + 9 Then Exit Do    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Loop While BetActive

End Sub



Public Sub DRAWfichesPilesAt(X#, Y#, Amount As Long)
    Dim XX#, YY#, ZZ#
    XX = TableScreenX + X * TableCellW * 0.5
    YY = TableScreenY + TableTopBorder + Y * TableCellH * 0.5

    For ZZ = 0 To Amount - 1
        'vbCyan
        CC.Arc XX, YY - ZZ * 2.5, FicheRadius
        CC.Fill True, Cairo.CreateSolidPatternLng(vbCyan)
        CC.SetSourceColor 0
        CC.Stroke
    Next

    CC.DrawText XX - FicheRadius, YY - FicheRadius * 0.5 - ZZ * 2.5, FicheRadius * 2, FicheRadius * 2, CStr(Amount), , True

End Sub

Public Sub DRAWBets()
    Dim X#
    Dim Y#
    Dim V&

    CC.SelectFont "Arial", 12, vbBlack, True

    For Y = 0 To 10
        For X = 3 To 30
            V = FichesPlacedAt(X, Y)
            If V Then
                DRAWfichesPilesAt X, Y, V
            End If
        Next
    Next


    ' HIGHLIGH MOUSE position OK
    If BetActive Then
        If BetPosInsideBounds(BetPosX, BetPosY) Then
            If WINTABLEMultiplier(BetPosX, BetPosY) <> 0 Then
                CC.SetSourceRGBA 1, 1, 0, 0.6
                CC.Arc BetMouseX, BetMouseY, FicheRadius
                CC.Fill
                'CC.TextOut BetMouseX + 12, BetMouseY + 12, BetPosX & "," & BetPosY & "         " & WINTABLEMultiplier(BetPosX, BetPosY)
                'CC.TextOut BetMouseX, BetMouseY + 24, "X " & WINTABLEMultiplier(BetPosX, BetPosY)
            End If
        End If
    End If

End Sub


Private Sub HIlight(S As String)
    Dim N&
    Dim X#, Y#

    N = Val(S)

    If N > 0 Then
        X = TableScreenX + TableCellW * (2 + ((N - 1) \ 3))
        Y = TableScreenY + TableTopBorder + TableCellH * 2 - TableCellH * ((N - 1) Mod 3)

        CC.SetSourceRGBA 1, 1, 0, 0.25
        CC.Rectangle X, Y, TableCellW, TableCellH
        CC.Fill
    Else

        If N = 0 Then              ' ZERO
            X = TableScreenX + TableCellW * 1
            Y = TableScreenY + TableTopBorder + TableCellH * 1.5
            CC.SetSourceRGBA 1, 1, 0, 0.25
            CC.Rectangle X, Y, TableCellW, TableCellH * 1.5
            CC.Fill
        Else                       '    -1    DOBLE ZERO
            X = TableScreenX + TableCellW * 1
            Y = TableScreenY + TableTopBorder
            CC.SetSourceRGBA 1, 1, 0, 0.25
            CC.Rectangle X, Y, TableCellW, TableCellH * 1.5
            CC.Fill
        End If
    End If

End Sub

Public Sub HILightBET()
    Dim X&, Y&
    Dim S()       As String
    Dim I         As Long

    For Y = 0 To 10
        For X = 3 To 30
            If FichesPlacedAt(X, Y) Then
                S = Split(WINTABLENumbersList(X, Y), ",")
                For I = 0 To UBound(S)
                    HIlight S(I)
                Next
            End If
        Next
    Next

    If BetPosInsideBounds(BetPosX, BetPosY) Then
        If WINTABLEMultiplier(BetPosX, BetPosY) <> 0 Then
            S = Split(WINTABLENumbersList(BetPosX, BetPosY), ",")
            For I = 0 To UBound(S)
                HIlight S(I)
            Next
        End If
    End If

End Sub



Public Sub MANAGEBETS()
    Dim X&, Y&, V&
    Dim WIN       As Long
    Dim S()       As String
    Dim I&

    Dim DoNOTWin  As Boolean


    fMain.txtWIN = ""
    fMain.lWin = "WIN: 0"

    TotalWin = 0


    For Y = 0 To 10
        For X = 3 To 30
            V = FichesPlacedAt(X, Y)
            If V Then
                S = Split(WINTABLENumbersList(X, Y), ",")
                DoNOTWin = True
                For I = 0 To UBound(S)
                    If NumberExtracted = Val(S(I)) Then DoNOTWin = False
                Next
                If DoNOTWin Then AnimateFichesOUT X, Y
            End If
        Next
    Next




    For Y = 0 To 10
        For X = 3 To 30
            V = FichesPlacedAt(X, Y)
            If V Then
                S = Split(WINTABLENumbersList(X, Y), ",")
                For I = 0 To UBound(S)
                    If NumberExtracted = Val(S(I)) Then
                        WIN = V * WINTABLEMultiplier(X, Y)
                        BUDGET = BUDGET + WIN
                        'fMain.txtWIN = fMain.txtWIN & V & " x " & WINTABLEMultiplier(X, Y) & " = " & WIN & vbCrLf
                        fMain.txtWIN = V & " x " & WINTABLEMultiplier(X, Y) & " = " & WIN & vbCrLf & fMain.txtWIN
                        TotalWin = TotalWin + WIN
                        fMain.lWin = "WIN: " & TotalWin
                        fMain.txtWIN.Refresh
                        AnimateFichesIN X, Y, WIN
                    End If
                Next
            End If
        Next
    Next

    fMain.txtWIN = "WIN: " & TotalWin & " (Total)" & vbCrLf & fMain.txtWIN

    fMain.txtWIN.Refresh
    Sleep 1000

End Sub




Private Sub AnimateFichesOUT(X&, Y&)

    Dim I         As Long
    Dim DX#, DY#
    Dim D#

    FichesOUTX = X
    FichesOUTY = Y

    DX = (X - 15) * 0.1
    '    DY = (Y + 1) * 0.1
    DY = (Y + 4) * 0.1

    D = Sqr(DX * DX + DY * DY)
    DX = DX / D * 0.08
    DY = DY / D * 0.08

    FichesOUTAmount = FichesPlacedAt(X, Y)
    FichesPlacedAt(X, Y) = 0

    FichesOUTAnim = True

    SOUNDSPLAYER.PlaySound "movement-swipe-whoosh-1-186575.MP3", 0, -1500

    For I = 0 To 5000
        FichesOUTX = FichesOUTX - DX
        FichesOUTY = FichesOUTY - DY
        DX = DX * 1.9
        DY = DY * 1.9

        DRAWALL
        Sleep 20

        If FichesOUTY < -FicheRadius Then Exit For
        If FichesOUTX < -FicheRadius Then Exit For

    Next

    FichesOUTAnim = False


End Sub


Private Sub AnimateFichesIN(X&, Y&, Amount As Long)
    Dim XtoReach  As Double
    Dim YtoReach  As Double
    Dim DX#, DY#, D#
    Dim I         As Long

    FichesINX = 15
    FichesINY = -1

    FichesINAmount = Amount

    XtoReach = X
    YtoReach = Y


    DX = XtoReach - FichesINX
    DY = YtoReach - FichesINY
    D = Sqr(DX * DX + DY * DY)
    DX = DX * 0.2
    DY = DY * 0.2


    FichesINAnim = True
    SOUNDSPLAYER.PlaySound "correct-2-46134.MP3", 0, -1000
    'SOUNDSPLAYER.PlaySound "item-pick-up-38258.MP3", 0, -1000


    For I = 0 To 5000
        FichesINX = FichesINX + DX
        FichesINY = FichesINY + DY

        If DX * DX + DY * DY > 0.05 Then
            DX = DX * 0.8
            DY = DY * 0.8
        End If

        DRAWALL
        Sleep 20

        D = (XtoReach - FichesINX) * (XtoReach - FichesINX) + _
            (YtoReach - FichesINY) * (YtoReach - FichesINY)

        If D < 0.05 Then Exit For

    Next

    FichesINAnim = False

    FichesPlacedAt(X, Y) = Amount

    DRAWALL
    Sleep 400
    If TURBO Then ReDim FichesPlacedAt(30, 10)

End Sub

