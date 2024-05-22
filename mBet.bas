Attribute VB_Name = "mBet"
Option Explicit


Public BetMouseX#
Public BetMouseY#

Public BetPosX    As Long
Public BetPosY    As Long


Public BetActive  As Boolean

Public TableW     As Double
Attribute TableW.VB_VarUserMemId = 1073741827
Public TableH     As Double
Attribute TableH.VB_VarUserMemId = 1073741828
Public TableX     As Double
Attribute TableX.VB_VarUserMemId = 1073741829
Public TableY     As Double
Attribute TableY.VB_VarUserMemId = 1073741830
Public TBO        As Double
Attribute TBO.VB_VarUserMemId = 1073741831
Public TcX        As Double
Attribute TcX.VB_VarUserMemId = 1073741832
Public TcY        As Double
Attribute TcY.VB_VarUserMemId = 1073741833

Public FICHEScount() As Double
Public WINTABLE() As Double
Public WINTABLEString() As String


Public Const FicheRadius As Double = 12

Public Budget     As Long


Public Sub SETUPWINTABLE()
    Dim X&, Y&
    Dim X2&, Y2&
    Dim S         As String
    Dim I         As Long


    Dim N&
    ReDim WINTABLE(30, 10)
    ReDim WINTABLEString(30, 10)

    N = 0
    For X = 5 To 27 Step 2
        For Y = 5 To 1 Step -2
            WINTABLE(X, Y) = 36    ' 1 NUMBER
            N = N + 1
            WINTABLEString(X, Y) = N
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
            WINTABLE(X, Y) = 18    '2 NUMBERS vert
            WINTABLEString(X, Y) = S
        Next
    Next


    For X = 6 To 26 Step 2
        For Y = 5 To 1 Step -2
            X2 = (X - 6) / 2 * 3
            Y2 = (-Y + 7) / 2
            N = X2 + Y2
            S = N & "," & N + 3
            'MsgBox X2 & ", " & Y2 & "  " & S
            WINTABLE(X, Y) = 18    '2 NUMBERS hor
            WINTABLEString(X, Y) = S
        Next
    Next


    For X = 5 To 27 Step 2
        WINTABLE(X, 6) = 12        '3 Numbers vert
        WINTABLE(X, 0) = 12        '3 Numbers vert (SAME)
        N = (X - 5) / 2 * 3 + 1
        S = N & "," & N + 1 & "," & N + 2
        WINTABLEString(X, 0) = S
        WINTABLEString(X, 6) = S

    Next

    For X = 6 To 26 Step 2
        For Y = 2 To 4 Step 2
            X2 = (X - 6) / 2 * 3
            Y2 = (-Y + 4) / 2
            N = X2 + Y2 + 1
            S = N & "," & N + 1 & "," & N + 3 & "," & N + 4
            'MsgBox S
            WINTABLE(X, Y) = 9     '4 NUMBERS
            WINTABLEString(X, Y) = S
        Next
    Next


    For X = 6 To 26 Step 2
        WINTABLE(X, 6) = 6         '6 Numbers vert
        WINTABLE(X, 0) = 6         '6 Numbers vert (SAME)
        X2 = (X - 6) / 2 * 3
        N = X2 + 1
        S = N
        For I = N + 1 To N + 5
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLEString(X, 6) = S
        WINTABLEString(X, 0) = S
    Next

    X = 29
    For Y = 5 To 1 Step -2
        WINTABLE(X, Y) = 3         '12 Numbers vert
        N = (5 - Y) / 2 + 1
        S = N
        For I = 1 To 11
            S = S & "," & N + I * 3
        Next
        'MsgBox S
        WINTABLEString(X, Y) = S
    Next

    ' ZERO and 00
    WINTABLE(3, 1) = 36
    WINTABLE(3, 2) = 36
    WINTABLE(3, 4) = 36
    WINTABLE(3, 5) = 36

    WINTABLEString(3, 1) = "-1"
    WINTABLEString(3, 2) = "-1"
    WINTABLEString(3, 4) = "0"
    WINTABLEString(3, 5) = "0"


    '---------------------------

    For X = 5 To 11
        WINTABLE(X, 7) = 3         ' 1st 12
        S = "1"
        For I = 2 To 12
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLEString(X, 7) = S
    Next

    For X = 13 To 19
        WINTABLE(X, 7) = 3         ' 2nd 12

        S = "13"
        For I = 14 To 24
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLEString(X, 7) = S

    Next

    For X = 21 To 27
        WINTABLE(X, 7) = 3         ' 3rd 12
        S = "25"
        For I = 26 To 36
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLEString(X, 7) = S

    Next
    '---------------------------
    For X = 5 To 7
        WINTABLE(X, 9) = 2         ' Manque

        S = "1"
        For I = 2 To 18
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLEString(X, 9) = S
    Next
    '---------------------------
    For X = 9 To 11
        WINTABLE(X, 9) = 2         'Pair
        S = "2"
        For I = 4 To 36 Step 2
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLEString(X, 9) = S
    Next
    '---------------------------
    For X = 13 To 15
        WINTABLE(X, 9) = 2         'Red
        WINTABLEString(X, 9) = "1,3,5,7,9,12,14,16,18,19,21,23,25,27,30,32,34,36"
    Next

    '---------------------------
    For X = 17 To 19
        WINTABLE(X, 9) = 2         'Black
        WINTABLEString(X, 9) = "2,4,6,8,10,11,13,15,17,20,22,24,26,28,29,31,33,35"
    Next
    '---------------------------

    For X = 21 To 23
        WINTABLE(X, 9) = 2         'Impair
        S = "1"
        For I = 3 To 35 Step 2
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLEString(X, 9) = S
    Next
    '---------------------------
    For X = 25 To 27
        WINTABLE(X, 9) = 2         'pass
        S = "19"
        For I = 20 To 36
            S = S & "," & I
        Next
        'MsgBox S
        WINTABLEString(X, 9) = S
    Next
    '---------------------------



    Budget = 500

End Sub



Public Sub BET()
    Dim T#

    ReDim FICHEScount(30, 10)

    BetActive = True

    fMain.lBudget = "Budget: " & Budget

    T = Timer
    Do

        DoEvents

        DRAWALL True


        If Timer > T + 10 Then Exit Do    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Loop While BetActive

End Sub





Public Sub DRAWBets()
    Dim X&
    Dim Y&
    Dim V&

    Dim XX#, YY#

    CC.SelectFont "Arial", 12, vbBlack, True


    For Y = 0 To 10
        For X = 3 To 30
            V = FICHEScount(X, Y)

            If V Then
                XX = TableX + X * TcX * 0.5
                YY = TableY + TBO + Y * TcY * 0.5

                CC.SetSourceColor vbCyan
                CC.Arc XX, YY, FicheRadius
                CC.Fill
                CC.DrawText XX - FicheRadius, YY - FicheRadius * 0.7, FicheRadius * 2, FicheRadius * 2, CStr(V), , True

            End If
        Next
    Next



    If BetActive Then

        If BetMouseX > TableX - FicheRadius Then
            If BetMouseY > TableY - FicheRadius Then
                If BetMouseX + FicheRadius < TableW + TableX Then
                    If BetMouseY + FicheRadius < TableH + TableY Then
                        If WINTABLE(BetPosX, BetPosY) <> 0 Then
                            CC.SetSourceRGBA 1, 1, 0, 0.6
                            CC.Arc BetMouseX, BetMouseY, FicheRadius
                            CC.Fill
                            '                            CC.TextOut BetMouseX + 12, BetMouseY + 12, BetPosX & "," & BetPosY & "         " & WINTABLE(BetPosX, BetPosY)
                        End If
                    End If

                End If
            End If
        End If
    End If

End Sub


Private Sub HIlight(S As String)
    Dim N&
    Dim X#, Y#

    N = Val(S)

    If N > 0 Then
        X = TableX + TcX * (2 + ((N - 1) \ 3))
        Y = TableY + TBO + TcY * 2 - TcY * ((N - 1) Mod 3)

        CC.SetSourceRGBA 1, 1, 0, 0.25
        CC.Rectangle X, Y, TcX, TcY
        CC.Fill
    Else

        If N = 0 Then
            X = TableX + TcX * 1
            Y = TableY + TBO + TcY * 1.5
            CC.SetSourceRGBA 1, 1, 0, 0.25
            CC.Rectangle X, Y, TcX, TcY * 1.5
            CC.Fill
        Else                       '    -1          '00
            X = TableX + TcX * 1
            Y = TableY + TBO
            CC.SetSourceRGBA 1, 1, 0, 0.25
            CC.Rectangle X, Y, TcX, TcY * 1.5
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
            If FICHEScount(X, Y) Then
                S = Split(WINTABLEString(X, Y), ",")
                For I = 0 To UBound(S)
                    HIlight S(I)
                Next
            End If
        Next
    Next

    If BetMouseX > TableX - FicheRadius Then
        If BetMouseY > TableY - FicheRadius Then
            If BetMouseX + FicheRadius < TableW + TableX Then
                If BetMouseY + FicheRadius < TableH + TableY Then
                    If WINTABLE(BetPosX, BetPosY) <> 0 Then
                        S = Split(WINTABLEString(BetPosX, BetPosY), ",")
                        For I = 0 To UBound(S)
                            HIlight S(I)
                        Next
                    End If
                End If
            End If
        End If
    End If

End Sub



Public Sub MANAGEBETS()
    Dim X&, Y&, V&
    Dim WIN       As Long
    Dim S()       As String
    Dim I&
    Dim TOT       As Long


    fMain.txtWIN = ""

    For Y = 0 To 10
        For X = 3 To 30
            V = FICHEScount(X, Y)
            If V Then

                S = Split(WINTABLEString(X, Y), ",")
                For I = 0 To UBound(S)
                    If NumberExtracted = Val(S(I)) Then
                        WIN = V * WINTABLE(X, Y)
                        Budget = Budget + WIN
                        fMain.txtWIN = fMain.txtWIN & V & " x " & WINTABLE(X, Y) & " = " & WIN & vbCrLf
                        TOT = TOT + WIN
                    End If
                Next
            End If
        Next
    Next

    fMain.txtWIN = "WIN: " & TOT & "(Total)" & vbCrLf & fMain.txtWIN

    fMain.txtWIN.Refresh


End Sub
