Attribute VB_Name = "mMain"
Option Explicit

Private SRF       As cCairoSurface
Private WheelSRF  As cCairoSurface
Private TableSRF  As cCairoSurface

Private CC        As cCairoContext

Private CX        As Double
Private CY        As Double

Private WheelImageRadius As Double

Private WheelANG  As Double
Private WheelANGSpeed As Double

'Private CurrTICK  As Long
'Private TickAnim  As Long
'Private TickDRAW  As Long

Private TEMPO     As clsTick
Private tDRAW     As Long
Private tCompute  As Long
Private t1sec     As Long
Private ComputedFPS As Double
Private DrawFPS   As Double


Private BallX     As Double
Private BallY     As Double
Private BallVX    As Double
Private BallVY    As Double
Private Const R   As Double = 7    'BallRadius

Private OuterRadius As Double
Private Const InnerRadius As Double = 159

Private Const PI2 As Double = 6.28318530717959
Private Const PI  As Double = 3.14159265358979
Public Const PIh  As Double = 1.5707963267949

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private SLOTn(0 To 37) As Long
Private STATFreq(0 To 37) As Long
Private STATRita(0 To 37) As Long

Public NSPINS     As Long

Public TURBO      As Boolean
Public SoundMODE  As Long

Private BallSTOPCount As Long

Private CNT       As Long

Public RenderDev  As cMMDevice


Public Sub SETUP(Optional andLAUNCH As Boolean = False)
    Randomize Timer

    Set WheelSRF = Cairo.ImageList.AddImage("WHEEL", App.Path & "\RouletteWheel.png")
    Set TableSRF = Cairo.ImageList.AddImage("TABLE", App.Path & "\AmericanTable.png")

    WheelImageRadius = WheelSRF.Width * 0.5
    CX = WheelImageRadius + 10
    CY = WheelImageRadius + 10

    OuterRadius = WheelImageRadius - 14    '24

    fMain.Text1.Left = CX * 2 + 5
    fMain.Text2.Left = fMain.Text1.Width + fMain.Text1.Left + 5
    fMain.Text3.Left = fMain.Text2.Width + fMain.Text2.Left + 5





    Set SRF = Cairo.CreateSurface(WheelImageRadius * 2 + 20 + TableSRF.Width, WheelImageRadius * 2 + 20, ImageSurface)
    Set CC = SRF.CreateContext
    CC.AntiAlias = CAIRO_ANTIALIAS_FAST
    CC.SetLineWidth 1
    CC.SetSourceColor vbWhite: CC.Paint


    SLOTn(0) = 33: SLOTn(1) = 7: SLOTn(2) = 17: SLOTn(3) = 5: SLOTn(4) = 22: SLOTn(5) = 34: SLOTn(6) = 15: SLOTn(7) = 3
    SLOTn(8) = 24: SLOTn(9) = 36: SLOTn(10) = 13: SLOTn(11) = 1: SLOTn(12) = -1: SLOTn(13) = 27: SLOTn(14) = 10: SLOTn(15) = 25
    SLOTn(16) = 29: SLOTn(17) = 12: SLOTn(18) = 8: SLOTn(19) = 19: SLOTn(20) = 31: SLOTn(21) = 18: SLOTn(22) = 6: SLOTn(23) = 21
    SLOTn(24) = 26: SLOTn(25) = 16: SLOTn(26) = 4: SLOTn(27) = 23: SLOTn(28) = 35: SLOTn(29) = 14: SLOTn(30) = 2: SLOTn(31) = 0
    SLOTn(32) = 20: SLOTn(33) = 9: SLOTn(34) = 28: SLOTn(35) = 32: SLOTn(36) = 11: SLOTn(37) = 30



    Dim I&
    Dim S$
    Dim FN$


    For I = 0 To 37
        FN = Slot2MP3(I, S)
        If Dir(FN) = vbNullString Then
            GoogleSpeakCreateMP3_2 FN, "      " & S, "fr"
            PlayMP3 FN
        End If
    Next
    '

    FN = App.Path & "\Sounds\Faites vos jeux.MP3"
    If Dir(FN) = vbNullString Then
        GoogleSpeakCreateMP3_2 FN, "      " & "Faites vos jeux", "fr"
        PlayMP3 FN
    End If

    FN = App.Path & "\Sounds\Les Jeux sont faits.MP3"
    If Dir(FN) = vbNullString Then
        GoogleSpeakCreateMP3_2 FN, "      " & "Les Jeux sont faits", "fr"
        PlayMP3 FN
    End If

    FN = App.Path & "\Sounds\Rien ne va plus.MP3"
    If Dir(FN) = vbNullString Then
        GoogleSpeakCreateMP3_2 FN, "      " & "Rien ne va plus", "fr"
        PlayMP3 FN
    End If



    SETUPSOUND

    If andLAUNCH Then LAUNCH





End Sub


Public Sub LAUNCH()

    '    WheelANGSpeed = 0.25 + (Rnd * 2 - 1) * 0.07
    WheelANGSpeed = 0.2 + (Rnd * 2 - 1) * 0.05

    BallX = -0
    BallY = -(OuterRadius - R) + 4 + Rnd * 14
    '    BallVX = Rnd * 8
    BallVX = (Rnd * 2 - 1) * 8

    BallVY = 0

    While WheelANG > PI2: WheelANG = WheelANG - PI2: Wend




    NSPINS = NSPINS + 1

    If SoundMODE > 1 Then PlayMP3 App.Path & "\Sounds\Faites vos jeux.MP3"
    'If SoundMODE > 1 Then SOUNDFait.PLAY

    '-<<<<--------- WAIT BETS

    'If SoundMODE > 1 Then PlayAsync App.Path & "\Sounds\Les Jeux sont faits.MP3"
    If SoundMODE > 1 Then SOUNDLeJeux.PLAY


    WHEELLOOP

End Sub
Public Function Slot2Number(Slot As Long, Optional JustNumber As Boolean = False) As String
    Dim Result&
    Result = SLOTn(Slot)
    If Result < 0 Then Slot2Number = "00" Else: Slot2Number = CStr(Result)
    If Len(Slot2Number) = 1 Then Slot2Number = " " & Slot2Number

    If Not (JustNumber) Then
        If Result > 0 Then
            Slot2Number = Slot2Number & IIf((Slot Mod 2), " Rouge", " Noir ")
            Slot2Number = Slot2Number & IIf((Val(Result) Mod 2), " Impair", " Pair  ")
            Slot2Number = Slot2Number & IIf((Val(Result) <= 18), " Manque", " Passe ")
        Else
            Slot2Number = Slot2Number & Space(20)
        End If
    End If
End Function
Public Function Slot2MP3(Slot As Long, ByRef ToSpeak As String) As String


    ToSpeak = Slot2Number(Slot)
    ToSpeak = Left$(ToSpeak, 8) & " " & Replace(Right$(ToSpeak, Len(ToSpeak) - 8), " ", ".")
    If Left$(ToSpeak, 2) = "00" Then ToSpeak = "doubler 0"
    Slot2MP3 = App.Path & "\Sounds\" & ToSpeak & ".MP3"


End Function

Public Function Number2Slot(N As Long) As Long
    Dim I         As Long
    For I = 0 To 37
        If SLOTn(I) = N Then
            Number2Slot = I
            Exit For
        End If
    Next
End Function


Private Sub ShowResult()
    Dim N         As Long
    Dim Result    As Long
    Dim S         As String
    Dim I&

    N = (-WheelANG + Atan2(BallX, BallY)) / PI2 * 38
    While N < 0: N = N + 38: Wend
    N = N Mod 38

    STATFreq(N) = STATFreq(N) + 1

    For I = 0 To 37
        STATRita(I) = STATRita(I) + 1
    Next
    STATRita(N) = 0



    Result = SLOTn(N)

    S = Slot2Number(N)
    S = S & "  (" & STATFreq(N) & ")"
    fMain.Text1 = S & vbCrLf & fMain.Text1

    ' Test Text Length
    If Len(fMain.Text1) > 1600 Then fMain.Text1 = Left$(fMain.Text1, 1600)


    UPDATESTAT


    If SoundMODE = 1 Then
        SoundSLOT(N).PLAY
    Else
        If SoundMODE <> 0 Then PlayMP3 Slot2MP3(N, S)
        'If SoundMODE <> 0 Then SoundSLOT(N).PLAY
    End If

    LAUNCH

End Sub


Public Sub WHEELLOOP()

    If TURBO Then
        ComputedFPS = 500          '400
        DrawFPS = 6                '8
    Else
        ComputedFPS = 100
        DrawFPS = 30               '100
    End If

    '    If TEMPO Is Nothing Then       '
    Set TEMPO = New clsTick
    tDRAW = TEMPO.Add(DrawFPS)
    tCompute = TEMPO.Add(ComputedFPS)
    t1sec = TEMPO.Add(1)
    '    Else
    '        TEMPO.Remove tDRAW
    '        TEMPO.Remove tCompute
    '        TEMPO.Remove t1sec
    '        tDRAW = TEMPO.Add(DrawFPS)
    '        tCompute = TEMPO.Add(ComputedFPS)
    '        t1sec = TEMPO.Add(1)
    '    End If

    BallSTOPCount = 0
    CNT = 0
    Do

        Select Case TEMPO.WaitForNext

            Case tCompute

                SIMULATE

                If WheelANGSpeed < 0 Then
                    '                    Exit Do
                    WheelANGSpeed = 0
                End If
                If BallSTOPCount > 50 Then Exit Do


                CNT = CNT + 1
                If Not (TURBO) Then
                    If CNT > 450 Then
                        CNT = -100000000#
                        ' PlayMP3 App.Path & "\Sounds\Rien ne va plus.MP3"
                        'If SoundMODE > 1 Then PlayAsync App.Path & "\Sounds\Rien ne va plus.MP3"
                        If SoundMODE > 1 Then SOUNDRien.PLAY


                    End If
                End If

            Case tDRAW
                DRAWALL
                DoEvents

            Case t1sec
                fMain.Caption = "  Computed FPS:" & TEMPO.Count(tCompute) & " DrawnFPS:" & TEMPO.Count(tDRAW)
                TEMPO.ResetCount (tCompute)
                TEMPO.ResetCount (tDRAW)
                'TEMPO.ResetCount (t1sec)
        End Select

    Loop While True

    ShowResult

End Sub

Public Sub DRAWALL()

    '   CC.SetSourceColor vbWhite: CC.Paint

    CC.Save
    CC.TranslateDrawings CX, CY
    CC.RotateDrawings WheelANG
    CC.RenderSurfaceContent WheelSRF, -WheelImageRadius, -WheelImageRadius

    'DEBUG .............
    '        Dim A#, CA#, SA#
    '        Const angSTeP     As Double = 0.165346981767884
    '    For A = -0.07 To PI2 Step angSTeP
    '            CA = Cos(A)
    '            SA = Sin(A)
    '            CC.DrawLine CA * 175, SA * 175, CA * 150, SA * 150, , 4, vbRed
    '        Next
    '    ' ...............

    CC.Restore

    CC.ARC BallX + CX, BallY + CY, R + 1
    CC.Fill True, Cairo.CreateSolidPatternLng(vbWhite)
    CC.SetSourceColor 0
    CC.Stroke



    CC.RenderSurfaceContent TableSRF, WheelImageRadius * 2 + 20, 0

    fMain.Picture = SRF.Picture
    '    SRF.DrawToDC fMain.hDC
    'fMain.Refresh

End Sub

'Private Sub Reflect(X#, Y#, ByVal WallX#, ByVal WallY#)
'    Dim D#
'    D = X * WallX + Y * WallY
'    X = X - WallX * D * 2
'    Y = Y - WallY * D * 2
'End Sub
Private Sub Project(X#, Y#, ByVal v2X#, ByVal v2Y#)
    Dim D         As Double
    D = X * v2X + Y * v2Y
    X = v2X * D
    Y = v2Y * D
End Sub


Public Function Atan2(ByVal DX As Double, ByVal DY As Double) As Double
    If DX Then Atan2 = Atn(DY / DX) + PI * (DX < 0#) Else Atan2 = -PIh - (DY > 0#) * PI
End Function


Private Sub CalcDistFromLineAndNormal(ByVal PX#, ByVal PY#, ByVal AX#, ByVal AY#, ByVal BX#, ByVal BY#, ByVal InvABlen2#, rDIST#, rNX#, rNY#)    ', rPosX#, rPosY#)
    Dim PAX#, PAY#, H#
    Dim bAX#, bAY#
    Dim DX#, DY#

    PAX = PX - AX
    PAY = PY - AY
    bAX = BX - AX
    bAY = BY - AY

    H = (PAX * bAX + PAY * bAY) * InvABlen2
    If H > 1# Then H = 1#
    If H < 0# Then H = 0#

    DX = PAX - bAX * H
    DY = PAY - bAY * H

    rDIST = Sqr(DX * DX + DY * DY)

    rNX = DX                       ' / rDIST 'Will be normalized later
    rNY = DY                       ' / rDIST
    '    rPosX = AX + bAX * H
    '    rPosY = AY + bAY * H

End Sub


Public Sub SIMULATE()

    Dim DistFromCenter#
    Dim invDFC#

    Dim DX#, DY#

    Const WallSPEEDK As Double = 0.002

    WheelANG = WheelANG + WheelANGSpeed

    '    WheelANGSpeed = WheelANGSpeed * 0.997
    WheelANGSpeed = WheelANGSpeed * 0.9974

    If WheelANGSpeed > 0.000005 Then WheelANGSpeed = WheelANGSpeed - 0.000005

    DistFromCenter = Sqr(BallX * BallX + BallY * BallY)
    invDFC = 1# / DistFromCenter

    DX = BallX * invDFC
    DY = BallY * invDFC

    If DistFromCenter < 183 Then   ' CHECK CHELL SLOT
        CheckCOLLISIONwihtSLOTS DistFromCenter
    End If

    If DistFromCenter > (OuterRadius - R) Then    'Beyond OUTER Radius --->  Reflect Velocity

        COLLISIONResponse BallVX, BallVY, -DY * DistFromCenter * WallSPEEDK * WheelANGSpeed, DX * DistFromCenter * WallSPEEDK * WheelANGSpeed, DX, DY
        DistFromCenter = (OuterRadius - R)
        BallX = DX * DistFromCenter
        BallY = DY * DistFromCenter
    End If

    If DistFromCenter < InnerRadius Then    'Beyond INNER Radius --->  Reflect Velocity


        COLLISIONResponse BallVX, BallVY, -DY * DistFromCenter * WallSPEEDK * WheelANGSpeed, DX * DistFromCenter * WallSPEEDK * WheelANGSpeed, -DX, -DY
        DistFromCenter = InnerRadius
        BallX = DX * DistFromCenter
        BallY = DY * DistFromCenter
    End If

    '
    BallVX = BallVX - DY * DistFromCenter * WheelANGSpeed * 0.003    '0.003    ' Force induced by spinning wheel Floor
    BallVY = BallVY + DX * DistFromCenter * WheelANGSpeed * 0.003

    BallVX = BallVX - DX * 0.05    ' 0.11    '0.15    '0.25    'Toward Center (Like CONE) [Wheel Slope]
    BallVY = BallVY - DY * 0.05    ' 0.11    '0.15

    BallVX = BallVX * 0.997        '.995        'GLOABAL Friction
    BallVY = BallVY * 0.997

    BallX = BallX + BallVX         'Update Position
    BallY = BallY + BallVY


    If (BallVX * BallVX + BallVY * BallVY) < 0.01 Then BallSTOPCount = BallSTOPCount + 1 Else: BallSTOPCount = 0


End Sub



Private Sub CheckCOLLISIONwihtSLOTS(DFC#)
    Dim A         As Double
    Dim CA#, SA#
    Dim Penetration#
    Const InvLineLengthSquared As Double = 0.0016    '1 / (25 * 25)  '175-150
    Const angSTeP As Double = 0.165346981767884    ' 2 PI / 38

    Dim rDIST#, rNX#, rNY#
    '    Dim rLX#, rLY#
    Dim DX#, DY#
    Dim wVX#, wVY#
    Dim TVX#, TVY#

    'Dim BALLA As Double
    'BALLA = Atan2(BallX, BallY)


    For A = -0.07 To PI2 Step angSTeP

        CA = Cos(A + WheelANG)
        SA = Sin(A + WheelANG)

        '        CalcDistFromLineAndNormal BallX + BallVX, BallY + BallVY, CA * 150, SA * 150, CA * 175, SA * 175, InvLineLengthSquared, rDIST, rNX, rNY ', rLX, rLY
        CalcDistFromLineAndNormal BallX, BallY, CA * 150, SA * 150, CA * 175, SA * 175, InvLineLengthSquared, rDIST, rNX, rNY    ', rLX, rLY
        '
        If rDIST < R + 3 Then      '4


            rNX = rNX / rDIST
            rNY = rNY / rDIST


            wVX = -SA * DFC * WheelANGSpeed * 1#    '* 1.31
            wVY = CA * DFC * WheelANGSpeed * 1#    '* 1.31

            '            TVX = BallVX
            '            TVY = BallVY

            COLLISIONResponse BallVX, BallVY, _
                              wVX, wVY, rNX, rNY

            '            BallX = BallX - TVX
            '            BallY = BallY - TVY

            Penetration = ((R + 3) - rDIST)
            BallX = BallX + rNX * Penetration    ' * 2
            BallY = BallY + rNY * Penetration    ' * 2

            '1 Step forward
            BallX = BallX + BallVX
            BallY = BallY + BallVY
            Exit For
        End If
    Next

End Sub


Private Sub COLLISIONResponse(VX1, VY1, VX2, VY2, nDX, nDY)

    Const Elasticity As Double = 0.86    '0.7
    Const Friction As Double = 0.975    '0.9

    Const MassI   As Double = 1
    Const MassJ   As Double = 999
    Const InvMassSum As Double = 0.001
    Const MassDiff As Double = -998

    Dim parIx#, parIy#             'Parallel VEL for V1
    Dim perpIx#, perpIy#           'Perpendicular VEL for V1
    Dim parJx#, parJy#             'Parallel VEL for V2
    Dim perpJx#, perpJy#           'Perpendicular VEL for V2

    parIx = VX1: parIy = VY1
    parJx = VX2: parJy = VY2

    '    'decompose velocities along collision direction (Parallel and Perpendicular)
    Project parIx, parIy, nDX, nDY
    Project parJx, parJy, nDX, nDY
    perpIx = VX1 - parIx
    perpIy = VY1 - parIy
    perpJx = VX2 - parJx
    perpJy = VY2 - parJy
    '-------------------------------

    VX1 = (parIx * MassDiff + parJx * 2 * MassJ) * InvMassSum
    VY1 = (parIy * MassDiff + parJy * 2 * MassJ) * InvMassSum
    VX2 = (parJx * -MassDiff + parIx * 2 * MassI) * InvMassSum
    VY2 = (parJy * -MassDiff + parIy * 2 * MassI) * InvMassSum

    'Apply Elasticity and friction
    VX1 = VX1 * Elasticity + perpIx * Friction
    VY1 = VY1 * Elasticity + perpIy * Friction
    VX2 = VX2 * Elasticity + perpJx * Friction
    VY2 = VY2 * Elasticity + perpJy * Friction

End Sub

Private Sub UPDATESTAT()
    Dim I         As Long
    Dim j         As Long
    Dim High      As Long
    Dim Max       As Long
    Dim MAXALL    As Long

    fMain.Text2 = "Rounds:" & NSPINS & "  Frequencies   " & Format$(100 / 38, "00.00") & vbCrLf & vbCrLf
    fMain.Text3 = "Late Numbers:" & vbCrLf & vbCrLf & vbCrLf

    For j = 1 To 38
        Max = -1000000000
        For I = 0 To 37
            If STATFreq(I) > Max Then
                High = I
                Max = STATFreq(I)
            End If
        Next

        'fMain.Text2 = fMain.Text2 & Slot2Number(High, True) & "    " & STATFreq(High) & "" & vbCrLf
        fMain.Text2 = fMain.Text2 & Slot2Number(High, True) & "  " & Format$(100 * STATFreq(High) / NSPINS, "00.000") & "%  (" & STATFreq(High) & ")" & vbCrLf
        STATFreq(High) = -STATFreq(High) - 1
    Next
    'RESTORE
    For I = 0 To 37
        STATFreq(I) = -STATFreq(I) - 1
    Next



    For I = 0 To 37
        If STATRita(I) > MAXALL Then MAXALL = STATRita(I)
    Next

    For j = 1 To 38
        Max = -1000000000
        For I = 0 To 37
            If STATRita(I) > Max Then
                High = I
                Max = STATRita(I)
            End If
        Next

        fMain.Text3 = fMain.Text3 & Slot2Number(High, True) & "  " & Format$(100 * STATRita(High) / MAXALL, "00.000") & "%  (" & STATRita(High) & ")" & vbCrLf
        STATRita(High) = -STATRita(High) - 1
    Next


    'RESTORE
    For I = 0 To 37
        STATRita(I) = -STATRita(I) - 1
    Next


    fMain.Cls
    For I = 0 To 37
        fMain.Line (I * 12 + 12, fMain.ScaleHeight - 2)-(I * 12 + 12, -2 + fMain.ScaleHeight - 2300 * STATFreq(I) / NSPINS), vbBlack
        fMain.Line (I * 12 + 15, fMain.ScaleHeight - 2)-(I * 12 + 15, -2 + fMain.ScaleHeight - 120 * STATRita(I) / MAXALL), vbGreen
    Next




End Sub
