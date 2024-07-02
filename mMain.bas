Attribute VB_Name = "mMain"
Option Explicit

Private SRF       As cCairoSurface
Private WheelSRF  As cCairoSurface
Private TableSRF  As cCairoSurface

Public CC         As cCairoContext

Private WheelScreenCenterX As Double
Private WheelScreenCenterY As Double

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
Private DistFromCenter As Double

Private Const BallRadius As Double = 8.25    '8  ' 7
Private Const InnerRadius As Double = 150 + BallRadius    '159
Private Const OuterRadius As Double = 293 - BallRadius

Private Const PI2 As Double = 6.28318530717959
Private Const PI  As Double = 3.14159265358979
Public Const PIh  As Double = 1.5707963267949

Private Const PI28 As Double = PI2 / 8
Private Const InvPI2 As Double = 1 / 6.28318530717959


'Private Declare Function GetTickCount Lib "kernel32" () As Long

Private SLOTN(0 To 37) As Long
Private STATFreq(0 To 37) As Long
Private STATLate(0 To 37) As Long

Public ROUNDS     As Long

Public TURBO      As Boolean
Public SoundMODE  As Long

Private BallIsSTILLCount As Long

Private CNT       As Long

Public RenderDev  As cMMDevice

Public SOUNDSPLAYER As cSounds

Public NumberExtracted As Long

Private Const WallSPEEDK As Double = 1

Private SrfFlat   As cCairoSurface

Public DoHighlightBALLPosition As Boolean


Public Function RndM(Optional ByVal Number As Long) As Double
'https://www.vbforums.com/showthread.php?899623-Random-repeatable-numbers-that-do-NOT-depend-on-prior-values-for-the-next-result&p=5600596&viewfull=1#post5600596
    Static Ri     As Double

    If Number Then Ri = Number

    Ri = Ri * (1.241 + (Ri > 983732.3)) + 1.737
    RndM = Ri - Int(Ri)
    
End Function
Public Sub SETUP(Optional andLAUNCH As Boolean = False)
'    Randomize Timer
    RndM Timer
    

    Set WheelSRF = Cairo.ImageList.AddImage("WHEEL", App.Path & "\RouletteWheel.png")
    Set TableSRF = Cairo.ImageList.AddImage("TABLE", App.Path & "\AmericanTable.png")



    WheelImageRadius = WheelSRF.Width * 0.5
    WheelScreenCenterX = WheelImageRadius + 10
    WheelScreenCenterY = WheelImageRadius + 10

    '    OuterRadius = WheelImageRadius - 14    '24



    ''' Chenge external White with green
    With WheelSRF.CreateContext
        .SetSourceRGB 0, 0.5, 0: .Paint
        .Arc WheelImageRadius, WheelImageRadius, WheelImageRadius
        .Clip
        .Paint , Cairo.ImageList.AddImage("WHEEL", App.Path & "\RouletteWheel.png").CreateSurfacePattern
    End With
    '----------------------








    TableW = TableSRF.Width
    TableH = TableSRF.Height
    TableScreenX = WheelImageRadius * 2 + 20
    TableScreenY = 0

    TableTopBorder = TableH * 0.05    '0.025
    TableCellW = TableW / 15.5
    TableCellH = (TableH - TableTopBorder * 2) / 5


    fMain.Text1.Left = TableScreenX + TableCellW * 2    'WheelScreenCenterX * 2
    fMain.Text2.Left = fMain.Text1.Width + fMain.Text1.Left + 5
    fMain.Text3.Left = fMain.Text2.Width + fMain.Text2.Left + 5


    fMain.PICpanel.Left = TableScreenX + TableCellW * 2
    fMain.PICpanel.Top = TableScreenY + TableH
    fMain.PICpanel.Width = TableW - TableCellW * 3.5


    fMain.chkTurbo.Left = fMain.PICpanel.Width - fMain.chkTurbo.Width - 5
    fMain.Label1.Left = fMain.PICpanel.Width - fMain.Label1.Width - 5
    fMain.cmbSound.Left = fMain.PICpanel.Width - fMain.cmbSound.Width - 5


    fMain.lBudget.Width = fMain.PICpanel.Width / 4


    fMain.lBet.Left = fMain.lBudget.Left + fMain.lBudget.Width
    fMain.lBet.Width = fMain.lBudget.Width
    fMain.lWin.Width = fMain.lBudget.Width

    fMain.lWin.Left = fMain.lBet.Left + fMain.lBet.Width


    Set SRF = Cairo.CreateSurface(WheelImageRadius * 2 + 20 + TableSRF.Width, WheelImageRadius * 2 + 20 + 200, ImageSurface)
    Set CC = SRF.CreateContext
    CC.AntiAlias = CAIRO_ANTIALIAS_FAST
    CC.SetLineWidth 1
    'CC.SetSourceColor vbWhite: CC.Paint
    CC.SetSourceRGB 0, 0.5, 0: CC.Paint




    SLOTN(0) = 33: SLOTN(1) = 7: SLOTN(2) = 17: SLOTN(3) = 5: SLOTN(4) = 22: SLOTN(5) = 34: SLOTN(6) = 15: SLOTN(7) = 3
    SLOTN(8) = 24: SLOTN(9) = 36: SLOTN(10) = 13: SLOTN(11) = 1: SLOTN(12) = -1: SLOTN(13) = 27: SLOTN(14) = 10: SLOTN(15) = 25
    SLOTN(16) = 29: SLOTN(17) = 12: SLOTN(18) = 8: SLOTN(19) = 19: SLOTN(20) = 31: SLOTN(21) = 18: SLOTN(22) = 6: SLOTN(23) = 21
    SLOTN(24) = 26: SLOTN(25) = 16: SLOTN(26) = 4: SLOTN(27) = 23: SLOTN(28) = 35: SLOTN(29) = 14: SLOTN(30) = 2: SLOTN(31) = 0
    SLOTN(32) = 20: SLOTN(33) = 9: SLOTN(34) = 28: SLOTN(35) = 32: SLOTN(36) = 11: SLOTN(37) = 30



    '------------ CREATE SOUNDS ---------------
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
    '------------ END CREATE SOUNDS ---------------


    Set SOUNDSPLAYER = New cSounds



    SETUPWINTABLE


    SETUPFLAT

    If andLAUNCH Then LAUNCH


End Sub


Public Sub LAUNCH()

    '    WheelANGSpeed = 0.25 + (rndm * 2 - 1) * 0.07
    WheelANGSpeed = 0.2 + (RndM * 2 - 1) * 0.07 '0.05


    BallX = (RndM * 2 - 1) * 5 '3
    BallY = -(OuterRadius - BallRadius) + 4 + RndM * 14
    '    BallVX = rndm * 8
    '    BallVX = (rndm * 2 - 1) * 8
    BallVX = (RndM * 3 - 2) * 8

    BallVY = 0

    While WheelANG > PI2: WheelANG = WheelANG - PI2: Wend


    ROUNDS = ROUNDS + 1

    'If SoundMODE > 1 Then PlayMP3 App.Path & "\Sounds\Faites vos jeux.MP3"
    If SoundMODE > 1 Then SOUNDSPLAYER.PlaySound "Faites vos jeux.MP3", 0, 0, 1000
DoHighlightBALLPosition = False

    TotalBet = 0
    fMain.lBet = "Bet: 0"
    fMain.lWin = "WIN: 0"
    fMain.lBudget = "Budget: " & BUDGET

    '-<<<<--------- WAIT BETS
    If Not (TURBO) Then BET
    '--------------------------

    'If SoundMODE > 1 Then PlayAsync App.Path & "\Sounds\Les Jeux sont faits.MP3"
    If SoundMODE > 1 Then SOUNDSPLAYER.PlaySound "Les Jeux sont faits.MP3"

DoHighlightBALLPosition = True


    WHEELLOOP

End Sub
Public Function Slot2Number(Slot As Long, Optional JustNumber As Boolean = False) As String
    Dim Result&
    Result = SLOTN(Slot)
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
        If SLOTN(I) = N Then
            Number2Slot = I
            Exit For
        End If
    Next
End Function


Private Sub ShowResult()
    Dim N         As Long
    Dim S         As String
    Dim I&

    N = (-WheelANG + Atan2(BallX, BallY)) / PI2 * 38
    While N < 0: N = N + 38: Wend
    N = N Mod 38

    STATFreq(N) = STATFreq(N) + 1

    For I = 0 To 37
        STATLate(I) = STATLate(I) + 1
    Next
    STATLate(N) = 0

    NumberExtracted = SLOTN(N)

    S = Slot2Number(N)
    S = S & "  (" & STATFreq(N) & ")"
    fMain.Text1 = S & vbCrLf & fMain.Text1

    ' Test Text Length
    If Len(fMain.Text1) > 1600 Then fMain.Text1 = Left$(fMain.Text1, 1600)


    UPDATESTAT




    If SoundMODE = 1 Then
        'SoundSLOT(N).PLAY
        Slot2MP3 N, S
        SOUNDSPLAYER.PlaySound S & ".MP3"

    Else
        'If SoundMODE <> 0 Then PlayMP3 Slot2MP3(N, S)
        If SoundMODE <> 0 Then
            Slot2MP3 N, S
            SOUNDSPLAYER.PlaySound S & ".MP3", , , 3000
        End If
    End If

    MANAGEBETS


    LAUNCH

End Sub


Public Sub WHEELLOOP()

    If TURBO Then
        ComputedFPS = 650          ' 500          '400
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

    BallIsSTILLCount = 0
    CNT = 0



    Do

        Select Case TEMPO.WaitForNext

            Case tCompute

                SIMULATE

                If WheelANGSpeed < 0 Then WheelANGSpeed = 0

                If BallIsSTILLCount > 50 Then Exit Do

                CNT = CNT + 1
                ' If Not (TURBO) Then
                If CNT > 450 Then
                    CNT = -100000000#
                    ' PlayMP3 App.Path & "\Sounds\Rien ne va plus.MP3"
                    'If SoundMODE > 1 Then PlayAsync App.Path & "\Sounds\Rien ne va plus.MP3"
                    If SoundMODE > 1 Then SOUNDSPLAYER.PlaySound "Rien ne va plus.MP3"

                    BetActive = False
                End If
                'End If

            Case tDRAW
                DRAWALL
                
                DoEvents

            Case t1sec
                '                fMain.Caption = "  Computed FPS:" & TEMPO.Count(tCompute) & " DrawnFPS:" & TEMPO.Count(tDRAW)
                TEMPO.ResetCount (tCompute)
                TEMPO.ResetCount (tDRAW)
                TEMPO.ResetCount (t1sec)
        End Select

    Loop While True

    ShowResult

End Sub

Public Sub DRAWALL(Optional DoHighlight As Boolean, Optional SleepMS As Long)
    Dim N         As Long

    '   CC.SetSourceColor vbWhite: CC.Paint

    CC.Save
    CC.TranslateDrawings WheelScreenCenterX, WheelScreenCenterY
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

    CC.Arc BallX + WheelScreenCenterX, BallY + WheelScreenCenterY, BallRadius    '+ 1
    CC.Fill True, Cairo.CreateSolidPatternLng(vbWhite)
    CC.SetSourceColor 0
    CC.Stroke



    CC.RenderSurfaceContent TableSRF, TableScreenX, TableScreenY



    DRAWBets

    If FichesOUTAnim Then DRAWfichesPilesAt FichesOUTX, FichesOUTY, FichesOUTAmount

    If FichesINAnim Then DRAWfichesPilesAt FichesINX, FichesINY, FichesINAmount



    If DoHighlight Then HILightBET    'HIlight CStr(Timer * 7 Mod 36 + 1)




    '''    CC.RenderSurfaceContent SrfFlat, 5, WheelImageRadius * 2 + 20
    '''    Dim X#, Y#
    '''    Y = WheelImageRadius * 2 + 20
    '''    Y = Y + 240 * (1 - DistFromCenter / WheelImageRadius)
    '''    X = (-WheelANG + Atan2(BallX, BallY))
    '''    While X < 0: X = X + PI2: Wend
    '''    X = X / PI2 * 720
    '''    CC.Arc X, Y, BallRadius * 0.78
    '''    CC.Fill True, Cairo.CreateSolidPatternLng(vbWhite)
    '''    CC.SetSourceColor 0
    '''    CC.Stroke

    If DoHighlightBALLPosition Then
        N = (-WheelANG + Atan2(BallX, BallY)) * InvPI2 * 38
        While N < 0: N = N + 38: Wend
        N = N Mod 38
        HIlight CStr(SLOTN(N)), 0.33
    End If


    fMain.Picture = SRF.Picture
    '    SRF.DrawToDC fMain.hDC
    'fMain.Refresh
If SleepMS Then New_c.SleepEx SleepMS

End Sub

'Private Sub Reflect(X#, Y#, ByVal WallX#, ByVal WallY#)
'    Dim D#
'    D = X * WallX + Y * WallY
'    X = X - WallX * D * 2
'    Y = Y - WallY * D * 2
'End Sub
Private Sub Project(X#, Y#, ByVal NX#, ByVal NY#)
    Dim D         As Double
    D = X * NX + Y * NY
    X = NX * D
    Y = NY * D
End Sub


Public Function Atan2(ByVal DX As Double, ByVal DY As Double) As Double
    If DX Then Atan2 = Atn(DY / DX) + PI * (DX < 0#) Else Atan2 = -PIh - (DY > 0#) * PI
End Function


Private Sub CalcDistFromLineAndNormal2(ByVal PX#, ByVal PY#, ByVal AX#, ByVal AY#, ByVal BX#, ByVal BY#, ByVal InvABlen2#, rDIST#, rNX#, rNY#)    ', rPosX#, rPosY#)
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
    rDIST = (DX * DX + DY * DY)    ' Will be SQR Later

    rNX = DX                       ' / rDIST 'Will be normalized later
    rNY = DY                       ' / rDIST
End Sub


Public Sub SIMULATE()


    Dim invDFC#

    Dim DX#, DY#
    Dim A#



    WheelANG = WheelANG + WheelANGSpeed

    '    WheelANGSpeed = WheelANGSpeed * 0.997
    WheelANGSpeed = WheelANGSpeed * 0.9974

    '    If WheelANGSpeed > 0.000005 Then WheelANGSpeed = WheelANGSpeed - 0.000005
    If WheelANGSpeed > 0.0000025 Then WheelANGSpeed = WheelANGSpeed - 0.0000025

    DistFromCenter = Sqr(BallX * BallX + BallY * BallY)
    invDFC = 1# / DistFromCenter

    DX = BallX * invDFC
    DY = BallY * invDFC


    If DistFromCenter < 183 Then   ' CHECK CHELL SLOT
        CheckCOLLISIONwihtSLOTS
    End If

    If DistFromCenter > OuterRadius Then    'Beyond OUTER Radius --->  Reflect Velocity
        COLLISIONResponse BallVX, BallVY, -DY * DistFromCenter * WallSPEEDK * WheelANGSpeed, DX * DistFromCenter * WallSPEEDK * WheelANGSpeed, DX, DY
        DistFromCenter = OuterRadius
        BallX = DX * DistFromCenter
        BallY = DY * DistFromCenter
    End If

    If DistFromCenter < InnerRadius Then    'Beyond INNER Radius --->  Reflect Velocity
        COLLISIONResponse BallVX, BallVY, -DY * DistFromCenter * WallSPEEDK * WheelANGSpeed, DX * DistFromCenter * WallSPEEDK * WheelANGSpeed, -DX, -DY
        DistFromCenter = InnerRadius
        BallX = DX * DistFromCenter
        BallY = DY * DistFromCenter
    End If

    '--------ROMBUS  deviation
    If DistFromCenter > 267.5 Then
        If DistFromCenter < 269.5 Then
            A = Atan2(BallX, BallY)
            While A < 0: A = A + PI2: Wend
            A = (-WheelANG + A)
            While A < 0#: A = A + PI2: Wend
            While A > PI28: A = A - PI28: Wend
            A = A + 0.006
            If A > 0.29 And A < 0.49 Then
                BallVY = BallVY + (RndM * 2 - 1) * (WheelANGSpeed + 0.01) * 20
                BallVX = BallVX + (RndM * 2 - 1) * (WheelANGSpeed + 0.01) * 20
            End If
        End If
    End If
    '------------------


    '
    BallVX = BallVX - DY * DistFromCenter * WheelANGSpeed * 0.003    '0.003    ' Force induced by spinning wheel Floor
    BallVY = BallVY + DX * DistFromCenter * WheelANGSpeed * 0.003

    BallVX = BallVX - DX * 0.055   '0.05    '    '0.15    '0.25    'Toward Center (Like CONE) [Wheel Slope]
    BallVY = BallVY - DY * 0.055   '0.05    '    '0.15

    BallVX = BallVX * 0.997        '.995        'GLOABAL Friction
    BallVY = BallVY * 0.997

    BallX = BallX + BallVX         'Update Position
    BallY = BallY + BallVY


    If (BallVX * BallVX + BallVY * BallVY) < 0.01 Then BallIsSTILLCount = BallIsSTILLCount + 1 Else: BallIsSTILLCount = 0


End Sub



Private Sub CheckCOLLISIONwihtSLOTS()
    Dim A         As Double
    Dim CA#, SA#
    Dim Penetration#
    Const angSTeP As Double = 0.165346981767884    ' 2 PI / 38
    Const SlotThick As Double = 1.8
    Const rIN     As Double = 150
    Const rOUT    As Double = 173.5    '175
    Const InvLineLengthSquared As Double = 1 / ((rOUT - rIN) * (rOUT - rIN))

    Const MinDist As Double = (BallRadius + SlotThick)
    Const MinDist2 As Double = MinDist * MinDist

    Dim RetDIST#, rNX#, rNY#

    Dim wVX#, wVY#

    '    For A = -0.07 To PI2 Step angSTeP
    For A = -0.065 To PI2 Step angSTeP

        CA = Cos(A + WheelANG)
        SA = Sin(A + WheelANG)

        CalcDistFromLineAndNormal2 BallX, BallY, CA * rIN, SA * rIN, CA * rOUT, SA * rOUT, InvLineLengthSquared, RetDIST, rNX, rNY
        '
        If RetDIST < MinDist2 Then
            RetDIST = Sqr(RetDIST)
            rNX = rNX / RetDIST
            rNY = rNY / RetDIST

            wVX = -SA * DistFromCenter * WheelANGSpeed * WallSPEEDK
            wVY = CA * DistFromCenter * WheelANGSpeed * WallSPEEDK


            COLLISIONResponse BallVX, BallVY, _
                              wVX, wVY, rNX, rNY

            Penetration = (MinDist - RetDIST)
            BallX = BallX + rNX * Penetration
            BallY = BallY + rNY * Penetration

            '1 Step forward
            BallX = BallX + BallVX
            BallY = BallY + BallVY
            Exit For
        End If
    Next

End Sub


Private Sub COLLISIONResponse(VX1, VY1, VX2, VY2, nDX, nDY)

    Const Elasticity As Double = 0.85 ' 0.86
    Const Friction As Double = 0.98 ' 0.975    '0.9

    Const MassI   As Double = 1
    Const MassJ   As Double = 999
    Const InvMassSum As Double = 0.001
    Const MassDiff As Double = -998

    Dim parIx#, parIy#             'Parallel      to nDX,nDY  VEL for V1
    Dim perpIx#, perpIy#           'Perpendicular to nDX,nDY  VEL for V1
    Dim parJx#, parJy#             'Parallel      to nDX,nDY  VEL for V2
    Dim perpJx#, perpJy#           'Perpendicular to nDX,nDY  VEL for V2

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

    Dim Volume    As Long
    Dim DX#, DY#

    If Not (TURBO) Then
        DX = parIx - parJx
        DY = parIy - parJy
        Volume = -7000 + Log(1 + (DX * DX + DY * DY) ^ 0.125) * 8000
        If Volume > -0 Then Volume = -0
        If Volume > -4000 Then SOUNDSPLAYER.PlaySound "Ball.MP3", , Volume

        'Debug.Print Volume
    End If



End Sub

Private Sub UPDATESTAT()
    Dim I         As Long
    Dim j         As Long
    Dim High      As Long
    Dim Max       As Long
    Dim MAXALL    As Long

    fMain.Text2 = "Rounds:" & ROUNDS & "  Frequencies   " & format$(100 / 38, "00.00") & vbCrLf & vbCrLf
    fMain.Text3 = "Late Numbers:" & vbCrLf & vbCrLf
    For j = 1 To 38
        Max = -1000000000
        For I = 0 To 37
            If STATFreq(I) > Max Then
                High = I
                Max = STATFreq(I)
            End If
        Next

        'fMain.Text2 = fMain.Text2 & Slot2Number(High, True) & "    " & STATFreq(High) & "" & vbCrLf
        fMain.Text2 = fMain.Text2 & Slot2Number(High, True) & "  " & format$(100 * STATFreq(High) / ROUNDS, "00.000") & "%  (" & STATFreq(High) & ")" & vbCrLf
        STATFreq(High) = -STATFreq(High) - 1
    Next
    'RESTORE
    For I = 0 To 37
        STATFreq(I) = -STATFreq(I) - 1
    Next



    For I = 0 To 37
        If STATLate(I) > MAXALL Then MAXALL = STATLate(I)
    Next

    For j = 1 To 38
        Max = -1000000000
        For I = 0 To 37
            If STATLate(I) > Max Then
                High = I
                Max = STATLate(I)
            End If
        Next

        fMain.Text3 = fMain.Text3 & Slot2Number(High, True) & "  " & format$(100 * STATLate(High) / MAXALL, "00.000") & "%  (" & STATLate(High) & ")" & vbCrLf
        STATLate(High) = -STATLate(High) - 1
    Next


    'RESTORE
    For I = 0 To 37
        STATLate(I) = -STATLate(I) - 1
    Next


    '---- basic Histogram
    '    fMain.Cls
    '    For I = 0 To 37
    '        fMain.Line (I * 12 + 12, fMain.ScaleHeight - 2)-(I * 12 + 12, -2 + fMain.ScaleHeight - 2300 * STATFreq(I) / ROUNDS), vbBlack
    '        fMain.Line (I * 12 + 15, fMain.ScaleHeight - 2)-(I * 12 + 15, -2 + fMain.ScaleHeight - 120 * STATLate(I) / MAXALL), vbGreen
    '    Next
    '-------------

    fMain.Refresh
    DRAWALL


End Sub

Private Sub SETUPFLAT()
    Dim BW()      As Long
    Dim BF()      As Long
    Dim X#, Y#
    Dim A#
    Dim R#
    Dim ASt#

    Set SrfFlat = Cairo.CreateSurface(720, 241, ImageSurface)

    SrfFlat.BindToArrayLong BF
    WheelSRF.BindToArrayLong BW

    ASt = PI2 / 720

    'For A = 0# To PI2 Step PI2 / 699
    For A = 0.05 To PI2 + 0.05 - ASt Step ASt
        Y = 0
        For R = WheelImageRadius - 0.5 To 0 Step -1
            BF(X, Y) = BW(WheelImageRadius + Cos(A) * R, WheelImageRadius + Sin(A) * R)
            Y = Y + 240 / WheelImageRadius
        Next
        X = X + 1
    Next


    '    For Y = 0 To 240
    '        R = 240 - Y / 241 * WheelImageRadius
    '        For X = 0 To 719
    '            A = X / 720 * PI2
    '            BF(X, Y) = BW(WheelImageRadius + Cos(A) * R, WheelImageRadius + Sin(A) * R)
    '        Next
    '    Next





    SrfFlat.ReleaseArrayLong BF
    WheelSRF.ReleaseArrayLong BW


End Sub
