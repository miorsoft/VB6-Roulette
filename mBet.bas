Attribute VB_Name = "mBet"
Option Explicit


Public BetMouseX#
Public BetMouseY#

Public BetPosX        As Long
Public BetPosY        As Long


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

Public Const FicheRadius As Double = 12


Public Sub BET()
    Dim T#

    ReDim FICHEScount(30, 10)

    BetActive = True

    T = Timer
    Do

        BetManage
        DoEvents

        DRAWALL


        If Timer > T + 9 Then Exit Do
    Loop While BetActive

End Sub

Public Sub BetManage()

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
                        CC.SetSourceRGBA 1, 1, 0, 0.6
                        CC.Arc BetMouseX, BetMouseY, FicheRadius
                        CC.Fill
                    End If
                End If
            End If
        End If
    End If


End Sub
