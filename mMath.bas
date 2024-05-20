Attribute VB_Name = "mMath"
Option Explicit

Public Const PI   As Double = 3.14159265358979
Public Const PIh  As Double = 1.5707963267949
Public Const PIh2 As Double = 0.785398163397448

Public Function Atan2(ByVal DX As Double, ByVal DY As Double) As Double
    If DX Then Atan2 = Atn(DY / DX) + PI * (DX < 0#) Else Atan2 = -PIh - (DY > 0#) * PI
End Function


