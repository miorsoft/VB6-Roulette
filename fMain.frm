VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Roulette"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23610
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   23610
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbSound 
      Height          =   375
      Left            =   7680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   10560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DrawTable"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   9600
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   16920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
   Begin VB.CheckBox chkTurbo 
      Caption         =   "TURBO"
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9480
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   14160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   9720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   " Sound"
      Height          =   255
      Left            =   7680
      TabIndex        =   6
      Top             =   10320
      Width           =   1815
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkTurbo_Click()
    TURBO = chkTurbo.Value = vbChecked
    If TURBO Then
        cmbSound.ListIndex = 0
    Else
        cmbSound.ListIndex = 2
    End If

End Sub



Private Sub cmbSound_Click()
    SoundMODE = cmbSound.ListIndex
End Sub

Private Sub Form_Activate()
    SETUP True
End Sub

Private Sub Form_Load()


    cmbSound.AddItem "No Sound"
    cmbSound.AddItem "Just Number"
    cmbSound.AddItem "All"
    cmbSound.ListIndex = 2


    ScaleMode = vbPixels

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BetActive Then
        X = X - TableX
        BetPosX = Round(X / TcX * 2)
        X = BetPosX * TcX * 0.5 + TableX

        Y = Y - TableY - TBO
        BetPosY = Round(Y / TcY * 2)
        Y = BetPosY * TcY * 0.5 + TableY + TBO

        BetMouseX = X
        BetMouseY = Y

    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BetActive Then
        If Button = 1 Then FICHEScount(BetPosX, BetPosY) = FICHEScount(BetPosX, BetPosY) + 1
        If Button = 2 Then FICHEScount(BetPosX, BetPosY) = FICHEScount(BetPosX, BetPosY) - 1: If FICHEScount(BetPosX, BetPosY) < 0 Then FICHEScount(BetPosX, BetPosY) = 0
    End If

End Sub

Private Sub Form_Resize()
    If fMain.WindowState <> vbMinimized Then
        fMain.Text1.Top = fMain.chkTurbo.Top
        fMain.Text2.Top = fMain.chkTurbo.Top
        fMain.Text3.Top = fMain.chkTurbo.Top

        fMain.Text1.Height = fMain.ScaleHeight - fMain.Text1.Top
        fMain.Text2.Height = fMain.Text1.Height
        fMain.Text3.Height = fMain.Text1.Height
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SOUNDSPLAYER = Nothing

    End
End Sub



Private Sub Command1_Click()
    '*******************************************

    'Draw Table

    Dim W#, H#
    Dim X#, Y#
    Dim X1#, Y1#
    Dim X2#, Y2#
    Dim CellX#, CellY#
    Dim BORDER#
    Dim tSRF      As cCairoSurface
    Dim tCC       As cCairoContext
    Dim I         As Long

Const SCALA As Double = 0.74



    Const FontName As String = "Times New Roman"
    Const FontSize As Double = 28 * SCALA

    SETUP

    Set tSRF = Cairo.CreateSurface(1100 * SCALA, 631 * SCALA, ImageSurface)
    Set tCC = tSRF.CreateContext

    W = tSRF.Width
    H = tSRF.Height

    BORDER = H * 0.028 '0.025
    CellX = W / 15.5
    CellY = (H - BORDER * 2) / 5



    With tCC

        .SelectFont FontName, FontSize, vbWhite, True

        .SetSourceRGB 0, 0.5, 0: .Paint


        For X = 2 To 15
            .DrawLine X * CellX, BORDER, _
                      X * CellX, BORDER + CellY * 3, , 2, vbWhite
        Next

        For Y = 0 To 3
            .DrawLine 2 / 15.5 * W, BORDER + Y * CellY, _
                      15 / 15.5 * W, BORDER + Y * CellY, , 2, vbWhite
        Next

        For Y = 4 To 5
            .DrawLine 2 / 15.5 * W, BORDER + Y * CellY, _
                      14 / 15.5 * W, BORDER + Y * CellY, , 2, vbWhite
        Next

        For X = 2 To 15 Step 4
            .DrawLine X * CellX, BORDER + CellY * 3, _
                      X * CellX, BORDER + CellY * 4, , 2, vbWhite
        Next

        For X = 2 To 15 Step 2
            .DrawLine X * CellX, BORDER + CellY * 4, _
                      X * CellX, BORDER + CellY * 5, , 2, vbWhite
        Next

        '---------------------------------
        .DrawText 2 * CellX, BORDER + CellY * 3, CellX * 4, CellY, "1er 12", True, vbCenter, , True
        .DrawText 6 * CellX, BORDER + CellY * 3, CellX * 4, CellY, "2ème 12", True, vbCenter, , True
        .DrawText 10 * CellX, BORDER + CellY * 3, CellX * 4, CellY, "3ème 12", True, vbCenter, , True


        .DrawText 2 * CellX, BORDER + CellY * 4, CellX * 2, CellY, "1 au 18 manque", , vbCenter, , True
        .DrawText 4 * CellX, BORDER + CellY * 4, CellX * 2, CellY, "Pair", True, vbCenter, , True: .SelectFont FontName, FontSize, vbRed, True
        .DrawText 6 * CellX, BORDER + CellY * 4, CellX * 2, CellY, "Rouge", True, vbCenter, , True: .SelectFont FontName, FontSize, vbBlack, True
        .DrawText 8 * CellX, BORDER + CellY * 4, CellX * 2, CellY, "Noir", True, vbCenter, , True: .SelectFont FontName, FontSize, vbWhite, True
        .DrawText 10 * CellX, BORDER + CellY * 4, CellX * 2, CellY, "Impair", True, vbCenter, , True
        .DrawText 12 * CellX, BORDER + CellY * 4, CellX * 2, CellY, "19 au 36 passe", , vbCenter, , True


        '........... NUMBERS
        X = CellX * 2.5
        Y = BORDER + CellY * 2.5
        For I = 1 To 36

            If InStr(Slot2Number(Number2Slot(I)), "Rouge") Then
                .SetSourceColor vbRed
            Else
                .SetSourceColor vbBlack
            End If

            .Ellipse X, Y, CellX * 0.7, CellY * 0.65: .Fill

            .Save
            .TranslateDrawings X, Y
            .RotateDrawings -PIh
            .DrawText -CellY * 0.5, -CellX * 0.5, CellY, CellX, CStr(I), True, vbCenter, , True
            .Restore

            Y = Y - CellY
            If I Mod 3 = 0 Then
                X = X + CellX
                Y = BORDER + CellY * 2.5
            End If

        Next
        '----------------------------

        .SetSourceColor vbWhite

        .Save
        .TranslateDrawings CellX * 1.5, BORDER + CellY * 0.75
        .RotateDrawings -PIh
        .DrawText -CellY * 0.75, -CellX * 0.5, CellY * 1.5, CellX, "00", True, vbCenter, , True
        .Rectangle -CellY * 0.75, -CellX * 0.5, CellY * 1.5, CellX: .Stroke
        .Restore

        .Save
        .TranslateDrawings CellX * 1.5, BORDER + CellY * 2.25
        .RotateDrawings -PIh
        .DrawText -CellY * 0.75, -CellX * 0.5, CellY * 1.5, CellX, "0", True, vbCenter, , True
        .Rectangle -CellY * 0.75, -CellX * 0.5, CellY * 1.5, CellX: .Stroke
        .Restore

    End With


    '--------------------------------------------------------------------

    tSRF.WriteContentToPngFile App.Path & "\AmericanTable.png"
End Sub

