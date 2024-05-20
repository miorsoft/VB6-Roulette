VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Roulette"
   ClientHeight    =   11340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19785
   LinkTopic       =   "Form1"
   ScaleHeight     =   11340
   ScaleWidth      =   19785
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "DrawTable"
      Height          =   495
      Left            =   8160
      TabIndex        =   4
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   16920
      MultiLine       =   -1  'True
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
      Top             =   10080
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   14160
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   9720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4335
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

End Sub


Private Sub Form_Activate()
    SETUP True
End Sub

Private Sub Form_Load()
    ScaleMode = vbPixels

End Sub

Private Sub Form_Unload(Cancel As Integer)
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


    SETUP

    Set tSRF = Cairo.CreateSurface(1200, 640, ImageSurface)
    Set tCC = tSRF.CreateContext

    W = tSRF.Width
    H = tSRF.Height

    BORDER = H * 0.05

    CellX = W / 15.5
    CellY = H * 2 / 9



    With tCC

        .SelectFont "Arial", 18, vbWhite, True

        .SetSourceRGB 0, 0.7, 0: .Paint


        For X = 0 To 15
            If X >= 2 Then
                .DrawLine X * CellX, BORDER, _
                          X * CellX, BORDER + CellY * 3, , 2, vbWhite
            End If
        Next

        For Y = 0 To 3
            .DrawLine 2 / 15.5 * W, BORDER + Y * CellY, _
                      15 / 15.5 * W, BORDER + Y * CellY, , 2, vbWhite
        Next



        X = CellX * 2
        Y = BORDER + CellY * 3

        '        .SetSourceColor vbRed:         .Arc X, Y, 20: .Fill '---- TEST

        X = CellX * 2.5
        Y = BORDER + CellY * 2.5
        For I = 1 To 36

            If InStr(Slot2Number(Number2Slot(I)), "Rouge") Then
                .SetSourceColor vbRed
            Else
                .SetSourceColor vbBlack
            End If

            .Ellipse X, Y, CellX * 0.7, CellY * 0.5: .Fill

            .Save
            .TranslateDrawings X, Y
            .RotateDrawings -PIh
            '            .TextOut X - 12, Y - 12, CStr(I)
            .TextOut -12, -18, CStr(I)
            .Restore

            Y = Y - CellY
            If I Mod 3 = 0 Then
                X = X + CellX
                Y = BORDER + CellY * 2.5
            End If

        Next


    End With


'--------------------------------------------------------------------

    tSRF.WriteContentToPngFile App.Path & "\AmericanTable.png"
End Sub

