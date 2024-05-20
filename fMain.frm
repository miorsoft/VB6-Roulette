VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Roulette"
   ClientHeight    =   11340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18750
   LinkTopic       =   "Form1"
   ScaleHeight     =   11340
   ScaleWidth      =   18750
   StartUpPosition =   1  'CenterOwner
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
      Left            =   15240
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
      Left            =   12480
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
      Width           =   2655
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
    SETUP
End Sub

Private Sub Form_Load()
    ScaleMode = vbPixels

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
