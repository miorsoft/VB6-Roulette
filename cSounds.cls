VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Raw / Not Elagant  use of DirectSound_Functions.bas to play multiple sounds....

'(By theTrick?)

Private Type tMyBuff
    Buffer        As IDirectSoundBuffer8
    FileName      As String
End Type

Private DirectSound As DirectSound8

Private Const BUFFLEN As Long = 31

Private SoundsSET() As tMyBuff

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Class_Initialize()
    Dim desc      As DSBUFFERDESC
    Dim I         As Long

    Set DirectSound = New DirectSound8
    DirectSound.Initialize ByVal 0
    DirectSound.SetCooperativeLevel fMain.hWnd, DSSCL_NORMAL


    ReDim SoundsSET(BUFFLEN)

    desc.dwFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN    'Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLFX


    For I = 0 To BUFFLEN
        SoundsSET(I).FileName = "Ball.MP3"
        Set SoundsSET(I).Buffer = DSCreateSoundBufferFromFile(DirectSound, App.Path & "\Sounds\" & SoundsSET(I).FileName, desc)
    Next


End Sub



Public Sub PlaySoundOLD(S As String, Optional Pan As Long = 0, Optional Volume As Long = 0)
    Dim desc      As DSBUFFERDESC
    Dim I         As Long

    desc.dwFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN


    For I = 0 To BUFFLEN

        With SoundsSET(I)
            If .FileName = S Then
                If .Buffer.GetStatus <> DSBSTATUS_PLAYING Then
                    .Buffer.SetPan Pan
                    .Buffer.SetVolume Volume
                    .Buffer.Play 0, 0, 0
                    Exit For
                End If
            Else
                If .Buffer.GetStatus <> DSBSTATUS_PLAYING Then
                    Set .Buffer = DSCreateSoundBufferFromFile(DirectSound, App.Path & "\Sounds\" & S, desc)

                    .Buffer.SetPan Pan
                    .Buffer.SetVolume Volume
                    .Buffer.Play 0, 0, 0

                    .FileName = S
                    Exit For
                End If
            End If
        End With

    Next

End Sub



'Volume from 0 to -10000


Public Sub PlaySound(S As String, Optional Pan As Long = 0, Optional Volume As Long = 0, Optional andSleep As Long = 0)
    Dim desc      As DSBUFFERDESC
    Dim I         As Long
    Dim CanExitSub As Long

    desc.dwFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN

    For I = 0 To BUFFLEN
        With SoundsSET(I)
            If .Buffer.GetStatus <> DSBSTATUS_PLAYING Then
                If .FileName = S Then
                    .Buffer.SetPan Pan
                    .Buffer.SetVolume Volume
                    .Buffer.Play 0, 0, 0
                    CanExitSub = -1
                    Exit For
                End If
            End If
        End With
    Next

    If CanExitSub Then Exit Sub

    For I = 0 To BUFFLEN
        With SoundsSET(I)
            If .Buffer.GetStatus <> DSBSTATUS_PLAYING Then
                Set .Buffer = DSCreateSoundBufferFromFile(DirectSound, App.Path & "\Sounds\" & S, desc)
                .Buffer.SetPan Pan
                .Buffer.SetVolume Volume
                .Buffer.Play 0, 0, 0
                .FileName = S
                Exit For
            End If
        End With
    Next

    If andSleep Then Sleep andSleep

End Sub

Private Sub Class_Terminate()
    Dim I         As Long

    For I = 0 To BUFFLEN
        Set SoundsSET(I).Buffer = Nothing
    Next

    Set DirectSound = Nothing

End Sub

''Public Property Get DataSource() As DataSource
''
''End Property
''
''Public Property Set DataSource(ByVal objDataSource As DataSource)
''
''End Property
''
''Public Property Get DataMember() As DataMember
''
''End Property
''
''Public Property Let DataMember(ByVal DataMember As DataMember)
''
''End Property
