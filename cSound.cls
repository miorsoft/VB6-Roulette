VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mSoundRes As cSoundResource
Attribute mSoundRes.VB_VarHelpID = -1
Private WithEvents SAV As cSimpleAudioVolume
Attribute SAV.VB_VarHelpID = -1

Public Sub InitResourceFromMP3(FileName As String, RenderEndPoint As cMMDevice, ByVal PlayFreq As Long, Optional ByVal SkipLeadingSilence As Boolean)
    Set mSoundRes = New cSoundResource
    mSoundRes.InitFromMP3 FileName, RenderEndPoint, PlayFreq, SkipLeadingSilence
    mSoundRes.Repeat = False
End Sub


Public Sub PLAY()
    '    Set SAV = mSoundRes.ARC.AudioClientBase.GetService(SimpleAudioVolume)
    '    SAV.SetMasterVolume 0.8
    '    mSoundRes.ARC.ResetStream
    mSoundRes.ARC.StartStream
    Do
        DoEvents
    Loop While mSoundRes.SecondsPlayed And (mSoundRes.SecondsPlayed < mSoundRes.SecondsTotal)
End Sub




