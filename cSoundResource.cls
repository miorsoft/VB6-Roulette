VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSoundResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit                    'a Class for playing PCM-Resources in 16Bit, Stereo (initializable from MP3-resources or with "Raw-16Bit-Stereo-WAV-Data")

Event Progress(ByVal CurFrameIdx As Long, ByVal TotalFrames As Long, ByVal Percent As Double)

Public WithEvents ARC As cAudioRenderClient, Repeat As Boolean
Attribute ARC.VB_VarUserMemId = 1073938432

Private mPlayFreq As Long, OrigFreq As Long, Channels As Long, ID3Offs As Long
Attribute OrigFreq.VB_VarUserMemId = 1073938433
Attribute Channels.VB_VarUserMemId = 1073938433
Attribute ID3Offs.VB_VarUserMemId = 1073938433
Private FramesTotal As Long, FrameData() As Long, FrameIdxToCopyFrom As Long
Attribute FramesTotal.VB_VarUserMemId = 1073938437
Attribute FrameData.VB_VarUserMemId = 1073938437
Attribute FrameIdxToCopyFrom.VB_VarUserMemId = 1073938437

Public Sub InitFromMP3(FileNameOrByteArray, RenderEndPoint As cMMDevice, ByVal PlayFreq As Long, Optional ByVal SkipLeadingSilence As Boolean)
    Dim Src() As Byte, Dst() As Long, DstBytesUsed As Long, msEstimated As Long
    Src = FileNameOrByteArray
    If VarType(FileNameOrByteArray) = vbString Then Src = New_c.FSO.ReadByteContent(CStr(FileNameOrByteArray))
    With New_c.MP3Resource
        .GetMP3Info Src, OrigFreq, Channels, ID3Offs, msEstimated
        ReDim Dst(0 To msEstimated / 1000 * OrigFreq * 1.5)
        .OpenConverter OrigFreq, Channels, 16
        .ConvertBuffer VarPtr(Src(ID3Offs)), UBound(Src) + 1 - ID3Offs, VarPtr(Dst(0)), 4 * (UBound(Dst) + 1), DstBytesUsed, True
    End With
    ReDim Preserve Dst(DstBytesUsed \ 4 - 1)
    InitWithFrames Dst, RenderEndPoint, PlayFreq, SkipLeadingSilence
End Sub

'a "Frame" in this Class, is "a complete 16Bit-Stereo-Sample" (two bytes for each stereo-channel = 4Bytes per Sample)
'therefore we use Long-Arrays to store these "Frames" ... and the amount of frames divided by the Wav-Frequency,
'gives the time of this "Frames-Resource" in seconds (e.g. a FramesData-Array with a length of 88200 equals 2 seconds of sound (when a Freq of 44100 is used)
Private Sub InitWithFrames(Frames() As Long, RenderEndPoint As cMMDevice, ByVal PlayFreq As Long, Optional ByVal SkipLeadingSilence As Boolean)
    FrameData = Frames
    If SkipLeadingSilence Then     'this is useful for sound-resources (like e.g. from a DrumKit), which don't start their sound immediately
        Dim I As Long, j As Long, F As Long
        For I = 0 To UBound(Frames)
            F = Frames(I) And &HFFFF&: If F > 32767 Then F = F - 65536
            If Abs(F) > 512 Then Do Until I > UBound(Frames): FrameData(j) = Frames(I): I = I + 1: j = j + 1: Loop
    Next
    If j Then ReDim Preserve FrameData(j - 1)
End If

FramesTotal = UBound(FrameData) + 1
mPlayFreq = PlayFreq

Set ARC = New_c.AudioRenderClient
ARC.InitOn RenderEndPoint, PlayFreq, 2
End Sub

Public Property Get SecondsTotal() As Double
    If ARC Is Nothing Then Exit Property
    SecondsTotal = FramesTotal / mPlayFreq
End Property
Public Property Get SecondsPlayed() As Double
    If ARC Is Nothing Then Exit Property
    SecondsPlayed = FrameIdxCurrent / FramesTotal * SecondsTotal
End Property

Public Property Get FrameIdxCurrent() As Long

    If ARC Is Nothing Then Exit Property
    FrameIdxCurrent = (FrameIdxToCopyFrom - ARC.CalcFrameCountAheadOfClientSrc + 1000 * FramesTotal) Mod FramesTotal
End Property

Private Sub ARC_FillNextBuffer(ByVal pBuf As Long, ByVal FramesAvailable As Long, FramesFilled As Long, BufferFlags As RC6.AUDCLNT_BUFFERFLAGS)
    FramesFilled = IIf(FramesAvailable > FramesTotal - FrameIdxToCopyFrom, FramesTotal - FrameIdxToCopyFrom, FramesAvailable)
    If FramesFilled Then
        New_c.MemCopy pBuf, VarPtr(FrameData(FrameIdxToCopyFrom)), FramesFilled * 4
        FrameIdxToCopyFrom = FrameIdxToCopyFrom + FramesFilled
    ElseIf Repeat Then
        FrameIdxToCopyFrom = 0
    End If
    RaiseProgress
End Sub

Private Sub ARC_Reset()
    FrameIdxToCopyFrom = 0
    RaiseProgress
End Sub

Private Sub RaiseProgress()
    Dim CurIdx    As Long: CurIdx = FrameIdxCurrent
    RaiseEvent Progress(CurIdx, FramesTotal, CurIdx / FramesTotal)
End Sub

