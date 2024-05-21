Attribute VB_Name = "modGoogleSpeak"
Option Explicit
'-----------------------------------------------------------------------------------------------------
'Autor: Leandro Ascierto
'Web: www.leandroascierto.com.ar
'Abreviaturas
'de, da, es, fi, fr, en, it, nl, pl, pt, sv"
'Alemán , Danés, Español, Finlandia, Francés, Inglés, Italiano, Neerlandés, Polaco, Portugués, Sueco
'----------------------------------------------------------------------------------------------------

'Environ("Temp")

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function GoogleSpeakLess100chars(ByVal sText As String, Optional ByVal Language As String = "es", Optional ByVal bDoEvents As Boolean = True) As Boolean
    'Leandro Ascierto
    On Error Resume Next
    Dim sTempPath As String, ML As String
    Dim FileLength As Long

    'sText = Replace(sText, vbCrLf, " ")

    If Len(sText) > 100 Then Exit Function

    sTempPath = App.Path & "\Sounds\.MP3"

    If URLDownloadToFile(0&, "https://translate.google.com/translate_tts?tl=" & Language & "&q=" & sText & "&client=tw-ob", sTempPath, 0&, 0&) = 0 Then

        If mciSendString("open " & Chr$(34) & sTempPath & Chr$(34) & " type MpegVideo" & " alias myfile", 0&, 0&, 0&) = 0 Then
            ML = String(30, 0)
            Call mciSendString("status myfile length ", ML, 30, 0&)
            FileLength = Val(ML)
            If FileLength Then
                If mciSendString("play myFile", 0&, 0&, 0&) = 0 Then
                    Do While mciSendString("status myfile position ", ML, 30, 0&) = 0
                        If Val(ML) = FileLength Then GoogleSpeakLess100chars = True: Exit Do
                        If bDoEvents Then DoEvents
                    Loop
                End If
            End If
            Call mciSendString("close myfile", 0&, 0&, 0&)
        End If

        Kill sTempPath
    End If

End Function

Private Function GoogleSpeakCreateMP3(Number As Long, ByVal sText As String, Optional ByVal Language As String = "es", Optional ByVal bDoEvents As Boolean = True) As Boolean

    On Error Resume Next
    Dim FilePathName As String
    Dim ML        As String
    Dim FileLength As Long

    'sText = Replace(sText, vbCrLf, " ")

    If Len(sText) > 100 Then Exit Function

    FilePathName = App.Path & "\Sounds\" & Format(Number, "000") & ".MP3"

    URLDownloadToFile 0&, "https://translate.google.com/translate_tts?tl=" & Language & "&q=" & sText & "&client=tw-ob", FilePathName, 0&, 0&

End Function

Public Function GoogleSpeakCreateMP3_2(FilePathName As String, ByVal sText As String, Optional ByVal Language As String = "es", Optional ByVal bDoEvents As Boolean = True) As Boolean

    On Error Resume Next
    '    Dim FilePathName As String
    Dim ML        As String
    Dim FileLength As Long

    'sText = Replace(sText, vbCrLf, " ")

    If Len(sText) > 100 Then Exit Function

    '    FilePathName = App.Path & "\Sounds\" & Format(Number, "000") & ".MP3"

    URLDownloadToFile 0&, "https://translate.google.com/translate_tts?tl=" & Language & "&q=" & sText & "&client=tw-ob", FilePathName, 0&, 0&

End Function

'reexre:
Public Sub SpeekMoreThan100(ByRef sText As String, Lang As String)
    Dim S         As String
    Dim SS()      As String
    Dim NS        As Long
    Dim M         As String
    Dim I         As Long
    Dim II        As Long
    Dim II2       As Long


    S = sText

    S = Right$(S, Len(S) - 2)      'remove 1st 2 char (to remove first vbCrLf)

    S = Replace(S, "&", "and")
    S = Replace(S, Chr$(34), "")
    S = Replace(S, "=", " equal to ")


    S = Replace(S, vbCrLf, ". ")
    S = Replace(S, vbLf, ". ")
    S = Replace(S, vbCr, ". ")



    If Len(S) > 100 Then
        Do

            I = I + 1
            If I > Len(S) Then
                NS = NS + 1
                ReDim Preserve SS(NS)
                SS(NS) = S
Debug.Print SS(NS) & "-----------"
                Exit Do
            End If

            M = Mid$(S, I, 1)

            If M = "." Or M = ":" Or M = "," Or M = ";" Or M = "!" Then II = I
            If M = "(" Or M = ")" Or M = "[" Or M = "]" Then II = I

            If M = " " And I < 100 Then II2 = I
            If I >= 100 Then


                NS = NS + 1
                ReDim Preserve SS(NS)
                If II = 0 Then II = II2
                SS(NS) = Left$(S, II)
                S = Right$(S, Len(S) - II)
Debug.Print SS(NS) & "-----------"
                If Len(SS(NS)) = 0 Then Exit Do
                I = 0
                II = 0
                II2 = 0
            End If
        Loop While True

    Else
        ReDim SS(1): NS = 1
        SS(1) = S
    End If


    'For I = 1 To NS
    '    GoogleSpeak SS(I), Lang
    'Next

    If NS = 1 Then
        GoogleSpeakLess100chars SS(1), Lang
    Else
        'fmain.PB.Max = NS
        For I = 1 To NS
            fMain.Caption = "creating portion  " & I & " / " & NS
            GoogleSpeakCreateMP3 I, SS(I), Lang
            'fmain.PB.Value = I
            DoEvents
        Next
        PlayMP3 (JoinMP3(NS))
    End If

    fMain.Caption = "Ready."



End Sub

Private Function JoinMP3(N As Long) As String

    Dim I         As Long
    Dim Names()   As String
    Dim MP3()     As String
    Dim fLen      As Long
    Dim allMP3    As String
    Dim ALLMP3FN  As String

    ALLMP3FN = App.Path & "\ALLMP3.MP3"
    If Dir(ALLMP3FN) <> vbNullString Then Kill ALLMP3FN


    ReDim Names(N)
    ReDim MP3(N)

    fMain.Caption = "Joining MP3s...."

    For I = 1 To N
        Names(I) = App.Path & "\Sounds\" & Format(I, "000") & ".MP3"
    Next

    allMP3 = vbNullString
    For I = 1 To N
        Open Names(I) For Binary As #1
        MP3(I) = String(LOF(1), 0)
        Get #1, 1, MP3(I)
        Close 1
        Kill Names(I)
        allMP3 = allMP3 & MP3(I)
    Next


    Open ALLMP3FN For Binary As #1
    Put #1, 1, allMP3
    Close #1


    JoinMP3 = ALLMP3FN

End Function

Public Sub PlayMP3(FN As String)   ', Optional bDoEvents As Boolean = True)

    Dim ML        As String
    Dim FileLength As Long

    '    fMain.Caption = "Speaking... " & FN

    If mciSendString("open " & Chr$(34) & FN & Chr$(34) & " type MpegVideo" & " alias myfile", 0&, 0&, 0&) = 0 Then
        ML = String(30, 0)
        Call mciSendString("status myfile length ", ML, 30, 0&)
        FileLength = Val(ML)
        'fmain.PB.Max = FileLength \ 10
        If FileLength Then
            If mciSendString("play myFile", 0&, 0&, 0&) = 0 Then
                Do While mciSendString("status myfile position ", ML, 30, 0&) = 0
                    'If Val(ML) = FileLength Then GoogleSpeak = True: Exit Do
                    If Val(ML) = FileLength Then Exit Do
                    'If bDoEvents Then DoEvents
                    'fmain.PB.Value = Val(ML) \ 10
                    '                    DoEvents
                    '                    DoEvents
                    '                    DoEvents
                Loop
            End If
        End If
        Call mciSendString("close myfile", 0&, 0&, 0&)
    End If


End Sub
