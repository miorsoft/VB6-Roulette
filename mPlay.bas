Attribute VB_Name = "mPlay"
''https://microsoft.public.vb.general.discussion.narkive.com/AQLnyI4k/how-do-you-code-asynchronous-mp3-playing-with-vb6-api

Option Explicit

Private Const WS_CHILD = &H40000000

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Function SendMCIString(ByVal Cmd As String, Optional ByVal CheckErr As Boolean = False) As Boolean
    Dim Ret       As Long
    Dim Buff      As String * 260
    Ret = mciSendString(Cmd, 0&, 0&, 0&)
    If Ret Then
        If CheckErr Then
            mciGetErrorString Ret, Buff, 260
            MsgBox Left$(Buff, InStr(Buff, Chr$(0)))
        End If
    End If
    SendMCIString = CBool(Ret)
End Function

Private Function ShortPath(ByVal LongPath As String) As String
    Dim Ret       As Long
    Dim Buff      As String * 260
    Ret = GetShortPathName(LongPath, Buff, 260)
    ShortPath = Left$(Buff, Ret)
End Function

Public Sub PlayAsync(ByVal MP3Path As String)
    Dim CmdString As String
    Dim hWnd      As Long
    hWnd = fMain.hWnd


    SendMCIString "close all", True


    MP3Path = ShortPath(MP3Path)
    If Len(MP3Path) Then
        CmdString = "open  " & MP3Path & _
                    " type MPEGVideo Alias mp3 parent " _
                    & hWnd & " Style " & WS_CHILD
        If Not SendMCIString(CmdString, True) Then
            SendMCIString "play mp3", False
        End If
    Else
        MsgBox "Problem with Mp3 Path"
    End If
End Sub

