Attribute VB_Name = "mSound"
Option Explicit


Public SoundSLOT(0 To 37) As cSound

Public SOUNDFait  As cSound
Public SOUNDLeJeux As cSound
Public SOUNDRien  As cSound

Public Sub SETUPSOUND()

    Dim I         As Long
    Dim S$
    Dim tmp$

    'instantiate the Default-Audio-RenderDevice
    Set RenderDev = New_c.MMDeviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia)

    For I = 0 To 37
        Set SoundSLOT(I) = New cSound
        S = Slot2MP3(I, tmp)
        SoundSLOT(I).InitResourceFromMP3 S, RenderDev, 11050
        '        SoundSLOT(I).PLAY
    Next

    Set SOUNDFait = New cSound
    Set SOUNDLeJeux = New cSound
    Set SOUNDRien = New cSound

    SOUNDFait.InitResourceFromMP3 App.Path & "\Sounds\Faites vos jeux.MP3", RenderDev, 11050
    'SOUNDFait.PLAY
    SOUNDLeJeux.InitResourceFromMP3 App.Path & "\Sounds\Les Jeux sont faits.MP3", RenderDev, 11050
    'SOUNDLeJeux.PLAY
    SOUNDRien.InitResourceFromMP3 App.Path & "\Sounds\Rien ne va plus.MP3", RenderDev, 11050
    'SOUNDRien.PLAY


End Sub

