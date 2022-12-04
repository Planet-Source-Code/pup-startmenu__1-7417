Attribute VB_Name = "modSound"
Option Explicit

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_SYNC = &H0 ' Don't return until sound ends (default).
Public Const SND_ASYNC = &H1 ' Return immediately after the sound starts.
Public Const SND_NODEFAULT = &H2 ' If the sound file is not found, do NOT play default sound.
Public Const SND_MEMORY = &H4 ' Play a sound from a buffer in memory.
Public Const SND_LOOP = &H8 ' Loop sound continuously (used with SND_ASYNC)
Public Const SND_NOSTOP = &H10 ' Don't stop current sound to play another.

Public Sub s_Playsound(strName As String)
  Dim a As String
  
  a = GetSetting("BoS", "BInterface", "Skin", "BoS Standard")
  If a <> "BoS Standard" Then
    strName = App.path & "\skins\" & a & "\" & strName & ".wav"
  Else
    strName = App.path & "\" & strName & ".wav"
  End If
  sndPlaySound strName, SND_ASYNC Or SND_NODEFAULT
End Sub

