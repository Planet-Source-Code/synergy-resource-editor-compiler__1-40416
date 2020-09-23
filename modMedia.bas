Attribute VB_Name = "ModMedia"
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NODEFAULT = &H2
Public Const SND_SYNC = &H0
Public Const SND_NOSTOP = &H10
Public Const SND_MEMORY = &H4

Public sndArray() As Byte

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public vidArray() As Byte

Public Sub PlaySound()
  'Beep Frequency, Duration
  For x = 0 To 2500 Step 25
    Beep x, 100
  Next
End Sub

