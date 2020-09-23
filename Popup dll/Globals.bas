Attribute VB_Name = "Globals"
Option Explicit

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private m_snd() As Byte
Private Const SND_ASYNC = &H1 ' play asynchronously
Private Const SND_MEMORY = &H4 ' lpszSoundName points to a memory file
Private Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Public Const Msn = 101  'Good idea to define your sound id's

Public AlertCount As Integer

Private Declare Function SetWindowPos Lib "user32" _
                                    (ByVal hwnd As Long, _
                                     ByVal hWndInsertAfter As Long, _
                                     ByVal X As Long, Y, _
                                     ByVal cx As Long, _
                                     ByVal cy As Long, _
                                     ByVal wFlags As Long) As Long
                                     
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" _
                                           (lpData As Any, _
                                      ByVal hModule As Long, _
                                      ByVal dwFlags As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

 Public Function PlaySound(ByVal SndID As Long) As Long
       Const Flags = SND_ASYNC Or SND_MEMORY
       m_snd = LoadResData(SndID, "CUSTOM")
       PlaySoundData m_snd(0), 0, Flags
 End Function
