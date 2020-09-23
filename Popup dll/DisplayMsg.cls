VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long

Private Type POINT_API
    X As Long
    Y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Enum IconType
    vbExclamation = 48
    vbCritical = 16
    vbInformation = 64
    vbNone = 0
End Enum


Public Sub popUpBalloon(Message As String, _
                        Optional title As String, _
                        Optional Icon As IconType, _
                        Optional showClose As Boolean = True, _
                        Optional autoCloseTime As Integer = 0, _
                        Optional fontSize As Integer = 8, _
                        Optional fontFace As String = "MS Sans Serif", _
                        Optional PutAtCurrentMousePos As Boolean = True, _
                        Optional XinPixels As Integer, _
                        Optional YinPixels As Integer, _
                        Optional Button1 As Boolean = False, _
                        Optional Caption1 As String, _
                        Optional Button2 As Boolean = False, _
                        Optional Caption2 As String, _
                        Optional Button3 As Boolean = False, _
                        Optional Caption3 As String, _
                        Optional FormX As Object, _
                        Optional vbMod As Boolean = True, _
                        Optional intBoxes As Integer = 1, _
                        Optional Style As String = "Balloon")

Dim winRect As RECT
Dim Dot As POINT_API
On Error GoTo Default

Dim Balloon As New frmBalloon
Set Balloon.objHost = FormX
Balloon.stStyle = Style

If title = "" Then title = App.title
title = title & "     "

With Balloon
    .lblMsg = Message
    .lblMsg.Font = fontFace
    .lblMsg.fontSize = fontSize
    .lblTitle.Font = fontFace
    .lblTitle.fontSize = fontSize
    .lblTitle = title
    .imgIcon = .imgIconXP(Icon \ 16)
    .imgIcon.Top = 360 + 60 * Abs(Icon <> 48)
    .imgIcon.Visible = (Icon <> 0)
    .lblTitle.Left = 360 * Abs(.imgIcon.Visible) + 120
    .Width = IIf(.lblMsg.Left + .lblMsg.Width > .lblTitle.Left + .lblTitle.Width, (.lblMsg.Left + .lblMsg.Width + 240), (.lblTitle.Left + .lblTitle.Width + 240))
    .Height = .lblMsg.Top + .lblMsg.Height + 180
    'if either button styles are chosen then re adjust form height
    
    If Style = "Input" Then
       .Width = .Width + .cmdOk.Width
       .cmdNo.Caption = "Cancel"
       'if 2 box's are chosen then re adjust form height
       .txtInput2.Visible = (intBoxes <> 1)
       .txtInput.Visible = True
       .txtInput.Width = .Width - 430
       .txtInput2.Width = .Width - 430
       .txtInput.Left = 240
       .txtInput2.Left = 240
       .txtInput.Top = .lblMsg.Top + .lblMsg.Height + 100
       .Height = .txtInput.Top + 450
       With Balloon.cmdOk
           .Visible = True
           .Top = 480
           .Left = (Balloon.Width - 1050)
       End With
       With Balloon.cmdNo
           .Visible = True
           .Top = 840
           .Left = (Balloon.Width - 1050)
       End With
       Select Case intBoxes
              Case 2
               .txtInput2.Top = .txtInput.Top + 430
               .Height = .txtInput2.Top + 420
       End Select
    End If
    
    If Style = "Balloon" Then
       .cmdOk.Caption = Caption1
       .cmdYes.Caption = Caption2
       .cmdNo.Caption = Caption3
       .imgX.Visible = showClose
       .Timer1.Interval = autoCloseTime * 1000
       .cmdOk.Visible = Button1
       If Button2 = True Then
          .cmdYes.Visible = True
          .cmdNo.Visible = True
       End If
       If Button3 = True Then
          .cmdYes.Visible = True
          .cmdNo.Visible = True
          .cmdOk.Visible = True
       End If
       If Button1 = True Or _
          Button2 = True Or _
          Button3 = True Then .Height = .Height + 600
       With Balloon.cmdOk
           .Top = Balloon.Height - 570
           .Left = (Balloon.Width \ 2) - (Balloon.cmdOk.Width \ 2)
       End With
       With Balloon.cmdYes
           .Top = Balloon.Height - 570
           .Left = (Balloon.Width \ 4) - (Balloon.cmdYes.Width \ 2)
       End With
       With Balloon.cmdNo
           .Top = Balloon.Height - 570
           .Left = (Balloon.Width \ 4) * 3 - (Balloon.cmdNo.Width \ 2)
       End With
    End If

End With

    If PutAtCurrentMousePos Then
        Call GetCursorPos(Dot)
        GoTo SHOWFORM
    ElseIf (XinPixels <> 0) And (YinPixels <> 0) Then
        Dot.X = XinPixels
        Dot.Y = YinPixels
        GoTo SHOWFORM
    End If

'if no position set then set to default(current cursor position)
Default:
    Call GetCursorPos(Dot)
SHOWFORM:
    Dim W As Integer, H As Integer
    W = (Dot.X * 15) - 240: H = Dot.Y * 15
    'see if balloon runs off screen (right)
    If W + Balloon.Width > Screen.Width - 50 Then
    Balloon.Left = Screen.Width - Balloon.Width - 50
    Else
    Balloon.Left = (Dot.X * 15) - 240
    End If
    'see if balloon runs off screen (bottom)
    If H + Balloon.Height > Screen.Height - 420 Then
    Balloon.Top = Screen.Height - Balloon.Height - 420
    Else
    Balloon.Top = Dot.Y * 15
    End If
    
    Select Case vbMod
           Case True:  Balloon.Show vbModal, FormX
           Case False: Balloon.Show
    End Select
    
End Sub

Public Function DisplayAlert(FormX As Object, MessageText As String, Duration As Long, Sound As Long) As String
On Error GoTo errtrap
Dim AlertBox As frmMsnStyle
Set AlertBox = New frmMsnStyle
    AlertBox.Display MessageText, Duration, Sound
    DisplayAlert = vbNullString
    Set AlertBox.objHost = FormX
Exit Function
errtrap:
    DisplayAlert = Err.Description
End Function
