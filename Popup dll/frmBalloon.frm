VERSION 5.00
Begin VB.Form frmBalloon 
   BackColor       =   &H00E1FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInput2 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00E1FFFF&
      Caption         =   "No"
      Height          =   285
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00E1FFFF&
      Caption         =   "Yes"
      Height          =   285
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E1FFFF&
      Caption         =   "Ok"
      Height          =   285
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   480
   End
   Begin VB.Image imgX 
      Height          =   270
      Left            =   2520
      Picture         =   "frmBalloon.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgClose 
      Height          =   270
      Index           =   2
      Left            =   2520
      Picture         =   "frmBalloon.frx":0432
      Top             =   720
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgClose 
      Height          =   270
      Index           =   1
      Left            =   2160
      Picture         =   "frmBalloon.frx":0864
      Top             =   720
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgClose 
      Height          =   270
      Index           =   0
      Left            =   2160
      Picture         =   "frmBalloon.frx":0C96
      Top             =   360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   3
      Left            =   1080
      Picture         =   "frmBalloon.frx":10C8
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   1
      Left            =   600
      Picture         =   "frmBalloon.frx":1652
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   4
      Left            =   1560
      Picture         =   "frmBalloon.frx":1BDC
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   120
      Top             =   420
      Width           =   240
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   1
      Top             =   450
      Width           =   585
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Corners 
      Height          =   150
      Index           =   3
      Left            =   3480
      Picture         =   "frmBalloon.frx":2166
      Top             =   1800
      Width           =   135
   End
   Begin VB.Image Corners 
      Height          =   135
      Index           =   2
      Left            =   0
      Picture         =   "frmBalloon.frx":22C0
      Top             =   1800
      Width           =   120
   End
   Begin VB.Image Corners 
      Height          =   480
      Index           =   1
      Left            =   3480
      Picture         =   "frmBalloon.frx":23DA
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Corners 
      Height          =   450
      Index           =   0
      Left            =   0
      Picture         =   "frmBalloon.frx":271C
      Top             =   0
      Width           =   555
   End
   Begin VB.Label Fraudy 
      BackColor       =   &H00FF00FF&
      Caption         =   "Label1"
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   27810
   End
End
Attribute VB_Name = "frmBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'======================================================================================
'======================================================================================
'       Message     - Required - The message you want to show
'       Title       - Optional - The title of the balloon; Default is the Project Title
'       Icon        - Optional - The icon you want to show;
'                                Critical, Exclamation or Information; Default is no Icon
'       ShowClose   - Optional - Want to show the close button; Default is True
'       autoCloseTime Optional - Specify nonzero time in seconds if you want to close
'                                the popUp automatically
'       fontSize    - Optional - Specifies the font size; Default is 8
'       fontFace    - Optional - Specifies a custom font; Default is MS Sans Serif
'       PutAtCurrentMousePos   - Optional; Default true; will place the balloon at the
'                                current cursor position
'       XinPixel    - Optional - Specifies the X-coordinate in pixel where ballon is to
'                                be placed;
'                                works only if PutAtCurrentMousePos=False
'       YinPixel    - Optional - Specifies the Y-coordinate in pixel where ballon is to
'                                be placed;
'                                works only if PutAtCurrentMousePos=False
'       Button1     - Optional - Button1 visible state true/false
'       Caption1    - Optional - Button1 Caption
'       Button2     - Optional - Button2 visible state true/false
'       Caption2    - Optional - Button2 Caption
'       Button3     - Optional - Button3 visible state true/false
'       Caption3    - Optional - Button3 Caption
'       FormX       - Optional - Name of the control (Having hwnd property otherwise current
'                                Cursor position will be used) where you want the balloon to
'                                be placed
'                                works only if PutAtCurrentMousePos=False
'       vbMod       - Optional - Sets vbModal to .hWnd true/false
'       intBoxes    - Optional - Number of Textboxes on the Inputbox Popup
'       Style       - Optional - Popup Balloon Style "Balloon"/"Input"
'=======================================================================================
'=======================================================================================

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public objHost As Object
Public stStyle As String


Private Sub createRegion()
Const RGN_OR = 2
Dim glSkinImage As Long
Dim glHeight    As Long
Dim glwidth     As Long
Dim lReturn     As Long
Dim lRgnTmp     As Long
Dim lSkinRgn    As Long
Dim lStart      As Long
Dim lRow        As Long
Dim lCol        As Long

Dim I, H, W As Integer
Me.ScaleMode = vbPixels
H = Me.ScaleHeight
W = Me.ScaleWidth

lSkinRgn = CreateRectRgn(0, 0, 0, 0)

lRgnTmp = CreateRectRgn(16, 0, 17, 1)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 1, 18, 2)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 2, 19, 3)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 3, 20, 4)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 4, 21, 5)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 5, 22, 6)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 6, 23, 7)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 7, 24, 8)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 8, 25, 9)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 9, 26, 10)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 10, 27, 11)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 11, 28, 12)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 12, 29, 14)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 13, 30, 15)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 14, 31, 16)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 15, 32, 17)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 16, 33, 18)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 17, 34, 19)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(16, 18, 35, 19)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)

'TOP LEFT CORNER
lRgnTmp = CreateRectRgn(5, 19, 6, 20)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(3, 20, 6, 22)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(2, 21, 6, 22)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(1, 22, 6, 24)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(0, 24, 6, 24)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)

'TOP RIGHT CORNER
lRgnTmp = CreateRectRgn(W - 6, 19, W - 5, 21)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(W - 6, 20, W - 3, 21)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(W - 6, 21, W - 2, 22)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(W - 6, 22, W - 1, 23)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(W - 6, 23, W - 1, 24)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)

'BOTTOM LEFT CORNER
lRgnTmp = CreateRectRgn(0, H - 6, 16, H - 5)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(1, H - 6, 16, H - 3)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(2, H - 6, 16, H - 2)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(3, H - 6, 16, H - 1)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)

'BOTTOM RIGHT CORNER

lRgnTmp = CreateRectRgn(W - 6, H - 6, W - 0, H - 5)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(W - 6, H - 5, W - 1, H - 3)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(W - 6, H - 3, W - 2, H - 2)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(W - 6, H - 2, W - 3, H - 1)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
lRgnTmp = CreateRectRgn(W - 6, H - 1, W - 5, H)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)


'TOP LINE
lRgnTmp = CreateRectRgn(5, 19, W - 6, 24)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
'BOTTOM LINE
lRgnTmp = CreateRectRgn(5, H - 6, W - 6, H)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)

lRgnTmp = CreateRectRgn(0, 24, W, H - 6)
lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
    
Call SetWindowRgn(Me.hwnd, lSkinRgn, False)
Me.ScaleMode = vbTwips
End Sub

Private Sub cmdNo_Click()
On Error Resume Next
   Select Case stStyle
      Case "Balloon": objHost.MsgBack cmdNo.Caption
      Case "Input": objHost.InputBack cmdNo.Caption
   End Select
   Unload Me
End Sub

Private Sub cmdOk_Click()
   On Error Resume Next
   Select Case stStyle
       Case "Balloon": objHost.MsgBack cmdOk.Caption
       Case "Input"
            If txtInput.Text = vbNullString Then Exit Sub
            If txtInput2.Visible = True Then
               If txtInput2.Text = vbNullString Then Exit Sub
            End If
            objHost.InputBack txtInput.Text & "," & txtInput2.Text
   End Select
   Unload Me
 End Sub

Private Sub cmdYes_Click()
On Error Resume Next
   objHost.MsgBack cmdYes.Caption
   Unload Me
   
End Sub

Private Sub Form_Click()
' Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgX.Picture = imgClose(0)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub Form_Resize()
If Me.Width < 3000 Then Me.Width = 3000
Corners(1).Left = Me.ScaleWidth - Corners(1).Width
Corners(3).Left = Me.ScaleWidth - Corners(3).Width
Corners(2).Top = Me.ScaleHeight - Corners(2).Height
Corners(3).Top = Me.ScaleHeight - Corners(3).Height
Me.Cls
'With cmdOK
'    .Top = Me.Height - 570
'    .Left = (Me.Width \ 2) - (cmdOK.Width \ 2)
'End With
'With cmdYes
'    .Top = Me.Height - 570
'    .Left = (Me.Width \ 4) - (cmdYes.Width \ 2)
'End With
'With cmdNo
'    .Top = Me.Height - 570
'    .Left = (Me.Width \ 4) * 3 - (cmdNo.Width \ 2)
'End With
Line (540, 285)-(Me.Width - 105, 285)
Line (120, Me.Height - 15)-(Me.Width - 105, Me.Height - 15)
Line (0, 405)-(0, Me.Height - 120)
Line (Me.Width - 15, 405)-(Me.Width - 15, Me.Height - 120)

imgX.Left = Me.ScaleWidth - (1.5 * imgX.Width) - 1
createRegion
MakeTopMost Me.hwnd
End Sub

Private Sub imgX_Click()
objHost.MsgBack "MiscEv"
    Unload Me
End Sub

Private Sub imgX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then imgX.Picture = imgClose(2) ' X_Dn.Picture
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        imgX.Picture = imgClose(2)
    Else
        imgX.Picture = imgClose(1)
    End If
End Sub

Private Sub imgX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then imgX.Picture = imgClose(1) 'imgX_Up.Picture
End Sub

Private Sub lblMsg_Click()
    'Unload Me
End Sub

Private Sub lblMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgX.Picture = imgClose(0)
End Sub

Private Sub lblTitle_Click()
    'Unload Me
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgX.Picture = imgClose(0)
End Sub

Private Sub Timer1_Timer()
    If Timer1.Interval = 0 Then Timer1.Enabled = False
    objHost.MsgBack "MiscEv"
    Unload Me
End Sub
