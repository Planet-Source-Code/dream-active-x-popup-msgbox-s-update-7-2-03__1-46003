VERSION 5.00
Begin VB.Form frmMsnStyle 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2745
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAlert 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   0
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00FAC2A5&
         Caption         =   "Ok"
         Height          =   255
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblAlert 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Alert Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   180
         Left            =   1680
         MouseIcon       =   "frmMsnStyle.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmMsnStyle.frx":030A
         Top             =   120
         Width           =   195
      End
   End
   Begin VB.Image Image2 
      Height          =   180
      Left            =   2160
      MouseIcon       =   "frmMsnStyle.frx":0720
      MousePointer    =   99  'Custom
      Picture         =   "frmMsnStyle.frx":0A2A
      Top             =   1560
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image3 
      Height          =   180
      Left            =   2400
      MouseIcon       =   "frmMsnStyle.frx":0E1E
      MousePointer    =   99  'Custom
      Picture         =   "frmMsnStyle.frx":1128
      Top             =   1560
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmMsnStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' API Declarations
Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)
' Constants
Const SM_CXFULLSCREEN = 16   ' Width of window client area
Const SM_CYFULLSCREEN = 17   ' Height of window client area
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

' Declarations
Private ClsGradient As New CGradient
Private fX As Long
Private fY As Long
Private lngScaleX As Long
Private lngScaleY As Long
Private AlertIndex As Long

Public objHost As Object

Private Sub Form_Load()
MakeTopMost Me.hWnd
End Sub

Private Sub Image1_Click()
If AlertCount = AlertIndex Then AlertCount = 0
Me.Visible = False
End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image1.Picture = Image2.Picture
End Sub

Private Sub cmdOk_Click()
objHost.MsgBack ("User Clicked Ok On MSN Style Popup")
If AlertCount = AlertIndex Then AlertCount = 0
Me.Visible = False
End Sub

Private Sub lblAlert_Click()
    ' When user clicked the alertbox
objHost.MsgBack ("User Clicked On Alert Message")
End Sub

Private Sub lblAlert_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Show as hyperlink
    If lblAlert.FontUnderline = False Then
        lblAlert.FontUnderline = True
        lblAlert.ForeColor = RGB(0, 0, 255)
    End If
    
End Sub

Private Sub picBackground_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Show text
    If lblAlert.FontUnderline = True Then
        lblAlert.FontUnderline = False
        lblAlert.ForeColor = &H0
    End If
 
    Image1.Picture = Image3.Picture
    
End Sub

Public Sub Display(MessageText As String, Duration As Long, Sound As Long)

    Dim wFlags As Long, x As Long

    ' Increase the alert count
    AlertCount = AlertCount + 1
    If AlertCount >= 5 Then AlertCount = 1
    AlertIndex = AlertCount

    ' Set the message
    lblAlert.Caption = MessageText
    
    ' Set the duration
    tmrAlert.Interval = Duration

    ' Get the system metrics we need
    fX = GetSystemMetrics(SM_CXFULLSCREEN)
    fY = GetSystemMetrics(SM_CYFULLSCREEN)
    lngScaleX = Me.Width - Me.ScaleWidth
    lngScaleY = Me.Height - Me.ScaleHeight
    
    ' Size the form
    Me.Height = 90
    Me.Width = picBackground.Width + lngScaleX
    Me.Left = fX * Screen.TwipsPerPixelX - Me.Width - 200
    Me.Top = (fY * Screen.TwipsPerPixelY) - ((picBackground.Height + lngScaleY) * (AlertCount - 1)) + 200
    Me.Show
    
    ' Play sound
    PlaySound Sound
    ' Draw the gradient background
    With ClsGradient
        .Angle = -100
        .Color1 = RGB(61, 149, 255)
        .Color2 = RGB(255, 255, 255)
        .Draw picBackground
    End With
    picBackground.Refresh
    If Duration = 0 Then cmdOK.Visible = True
    ' Open the alert box
     Call tmOpen
 
        
End Sub

Private Sub tmOpen()
    Dim curHeight As Long
    Dim newHeight As Long
    curHeight = Me.Height
    Do Until curHeight >= picBackground.Height + lngScaleY
       DoEvents
        newHeight = curHeight + 10
        If newHeight > picBackground.Height + lngScaleY Then newHeight = picBackground.Height + lngScaleY
        Me.Height = Me.Height + (newHeight - curHeight)
        Me.Top = Me.Top - (newHeight - curHeight)
        curHeight = Me.Height
   Loop
        If cmdOK.Visible = True Then Exit Sub
        tmrAlert.Enabled = True
End Sub

Private Sub tmrAlert_Timer()
    ' Alert was displayed, now close it
    tmrAlert.Enabled = False
    Call tmClose
End Sub

Private Sub tmClose()
Dim curHeight As Long
    curHeight = Me.Height
    Do Until curHeight <= 120
       DoEvents
        Me.Height = curHeight - 5 '0
        Me.Top = Me.Top + 5 '0
        curHeight = Me.Height
    Loop
        ' Close form
        If AlertCount = AlertIndex Then AlertCount = 0
        Unload Me
End Sub
